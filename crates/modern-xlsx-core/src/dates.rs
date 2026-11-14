//! Excel date serials: the [`DateSystem`] (1900 or 1904), [`date_to_serial`] for
//! [`DateTimeComponents`], and `iso8601_to_serial` for the ISO 8601 text of
//! `t="d"` cells, which the worksheet parser calls.

use core::hint::cold_path;

use chrono::{Datelike, NaiveDate, NaiveDateTime, NaiveTime, Timelike};
use serde::{Deserialize, Serialize};

use crate::{ModernXlsxError, Result};

/// Which date-origin system the workbook uses.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum DateSystem {
    /// The default 1900 date system (used by Windows Excel).
    /// Day 1 = Jan 1, 1900. Includes the Lotus 1-2-3 leap-year bug at day 60.
    #[default]
    Date1900,
    /// The 1904 date system (used by older Mac Excel).
    /// Day 0 = Jan 1, 1904. No leap-year bug.
    Date1904,
}

/// Broken-down date/time components that [`date_to_serial`] converts to an Excel serial number.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DateTimeComponents {
    pub year: i32,
    pub month: u32,
    pub day: u32,
    pub hour: u32,
    pub minute: u32,
    pub second: u32,
    pub millisecond: u32,
}

/// The epoch for the 1900 date system: Dec 31, 1899 (serial 0).
const EPOCH_1900: NaiveDate = NaiveDate::from_ymd_opt(1899, 12, 31).unwrap();

/// The epoch for the 1904 date system: Jan 1, 1904 (serial 0).
const EPOCH_1904: NaiveDate = NaiveDate::from_ymd_opt(1904, 1, 1).unwrap();

/// The last date either date system can hold (serial 2,958,465 in the 1900
/// system).
const LAST_DATE: NaiveDate = NaiveDate::from_ymd_opt(9999, 12, 31).unwrap();

/// Convert date/time components back to an Excel serial number.
///
/// # Errors
///
/// `InvalidDate` when a time field is out of range (hour above 23, minute or
/// second above 59, millisecond above 999), the date is not a calendar date,
/// or it lies outside the date system: before 1899-12-31 (serial 0) in the
/// 1900 system, before 1904-01-01 in the 1904 system, or after 9999-12-31.
pub fn date_to_serial(dt: &DateTimeComponents, system: DateSystem) -> Result<f64> {
    if dt.hour > 23 || dt.minute > 59 || dt.second > 59 || dt.millisecond > 999 {
        cold_path();
        return Err(ModernXlsxError::InvalidDate(format!(
            "Failed to convert time {:02}:{:02}:{:02}.{:03}: a field is out of range",
            dt.hour, dt.minute, dt.second, dt.millisecond
        )));
    }
    let time_frac = (dt.hour as f64 * 3600.0
        + dt.minute as f64 * 60.0
        + dt.second as f64
        + dt.millisecond as f64 / 1000.0)
        / 86400.0;

    match system {
        DateSystem::Date1900 => date_to_serial_1900(dt, time_frac),
        DateSystem::Date1904 => date_to_serial_1904(dt, time_frac),
    }
}

/// The calendar date of `dt`, checked to lie in `epoch..=LAST_DATE`.
fn checked_date(dt: &DateTimeComponents, epoch: NaiveDate) -> Result<NaiveDate> {
    let date = NaiveDate::from_ymd_opt(dt.year, dt.month, dt.day).ok_or_else(|| {
        cold_path();
        ModernXlsxError::InvalidDate(format!(
            "Failed to construct date from {}-{:02}-{:02}: not a valid calendar date",
            dt.year, dt.month, dt.day
        ))
    })?;
    if !(epoch..=LAST_DATE).contains(&date) {
        cold_path();
        return Err(ModernXlsxError::InvalidDate(format!(
            "Failed to convert {date}: the date system holds {epoch} to {LAST_DATE}"
        )));
    }
    Ok(date)
}

fn date_to_serial_1900(dt: &DateTimeComponents, time_frac: f64) -> Result<f64> {
    // Special case: the fake Feb 29, 1900.
    if dt.year == 1900 && dt.month == 2 && dt.day == 29 {
        return Ok(60.0 + time_frac);
    }

    let days = (checked_date(dt, EPOCH_1900)? - EPOCH_1900).num_days();

    // Add 1 for serials >= 61 to re-introduce the Lotus bug gap.
    let serial = if days >= 60 { days + 1 } else { days };

    Ok(serial as f64 + time_frac)
}

fn date_to_serial_1904(dt: &DateTimeComponents, time_frac: f64) -> Result<f64> {
    let days = (checked_date(dt, EPOCH_1904)? - EPOCH_1904).num_days();
    Ok(days as f64 + time_frac)
}

/// The serial in `system` of the ISO 8601 text of a `t="d"` cell (ECMA-376
/// Part 1, `ST_CellType` value `d`): `YYYY-MM-DD`, `YYYY-MM-DDThh:mm:ss[.fff]`
/// or a time-only `Thh:mm:ss[.fff]` (a fraction of a day), each optionally
/// followed by `Z` or a `±hh:mm` offset that is dropped (the wall-clock value
/// is kept). Fractional seconds are truncated to the millisecond. Surrounding
/// ASCII white space is ignored.
///
/// `None` when the text has none of these forms or `date_to_serial` rejects
/// it (a date outside the date system, a field out of range).
pub(crate) fn iso8601_to_serial(s: &str, system: DateSystem) -> Option<f64> {
    let s = s.trim_ascii();
    let (date, time, rest) = if let Some(t) = s.strip_prefix('T') {
        let (time, rest) = NaiveTime::parse_and_remainder(t, "%H:%M:%S%.f").ok()?;
        let epoch = match system {
            DateSystem::Date1900 => EPOCH_1900,
            DateSystem::Date1904 => EPOCH_1904,
        };
        (epoch, time, rest)
    } else if s.contains('T') {
        let (dt, rest) = NaiveDateTime::parse_and_remainder(s, "%Y-%m-%dT%H:%M:%S%.f").ok()?;
        (dt.date(), dt.time(), rest)
    } else {
        let (date, rest) = NaiveDate::parse_and_remainder(s, "%Y-%m-%d").ok()?;
        (date, NaiveTime::MIN, rest)
    };
    if !is_zone_designator(rest) {
        return None;
    }
    let dt = DateTimeComponents {
        year: date.year(),
        month: date.month(),
        day: date.day(),
        hour: time.hour(),
        minute: time.minute(),
        second: time.second(),
        // A leap second (nanosecond >= 10^9) lands above 999 and is rejected.
        millisecond: time.nanosecond() / 1_000_000,
    };
    date_to_serial(&dt, system).ok()
}

/// Whether `rest` is empty, `Z`, or a `+hh:mm` / `-hh:mm` UTC offset.
fn is_zone_designator(rest: &str) -> bool {
    match rest.as_bytes() {
        [] | [b'Z'] => true,
        [b'+' | b'-', h1, h2, b':', m1, m2] => [h1, h2, m1, m2].iter().all(|b| b.is_ascii_digit()),
        _ => false,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_date_to_serial_roundtrip() {
        // Serial 45336.75 = Feb 14, 2024 at 18:00:00.
        let dt = DateTimeComponents {
            year: 2024,
            month: 2,
            day: 14,
            hour: 18,
            minute: 0,
            second: 0,
            millisecond: 0,
        };
        let serial = date_to_serial(&dt, DateSystem::Date1900).unwrap();
        assert!((serial - 45336.75).abs() < 1e-6);
    }

    #[test]
    fn test_date_to_serial_1900_lotus_leap_year() {
        let at = |month, day| DateTimeComponents {
            year: 1900,
            month,
            day,
            hour: 0,
            minute: 0,
            second: 0,
            millisecond: 0,
        };
        assert_eq!(
            date_to_serial(&at(2, 28), DateSystem::Date1900).unwrap(),
            59.0
        );
        assert_eq!(
            date_to_serial(&at(2, 29), DateSystem::Date1900).unwrap(),
            60.0
        );
        assert_eq!(
            date_to_serial(&at(3, 1), DateSystem::Date1900).unwrap(),
            61.0
        );
    }

    #[test]
    fn test_error_messages_have_context() {
        // Invalid calendar date
        let bad_dt = DateTimeComponents {
            year: 2024,
            month: 13,
            day: 1,
            hour: 0,
            minute: 0,
            second: 0,
            millisecond: 0,
        };
        let err = date_to_serial(&bad_dt, DateSystem::Date1900).unwrap_err();
        let msg = err.to_string();
        assert!(
            msg.contains("2024-13-01"),
            "expected formatted date in error, got: {msg}"
        );
        assert!(
            msg.contains("not a valid calendar date"),
            "expected calendar hint, got: {msg}"
        );
    }

    fn ymd_hms(year: i32, month: u32, day: u32, h: u32, m: u32, s: u32) -> DateTimeComponents {
        DateTimeComponents {
            year,
            month,
            day,
            hour: h,
            minute: m,
            second: s,
            millisecond: 0,
        }
    }

    #[test]
    fn date_to_serial_rejects_out_of_range_time_fields() {
        for (h, m, s, ms) in [(24, 0, 0, 0), (0, 60, 0, 0), (0, 0, 60, 0), (0, 0, 0, 1000)] {
            let dt = DateTimeComponents {
                millisecond: ms,
                ..ymd_hms(2024, 1, 15, h, m, s)
            };
            for system in [DateSystem::Date1900, DateSystem::Date1904] {
                let err = date_to_serial(&dt, system).unwrap_err();
                assert_eq!(err.code(), "INVALID_DATE", "{dt:?}");
            }
        }
        let last = DateTimeComponents {
            millisecond: 999,
            ..ymd_hms(2024, 1, 15, 23, 59, 59)
        };
        assert!(date_to_serial(&last, DateSystem::Date1900).unwrap() < 45307.0);
    }

    #[test]
    fn date_to_serial_rejects_dates_outside_the_date_system() {
        let d1900 = DateSystem::Date1900;
        let d1904 = DateSystem::Date1904;
        assert!(date_to_serial(&ymd_hms(1899, 12, 30, 0, 0, 0), d1900).is_err());
        assert_eq!(
            date_to_serial(&ymd_hms(1899, 12, 31, 12, 0, 0), d1900).unwrap(),
            0.5
        );
        assert!(date_to_serial(&ymd_hms(1903, 12, 31, 0, 0, 0), d1904).is_err());
        assert_eq!(
            date_to_serial(&ymd_hms(1904, 1, 1, 0, 0, 0), d1904).unwrap(),
            0.0
        );
        assert_eq!(
            date_to_serial(&ymd_hms(9999, 12, 31, 0, 0, 0), d1900).unwrap(),
            2_958_465.0
        );
        assert_eq!(
            date_to_serial(&ymd_hms(9999, 12, 31, 0, 0, 0), d1904).unwrap(),
            2_957_003.0
        );
        for system in [d1900, d1904] {
            assert!(date_to_serial(&ymd_hms(10_000, 1, 1, 0, 0, 0), system).is_err());
        }
    }

    #[test]
    fn iso8601_serial_reads_the_cell_type_d_forms() {
        let d1900 = DateSystem::Date1900;
        for (s, serial) in [
            ("2024-01-15", 45306.0),
            ("2024-01-15T12:00:00", 45306.5),
            ("2024-01-15T12:00:00Z", 45306.5),
            ("2024-01-15T12:00:00+05:30", 45306.5),
            ("2024-01-15T12:00:00-11:00", 45306.5),
            ("T06:00:00", 0.25),
        ] {
            assert_eq!(iso8601_to_serial(s, d1900), Some(serial), "{s}");
        }
        // Fractional seconds are kept to the millisecond.
        let frac = iso8601_to_serial("2024-01-15T12:00:00.2504", d1900).unwrap();
        assert!((frac - (45306.5 + 0.25 / 86_400.0)).abs() < 1e-9, "{frac}");
        assert_eq!(
            iso8601_to_serial("2024-01-15T12:00:00", DateSystem::Date1904),
            Some(45306.5 - 1462.0)
        );
        assert_eq!(
            iso8601_to_serial("T06:00:00", DateSystem::Date1904),
            Some(0.25)
        );
        for s in [
            "",
            "garbage",
            "2024-01-15T12:00",
            "2024-01-15 12:00:00",
            "2024-01-15T12:00:00+0530",
            "2024-01-15T12:00:00Zjunk",
            "2024-02-30",
            "1899-12-30",
            "10000-01-01",
            "T24:00:00",
        ] {
            assert_eq!(iso8601_to_serial(s, d1900), None, "{s}");
        }
    }
}
