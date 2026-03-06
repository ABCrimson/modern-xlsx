use core::hint::cold_path;

use chrono::{Datelike, NaiveDate};
use serde::{Deserialize, Serialize};

use crate::{ModernXlsxError, Result};

/// Which date-origin system the workbook uses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum DateSystem {
    /// The default 1900 date system (used by Windows Excel).
    /// Day 1 = Jan 1, 1900. Includes the Lotus 1-2-3 leap-year bug at day 60.
    Date1900,
    /// The 1904 date system (used by older Mac Excel).
    /// Day 0 = Jan 1, 1904. No leap-year bug.
    Date1904,
}

/// Broken-down date/time components extracted from a serial number.
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

/// Extract hour, minute, second, millisecond from the fractional part of a serial number.
#[inline]
fn fractional_to_time(frac: f64) -> (u32, u32, u32, u32) {
    let total_ms = (frac * 86_400_000.0).round() as u64;
    let ms = (total_ms % 1000) as u32;
    let total_secs = total_ms / 1000;
    let second = (total_secs % 60) as u32;
    let total_mins = total_secs / 60;
    let minute = (total_mins % 60) as u32;
    let hour = (total_mins / 60) as u32;
    (hour, minute, second, ms)
}

/// Convert an Excel serial date number to date/time components.
///
/// # 1900 system
///
/// - Serial 0 = Dec 31, 1899 (Excel quirk)
/// - Serial 1 = Jan 1, 1900
/// - Serial 60 = Feb 29, 1900 (fake Lotus 1-2-3 bug; 1900 was NOT a leap year)
/// - Serial 61 = March 1, 1900
/// - For serial >= 61, subtract 1 to compensate for the fake day 60.
///
/// # 1904 system
///
/// - Serial 0 = Jan 1, 1904
/// - No leap-year bug; straightforward day addition.
pub fn serial_to_date(serial: f64, system: DateSystem) -> Result<DateTimeComponents> {
    if serial < 0.0 {
        cold_path();
        return Err(ModernXlsxError::InvalidDate(format!(
            "Failed to convert serial number {serial}: value must be non-negative (negative serial numbers are not valid in Excel)"
        )));
    }

    let day_int = serial.floor() as i64;
    let frac = serial - (day_int as f64);
    let (hour, minute, second, millisecond) = fractional_to_time(frac);

    match system {
        DateSystem::Date1900 => serial_to_date_1900(day_int, hour, minute, second, millisecond),
        DateSystem::Date1904 => serial_to_date_1904(day_int, hour, minute, second, millisecond),
    }
}

fn serial_to_date_1900(
    day_int: i64,
    hour: u32,
    minute: u32,
    second: u32,
    millisecond: u32,
) -> Result<DateTimeComponents> {
    // Day 0 = Dec 31, 1899
    if day_int == 0 {
        return Ok(DateTimeComponents {
            year: 1899,
            month: 12,
            day: 31,
            hour,
            minute,
            second,
            millisecond,
        });
    }

    // Day 60 = the fake Feb 29, 1900 (Lotus 1-2-3 bug).
    if day_int == 60 {
        return Ok(DateTimeComponents {
            year: 1900,
            month: 2,
            day: 29,
            hour,
            minute,
            second,
            millisecond,
        });
    }

    // For serials >= 61, subtract 1 to compensate for the fake day 60.
    let adjusted = if day_int >= 61 { day_int - 1 } else { day_int };

    let epoch = EPOCH_1900;
    let date = epoch
        .checked_add_signed(chrono::Duration::days(adjusted))
        .ok_or_else(|| {
            cold_path();
            ModernXlsxError::InvalidDate(format!(
                "Failed to convert serial number {day_int}: date overflow — the value is too large for the calendar"
            ))
        })?;

    Ok(DateTimeComponents {
        year: date.year(),
        month: date.month(),
        day: date.day(),
        hour,
        minute,
        second,
        millisecond,
    })
}

fn serial_to_date_1904(
    day_int: i64,
    hour: u32,
    minute: u32,
    second: u32,
    millisecond: u32,
) -> Result<DateTimeComponents> {
    let epoch = EPOCH_1904;
    let date = epoch
        .checked_add_signed(chrono::Duration::days(day_int))
        .ok_or_else(|| {
            cold_path();
            ModernXlsxError::InvalidDate(format!(
                "Failed to convert serial number {day_int}: date overflow — the value is too large for the calendar"
            ))
        })?;

    Ok(DateTimeComponents {
        year: date.year(),
        month: date.month(),
        day: date.day(),
        hour,
        minute,
        second,
        millisecond,
    })
}

/// Convert date/time components back to an Excel serial number.
pub fn date_to_serial(dt: &DateTimeComponents, system: DateSystem) -> Result<f64> {
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

fn date_to_serial_1900(dt: &DateTimeComponents, time_frac: f64) -> Result<f64> {
    // Special case: the fake Feb 29, 1900.
    if dt.year == 1900 && dt.month == 2 && dt.day == 29 {
        return Ok(60.0 + time_frac);
    }

    // Special case: Dec 31, 1899 = serial 0.
    if dt.year == 1899 && dt.month == 12 && dt.day == 31 {
        return Ok(time_frac);
    }

    let date = NaiveDate::from_ymd_opt(dt.year, dt.month, dt.day).ok_or_else(|| {
        cold_path();
        ModernXlsxError::InvalidDate(format!(
            "Failed to construct date from {}-{:02}-{:02}: not a valid calendar date",
            dt.year, dt.month, dt.day
        ))
    })?;

    let epoch = EPOCH_1900;
    let days = (date - epoch).num_days();

    // Add 1 for serials >= 61 to re-introduce the Lotus bug gap.
    let serial = if days >= 60 { days + 1 } else { days };

    Ok(serial as f64 + time_frac)
}

fn date_to_serial_1904(dt: &DateTimeComponents, time_frac: f64) -> Result<f64> {
    let date = NaiveDate::from_ymd_opt(dt.year, dt.month, dt.day).ok_or_else(|| {
        cold_path();
        ModernXlsxError::InvalidDate(format!(
            "Failed to construct date from {}-{:02}-{:02}: not a valid calendar date",
            dt.year, dt.month, dt.day
        ))
    })?;

    let epoch = EPOCH_1904;
    let days = (date - epoch).num_days();

    Ok(days as f64 + time_frac)
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_serial_to_date_1900() {
        let dt = serial_to_date(1.0, DateSystem::Date1900).unwrap();
        assert_eq!(dt.year, 1900);
        assert_eq!(dt.month, 1);
        assert_eq!(dt.day, 1);
    }

    #[test]
    fn test_serial_to_date_known() {
        // Serial 45336 = Feb 14, 2024
        let dt = serial_to_date(45336.0, DateSystem::Date1900).unwrap();
        assert_eq!(dt.year, 2024);
        assert_eq!(dt.month, 2);
        assert_eq!(dt.day, 14);
    }

    #[test]
    fn test_serial_with_time() {
        // Serial 45336.75 = Feb 14, 2024 at 18:00:00
        let dt = serial_to_date(45336.75, DateSystem::Date1900).unwrap();
        assert_eq!(dt.year, 2024);
        assert_eq!(dt.month, 2);
        assert_eq!(dt.day, 14);
        assert_eq!(dt.hour, 18);
        assert_eq!(dt.minute, 0);
        assert_eq!(dt.second, 0);
    }

    #[test]
    fn test_serial_half_day() {
        let dt = serial_to_date(45336.5, DateSystem::Date1900).unwrap();
        assert_eq!(dt.hour, 12);
        assert_eq!(dt.minute, 0);
        assert_eq!(dt.second, 0);
    }

    #[test]
    fn test_lotus_bug_day_60() {
        // Day 60 = the fake Feb 29, 1900
        let dt = serial_to_date(60.0, DateSystem::Date1900).unwrap();
        assert_eq!(dt.year, 1900);
        assert_eq!(dt.month, 2);
        assert_eq!(dt.day, 29);
    }

    #[test]
    fn test_day_61_march_1() {
        // Day 61 = March 1, 1900
        let dt = serial_to_date(61.0, DateSystem::Date1900).unwrap();
        assert_eq!(dt.year, 1900);
        assert_eq!(dt.month, 3);
        assert_eq!(dt.day, 1);
    }

    #[test]
    fn test_1904_system() {
        // Day 0 = Jan 1, 1904
        let dt = serial_to_date(0.0, DateSystem::Date1904).unwrap();
        assert_eq!(dt.year, 1904);
        assert_eq!(dt.month, 1);
        assert_eq!(dt.day, 1);
    }

    #[test]
    fn test_date_to_serial_roundtrip() {
        let original_serial = 45336.75;
        let dt = serial_to_date(original_serial, DateSystem::Date1900).unwrap();
        let back = date_to_serial(&dt, DateSystem::Date1900).unwrap();
        assert!((back - original_serial).abs() < 1e-6);
    }

    #[test]
    fn test_serial_day_0() {
        // Day 0 = Dec 31, 1899
        let dt = serial_to_date(0.0, DateSystem::Date1900).unwrap();
        assert_eq!(dt.year, 1899);
        assert_eq!(dt.month, 12);
        assert_eq!(dt.day, 31);
    }

    #[test]
    fn test_error_messages_have_context() {
        // Negative serial should mention the value
        let err = serial_to_date(-5.0, DateSystem::Date1900).unwrap_err();
        let msg = err.to_string();
        assert!(msg.contains("-5"), "expected serial value in error, got: {msg}");
        assert!(
            msg.contains("non-negative"),
            "expected 'non-negative' hint, got: {msg}"
        );
        assert_eq!(err.code(), "INVALID_DATE");

        // Invalid calendar date
        let bad_dt = DateTimeComponents {
            year: 2024, month: 13, day: 1,
            hour: 0, minute: 0, second: 0, millisecond: 0,
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
}
