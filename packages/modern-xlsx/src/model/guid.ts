/**
 * A random RFC 9562 version-4 UUID in the form Excel requires of threaded
 * comment and person ids: upper-case, in braces,
 * `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`. Excel refuses a file whose
 * threaded comment or person ids have no braces, and drops a reply whose
 * `parentId` has none.
 */
export function excelGuid(): string {
  return `{${randomUuid().toUpperCase()}}`;
}

/** A random RFC 9562 version-4 UUID, lower-case, without braces. */
function randomUuid(): string {
  // randomUUID exists only in secure contexts; getRandomValues works everywhere.
  if (typeof crypto.randomUUID === 'function') return crypto.randomUUID();
  const bytes = crypto.getRandomValues(new Uint8Array(16));
  bytes[6] = ((bytes[6] ?? 0) & 0x0f) | 0x40;
  bytes[8] = ((bytes[8] ?? 0) & 0x3f) | 0x80;
  const hex = Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('');
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}
