# Security Policy

## Supported Versions

| Version | Supported |
|---------|-----------|
| 0.3.x   | Yes       |
| 0.2.x   | Yes       |
| < 0.2   | No        |

## Reporting a Vulnerability

If you discover a security vulnerability, please report it responsibly:

1. **Do not** open a public issue
2. Email the maintainers or use [GitHub Security Advisories](https://github.com/ABCrimson/modern-xlsx/security/advisories/new)
3. Include a description of the vulnerability and steps to reproduce

We will acknowledge receipt within 48 hours and aim to release a fix within 7 days for critical issues.

## Scope

modern-xlsx processes untrusted `.xlsx` files. Security-relevant areas include:

- **ZIP decompression** — handled by the `zip` crate with size limits
- **XML parsing** — SAX-style parsing via `quick-xml` (no entity expansion, no external DTDs)
- **Memory safety** — Rust core provides memory safety guarantees; the WASM sandbox provides additional isolation
