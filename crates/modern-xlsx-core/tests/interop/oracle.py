# Decrypt oracle for the encryption interop tests (tests/interop/main.rs).
#
# Purpose: msoffcrypto-tool, an independent implementation, decrypts a file that
# modern-xlsx encrypted and the plaintext must equal the expected bytes. The
# integrity check is on (the msoffcrypto CLI never checks the HMAC; the API
# checks it only with decrypt(..., verify_integrity=True), which this script
# passes),
# so a wrong dataIntegrity value fails here.
#
# Pinned: msoffcrypto-tool==6.1.0. Python standard library plus msoffcrypto only.
# CI: the `encryption-oracle` job installs the pin with pip and runs
#   MX_MSOFFCRYPTO_PYTHON=python cargo test -p modern-xlsx-core --release --locked \
#     --test interop -- --ignored
# which calls: python -I oracle.py <encrypted> <password> <expected plaintext>
import io
import sys

import msoffcrypto


def main() -> int:
    if len(sys.argv) != 4:
        print("usage: oracle.py <encrypted> <password> <expected>", file=sys.stderr)
        return 2
    encrypted, password, expected = sys.argv[1:]
    with open(encrypted, "rb") as f:
        office = msoffcrypto.OfficeFile(f)
        office.load_key(password=password, verify_password=True)
        out = io.BytesIO()
        office.decrypt(out, verify_integrity=True)
    with open(expected, "rb") as f:
        want = f.read()
    got = out.getvalue()
    if got != want:
        print(
            f"plaintext mismatch: got {len(got)} bytes, expected {len(want)} bytes",
            file=sys.stderr,
        )
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
