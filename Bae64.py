import base64
import re


class Base64ValidationError(Exception):
    """Raised when a string fails Base64 validation, with a reason."""
    pass


def validate_base64(value: str) -> bytes:
    """
    Strictly validate that `value` is a well-formed Base64 string and
    return the decoded bytes.

    Raises Base64ValidationError with a specific reason if validation fails.
    Raises TypeError if input is not a string.

    :param value: The string to validate.
    :return: Decoded bytes if valid.
    """
    # --- Type check ---
    if not isinstance(value, str):
        raise TypeError(f"Expected str, got {type(value).__name__}")

    if value == "":
        raise Base64ValidationError("Input is empty")

    s = value.strip()

    # --- Character set check ---
    if not re.fullmatch(r'[A-Za-z0-9+/]*={0,2}', s):
        raise Base64ValidationError("Contains characters outside the Base64 alphabet")

    # --- Length check ---
    if len(s) % 4 != 0:
        raise Base64ValidationError("Length is not a multiple of 4")

    # --- Padding placement check ---
    if '=' in s:
        core, _, pad = s.partition('=')
        # everything after first '=' must be only '=' characters
        if not re.fullmatch(r'=*', s[len(core):]):
            raise Base64ValidationError("Padding character '=' appears in an invalid position")

    # --- Decode attempt (strict mode) ---
    try:
        decoded = base64.b64decode(s, validate=True)
    except (ValueError, base64.binascii.Error) as e:
        raise Base64ValidationError(f"Decoding failed: {e}") from e

    # --- Round-trip check (canonical form) ---
    re_encoded = base64.b64encode(decoded).decode('ascii')
    if re_encoded != s:
        raise Base64ValidationError("Input is not canonical Base64 (round-trip mismatch)")

    return decoded


def is_valid_base64(value: str) -> bool:
    """
    Non-throwing convenience wrapper around validate_base64().
    Returns True/False instead of raising.
    """
    try:
        validate_base64(value)
        return True
    except (Base64ValidationError, TypeError):
        return False


# --- Example usage ---
if __name__ == "__main__":
    test_values = [
        "SGVsbG8gd29ybGQh",   # valid -> "Hello world!"
        "cGFzc3dvcmQ=",       # valid -> "password"
        "hunter2",            # invalid - plain password
        "not base64!!",       # invalid - bad chars
        "abcd=efgh",          # invalid - bad padding position
        "",                   # invalid - empty
        None,                 # invalid - wrong type
    ]

    for val in test_values:
        try:
            decoded = validate_base64(val)
            print(f"{val!r:30} -> VALID, decoded = {decoded!r}")
        except Base64ValidationError as e:
            print(f"{val!r:30} -> INVALID: {e}")
        except TypeError as e:
            print(f"{val!r:30} -> TYPE ERROR: {e}")
