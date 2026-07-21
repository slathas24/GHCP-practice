"""
Break-glass Shamir shard distributor.

Takes a shard-assignment batch JSON (one secret, split into shards, one shard
per holder) and for each holder:
  1. Extracts a per-holder payload (metadata + only their shard).
  2. Generates a unique random passphrase for that holder.
  3. Encrypts the payload with AES-256-GCM using a key derived from that
     passphrase (PBKDF2-HMAC-SHA256).
  4. Writes the encrypted file to shards_out/ and the passphrase to
     keys_out/, both permission-restricted to the current user.
  5. Emails the encrypted file (only) to the holder as an attachment.

Expected input JSON shape:
{
  "secret": {
    "secretid": "...",
    "secret_desc": "...",
    "totalshards": 3,
    "threshold": 2,
    "encoding": "base64",
    "created_at": "...",
    "rotated_at": "...",
    "status": "active"
  },
  "shard_assignments": [
    {
      "assignment_id": "assign-0042",
      "holderid": "asmith",
      "holder_name": "Alice Smith",
      "holder_email": "[email protected]",
      "shard_index": 1,
      "backend_type": "local",
      "backend_ref": "file-name",
      "shard_value": "base64-encoded-share-data",
      "created_at": "...",
      "updated_at": "..."
    },
    ...
  ]
}

IMPORTANT SECURITY NOTE
------------------------
This script emails the ENCRYPTED shard file, but deliberately does NOT email
the passphrase. Passphrases are written locally to keys_out/ and must be
delivered to each holder through a SEPARATE channel (phone call, SMS,
secrets manager, in person, etc.) - never via the same email as the .enc
file, or the encryption provides no real protection.

keys_out/ should be treated as highly sensitive: anyone with read access to
it can decrypt every shard. This local-folder approach is intended as an
interim/dev solution. For production break-glass use, prefer encrypting to
each holder's public key (GPG/age) so no shared passphrase ever needs to be
distributed at all.
"""

import json
import os
import stat
import secrets
import smtplib
import argparse
import sys
from email.message import EmailMessage

from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes


# ----------------------------
# Key derivation / passphrase generation
# ----------------------------

def derive_key(passphrase: str, salt: bytes, iterations: int = 480_000) -> bytes:
    """Derive a 256-bit AES key from a passphrase using PBKDF2-HMAC-SHA256."""
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=iterations,
    )
    return kdf.derive(passphrase.encode("utf-8"))


def generate_passphrase(num_bytes: int = 24) -> str:
    """Cryptographically strong, URL-safe random passphrase, unique per call."""
    return secrets.token_urlsafe(num_bytes)


# ----------------------------
# Batch validation
# ----------------------------

def validate_batch(batch: dict):
    if "secret" not in batch:
        raise ValueError("Missing required top-level field: secret")
    if "shard_assignments" not in batch:
        raise ValueError("Missing required top-level field: shard_assignments")

    secret = batch["secret"]
    required_secret_fields = [
        "secretid", "secret_desc", "totalshards", "threshold", "encoding",
    ]
    for field in required_secret_fields:
        if field not in secret:
            raise ValueError(f"Missing required 'secret' field: {field}")

    assignments = batch["shard_assignments"]
    if not isinstance(assignments, list) or not assignments:
        raise ValueError("'shard_assignments' must be a non-empty list")

    if len(assignments) != secret["totalshards"]:
        raise ValueError(
            f"secret.totalshards={secret['totalshards']} but shard_assignments "
            f"has {len(assignments)} entries"
        )

    if secret["threshold"] > secret["totalshards"]:
        raise ValueError("secret.threshold cannot be greater than secret.totalshards")

    required_assignment_fields = [
        "assignment_id", "holderid", "holder_name", "holder_email",
        "shard_index", "shard_value",
    ]
    seen_indices = set()
    seen_ids = set()
    seen_assignment_ids = set()

    for a in assignments:
        for field in required_assignment_fields:
            if field not in a or a[field] in (None, ""):
                raise ValueError(f"shard_assignments entry missing/empty field '{field}': {a}")

        if a["shard_index"] in seen_indices:
            raise ValueError(f"Duplicate shard_index: {a['shard_index']}")
        seen_indices.add(a["shard_index"])

        if a["holderid"] in seen_ids:
            raise ValueError(f"Duplicate holderid: {a['holderid']}")
        seen_ids.add(a["holderid"])

        if a["assignment_id"] in seen_assignment_ids:
            raise ValueError(f"Duplicate assignment_id: {a['assignment_id']}")
        seen_assignment_ids.add(a["assignment_id"])

    print("Batch validated successfully.")


# ----------------------------
# Per-holder payload extraction
# ----------------------------

def build_holder_payload(batch: dict, assignment_entry: dict) -> dict:
    """Metadata + ONLY this holder's shard - never include other holders' data."""
    secret = batch["secret"]
    return {
        "secretid": secret["secretid"],
        "secret_desc": secret["secret_desc"],
        "totalshards": secret["totalshards"],
        "threshold": secret["threshold"],
        "encoding": secret["encoding"],
        "status": secret.get("status"),
        "assignment_id": assignment_entry["assignment_id"],
        "holderid": assignment_entry["holderid"],
        "holder_name": assignment_entry["holder_name"],
        "shard_index": assignment_entry["shard_index"],
        "backend_type": assignment_entry.get("backend_type"),
        "backend_ref": assignment_entry.get("backend_ref"),
        "shard_value": assignment_entry["shard_value"],
    }


# ----------------------------
# AES-256-GCM encrypt / decrypt
# ----------------------------

def encrypt_payload(payload: dict, passphrase: str) -> bytes:
    """
    Returns: salt(16) || nonce(12) || ciphertext+tag
    Packed together so decryption has everything it needs from one blob.
    """
    plaintext = json.dumps(payload).encode("utf-8")
    salt = os.urandom(16)
    nonce = os.urandom(12)
    key = derive_key(passphrase, salt)
    aesgcm = AESGCM(key)
    ciphertext = aesgcm.encrypt(nonce, plaintext, associated_data=None)
    return salt + nonce + ciphertext


def decrypt_payload(encrypted_blob: bytes, passphrase: str) -> dict:
    """Reverse of encrypt_payload - for verification / holder-side tooling."""
    salt = encrypted_blob[:16]
    nonce = encrypted_blob[16:28]
    ciphertext = encrypted_blob[28:]
    key = derive_key(passphrase, salt)
    aesgcm = AESGCM(key)
    plaintext = aesgcm.decrypt(nonce, ciphertext, associated_data=None)
    return json.loads(plaintext.decode("utf-8"))


# ----------------------------
# Restricted local file writes
# ----------------------------

def write_restricted_file(path: str, data: bytes):
    """Write a file and restrict permissions to owner-only (rw-------)."""
    with open(path, "wb") as f:
        f.write(data)
    try:
        os.chmod(path, stat.S_IRUSR | stat.S_IWUSR)  # 600, Unix only
    except (AttributeError, NotImplementedError, OSError):
        pass  # not fully supported on all platforms (e.g. Windows) - fine to skip


def make_restricted_dir(path: str):
    os.makedirs(path, exist_ok=True)
    try:
        os.chmod(path, stat.S_IRWXU)  # 700, Unix only
    except (AttributeError, NotImplementedError, OSError):
        pass


# ----------------------------
# Email delivery (encrypted file only - key is NEVER emailed)
# ----------------------------

class EmailConfig:
    """Reads SMTP settings from environment variables - never hardcode credentials."""

    def __init__(self):
        self.smtp_host = os.environ["SMTP_HOST"]
        self.smtp_port = int(os.environ.get("SMTP_PORT", 587))
        self.smtp_user = os.environ["SMTP_USER"]
        self.smtp_password = os.environ["SMTP_PASSWORD"]
        self.from_addr = os.environ.get("SMTP_FROM", self.smtp_user)


def send_encrypted_shard_email(config: EmailConfig, assignment_entry: dict,
                                encrypted_blob: bytes):
    assignment_id = assignment_entry["assignment_id"]

    msg = EmailMessage()
    msg["Subject"] = f"Break-glass shard assignment ({assignment_id})"
    msg["From"] = config.from_addr
    msg["To"] = assignment_entry["holder_email"]

    msg.set_content(
        f"Hello {assignment_entry['holder_name']},\n\n"
        f"Attached is your encrypted break-glass shard for assignment "
        f"'{assignment_id}'.\n"
        f"Your shard index is {assignment_entry['shard_index']}.\n\n"
        f"The decryption passphrase will be provided to you separately, "
        f"through a different channel. Do not store the passphrase alongside "
        f"this file, and do not forward this email with the passphrase attached.\n\n"
        f"Please confirm receipt to the coordinator once you have decrypted "
        f"the file successfully.\n"
    )

    filename = f"shard_{assignment_entry['holderid']}_{assignment_id}.enc"
    msg.add_attachment(
        encrypted_blob,
        maintype="application",
        subtype="octet-stream",
        filename=filename,
    )

    with smtplib.SMTP(config.smtp_host, config.smtp_port) as server:
        server.starttls()
        server.login(config.smtp_user, config.smtp_password)
        server.send_message(msg)

    print(f"  -> Emailed encrypted shard to {assignment_entry['holder_email']} ({assignment_entry['holderid']})")


# ----------------------------
# Orchestration: process full batch
# ----------------------------

def process_and_send_batch(batch: dict, email_config: EmailConfig,
                            enc_dir: str = "shards_out",
                            key_dir: str = "keys_out",
                            send_email: bool = True) -> list:
    make_restricted_dir(enc_dir)
    make_restricted_dir(key_dir)

    results = []

    for assignment_entry in batch["shard_assignments"]:
        holder_id = assignment_entry["holderid"]
        assignment_id = assignment_entry["assignment_id"]
        print(f"Processing holder '{holder_id}' (assignment '{assignment_id}')...")

        # Unique passphrase per holder - generated fresh each iteration
        passphrase = generate_passphrase()

        payload = build_holder_payload(batch, assignment_entry)
        encrypted_blob = encrypt_payload(payload, passphrase)

        enc_filename = os.path.join(
            enc_dir, f"shard_{holder_id}_{assignment_id}.enc"
        )
        write_restricted_file(enc_filename, encrypted_blob)

        key_filename = os.path.join(
            key_dir, f"shard_{holder_id}_{assignment_id}.key"
        )
        write_restricted_file(key_filename, passphrase.encode("utf-8"))

        if send_email:
            send_encrypted_shard_email(email_config, assignment_entry, encrypted_blob)

        results.append({
            "holder_id": holder_id,
            "assignment_id": assignment_id,
            "email": assignment_entry["holder_email"],
            "enc_file": enc_filename,
            "key_file": key_filename,
        })

    return results


# ----------------------------
# Decryption verification
# ----------------------------

def verify_decryption(enc_file: str, passphrase: str) -> dict:
    """
    Attempt to decrypt an .enc file with the given passphrase.
    Returns the decrypted payload dict on success.
    Raises a clear, specific error on failure (bad passphrase, corrupted/
    tampered file, or malformed file) rather than failing silently.
    """
    if not os.path.exists(enc_file):
        raise FileNotFoundError(f"Encrypted file not found: {enc_file}")

    with open(enc_file, "rb") as f:
        blob = f.read()

    if len(blob) < 16 + 12 + 16:  # salt + nonce + minimum GCM tag
        raise ValueError(f"File is too short to be a valid encrypted shard: {enc_file}")

    try:
        return decrypt_payload(blob, passphrase)
    except Exception as e:
        # cryptography raises InvalidTag on wrong passphrase OR any tampering/corruption
        raise ValueError(
            "Decryption failed - either the passphrase is incorrect, or the "
            "file was corrupted/tampered with in transit."
        ) from e


def load_passphrase_from_key_file(key_file: str) -> str:
    """Read a passphrase out of a .key file, stripping any trailing newline."""
    if not os.path.exists(key_file):
        raise FileNotFoundError(f"Key file not found: {key_file}")
    with open(key_file) as f:
        return f.read().strip()


def decrypt_shard_file(enc_file: str, passphrase: str = None, key_file: str = None,
                        output_file: str = None) -> dict:
    """
    Reusable, importable decryption function - not tied to the CLI.

    Decrypts a single .enc shard file and returns the payload dict:
        {secretid, assignment_id, secret_desc, totalshards, threshold,
         encoding, holderid, name, shard_index, shard_value}

    :param enc_file: Path to the encrypted .enc file.
    :param passphrase: Passphrase string. Provide this OR key_file, not both.
    :param key_file: Path to a .key file containing the passphrase.
    :param output_file: If given, writes the decrypted payload as plaintext
                         JSON to this path. Use with care - a plaintext file
                         on disk defeats the purpose of encrypting it in the
                         first place; prefer using the returned dict directly
                         in memory and deleting any output_file promptly.
    :return: Decrypted payload as a dict.
    :raises ValueError: if neither/both of passphrase and key_file are given,
                         or if decryption fails (wrong passphrase / corrupted file).
    :raises FileNotFoundError: if enc_file or key_file doesn't exist.
    """
    if (passphrase is None) == (key_file is None):
        raise ValueError("Provide exactly one of 'passphrase' or 'key_file'")

    if key_file is not None:
        passphrase = load_passphrase_from_key_file(key_file)

    payload = verify_decryption(enc_file, passphrase)

    if output_file:
        with open(output_file, "w") as f:
            json.dump(payload, f, indent=2)
        try:
            os.chmod(output_file, stat.S_IRUSR | stat.S_IWUSR)  # 600, Unix only
        except (AttributeError, NotImplementedError, OSError):
            pass
        print(f"Decrypted payload written to {output_file} (plaintext - handle with care)")

    return payload


def run_decrypt_check(enc_file: str, passphrase: str, expected_holder_id: str = None,
                       output_file: str = None):
    """
    CLI-facing check: decrypts, prints the payload, and optionally verifies
    it matches the expected holder (catches accidental key/file mismatches).
    """
    try:
        payload = decrypt_shard_file(enc_file, passphrase=passphrase, output_file=output_file)
    except (FileNotFoundError, ValueError) as e:
        print(f"DECRYPTION FAILED: {e}")
        sys.exit(1)

    print("Decryption succeeded. Payload:")
    print(json.dumps(payload, indent=2))

    if expected_holder_id and payload.get("holderid") != expected_holder_id:
        print(
            f"\nWARNING: decrypted holderid '{payload.get('holderid')}' does not "
            f"match expected '{expected_holder_id}' - you may have paired the "
            f"wrong key with the wrong file."
        )
        sys.exit(1)

    print("\nOK: shard content verified.")


# ----------------------------
# CLI entry point
# ----------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Encrypt/distribute or verify Shamir shard assignments."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # --- distribute subcommand (original behavior) ---
    dist_parser = subparsers.add_parser(
        "distribute", help="Encrypt shards and email them to holders"
    )
    dist_parser.add_argument("batch_file", help="Path to the shard assignment batch JSON file")
    dist_parser.add_argument("--enc-dir", default="shards_out",
                              help="Directory to store encrypted shard files (default: shards_out)")
    dist_parser.add_argument("--key-dir", default="keys_out",
                              help="Directory to store passphrase files (default: keys_out)")
    dist_parser.add_argument("--no-email", action="store_true",
                              help="Skip sending emails; only generate encrypted files + keys locally")

    # --- decrypt subcommand (new) ---
    dec_parser = subparsers.add_parser(
        "decrypt", help="Verify that an encrypted shard file decrypts correctly"
    )
    dec_parser.add_argument("enc_file", help="Path to the .enc file to decrypt")
    group = dec_parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--passphrase", help="Passphrase, passed directly (visible in shell history - prefer --key-file)")
    group.add_argument("--key-file", help="Path to a .key file containing the passphrase")
    dec_parser.add_argument("--expect-holder", default=None,
                             help="Optional: assert the decrypted holderid matches this value")
    dec_parser.add_argument("--output", default=None,
                             help="Optional: write decrypted payload as plaintext JSON to this path "
                                  "(handle with care - defeats encryption once written)")

    args = parser.parse_args()

    if args.command == "decrypt":
        if args.key_file:
            with open(args.key_file) as f:
                passphrase = f.read().strip()
        else:
            passphrase = args.passphrase
        run_decrypt_check(args.enc_file, passphrase, expected_holder_id=args.expect_holder,
                           output_file=args.output)
        return

    # --- args.command == "distribute" ---
    if not os.path.exists(args.batch_file):
        print(f"Error: batch file '{args.batch_file}' not found.")
        sys.exit(1)

    with open(args.batch_file) as f:
        batch = json.load(f)

    try:
        validate_batch(batch)
    except ValueError as e:
        print(f"Batch validation failed: {e}")
        sys.exit(1)

    email_config = None
    if not args.no_email:
        try:
            email_config = EmailConfig()
        except KeyError as e:
            print(f"Missing required SMTP environment variable: {e}")
            sys.exit(1)

    results = process_and_send_batch(
        batch,
        email_config,
        enc_dir=args.enc_dir,
        key_dir=args.key_dir,
        send_email=not args.no_email,
    )

    print(f"\nCompleted: processed {len(results)} holder(s).")
    print(f"Encrypted files: {args.enc_dir}/")
    print(f"Passphrase files: {args.key_dir}/")
    print(
        "\nReminder: passphrases were NOT emailed. Retrieve each holder's "
        "passphrase from the key file and deliver it via a separate channel "
        "(phone, SMS, secrets manager, etc.) - never alongside the .enc file."
    )


if __name__ == "__main__":
    main()