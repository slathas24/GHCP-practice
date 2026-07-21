"""
shamir_reconstruct.py  (pycryptodome edition)

Uses Crypto.Protocol.SecretSharing.Shamir from pycryptodome for the actual
Shamir math - not a hand-rolled implementation.

IMPORTANT: pycryptodome's Shamir.split()/combine() only operate on EXACTLY
16-byte secrets (it's designed for splitting AES-128 keys). Since your
password is variable length, this module:
  1. PKCS7-pads the secret to a multiple of 16 bytes
  2. Splits it into 16-byte blocks
  3. Runs Shamir.split() independently on EACH block, using the same
     (threshold, total_shares) and the same share index per holder
  4. Concatenates each holder's per-block share bytes into one shard_value

So each holder still gets exactly one shard (one index, one shard_value) -
it's just that shard_value is now (16 * num_blocks) bytes instead of 16,
for secrets longer than 16 bytes. Reconstruction reverses this: splits each
holder's shard_value back into 16-byte pieces, runs Shamir.combine() per
block position, concatenates the blocks, then strips the PKCS7 padding.

Requires: pip install pycryptodome

Reconstruction CLI usage:
    python shamir_reconstruct.py holder_asmith.json holder_jdoe.json holder_bwong.json \\
        --threshold 3 --fingerprint <sha256-hex>

Each holder JSON file:
    {
      "secretid": "...",
      "assignment_id": "...",
      "holderid": "...",
      "shard_index": 1,
      "shard_value": "base64-encoded-share-bytes"
    }
"""

import sys
import json
import base64
import hashlib
import argparse
from typing import List, Tuple

from Crypto.Protocol.SecretSharing import Shamir
from Crypto.Util.Padding import pad, unpad

BLOCK_SIZE = 16  # pycryptodome's Shamir operates on exactly 16-byte blocks


# ============================================================
# Split (multi-block, for secrets of any length)
# ============================================================

def split_secret(secret: bytes, threshold: int, total_shares: int) -> List[Tuple[int, bytes]]:
    """
    Splits `secret` (any length) into `total_shares` shares, any `threshold`
    of which reconstruct it. Returns [(share_index, shard_value_bytes), ...].
    """
    if threshold < 2:
        raise ValueError("threshold must be >= 2")
    if total_shares < threshold:
        raise ValueError("total_shares must be >= threshold")

    padded = pad(secret, BLOCK_SIZE)
    blocks = [padded[i:i + BLOCK_SIZE] for i in range(0, len(padded), BLOCK_SIZE)]

    # per_block_shares[b] = list of n tuples (index, 16-byte share) for block b
    per_block_shares = [Shamir.split(threshold, total_shares, block) for block in blocks]

    # Reassemble per holder: concatenate each holder's share across all blocks
    combined = []
    for holder_pos in range(total_shares):
        index = per_block_shares[0][holder_pos][0]  # same index across all blocks
        shard_value = b"".join(per_block_shares[b][holder_pos][1] for b in range(len(blocks)))
        combined.append((index, shard_value))

    return combined


# ============================================================
# Combine (reverses the multi-block split above)
# ============================================================

def combine_shares(shares: List[Tuple[int, bytes]]) -> bytes:
    """
    shares: list of (share_index, shard_value_bytes) - shard_value_bytes
    length must be a multiple of 16 (one or more concatenated blocks).
    Requires at least `threshold` shares (this function doesn't know the
    threshold itself - too few shares produces a WRONG result silently,
    which is why validate_shards_match() + fingerprint verification matter).
    """
    if len(shares) < 2:
        raise ValueError("Need at least 2 shares to reconstruct")

    lengths = {len(s[1]) for s in shares}
    if len(lengths) > 1:
        raise ValueError(f"Share byte lengths don't match: {lengths} - shares are inconsistent")

    shard_len = lengths.pop()
    if shard_len % BLOCK_SIZE != 0:
        raise ValueError(
            f"Share length {shard_len} is not a multiple of {BLOCK_SIZE} - "
            f"shard_value looks corrupted or truncated"
        )

    indices = [s[0] for s in shares]
    if len(indices) != len(set(indices)):
        raise ValueError(f"Duplicate share indices: {indices}")

    num_blocks = shard_len // BLOCK_SIZE
    reconstructed_blocks = []

    for b in range(num_blocks):
        block_shares = [
            (idx, shard_value[b * BLOCK_SIZE:(b + 1) * BLOCK_SIZE])
            for idx, shard_value in shares
        ]
        reconstructed_blocks.append(Shamir.combine(block_shares))

    padded_secret = b"".join(reconstructed_blocks)

    try:
        return unpad(padded_secret, BLOCK_SIZE)
    except ValueError as e:
        # Wrong padding almost always means wrong/insufficient/corrupted shares
        raise ValueError(
            f"Padding check failed after reconstruction ({e}) - this strongly "
            f"suggests one or more shards is wrong, corrupted, or below threshold."
        )


# ============================================================
# Validation
# ============================================================

def validate_shards_match(shard_payloads: list, expected_threshold: int = None):
    if not shard_payloads:
        raise ValueError("No shards submitted.")

    secret_ids = {p["secretid"] for p in shard_payloads}
    if len(secret_ids) > 1:
        raise ValueError(
            f"Shards belong to DIFFERENT secrets: {secret_ids}. "
            f"Do not combine these - confirm each holder is using the correct file."
        )

    holder_ids = [p["holderid"] for p in shard_payloads]
    if len(holder_ids) != len(set(holder_ids)):
        raise ValueError(f"Duplicate holder submission detected: {holder_ids}")

    shard_indices = [p["shard_index"] for p in shard_payloads]
    if len(shard_indices) != len(set(shard_indices)):
        raise ValueError(f"Duplicate shard_index submitted: {shard_indices}")

    if expected_threshold is not None and len(shard_payloads) < expected_threshold:
        raise ValueError(
            f"Only {len(shard_payloads)} shard(s) submitted, but threshold is {expected_threshold}."
        )

    print(f"OK: {len(shard_payloads)} shard(s) validated for secretid '{shard_payloads[0]['secretid']}'.")
    return True


# ============================================================
# Reconstruction CLI
# ============================================================

def load_shard_file(path: str) -> dict:
    with open(path) as f:
        data = json.load(f)

    required = ["secretid", "holderid", "shard_index", "shard_value"]
    for field in required:
        if field not in data:
            raise ValueError(f"{path}: missing required field '{field}'")

    return data


def reconstruct_from_files(file_paths: List[str], expected_threshold: int = None,
                            secret_fingerprint: str = None) -> str:
    payloads = [load_shard_file(p) for p in file_paths]

    validate_shards_match(payloads, expected_threshold=expected_threshold)

    shares = []
    for p in payloads:
        try:
            share_bytes = base64.b64decode(p["shard_value"], validate=True)
        except Exception as e:
            raise ValueError(f"Holder '{p['holderid']}': shard_value is not valid base64: {e}")
        shares.append((p["shard_index"], share_bytes))

    secret_bytes = combine_shares(shares)

    if secret_fingerprint:
        actual = hashlib.sha256(secret_bytes).hexdigest()
        if actual != secret_fingerprint:
            raise ValueError(
                "FINGERPRINT MISMATCH - reconstructed secret does NOT match the "
                "expected fingerprint. Do not use this result; a shard is likely "
                "wrong or corrupted."
            )
        print("Fingerprint verified: reconstruction matches the original secret.")

    # secret_bytes are the raw bytes that were base64-decoded before sharding -
    # re-encode to hand back in the same form the user originally gave you
    return base64.b64encode(secret_bytes).decode("ascii")


def main():
    parser = argparse.ArgumentParser(
        description="Reconstruct a Shamir-split secret (pycryptodome) from holder shard JSON files."
    )
    parser.add_argument("shard_files", nargs="+",
                         help="Paths to holder shard JSON files (need >= threshold of them)")
    parser.add_argument("--threshold", type=int, default=None,
                         help="Expected minimum number of shards required (optional check)")
    parser.add_argument("--fingerprint", default=None,
                         help="Expected secret_fingerprint (sha256 hex) to verify against, if known")
    args = parser.parse_args()

    try:
        result_b64 = reconstruct_from_files(
            args.shard_files,
            expected_threshold=args.threshold,
            secret_fingerprint=args.fingerprint,
        )
    except (ValueError, FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Reconstruction failed: {e}")
        sys.exit(1)

    print("\nReconstructed secret (base64):")
    print(result_b64)
    print("\nThis value was not written to disk or logged. Handle with care.")


if __name__ == "__main__":
    main()