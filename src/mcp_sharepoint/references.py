"""Stateless document reference encoding."""

from __future__ import annotations

import base64
import binascii
import json
from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class DocumentReference:
    drive_id: str
    item_id: str

    def encode(self) -> str:
        payload = json.dumps(
            {"v": 1, "d": self.drive_id, "i": self.item_id},
            separators=(",", ":"),
        ).encode("utf-8")
        return base64.urlsafe_b64encode(payload).decode("ascii").rstrip("=")

    @classmethod
    def decode(cls, value: str) -> DocumentReference:
        if not isinstance(value, str) or not value or len(value) > 4_096:
            raise ValueError("Invalid document reference")
        try:
            padding = "=" * (-len(value) % 4)
            raw = base64.b64decode(value + padding, altchars=b"-_", validate=True)
            payload = json.loads(raw)
        except (binascii.Error, UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ValueError("Invalid document reference") from exc

        if not isinstance(payload, dict) or payload.get("v") != 1:
            raise ValueError("Unsupported document reference")
        drive_id = payload.get("d")
        item_id = payload.get("i")
        if not isinstance(drive_id, str) or not drive_id:
            raise ValueError("Invalid document reference")
        if not isinstance(item_id, str) or not item_id:
            raise ValueError("Invalid document reference")
        return cls(drive_id=drive_id, item_id=item_id)
