from __future__ import annotations

import pytest

from mcp_sharepoint.references import DocumentReference


def test_document_reference_round_trip() -> None:
    reference = DocumentReference(drive_id="drive+id", item_id="item/id==")

    assert DocumentReference.decode(reference.encode()) == reference


@pytest.mark.parametrize("value", ["", "not base64!", "e30", "W10"])
def test_invalid_document_reference_is_rejected(value: str) -> None:
    with pytest.raises(ValueError):
        DocumentReference.decode(value)
