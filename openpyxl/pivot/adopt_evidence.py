# paper-xlsx: managed Excel equivalence latch for foreign adoption

"""Adoption eligibility is gated on committed desktop Excel evidence.

The managed release matrix still has unset Excel build and marker-survival
fields. Until those are real, no foreign pivot may return ``eligible=True``.
Tests may patch ``excel_equivalence_proved``.
"""

from __future__ import annotations


def excel_equivalence_proved():
    """Return True only after managed Excel transcripts prove the serializer.

    Successful Paper reopen is not that proof.
    """
    return False
