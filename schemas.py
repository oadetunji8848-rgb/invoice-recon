"""Pydantic v2 domain models for the invoice-to-ledger reconciliation pipeline.

Money is held as ``Decimal`` end to end so that the validation tolerance
($0.01) is applied to exact values, never to binary floating point.
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal
from enum import Enum
from typing import Annotated, Literal

from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    field_validator,
    model_validator,
)
from pydantic_core import PydanticCustomError

TOLERANCE = Decimal("0.01")
CENT = Decimal("0.01")

Money = Annotated[Decimal, Field(max_digits=14, decimal_places=2)]


def _close(a: Decimal, b: Decimal, tol: Decimal = TOLERANCE) -> bool:
    return abs(a - b) <= tol


class ReasonCode(str, Enum):
    """Machine-readable reasons attached to every routing decision."""

    EXACT_MATCH = "EXACT_MATCH"
    FUZZY_VENDOR = "FUZZY_VENDOR"
    LOW_VENDOR_SIMILARITY = "LOW_VENDOR_SIMILARITY"
    DATE_DRIFT = "DATE_DRIFT"
    LOW_CONFIDENCE = "LOW_CONFIDENCE"
    MULTIPLE_CANDIDATES = "MULTIPLE_CANDIDATES"
    NO_LEDGER_MATCH = "NO_LEDGER_MATCH"
    MATH_ERROR = "MATH_ERROR"
    SCHEMA_ERROR = "SCHEMA_ERROR"
    PARSE_ERROR = "PARSE_ERROR"
    DUPLICATE_INVOICE_ID = "DUPLICATE_INVOICE_ID"


class LineItem(BaseModel):
    model_config = ConfigDict(frozen=True, str_strip_whitespace=True)

    description: str = Field(min_length=1)
    quantity: Decimal = Field(gt=0)
    unit_price: Money = Field(ge=0)
    amount: Money = Field(ge=0)

    @model_validator(mode="after")
    def _check_extension(self) -> "LineItem":
        expected = (self.quantity * self.unit_price).quantize(CENT)
        if not _close(expected, self.amount):
            raise PydanticCustomError(
                "invoice_math",
                "Line item '{desc}': qty x unit price = {expected}, but amount is {amount}",
                {"desc": self.description, "expected": str(expected), "amount": str(self.amount)},
            )
        return self


class InvoiceRecord(BaseModel):
    model_config = ConfigDict(frozen=True, str_strip_whitespace=True)

    invoice_id: str = Field(min_length=3, pattern=r"^[A-Z0-9][A-Z0-9\-/_.]+$")
    vendor_name: str = Field(min_length=2)
    invoice_date: date
    due_date: date | None = None
    currency: str = Field(default="USD", pattern=r"^[A-Z]{3}$")
    line_items: list[LineItem] = Field(min_length=1)
    subtotal: Money = Field(ge=0)
    tax: Money = Field(ge=0)
    total: Money = Field(gt=0)
    source_file: str | None = None

    @field_validator("invoice_id", mode="before")
    @classmethod
    def _normalise_id(cls, v: object) -> object:
        return v.strip().upper() if isinstance(v, str) else v

    @model_validator(mode="after")
    def _check_math(self) -> "InvoiceRecord":
        line_sum = sum((li.amount for li in self.line_items), Decimal("0"))
        if not _close(line_sum, self.subtotal):
            raise PydanticCustomError(
                "invoice_math",
                "Line items sum to {line_sum} but subtotal is {subtotal}",
                {"line_sum": str(line_sum), "subtotal": str(self.subtotal)},
            )
        if not _close(self.subtotal + self.tax, self.total):
            raise PydanticCustomError(
                "invoice_math",
                "Subtotal {subtotal} + tax {tax} = {expected}, but total is {total}",
                {
                    "subtotal": str(self.subtotal),
                    "tax": str(self.tax),
                    "expected": str(self.subtotal + self.tax),
                    "total": str(self.total),
                },
            )
        if self.due_date is not None and self.due_date < self.invoice_date:
            raise PydanticCustomError(
                "invoice_dates",
                "Due date {due} precedes invoice date {inv}",
                {"due": str(self.due_date), "inv": str(self.invoice_date)},
            )
        return self


class LedgerEntry(BaseModel):
    model_config = ConfigDict(frozen=True, str_strip_whitespace=True)

    txn_id: str = Field(min_length=1)
    posted_date: date
    description: str = Field(min_length=1)
    amount: Money = Field(gt=0)


class ExtractionFailure(BaseModel):
    """A PDF that could not be turned into a valid ``InvoiceRecord``."""

    source_file: str
    reason: ReasonCode
    invoice_id: str | None = None
    vendor_name: str | None = None
    total: Decimal | None = None
    errors: list[str] = Field(default_factory=list)


class ReconciliationResult(BaseModel):
    """One routing decision for one invoice (or one rejected PDF)."""

    status: Literal["approved", "needs_review"]
    confidence: float = Field(ge=0, le=100)
    reasons: list[ReasonCode] = Field(default_factory=list)
    detail: str = ""

    source_file: str
    invoice_id: str | None = None
    vendor_name: str | None = None
    invoice_date: date | None = None
    invoice_total: Decimal | None = None

    ledger_txn_id: str | None = None
    ledger_description: str | None = None
    ledger_date: date | None = None
    ledger_amount: Decimal | None = None

    vendor_similarity: float | None = Field(default=None, ge=0, le=100)
    date_delta_days: int | None = None

    @model_validator(mode="after")
    def _approved_needs_ledger(self) -> "ReconciliationResult":
        if self.status == "approved" and self.ledger_txn_id is None:
            raise ValueError("An approved result must reference a ledger transaction")
        return self
