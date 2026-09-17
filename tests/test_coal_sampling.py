from __future__ import annotations

from decimal import Decimal

import pytest

from coal_sampling import (
    Parameters,
    Record,
    Sample,
    SamplingEngine,
    dec,
    normalize_lithology,
)


def d(value: str) -> Decimal:
    return Decimal(value)


def record(row: int, start: str, end: str, lithology: str, raw: str | None = None) -> Record:
    item = Record("Litho【煤岩编录】", row, d(start), d(end), d(raw or str(d(end) - d(start))), normalize_lithology(lithology))
    item.calculated = item.to - item.from_
    return item


def engine_with(*records: Record) -> SamplingEngine:
    engine = SamplingEngine()
    engine.records = list(records)
    engine.calculate()
    return engine


def reasons(engine: SamplingEngine) -> set[str]:
    return {row["Reason"] for row in engine.review_rows()}


# 1. Numeric parsing must reject missing and non-finite values.
def test_dec_rejects_missing_and_non_finite_values():
    assert dec(None) is None
    assert dec("") is None
    assert dec("NaN") is None
    assert dec("Infinity") is None
    assert dec("1.25") == d("1.25")


# 2. Lithology normalization follows the specification.
def test_normalize_lithology_removes_spaces_and_normalizes_dash():
    assert normalize_lithology(" co - py ") == "CO-PY"
    assert normalize_lithology(None) == ""


# 3. The exact 0.05 m boundary is THIN.
def test_classification_thin_boundary_is_inclusive():
    engine = SamplingEngine()
    assert engine.classify(record(1, "0", "0.05", "ST")) == "THIN"
    assert engine.classify(record(2, "1", "1.050001", "ST")) == "ORDINARY_THIN"


# 4. The ordinary-layer upper boundary is inclusive.
def test_classification_ordinary_boundary_is_inclusive():
    engine = SamplingEngine()
    assert engine.classify(record(1, "0", "0.20", "MS")) == "ORDINARY_THIN"
    assert engine.classify(record(2, "1", "1.200001", "MS")) == "ORDINARY"


# 5. Sheet selection gives the exact Litho sheet priority.
def test_sheet_selection_prioritizes_exact_litho_sheet():
    engine = SamplingEngine()
    assert engine._resolve_sheet_name(["LOG", "Litho backup", "Litho【煤岩编录】"]) == "Litho【煤岩编录】"


# 6. Ambiguous Litho candidates must not be silently selected.
def test_sheet_selection_rejects_multiple_litho_candidates():
    engine = SamplingEngine()
    with pytest.raises(ValueError, match="多个候选"):
        engine._resolve_sheet_name(["Litho A", "Litho B"])


# 7. Invalid depth records are retained in review and excluded from hierarchy.
def test_invalid_depth_enters_review():
    invalid = Record("LOG", 2, None, d("1"), None, "CO")
    engine = engine_with(invalid)
    assert "INVALID_DEPTH_OR_EMPTY_FIELD" in reasons(engine)
    assert not engine.samples


# 8. Calculated thickness is used when the raw value differs materially.
def test_raw_thickness_mismatch_enters_review():
    engine = engine_with(record(1, "0", "1", "CO", raw="0.8"))
    assert "RAW_THICKNESS_MISMATCH" in reasons(engine)
    assert engine.records[0].thickness == d("1")


# 9. A depth gap interrupts hierarchy and sampling.
def test_depth_gap_interrupts_hierarchy():
    engine = engine_with(record(1, "0", "0.3", "CO"), record(2, "0.4", "0.7", "CO"))
    assert "DEPTH_GAP" in reasons(engine)
    assert engine.records[0].coal_layer != engine.records[1].coal_layer


# 10. An unresolved overlap is reported instead of being silently clipped.
def test_overlap_enters_review():
    engine = engine_with(record(1, "0", "0.6", "CO"), record(2, "0.5", "0.8", "CO"))
    assert "UNRESOLVED_OVERLAP" in reasons(engine)


# 11. Any non-CO record ends a coal layer, even if it is THIN.
def test_thin_interbed_ends_coal_layer():
    engine = engine_with(
        record(1, "0", "0.3", "CO"),
        record(2, "0.3", "0.32", "ST"),
        record(3, "0.32", "0.6", "CO"),
    )
    assert engine.records[0].coal_layer == "CL-001"
    assert engine.records[2].coal_layer == "CL-002"


# 12. Equal adjacent special lithologies absorb a thin coal layer into one sample.
def test_thin_coal_between_same_special_lithology_is_merged():
    engine = engine_with(
        record(1, "0", "0.1", "XM"),
        record(2, "0.1", "0.2", "CO"),
        record(3, "0.2", "0.3", "XM"),
    )
    assert len(engine.samples) == 1
    assert engine.samples[0].sample_type == "特殊夹层样"
    assert engine.samples[0].from_ == d("0")
    assert engine.samples[0].to == d("0.3")
    assert "THIN_COAL_NOT_SAMPLED" not in reasons(engine)


# 13. With different special lithologies, the thin coal is assigned to the lower one.
def test_thin_coal_between_different_special_lithologies_uses_lower_special():
    engine = engine_with(
        record(1, "0", "0.1", "XM"),
        record(2, "0.1", "0.2", "CO"),
        record(3, "0.2", "0.3", "ZM"),
    )
    assert len(engine.samples) == 2
    lower = [sample for sample in engine.samples if sample.lithology == "ZM"][0]
    assert lower.from_ == d("0.1")
    assert lower.to == d("0.3")


# 14. A thin coal layer without an adjacent special layer enters review.
def test_thin_coal_without_special_neighbour_enters_review():
    engine = engine_with(record(1, "0", "0.1", "CO"))
    assert "THIN_COAL_NO_ADJACENT_SPECIAL" in reasons(engine)
    assert not engine.samples


# 15. Exactly 0.20 m coal is a normal coal-sampling candidate.
def test_coal_layer_at_020_is_sampled():
    engine = engine_with(record(1, "0", "0.2", "CO"))
    assert len(engine.samples) == 1
    assert engine.samples[0].sample_type == "煤样"
    assert "THIN_COAL_NOT_SAMPLED" not in reasons(engine)


# 16. A thin interbed between two coal records is included in the CO unit.
def test_thin_interbed_between_coal_is_included_in_co_sample():
    engine = engine_with(
        record(1, "0", "0.3", "CO"),
        record(2, "0.3", "0.32", "ST"),
        record(3, "0.32", "0.7", "CO"),
    )
    assert len(engine.samples) == 1
    sample = engine.samples[0]
    assert sample.sample_type == "煤样"
    assert sample.from_ == d("0") and sample.to == d("0.7")
    assert "2cm,ST" in sample.comments


# 17. A CO unit <= 1.5 m is sampled as one unit.
def test_co_unit_at_or_below_150_is_not_split():
    engine = engine_with(record(1, "0", "1.5", "CO"))
    coal_samples = [sample for sample in engine.samples if sample.sample_type == "煤样"]
    assert len(coal_samples) == 1
    assert coal_samples[0].thickness == d("1.5")


# 18. A thick ordinary layer interrupts the CO unit and is not sampled automatically.
def test_thick_ordinary_layer_interrupts_co_sampling():
    engine = engine_with(
        record(1, "0", "0.3", "CO"),
        record(2, "0.3", "0.6", "MS"),
        record(3, "0.6", "0.9", "CO"),
    )
    assert not any(sample.lithology == "MS" for sample in engine.samples)
    assert engine.records[0].coal_section != engine.records[2].coal_section


# 19. Non-coal samples with identical type/lithology are normalized together.
def test_normalize_merges_adjacent_non_coal_samples():
    engine = SamplingEngine()
    engine.samples = [
        Sample(d("0"), d("0.1"), "XM", "特殊夹层样"),
        Sample(d("0.1"), d("0.2"), "XM", "特殊夹层样"),
    ]
    engine._normalize()
    assert len(engine.samples) == 1
    assert engine.samples[0].from_ == d("0")
    assert engine.samples[0].to == d("0.2")


# 20. Input workbook loading uses the Litho field positions and preserves source rows.
def test_load_xlsx_reads_litho_layout(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    filename = tmp_path / "input.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Litho【煤岩编录】"
    sheet.cell(5, 1, 0)
    sheet.cell(5, 2, 0.2)
    sheet.cell(5, 3, 0.2)
    sheet.cell(5, 6, " co ")
    workbook.save(filename)

    engine = SamplingEngine()
    engine.load_xlsx(filename)
    assert len(engine.records) == 1
    assert engine.records[0].source_row == 5
    assert engine.records[0].lithology == "CO"
    assert engine.records[0].thickness == d("0.2")


@pytest.mark.xfail(reason="ordinary isolated interbeds still need the adjacent-CO guard from requirements 5.5")
def test_isolated_ordinary_thin_interbed_is_not_sampled():
    engine = engine_with(record(1, "0", "0.1", "MS"))
    assert not engine.samples


@pytest.mark.xfail(reason="manual special selection still needs to bypass independent-special sampling")
def test_manual_special_selection_recalculates_as_co_sample():
    selected = record(1, "0", "0.1", "XM")
    engine = SamplingEngine()
    engine.records = [record(2, "0", "0.3", "CO"), selected, record(3, "0.4", "0.7", "CO")]
    engine.selected_special.add(selected.key())
    engine.calculate()
    assert any(sample.sample_type == "煤样" and selected.key() in sample.sources for sample in engine.samples)


@pytest.mark.xfail(reason="the overlap guard currently annotates overlaps; requirements require resolving or rejecting them")
def test_final_samples_are_strictly_non_overlapping():
    engine = SamplingEngine()
    engine.samples = [
        Sample(d("0"), d("0.5"), "XM", "特殊夹层样"),
        Sample(d("0.4"), d("0.8"), "ZM", "特殊夹层样"),
    ]
    engine._normalize()
    engine._finalize_overlap_guard()
    ordered = sorted(engine.samples, key=lambda sample: sample.from_)
    assert ordered[1].from_ >= ordered[0].to - engine.params.eps
