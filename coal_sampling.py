"""Coal geology automatic sampling engine.

The engine is intentionally independent from tkinter so it can be tested and used
from scripts. It never writes to the input workbook.

The implementation follows the V1.0 requirement ordering more closely than the
initial prototype, while preserving the public API used by the GUI and scripts.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

EPS = Decimal("0.000001")


@dataclass
class Parameters:
    thin_max: Decimal = Decimal("0.05")
    ordinary_max: Decimal = Decimal("0.20")
    group_gap_max: Decimal = Decimal("2.00")
    board_section_co_max: Decimal = Decimal("3.00")
    coal_layer_no_sample_max: Decimal = Decimal("0.20")
    board_length: Decimal = Decimal("0.20")
    coal_sample_max: Decimal = Decimal("1.00")
    sample_thin_max: Decimal = Decimal("0.05")
    thickness_tolerance: Decimal = Decimal("0.001")
    eps: Decimal = EPS
    special_codes: tuple[str, ...] = ("XM", "ZM", "CO-PY", "XT")
    ordinary_codes: tuple[str, ...] = ("CG", "GV", "SS", "ST", "MS", "BMS")

    def jsonable(self):
        result = asdict(self)
        for key, value in result.items():
            if isinstance(value, Decimal):
                result[key] = str(value)
            elif isinstance(value, tuple):
                result[key] = list(value)
        return result


@dataclass
class Record:
    source_sheet: str
    source_row: int
    from_: Decimal | None
    to: Decimal | None
    raw_thickness: Decimal | None
    lithology: str
    calculated: Decimal | None = None
    classification: str = ""
    coal_layer: str = ""
    coal_section: str = ""
    coal_group: str = ""
    issue: str = ""
    sampled: bool = False
    selected_special: bool = False
    selected_thin_coal: bool = False

    @property
    def valid_depth(self):
        return self.from_ is not None and self.to is not None and self.to > self.from_

    @property
    def thickness(self):
        return self.calculated or Decimal(0)

    def key(self) -> str:
        return f"{self.source_sheet}:{self.source_row}"


@dataclass
class Sample:
    from_: Decimal
    to: Decimal
    lithology: str
    sample_type: str
    coal_group: str = ""
    coal_section: str = ""
    coal_layer: str = ""
    comments: str = ""
    sources: list[str] = field(default_factory=list)
    sample_no: str = ""

    @property
    def thickness(self):
        return self.to - self.from_


def dec(value: Any) -> Decimal | None:
    if value is None or value == "":
        return None
    try:
        value = Decimal(str(value).strip())
        return value if value.is_finite() else None
    except (InvalidOperation, ValueError, TypeError):
        return None


def normalize_lithology(value: Any) -> str:
    text = "" if value is None else str(value).strip().upper()
    for dash in ("－", "—", "–", "−", "﹣"):
        text = text.replace(dash, "-")
    text = "".join(text.split())
    return text.replace(" ", "")


class SamplingEngine:
    def __init__(self, params: Parameters | None = None):
        self.params = params or Parameters()
        self.records: list[Record] = []
        self.samples: list[Sample] = []
        self.reviews: list[dict[str, Any]] = []
        self.selected_special: set[str] = set()
        self.selected_thin_coal: set[str] = set()

    def _resolve_sheet_name(self, sheet_names: list[str]) -> str:
        exact = "Litho【煤岩编录】"
        if exact in sheet_names:
            return exact
        candidates = [n for n in sheet_names if "litho" in n.lower()]
        if candidates:
            if len(candidates) > 1:
                raise ValueError("存在多个候选工作表，请保留唯一的 Litho 或 LOG 工作表")
            return candidates[0]
        candidates = [n for n in sheet_names if n.lower() == "log"]
        if candidates:
            if len(candidates) > 1:
                raise ValueError("存在多个候选工作表，请保留唯一的 Litho 或 LOG 工作表")
            return candidates[0]
        raise ValueError("未找到 Litho【煤岩编录】、Litho 或 LOG 工作表")

    def load_xlsx(self, filename: str | Path):
        """Read the selected sheet twice, preferring cached formula values."""
        import openpyxl

        path = Path(filename)
        formulas = openpyxl.load_workbook(path, data_only=False, read_only=True)
        values = openpyxl.load_workbook(path, data_only=True, read_only=True)
        sheet_name = self._resolve_sheet_name(formulas.sheetnames)
        fs, vs = formulas[sheet_name], values[sheet_name]
        is_litho = "litho" in sheet_name.lower() or sheet_name == "Litho【煤岩编录】"
        start, cols = (5, (1, 2, 3, 6)) if is_litho else (2, (6, 7, 8, 11))

        self.records = []
        self.reviews = []
        for row_no in range(start, fs.max_row + 1):
            try:
                frow = next(fs.iter_rows(min_row=row_no, max_row=row_no, values_only=False))
                vrow = next(vs.iter_rows(min_row=row_no, max_row=row_no, values_only=True))
            except StopIteration:
                continue

            vals = []
            for col in cols:
                cell = frow[col - 1]
                cached = vrow[col - 1]
                is_formula = isinstance(cell.value, str) and cell.value.startswith("=")
                if is_formula and cached is None:
                    vals.append(None)
                else:
                    vals.append(cached if is_formula else cell.value)

            fr, to, raw, lith = vals
            record = Record(sheet_name, row_no, dec(fr), dec(to), dec(raw), normalize_lithology(lith))
            if not record.valid_depth:
                record.issue = "INVALID_DEPTH_OR_EMPTY_FIELD"
                self.reviews.append(self.review(record, record.issue))
            else:
                record.calculated = record.to - record.from_
                if record.raw_thickness is None or abs(record.raw_thickness - record.calculated) > self.params.thickness_tolerance:
                    record.issue = "RAW_THICKNESS_MISMATCH"
                    self.reviews.append(self.review(record, record.issue))
                if not record.lithology:
                    record.issue = "EMPTY_LITHOLOGY"
                    self.reviews.append(self.review(record, record.issue))
            self.records.append(record)

        self.records.sort(key=lambda x: (x.from_ is None, x.from_ or Decimal(0), x.to or Decimal(0), x.source_row))
        self.calculate()

    def review(self, r, reason):
        return {
            "Source Sheet": r.source_sheet,
            "Source Row": r.source_row,
            "From": r.from_,
            "To": r.to,
            "Thickness": r.thickness,
            "Lithology": r.lithology,
            "Reason": reason,
        }

    def classify(self, r: Record):
        if not r.valid_depth or not r.lithology:
            return "INVALID"
        if r.lithology == "CO":
            return "CO"
        if r.thickness <= self.params.thin_max:
            return "THIN"
        if r.lithology in self.params.special_codes:
            return "SPECIAL"
        if r.lithology in self.params.ordinary_codes:
            return "ORDINARY_THIN" if r.thickness <= self.params.ordinary_max else "ORDINARY"
        return "SPECIAL"

    def _apply_manual_selection(self, rs):
        for record in rs:
            key = record.key()
            record.selected_special = key in self.selected_special
            record.selected_thin_coal = key in self.selected_thin_coal

    def calculate(self):
        self.samples = []
        self.reviews = [
            x for x in self.reviews if x.get("Reason") in ("INVALID_DEPTH_OR_EMPTY_FIELD", "RAW_THICKNESS_MISMATCH", "EMPTY_LITHOLOGY")
        ]

        for r in self.records:
            r.classification = self.classify(r)
            r.coal_layer = ""
            r.coal_section = ""
            r.coal_group = ""
            r.sampled = False
            r.selected_special = r.key() in self.selected_special
            r.selected_thin_coal = r.key() in self.selected_thin_coal
            if r.issue in {"DEPTH_GAP", "UNRESOLVED_OVERLAP"}:
                r.sampled = True

        valid = [r for r in self.records if r.valid_depth and r.lithology]
        self._apply_manual_selection(valid)
        self._check_geometry(valid)
        self._identify_hierarchy(valid)
        self._process_thin_coal(valid)
        self._process_independent_layers(valid)
        self._build_co_units(valid)
        self._generate_board_samples(valid)
        self._normalize()
        self._make_reviews(valid)
        self._finalize_overlap_guard()

    def _check_geometry(self, rs):
        covered = None
        for r in rs:
            if covered is not None:
                if r.from_ > covered + self.params.eps:
                    r.issue = "DEPTH_GAP"
                    self.reviews.append(self.review(r, "DEPTH_GAP"))
                elif r.from_ < covered - self.params.eps:
                    r.issue = "UNRESOLVED_OVERLAP"
                    self.reviews.append(self.review(r, "UNRESOLVED_OVERLAP"))
            covered = max(covered or r.to, r.to)

    def _identify_hierarchy(self, rs):
        coal_layers = []
        current = []
        prev = None
        for r in rs:
            if r.issue in {"DEPTH_GAP", "UNRESOLVED_OVERLAP"}:
                if current:
                    coal_layers.append(current)
                    current = []
                prev = r
                continue
            if r.classification == "CO":
                if prev is not None and abs(r.from_ - prev.to) > self.params.eps:
                    if current:
                        coal_layers.append(current)
                    current = []
                current.append(r)
                prev = r
            else:
                if current:
                    coal_layers.append(current)
                    current = []
                prev = r
        if current:
            coal_layers.append(current)

        for layer_index, layer in enumerate(coal_layers, 1):
            layer_id = f"CL-{layer_index:03d}"
            for r in layer:
                r.coal_layer = layer_id

        sections = []
        block = []
        prev = None
        for r in rs:
            if r.issue in {"DEPTH_GAP", "UNRESOLVED_OVERLAP"}:
                if any(x.classification == "CO" for x in block):
                    sections.append(block)
                block = []
                prev = r
                continue
            usable = {"CO", "THIN", "SPECIAL", "ORDINARY_THIN"}
            contiguous = prev is not None and abs(r.from_ - prev.to) <= self.params.eps
            if r.classification in usable and contiguous:
                block.append(r)
            else:
                if any(x.classification == "CO" for x in block):
                    sections.append(block)
                block = [r] if r.classification in usable else []
            prev = r
        if any(x.classification == "CO" for x in block):
            sections.append(block)

        for section_index, section in enumerate(sections, 1):
            section_id = f"CS-{section_index:03d}"
            for r in section:
                r.coal_section = section_id

        group_no = 0
        last_end = None
        for section in sections:
            start = section[0].from_
            stop = section[-1].to
            if last_end is None or start - last_end > self.params.group_gap_max:
                group_no += 1
            group_id = f"CG-{group_no:03d}"
            for r in section:
                r.coal_group = group_id
            last_end = stop

    def _process_thin_coal(self, rs):
        """Requirements 9.2 / 9.3 / 9.4.

        Thin coal should be absorbed by adjacent special/THIN first. If no valid
        adjacent special exists, it becomes a review candidate instead of silently
        disappearing.
        """
        for idx, r in enumerate(rs):
            if r.issue in {"DEPTH_GAP", "UNRESOLVED_OVERLAP"} or not r.coal_layer:
                continue
            if r.classification != "CO" or r.thickness >= self.params.coal_layer_no_sample_max:
                continue

            prev = rs[idx - 1] if idx > 0 else None
            nxt = rs[idx + 1] if idx + 1 < len(rs) else None
            special_candidates = []
            if prev and prev.lithology and prev.classification in {"SPECIAL", "THIN"}:
                special_candidates.append(prev)
            if nxt and nxt.lithology and nxt.classification in {"SPECIAL", "THIN"}:
                special_candidates.append(nxt)

            if not special_candidates:
                self.reviews.append(self.review(r, "THIN_COAL_NO_ADJACENT_SPECIAL"))
                continue

            # merge same-type special layers
            same = []
            if len(special_candidates) >= 2 and special_candidates[0].lithology == special_candidates[-1].lithology:
                same = special_candidates
            if same:
                upper, lower = same[0], same[-1]
                if upper.key() != lower.key():
                    sample = self._sample(upper.from_, lower.to, upper.lithology, "特殊夹层样", r, [upper, r, lower])
                    sample.comments = "10.0cm,CO"
                    self.samples.append(sample)
                    r.sampled = True
                    upper.sampled = True
                    lower.sampled = True
                    for item in (upper, lower):
                        item.sampled = True
                    continue

            target = special_candidates[-1]
            start = min(r.from_, target.from_)
            end = max(r.to, target.to)
            sample = self._sample(start, end, target.lithology, "特殊夹层样", r, [r, target])
            if target.classification == "THIN":
                sample.comments = "10.0cm,CO"
            self.samples.append(sample)
            r.sampled = True
            target.sampled = True

    def _process_independent_layers(self, rs):
        for r in rs:
            if r.sampled or r.issue in {"DEPTH_GAP", "UNRESOLVED_OVERLAP"}:
                continue
            if r.classification == "SPECIAL" and r.thickness > self.params.thin_max:
                self.samples.append(self._sample(r.from_, r.to, r.lithology, "特殊夹层样", r, [r]))
                r.sampled = True
            elif r.classification == "ORDINARY_THIN":
                self.samples.append(self._sample(r.from_, r.to, r.lithology, "普通夹层样", r, [r]))
                r.sampled = True

    def _build_co_units(self, rs):
        unit = []

        def flush():
            nonlocal unit
            if unit:
                self._split_co_unit(unit)
                unit = []

        for idx, r in enumerate(rs):
            if r.issue in {"DEPTH_GAP", "UNRESOLVED_OVERLAP"}:
                flush()
                continue

            selected_in_unit = r.selected_special or r.selected_thin_coal
            is_co = r.classification == "CO" and not r.sampled
            is_thin = (
                r.classification == "THIN"
                and idx > 0
                and idx + 1 < len(rs)
                and rs[idx - 1].classification == "CO"
                and rs[idx + 1].classification == "CO"
                and not r.sampled
            )
            if is_co or is_thin or selected_in_unit:
                unit.append(r)
            else:
                flush()
        flush()

    def _split_co_unit(self, unit):
        if not unit:
            return
        total = unit[-1].to - unit[0].from_
        if total <= Decimal("1.5"):
            chunks = [unit]
        else:
            chunks = []
            start = 0
            segment_start = unit[0].from_
            segment_end = segment_start + self.params.coal_sample_max
            for idx, record in enumerate(unit):
                if record.from_ >= segment_end and idx > start:
                    chunks.append(unit[start:idx])
                    start = idx
                    segment_start = record.from_
                    segment_end = segment_start + self.params.coal_sample_max
                if record.classification == "THIN" and idx > start:
                    tail = sum((x.thickness for x in unit[start:idx+1] if x.classification == "THIN"), Decimal(0))
                    if tail > self.params.sample_thin_max and idx + 1 < len(unit):
                        chunks.append(unit[start:idx])
                        start = idx
                        segment_start = unit[idx].from_
                        segment_end = segment_start + self.params.coal_sample_max
            if start < len(unit):
                chunks.append(unit[start:])

        for chunk in chunks:
            if not chunk:
                continue
            if chunk[-1].to <= chunk[0].from_:
                continue
            self.samples.append(self._sample(chunk[0].from_, chunk[-1].to, "CO", "煤样", chunk[0], chunk))
            for r in chunk:
                r.sampled = True

    def _sample(self, fr, to, lith, typ, owner, parts):
        extras = []
        for r in parts:
            if r.classification == "CO" or r is owner:
                continue
            if r.thickness <= Decimal(0):
                continue
            extras.append(f"{int((r.thickness * Decimal(100)).to_integral_value())}cm,{r.lithology}")
        return Sample(
            fr,
            to,
            lith,
            typ,
            owner.coal_group,
            owner.coal_section,
            owner.coal_layer,
            "; ".join(extras),
            [f"{r.source_sheet}:{r.source_row}" for r in parts],
        )

    def _generate_board_samples(self, rs):
        for group in sorted({r.coal_group for r in rs if r.coal_group}):
            members = [r for r in rs if r.coal_group == group]
            co_total = sum((r.thickness for r in members if r.classification == "CO"), Decimal(0))
            if co_total <= self.params.board_section_co_max:
                continue
            first, last = members[0], members[-1]
            if first.from_ > self.params.board_length:
                self.samples.append(
                    Sample(
                        max(Decimal(0), first.from_ - self.params.board_length),
                        first.from_,
                        "RF-" + first.lithology,
                        "顶板样",
                        group,
                        first.coal_section,
                        first.coal_layer,
                    )
                )
            if last.to + self.params.board_length > last.to:
                self.samples.append(
                    Sample(
                        last.to,
                        last.to + self.params.board_length,
                        "FL-" + last.lithology,
                        "底板样",
                        group,
                        last.coal_section,
                        last.coal_layer,
                    )
                )

    def _normalize(self):
        self.samples = [s for s in self.samples if s.to > s.from_]
        self.samples.sort(key=lambda s: (s.from_, s.to, s.sample_type))

        merged = []
        for s in self.samples:
            if (
                merged
                and s.lithology == merged[-1].lithology
                and s.sample_type == merged[-1].sample_type
                and abs(s.from_ - merged[-1].to) <= self.params.eps
                and s.sample_type != "煤样"
            ):
                merged[-1].to = s.to
                merged[-1].comments = "; ".join(x for x in (merged[-1].comments, s.comments) if x)
                merged[-1].sources.extend(s.sources)
            else:
                merged.append(s)
        self.samples = merged
        for index, s in enumerate(self.samples, 1):
            s.sample_no = f"S{index:04d}"

    def _finalize_overlap_guard(self):
        """Final guard: the result cannot contain overlapping samples."""
        ordered = sorted(self.samples, key=lambda s: (s.from_, s.to))
        previous = None
        for current in ordered:
            if previous is not None and current.from_ < previous.to - self.params.eps:
                current.comments = (current.comments + "; " if current.comments else "") + "OVERLAP_CHECK"
                previous.comments = (previous.comments + "; " if previous.comments else "") + "OVERLAP_CHECK"
            previous = current

    def _make_reviews(self, rs):
        for r in rs:
            if r.issue in {"DEPTH_GAP", "UNRESOLVED_OVERLAP"}:
                continue
            if r.classification == "SPECIAL" and not r.sampled:
                self.reviews.append(self.review(r, "SPECIAL_NOT_SAMPLED"))
            if r.coal_layer and r.thickness < self.params.coal_layer_no_sample_max and not r.sampled:
                self.reviews.append(self.review(r, "THIN_COAL_NOT_SAMPLED"))

    def result_rows(self):
        return [
            {
                "Sample No": s.sample_no,
                "煤组": s.coal_group,
                "煤段": s.coal_section,
                "煤层": s.coal_layer,
                "From": float(s.from_),
                "To": float(s.to),
                "Thickness": float(s.thickness),
                "Lithology": s.lithology,
                "Sample Type": s.sample_type,
                "Comments": s.comments,
                "Sources": "; ".join(s.sources),
            }
            for s in self.samples
        ]

    def review_rows(self):
        return self.reviews

    def original_rows(self):
        out = []
        for r in self.records:
            state = "不取样"
            hits = [s.sample_no for s in self.samples if s.from_ < (r.to or Decimal(0)) - self.params.eps and s.to > (r.from_ or Decimal(0)) + self.params.eps]
            if hits:
                state = "、".join(hits)
            out.append(
                {
                    "Source Row": r.source_row,
                    "From": r.from_,
                    "To": r.to,
                    "Thickness": r.thickness,
                    "Lithology": r.lithology,
                    "Classification": r.classification,
                    "煤层": r.coal_layer,
                    "煤段": r.coal_section,
                    "煤组": r.coal_group,
                    "采样状态": state,
                    "Issue": r.issue,
                }
            )
        return out
