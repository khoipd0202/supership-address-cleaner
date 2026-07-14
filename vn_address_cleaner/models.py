from dataclasses import dataclass, field


OUTPUT_HEADERS = ("POI", "Tên đường", "Cấp 4", "Phường/Xã", "Quận/Huyện", "Tỉnh/TP")


@dataclass(frozen=True)
class CleanResult:
    """A normalized row ready for business-facing Excel output."""

    poi: str = ""
    street: str = ""
    level4: str = ""
    ward: str = ""
    district: str = ""
    province: str = ""
    confidence: float = 0.0
    flags: tuple[str, ...] = field(default_factory=tuple)

    @property
    def has_detail(self) -> bool:
        return bool(self.poi or self.street or self.level4)

    def as_output_row(self) -> list[str]:
        return [self.poi, self.street, self.level4, self.ward, self.district, self.province]

    def as_component_rows(self) -> list[list[str]]:
        rows: list[list[str]] = []
        for poi in _split_values(self.poi, " | "):
            rows.append([poi, "", "", self.ward, self.district, self.province])
        for street in _split_values(self.street, " | "):
            rows.append(["", street, "", self.ward, self.district, self.province])
        for level4 in _split_values(self.level4, "; "):
            rows.append(["", "", level4, self.ward, self.district, self.province])
        return rows or [["", "", "", self.ward, self.district, self.province]]


def _split_values(value: str, separator: str) -> list[str]:
    seen: set[str] = set()
    values: list[str] = []
    for part in str(value or "").split(separator):
        part = part.strip()
        if not part or part in seen:
            continue
        seen.add(part)
        values.append(part)
    return values


@dataclass
class CleanStats:
    input_n: int = 0
    output_n: int = 0
    removed: int = 0
    full_admin: int = 0
    mapped_new: int = 0
    duplicates: int = 0
    review_n: int = 0

    def as_dict(self) -> dict[str, int]:
        return {
            "input_n": self.input_n,
            "output_n": self.output_n,
            "removed": self.removed,
            "full_admin": self.full_admin,
            "mapped_new": self.mapped_new,
            "duplicates": self.duplicates,
            "review_n": self.review_n,
        }
