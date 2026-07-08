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


@dataclass
class CleanStats:
    input_n: int = 0
    output_n: int = 0
    removed: int = 0
    full_admin: int = 0
    mapped_new: int = 0

    def as_dict(self) -> dict[str, int]:
        return {
            "input_n": self.input_n,
            "output_n": self.output_n,
            "removed": self.removed,
            "full_admin": self.full_admin,
            "mapped_new": self.mapped_new,
        }
