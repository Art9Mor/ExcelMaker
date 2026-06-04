from dataclasses import dataclass, field

@dataclass(slots=True)
class SpecItem:
    """
    Описание позиции спецификации.
    """

    number: str
    name: str
    unit: str = "шт"
    quantity: int | float = 1
    price: float = 0.0
    total: float = 0.0
    is_hidden: bool = False

@dataclass(slots=True)
class Section:
    """
    Описание раздела спецификации.
    """

    number: int
    title: str
    items: list[SpecItem] = field(default_factory=list)
    section_total: float = 0.0  # Сумма всех позиций в разделе

@dataclass(slots=True)
class SpecHeader:
    """
    Описание заголовочной информации спецификации.
    """

    project: str | None = None
    equipment_type: str | None = None
    doc_number: str | None = None
    customer: str | None = None
    calculation_date: str | None = None

@dataclass(slots=True)
class SpecDocument:
    """
    Описание документа спецификации.
    """

    source_sheet: str
    target_sheet: str
    header: SpecHeader = field(default_factory=SpecHeader)
    sections: list[Section] = field(default_factory=list)
    grand_total: float = 0.0

    @property
    def total_items(self) -> int:
        return sum(len(s.items) for s in self.sections)
