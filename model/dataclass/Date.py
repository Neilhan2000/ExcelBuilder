from dataclasses import dataclass


@dataclass(frozen=True)
class Date:
    year: int
    month: int
    day: int
