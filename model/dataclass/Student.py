from dataclasses import dataclass


@dataclass(frozen=True)
class Student:
    name: str
    month_fee: int
    insurance_fee: int
    postpone_fee: int
    child_property: str
    class_type: str
    leave_refund: int
    receive_date: str
