from dataclasses import dataclass


@dataclass(frozen=True)
class Fee:
    tuition_fee: str
    miscellaneous_fee_term: str
    lunch_fee: str
    dessert_fee: str
    material_fee: str
    activity_fee: str
    miscellaneous_fee_month: str
    first_child: str
    second_child: str
    third_child: str
    low_income_households: str
