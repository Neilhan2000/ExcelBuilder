from abc import ABC, abstractmethod
from model.Result import Result


class LoadExcelModel(ABC):
    @abstractmethod
    async def read_all_data_from_excel(self) -> Result:
        pass
