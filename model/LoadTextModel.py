from abc import ABC, abstractmethod
from model.Result import Result


class LoadTextModel(ABC):
    @abstractmethod
    async def read_all_text_from_note(self) -> Result:
        pass
