from abc import ABC, abstractmethod


class LoadTextModel(ABC):
    @abstractmethod
    async def read_all_text_from_note(self) -> str:
        pass
