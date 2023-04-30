from model.LoadTextModel import LoadTextModel
import aiofiles
from model.Result import Result, Success, Error
from model.dataclass.Fee import Fee

note_path = "note.txt"


class ExcelModel(LoadTextModel):
    _instance = None

    # implement singleton pattern
    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super().__new__(cls, *args, **kwargs)
        return cls._instance

    async def read_all_text_from_note(self) -> Result:
        async with aiofiles.open(file=note_path, mode="r", encoding="utf-8-sig") as file:
            try:
                cleaned_text_list = [item.rstrip('\n').split(':', 1)[1] for item in await file.readlines()]
                print(f"Read result successfully: data = \n{cleaned_text_list}")
                return Success(
                    data=Fee(*cleaned_text_list)
                )

            except Exception as error:
                if file:
                    await file.close()
                print(f"Exception: {error}")
                return Error(exception=error)

            finally:
                if file:
                    await file.close()
                    print("File was closed successfully")

    # TODO: load student from excel
    def load_students_from_excel(self):
        print("load students")
