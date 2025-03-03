import json


class ProgramsData:
    date: str = ""
    cipher: str = ""
    programs: dict[str, dict] = {}
    total_time: str = ""

    def to_dict(self) -> dict:
        return {"date": self.date, "cipher": self.cipher, "programs": self.programs, "total_time": self.total_time}


class ProgramsDataEncoder(json.JSONEncoder):

    def default(self, obj):
        if isinstance(obj, ProgramsData):
            return obj.__dict__
        return json.JSONEncoder.default(self, obj)
