class Result:
    pass


class Success(Result):
    def __init__(self, data: any):
        self.data = data


class Error(Result):
    def __init__(self, exception: Exception):
        self.exception = exception
