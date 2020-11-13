class School:
    def __init__(self, school_id: int, name: str):
        self.school_id = school_id
        self.name = name

    @property
    def school_id(self) -> int:
        return self._school_id

    @school_id.setter
    def school_id(self, school_id: int) -> None:
        if type(school_id) != int:
            raise TypeError(f'School ID must be an integer, not {type(school_id)}')
        self._school_id = school_id

    @property
    def name(self) -> str:
        return self._name

    @name.setter
    def name(self, name: str) -> None:
        if type(name) != str:
            raise TypeError(f'School name must be a string, not {type(name)}')
        self._name = name

    def __str__(self) -> str:
        return self.name
