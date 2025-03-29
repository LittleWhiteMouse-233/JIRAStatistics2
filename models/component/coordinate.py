from dataclasses import dataclass
import itertools


@dataclass(slots=True, frozen=True)
class CoordinateValue:
    value: str | int | list[str | int] | tuple[str | int] | None
    label: str | None

    @property
    def iterable_value(self):
        if type(self.value) in (list, tuple):
            return self.value
        else:
            return (self.value,)


class CoordinateSet:
    def __init__(self, row_coord_series: tuple[CoordinateValue], col_coord_series: tuple[CoordinateValue]):
        self.__row_cs = row_coord_series
        self.__col_cs = col_coord_series

    @property
    def cs_label(self):
        row_label = '-'.join([rl for rl in map(lambda x: x.label, self.__row_cs) if rl is not None])
        col_label = '-'.join([cl for cl in map(lambda x: x.label, self.__col_cs) if cl is not None])
        if row_label or col_label:
            return f'[{row_label}, {col_label}]'
        else:
            return None

    @property
    def generator(self):
        row_combination = itertools.product(*list(map(lambda x: x.iterable_value, self.__row_cs)))
        col_combination = itertools.product(*list(map(lambda x: x.iterable_value, self.__col_cs)))
        return itertools.product(row_combination, col_combination)


class CoordinateCache:
    def __init__(self):
        # row coordinate series
        self.__row_cs = []
        # column coordinate series
        self.__col_cs = []

    @staticmethod
    def __add_coord_into_series(coord: dict[str, ...] | int | str | tuple | list | None,
                                coord_series: list[list[CoordinateValue]]):
        if type(coord) is dict:
            coord_series.append(list(map(lambda k, v: CoordinateValue(label=k, value=v), coord.keys(), coord.values())))
        else:
            coord_series.append([CoordinateValue(label=None, value=coord)])

    def add_row_coord(self, coord: dict[str, ...] | int | str | tuple | list | None):
        self.__add_coord_into_series(coord, self.__row_cs)

    def add_multi_row_coord(self, *args: dict[str, ...] | int | str | tuple | list | None):
        for arg in args:
            self.add_row_coord(arg)

    def add_col_coord(self, coord: dict[str, ...] | int | str | tuple | list | None):
        self.__add_coord_into_series(coord, self.__col_cs)

    def add_multi_col_coord(self, *args: dict[str, ...] | int | str | tuple | list | None):
        for arg in args:
            self.add_col_coord(arg)

    @staticmethod
    def __grouping(coord_series: list[list[CoordinateValue]]):
        coord_group = itertools.product(*coord_series)
        for coord_set in coord_group:
            coord_set:tuple[CoordinateValue]
            yield coord_set

    def grouping(self):
        for row_coord_set in self.__grouping(self.__row_cs):
            for col_coord_set in self.__grouping(self.__col_cs):
                yield CoordinateSet(row_coord_series=row_coord_set, col_coord_series=col_coord_set, )


if __name__ == '__main__':
    cache = CoordinateCache()
    cache.add_row_coord({'r1.0': 0, 'r1.1': 1})
    cache.add_row_coord(('r2.a', 'r2.b'))
    cache.add_row_coord({'r3.10': 'abc'})
    cache.add_col_coord('c1.200def')
    cache.add_col_coord(['c2.c', 'c2.d'])
    cache.add_col_coord({'c3.9': 9, 'c3.8': 8})
    for cs in cache.grouping():
        print('coord_set label:', cs.cs_label)
        for c in cs.generator:
            print(c)
