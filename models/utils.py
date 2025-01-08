import re
import pandas as pd
from datetime import datetime
from dateutil.parser import parse
import math
from wcwidth import wcswidth


def clean_string(string: str):
    string = re.sub(r'&amp;', '&', string)
    string = re.sub(r'<br/>', '\n', string)
    string = re.sub(r'[\n\r]+', '\n', string)
    string = re.sub(r'[ \f\t\v]+', ' ', string)
    string = re.sub(r'^\s+', '', string)
    string = re.sub(r'\s+$', '', string)
    # string = re.sub(r'<.*?>', '', string)
    return string


def parse_timestring(timestring: str, time_format: str = None):
    if time_format:
        return datetime.strptime(timestring, time_format)
    else:
        return parse(timestring)


def concat_single_value(centre: pd.Series | pd.DataFrame, left: list = None, right: list = None, repeat: bool = True,
                        columns: list[str] = None):
    def item2series(x):
        return pd.Series([x] * (centre.shape[0] if repeat else 1))

    concat_list = []
    if left:
        for l_item in left:
            concat_list.append(item2series(l_item))
    concat_list.append(centre)
    if right:
        for r_item in right:
            concat_list.append(item2series(r_item))
    concat_df = pd.concat(concat_list, axis=1)
    if columns:
        concat_df.columns = columns
    else:
        concat_df.columns = list(range(concat_df.shape[1]))
    return concat_df


def specific_length_string(origin: str, length: int = 80, suffix: str = '...'):
    shorter = ''
    for i in range(len(origin)):
        shorter = origin[:i + 1]
        if wcswidth(shorter) > length:
            shorter = origin[:i] + suffix
            break
    tab = '\t' * (math.ceil((length + len(suffix)) / 4) - math.floor(wcswidth(shorter) / 4))
    return shorter + tab
