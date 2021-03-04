"""Package to load and parse tables in a Prism pzfx file."""

import xml.etree.ElementTree as ET
import pandas as pd
from itertools import count, chain, cycle
import numpy as np

__version__ = '0.3'


class PrismFileLoadError(Exception):
    pass


def _get_all_text(element):
    s = ''
    for c in element.iter():
        if c.text is not None:
            s += c.text
    return s


def _subcolumn_to_numpy(subcolumn):
    try:
        data = []
        for d in subcolumn.findall('d'):
            if not (('Excluded' in d.attrib) and (d.attrib['Excluded'] == '1')):
                if _get_all_text(d) == '':
                    data.append(None)
                else:
                    data.append(float(_get_all_text(d)))
            else:
                data.append(np.nan)
    except Exception as a:  # If data can't be read silently fail
        print("Couldn't Read a column in the file because: %s" % a)
        data = None

    return np.array(data)


def _parse_xy_table(table):
    xformat = table.attrib['XFormat']
    try:
        yformat = table.attrib['YFormat']
    except KeyError:
        yformat = None
    evformat = table.attrib['EVFormat']

    xscounter = count()
    xsubcolumn_names = lambda: str(next(xscounter))
    if yformat == 'SEN':
        yslist = cycle(['Mean', 'SEM', 'N'])
        ysubcolumn_names = lambda: next(yslist)
    elif yformat == 'upper-lower-limits':
        yslist = cycle(['Mean', 'Lower', 'Upper'])
        ysubcolumn_names = lambda: next(yslist)
    else:
        yscounter = count()
        ysubcolumn_names = lambda: str(next(yscounter))

    columns = {}
    for xcolumn in chain(table.findall('XColumn'), table.findall('XAdvancedColumn')):
        xcolumn_name = _get_all_text(xcolumn.find('Title'))
        for subcolumn in xcolumn.findall('Subcolumn'):
            subcolumn_name = xcolumn_name + '_' + xsubcolumn_names()
            columns[subcolumn_name] = _subcolumn_to_numpy(subcolumn)
    for ycolumn in chain(table.findall('YColumn'), table.findall('YAdvancedColumn')):
        ycolumn_name = _get_all_text(ycolumn.find('Title'))
        for subcolumn in ycolumn.findall('Subcolumn'):
            subcolumn_name = ycolumn_name + '_' + ysubcolumn_names()
            columns[subcolumn_name] = _subcolumn_to_numpy(subcolumn)

    maxlength = max([v.shape[0] if v.shape != () else 0 for v in columns.values()])
    for k, v in columns.items():
        if v.shape != ():
            if v.shape[0] < maxlength:
                columns[k] = np.pad(v, (0, maxlength - v.shape[0]), mode='constant', constant_values=np.nan)
        else:
            columns[k] = np.pad(v, (0, maxlength - 0), mode='constant', constant_values=np.nan)

    return pd.DataFrame(columns)


def _parse_table_to_dataframe(table):
    # table_id = table.attrib['ID']
    tabletype = table.attrib['TableType']

    if tabletype == 'XY' or tabletype == 'TwoWay' or tabletype == 'OneWay':
        df = _parse_xy_table(table)
    else:
        raise PrismFileLoadError('Cannot parse %s tables for now!' % tabletype)

    return df


def read_pzfx(filename):
    """Open and parse the Prism pzfx file given in `filename`.
    Returns a dictionary containing table names as keys and pandas DataFrames as values."""
    tree = ET.parse(filename)
    root = tree.getroot()
    if root.tag != 'GraphPadPrismFile':
        raise PrismFileLoadError('Not a Prism file!')
    if root.attrib['PrismXMLVersion'] != '5.00':
        raise PrismFileLoadError('Can only load Prism files with XML version 5.00!')

    tables = {_get_all_text(table.find('Title')): _parse_table_to_dataframe(table)
              for table in root.findall('Table')}

    return tables


def convert_pzfx_to_excel(tables, output_filename):
    """Takes a `tables` dict (from `read_pzfx`) and writes to an xlsx file, returns nothing."""
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        for df_name, df in tables.items():
            for c in '\\/*[]:?':
                df_name = df_name.replace(c, '')
            df.to_excel(writer, sheet_name=df_name)
    return
