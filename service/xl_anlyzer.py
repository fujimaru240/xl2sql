import re
import openpyxl

from const.cell_pos import CellPos
from const.data_type import DataType
from const.query_templ import QueryTempl
from const.func_templ import FuncTempl
from const.db_kind import DBKind
from util.re_utils import REUtils
from util.xls_utils import XlsUtils

class XlAnalyzer:
  _xls_utils = None

  def __init__(self, file_name, sheet_name):
    """コンストラクタ"""
    self._xls_utils = XlsUtils()
    self._xls_utils.open_sheet(file_name, sheet_name)

  def get_data_type(self, col, use_float=False):
    """データタイプを取得する"""
    data_type = self._xls_utils.get_cell(row=CellPos.DATA_TYPE_FIRST[0], col=col)
    data_type = r'{0}'.format(data_type)
    for key, patterns in DataType.TYPES.items():
      if not isinstance(patterns, tuple):
        if REUtils.is_match(data_type, patterns):
          if use_float and self.check_float(key, data_type):
            return DataType.TYPE_FLOAT
          return key
        continue
      for pattern in patterns:
        if REUtils.is_match(data_type, pattern):
          if use_float and self.check_float(key, data_type):
            return DataType.TYPE_FLOAT
          return key

  def check_float(self, key, data_type):
    if key != DataType.TYPE_NUM:
      return False
    float_patterns = DataType.TYPES_EXT[DataType.TYPE_FLOAT]
    if not isinstance(float_patterns, tuple):
      if REUtils.is_match(data_type, float_patterns):
        return True
    for pattern in float_patterns:
      if REUtils.is_match(data_type, pattern):
        return True
    return False

  def get_colmn_name(self, col):
    """列名を取得する"""
    return self._xls_utils.get_cell(row=CellPos.COLUMNS_FIRST[0], col=col)

  def get_field_range(self):
    """対象フィールドの範囲を取得する"""
    return XlAnalyzer.get_range(
      CellPos.COLUMNS_FIRST[1],
      self._xls_utils.get_cell(pos=CellPos.COL_COUNT)
    )

  def get_record_range(self):
    """対象レコードの範囲を取得する"""
    return XlAnalyzer.get_range(
      CellPos.DATA_ROWS_FIRST[0],
      self._xls_utils.get_cell(pos=CellPos.ROW_COUNT)
    )

  @staticmethod
  def get_range(st_pos, add_pos):
    """範囲を取得する"""
    return range(int(st_pos), int(st_pos) + add_pos)

  def convert_data_for_sql(self, data, data_type, col_name=None):
    """データを変換する(SQL用)"""
    seq = self._xls_utils.get_cell(pos=CellPos.SEQUENCE)
    seq_target = self._xls_utils.get_cell(pos=CellPos.SEQUENCE_TARGET_COL)

    # シーケンスの場合
    if seq and seq_target == col_name:
      return self.get_sequence(seq)

    # nullの場合
    if "null" == str(data).lower() or data is None:
      return XlAnalyzer.get_nvl(data, data_type)

    # 型データの変換
    return self.convert_type_data_for_sql(data, data_type)

  def convert_data(self, data, data_type, col_name=None):
    """データを変換する"""
    seq = self._xls_utils.get_cell(pos=CellPos.SEQUENCE)
    seq_target = self._xls_utils.get_cell(pos=CellPos.SEQUENCE_TARGET_COL)

    # nullの場合
    if "null" == str(data).lower() or data is None:
      return None

    # 型データの変換
    return self.convert_type_data(data, data_type)

  def convert_type_data_for_sql(self, data, data_type):
    """型データを変換する(SQL用)"""
    if DataType.TYPE_STR == data_type:
      # 文字列 (シングルクォートで囲む)
      return "\'{0}\'".format(data)
    elif DataType.TYPE_DTTM == data_type:
      # 日付 (日付型に変換)
      return self.get_todate(data)
    elif DataType.TYPE_NUM == data_type:
      # 数値 (そのまま返す)
      return str(data)
    else:
      return str(data)

  def convert_type_data(self, data, data_type):
    """型データを変換する"""
    if DataType.TYPE_FLOAT == data_type:
      return float(data)
    elif DataType.TYPE_NUM == data_type:
      return int(data)
    else:
      return str(data)

  @staticmethod
  def get_nvl(data, data_type):
    """null値を取得する"""
    if DataType.TYPE_NUM == data_type:
      return int(0)
    else:
      return "NULL"

  def get_todate(self, data):
    """Date型に変換した値を取得する"""
    db_kind = self._xls_utils.get_cell(pos=CellPos.DB_KIND)
    date_format = self._xls_utils.get_cell(pos=CellPos.DATE_FORMAT)

    if db_kind in {DBKind.DBNM_ORACLE, DBKind.DBNM_POSTRE}:
      # oracle, postgresql
      return FuncTempl.TO_DATE.format(data, date_format)
    if db_kind == DBKind.DBNM_MYSQL:
      # mysql
      return FuncTempl.DATE_FORMAT.format(data, date_format)
    else:
      return None

  def get_sequence(self, seq):
    """シーケンスを取得する"""
    db_kind = self._xls_utils.get_cell(pos=CellPos.DB_KIND)

    if db_kind == DBKind.DBNM_ORACLE:
      # oracle
      return "{0}.NEXTVAL".format(seq)
    elif db_kind == DBKind.DBNM_POSTRE:
      # postgresql
      return "NEXTVAL(\'{0}\')".format(seq)
    else:
      return None

  def create_insert_sql(self, target_row, target_col=None, fix_value=None):
    """insert文を生成する"""
    columns = ""
    record = ""
    for col in self.get_field_range():
      # データ型を取得
      data_type = self.get_data_type(col)
      # 列名を取得
      col_name = self.get_colmn_name(col)
      # データを取得
      if fix_value and target_col == col_name:
        data = fix_value
      else:
        data = self._xls_utils.get_cell(row=target_row, col=col)
      # データを変換
      record += self.convert_data_for_sql(data, data_type, col_name) + ","
      columns += col_name + ","

    columns = columns[:-1]
    record = record[:-1]

    table_name = self._xls_utils.get_cell(pos=CellPos.TABLE_NAME)
    return QueryTempl.SQL_INSERT.format(table_name, columns, record)

  def create_dict(self, target_row, target_col=None, fix_value=None):
    """dictを生成する"""
    record = {}
    for col in self.get_field_range():
      # データ型を取得
      data_type = self.get_data_type(col, use_float=True)
      # 列名を取得
      col_name = self.get_colmn_name(col)
      # データを取得
      if fix_value and target_col == col_name:
        data = fix_value
      else:
        data = self._xls_utils.get_cell(row=target_row, col=col)
      # データを変換
      coverted_data = self.convert_data(data, data_type, col_name)
      if coverted_data:
        record[col_name] = coverted_data

    return record

  def get_output_file_name(self, ext=None):
    """出力ファイル用のファイル名を取得する"""
    file_name = self._xls_utils.get_cell(pos=CellPos.FILE_NAME)
    if ext:
      return '{0}.{1}'.format(file_name, ext)
    return file_name

  def read_bulk_setting(self, file_name):
    """複製データ群設定を読み込む"""
    target_sheet = self._xls_utils.get_cell(pos=CellPos.SET_BULK_COPY)
    if target_sheet:
      return XlAnalyzer(file_name, target_sheet)

  def get_bulk_data(self):
    """複製データ群を取得する"""

    bulk_data = {}
    col_name = self._xls_utils.get_cell(pos=CellPos.BULK_DATA_COL)
    bulk_data[col_name] = []

    count = self._xls_utils.get_cell(pos=CellPos.BULK_DATA_COUNT)
    st_pos = CellPos.BULK_DATA_START[0]
    for row in range(int(st_pos), int(st_pos)+count):
      value = self._xls_utils.get_cell(row=row, col=CellPos.BULK_DATA_START[1])
      bulk_data[col_name].append(value)

    return bulk_data