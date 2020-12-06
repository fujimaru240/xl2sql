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
  def __init__(self, file_name, sheet_name):
    """コンストラクタ"""
    XlsUtils().open_sheet(file_name, sheet_name)

  def get_data_type(self, col):
    """データタイプを取得する"""
    data_type = XlsUtils().get_cell(row=CellPos.DATA_TYPE_FIRST[0], col=col)
    data_type = r'{0}'.format(data_type)
    for key, patterns in DataType.TYPES.items():
      if not isinstance(patterns, tuple):
        if REUtils.is_match(data_type, patterns):
          return key
        continue
      for pattern in patterns:
        if REUtils.is_match(data_type, pattern):
          return key

  def get_colmn_name(self, col):
    """列名を取得する"""
    return XlsUtils().get_cell(row=CellPos.COLUMNS_FIRST[0], col=col)

  def get_field_range(self):
    """対象フィールドの範囲を取得する"""
    return XlAnalyzer.get_range(
      CellPos.COLUMNS_FIRST[1],
      XlsUtils().get_cell(pos=CellPos.COL_COUNT)
    )

  def get_record_range(self):
    """対象レコードの範囲を取得する"""
    return XlAnalyzer.get_range(
      CellPos.DATA_ROWS_FIRST[0],
      XlsUtils().get_cell(pos=CellPos.ROW_COUNT)
    )

  @staticmethod
  def get_range(st_pos, add_pos):
    """範囲を取得する"""
    return range(int(st_pos), int(st_pos) + add_pos)

  def convert_data(self, data, data_type, col_name=None):
    """データを変換する"""
    seq = XlsUtils().get_cell(pos=CellPos.SEQUENCE)
    seq_target = XlsUtils().get_cell(pos=CellPos.SEQUENCE_TARGET_COL)

    # シーケンスの場合
    if seq and seq_target == col_name:
      return self.get_sequence(seq)

    # nullの場合
    if "null" == str(data).lower() or data is None:
      return XlAnalyzer.get_nvl(data, data_type)

    # 型データの変換
    return self.convert_type_data(data, data_type)

  def convert_type_data(self, data, data_type):
    """型データを変換する"""
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

  @staticmethod
  def get_nvl(data, data_type):
    """null値を取得する"""
    if DataType.TYPE_NUM == data_type:
      return int(0)
    else:
      return "NULL"

  def get_todate(self, data):
    """Date型に変換した値を取得する"""
    db_kind = XlsUtils().get_cell(pos=CellPos.DB_KIND)
    date_format = XlsUtils().get_cell(pos=CellPos.DATE_FORMAT)

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
    db_kind = XlsUtils().get_cell(pos=CellPos.DB_KIND)

    if db_kind == DBKind.DBNM_ORACLE:
      # oracle
      return "{0}.NEXTVAL".format(seq)
    elif db_kind == DBKind.DBNM_POSTRE:
      # postgresql
      return "NEXTVAL(\'{0}\')".format(seq)
    else:
      return None

  def create_insert_sql(self, target_row):
    """insert文を生成する"""
    columns = ""
    record = ""
    for col in self.get_field_range():
      # データ型を取得
      data_type = self.get_data_type(col)
      # 列名を取得
      col_name = self.get_colmn_name(col)
      # データを取得
      data = XlsUtils().get_cell(row=target_row, col=col)
      # データを変換
      record += self.convert_data(data, data_type, col_name) + ","
      columns += col_name + ","

    columns = columns[:-1]
    record = record[:-1]

    table_name = XlsUtils().get_cell(pos=CellPos.TABLE_NAME)
    return QueryTempl.SQL_INSERT.format(table_name, columns, record)

  def get_output_file_name(self, ext=None):
    """出力ファイル用のファイル名を取得する"""
    file_name = XlsUtils().get_cell(pos=CellPos.FILE_NAME)
    if ext:
      return '{0}.{1}'.format(file_name, ext)
    return file_name