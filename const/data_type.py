class DataType:
  TYPE_STR = "string"
  TYPE_NUM = "number"
  TYPE_DTTM = "datetime"
  
  TYPES = {
    TYPE_STR: ('.*CHAR.*', 'TEXT'),
    TYPE_NUM: ('INT.*', 'LONG', 'FLOAT', 'NUMBER.*', 'REAL', 'DECIMAL'),
    TYPE_DTTM: ('DATE.*', 'TIME.*')
  }
