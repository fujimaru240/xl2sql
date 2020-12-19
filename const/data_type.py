class DataType:
  TYPE_STR = "string"
  TYPE_NUM = "number"
  TYPE_DTTM = "datetime"
  TYPE_FLOAT = "float"
  
  TYPES = {
    TYPE_STR: ('.*CHAR.*', 'TEXT'),
    TYPE_NUM: ('INT.*', 'LONG', 'FLOAT', 'NUMBER.*', 'REAL', 'DECIMAL'),
    TYPE_DTTM: ('DATE.*', 'TIME.*'),
  }

  TYPES_EXT = {
    TYPE_FLOAT: '.+\s*,\s*[1-9]{1}[0-9]*'
  }
