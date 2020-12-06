import re

class REUtils:

  @staticmethod
  def is_match(target, pattern):
    #大文字小文字の区別をしない
    repatter = re.compile(pattern, re.IGNORECASE)
    if repatter.match(target):
      return True
    return False
