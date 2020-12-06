import os

class ComUtils:

  @staticmethod
  def is_equal_to_either(data, *targets):
    """対象データがいずれかに一致するか判定する"""
    for target in targets:
      if target == data:
        return True
    return False

  @staticmethod
  def is_valid(title, *target_words):
    for word in target_words:
      if word in title:
        return True
    return False

  @staticmethod
  def extract_file_name(file_path):
    return os.path.splitext(os.path.basename(file_path))[0]