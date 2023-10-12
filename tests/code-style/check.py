import os

exclude_dirs = set(['vendor', 'externs'])
exclude_files = set(['jquery_native.js'])

def get_string_from_list(list):
  return "\n".join(list)

def get_files_by_ext(ext):
  result = []
  for root, dirs, files in os.walk('.', True):
    dirs[:] = [d for d in dirs if d not in exclude_dirs]
    files[:] = [f for f in files if f not in exclude_files]
    for file in files:
      if file.endswith(ext):
        result.append(os.path.join(root, file))
  return result

def find_string_in_file(file_path, str):
  with open(file_path, 'rb') as file:
    for line_number, line in enumerate(file):
      if str in line:
        return line_number
  return -1

def get_last_symbol_in_file(file_path):
  with open(file_path, 'rb') as file:
    try:  # catch OSError in case of a one line file 
      file.seek(-1, os.SEEK_END)
    except OSError:
      file.seek(0)
    return file.read(1)
  return b''

def check_file_without_license(files):
  files_without_license = []
  license_header = b'Copyright Ascensio System'
  for file in files:
    if -1 == find_string_in_file(file, license_header):
      files_without_license.append(file)
  if files_without_license:
    raise Exception("Files without license:\n" + get_string_from_list(files_without_license))

def check_file_without_latvian_address(files):
  files_without_latvian_address = []
  latvian_address = b'LV-1050'
  for file in files:
    if -1 == find_string_in_file(file, latvian_address):
      files_without_latvian_address.append(file)
  if files_without_latvian_address:
    raise Exception("Files without latvian adress:\n" + get_string_from_list(files_without_latvian_address))

def check_file_without_lf_ending(files):
  files_without_lf_ending = []
  crlf_ending = b'\r'
  for file in files:
    if -1 != find_string_in_file(file, crlf_ending):
      files_without_lf_ending.append(file)
  if files_without_lf_ending:
    raise Exception("Files without lf ending:\n" + get_string_from_list(files_without_lf_ending))

def check_file_without_newline(files):
  files_without_new_line = []
  for file in files:
    if b'\n' != get_last_symbol_in_file(file):
      files_without_new_line.append(file)
  if files_without_new_line:
    raise Exception("Files without newline:\n" + get_string_from_list(files_without_new_line))

def check_code_style():
  files_js = get_files_by_ext(".js")
  check_file_without_license(files_js)
  check_file_without_latvian_address(files_js)
  check_file_without_lf_ending(files_js)
  check_file_without_newline(files_js)

check_code_style()
