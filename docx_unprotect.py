import os
import zipfile
import shutil
import argparse
import tempfile





VERSION = "1.0"
NAME = "Docx Unprotect"


def delete_folder(folder_path):
    shutil.rmtree(folder_path, ignore_errors=True)


def generate_temp_folder_name():
    temp_dir = tempfile.mkdtemp()
    return temp_dir


def create_docx(docx_path, temp_folder):
    out_docx = zipfile.ZipFile(docx_path, 'w')
    for folder, subfolders, files in os.walk(temp_folder):
        for file in files:
            out_docx.write(os.path.join(folder, file), os.path.relpath(os.path.join(folder, file), temp_folder), compress_type = zipfile.ZIP_DEFLATED)
    out_docx.close()


def unzip_docx(docx_path, temp_folder):
    in_docx = zipfile.ZipFile(docx_path)
    in_docx.extractall(temp_folder)
    in_docx.close()


def read_binary_file(file_path):
    fh = open(file_path, 'rb')
    binary_file = fh.read()
    fh.close()
    return binary_file


def remove_protection(temp_folder):
    with open(os.path.join(os.path.join(temp_folder, "word"), "settings.xml"), 'r') as word_xml_file:
        filedata = word_xml_file.read()
        
        first_split = filedata.split("<w:documentProtection", 1)
        first_piece = first_split[0]
        second_split = first_split[1].split("/>", 1)
        second_piece = second_split[1]
        filedata = "".join([first_piece, second_piece])

    with open(os.path.join(os.path.join(temp_folder, "word"), "settings.xml"), 'w') as file:
      file.write(filedata)


def main_unprotect(protected, unprotected):
    temp_folder = generate_temp_folder_name()

    unzip_docx(protected, temp_folder)
    remove_protection(temp_folder)
    create_docx(unprotected, temp_folder)
    delete_folder(temp_folder)


if __name__ == "__main__":
    print(fr'''
  ,_   _,  _, ,  ,    ,  ,,  , ,_ ,_   _,  ___,_, _,___, 
  | \,/ \,/   \_/     |  ||\ | |_)|_) / \,' | /_,/ ' |   
 _|_/'\_/'\_  / \    '\__||'\|'| '| \'\_/   |'\_'\_  |   v{VERSION}
'     '     `'   `       `'  ` '  '  `'     '   `  ` '   
    ''')

    parser = argparse.ArgumentParser(description=f"{NAME}: remove password-based editing restrictions from DOCX files.")
    parser.add_argument("-v", "--version", help="show program version", action="store_true")
    parser.add_argument("-p", "--protected", help="path of input protected docx file")
    parser.add_argument("-u", "--unprotected", help="path of output unprotected docx file")
    args = parser.parse_args()

    if args.version:
        exit()

    if not args.protected or not args.unprotected:
        parser.print_help()
        exit()

    protected_docx = args.protected
    unprotected_docx = args.unprotected

    main_unprotect(protected=protected_docx, unprotected=unprotected_docx)
