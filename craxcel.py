import argparse
import errno
import os
import shutil
import sys
import zipfile
from lxml import etree

ZIP = '.zip'
TEMP_DIR = '/temp'
BACKUP_EXT = ' - backup'
WORKBOOK_REL = '/xl/workbook.xml'
WORKSHEETS_REL = '/xl/worksheets/'
WORKBOOK_TAG = 'workbookProtection'
WORKSHEET_TAG = 'sheetProtection'

class ExcelFile():
    def __init__(self,filepath):
        self.filepath = filepath
        self.folderpath = os.path.dirname(filepath)
        self.extension = os.path.splitext(filepath)[1]
        self.unpacked_folderpath = self.folderpath + TEMP_DIR
        self.zipped_filepath = self.filepath.replace(self.extension, ZIP)
        self.workbook_xml = self.unpacked_folderpath + WORKBOOK_REL
        self.worksheet_folder = self.unpacked_folderpath + WORKSHEETS_REL

    def _get_backup_filepath(self,backup_suffix=BACKUP_EXT):
        backup_ext = backup_suffix + self.extension

        return self.filepath.replace(self.extension,backup_ext)       

    def _change_extension_to_zip(self):
        os.rename(self.filepath,self.zipped_filepath)

    def _delete_zip(self):
        os.remove(self.zipped_filepath)

    def _change_extension_to_original(self):
        os.rename(self.zipped_filepath,self.filepath)

    def _get_worksheets(self):
        return _get_file_listing(self.worksheet_folder)

    def backup(self):
        shutil.copyfile(self.filepath,self._get_backup_filepath())
        print('Backup created...')

    def unpackage(self):
        self._change_extension_to_zip()
        zipfile.ZipFile(self.zipped_filepath,'r').extractall(self.unpacked_folderpath)
        self._delete_zip()
        print('File unpacked...')

    def unprotect_workbook(self):
        wb = WorkbookXML(self.workbook_xml)
        wb._remove_protection_element()
        print('Workbook unprotected...')

    def unprotect_worksheets(self):
        sheet_count = 0
        for ws_file in self._get_worksheets():
            ws = WorksheetXML(ws_file)
            ws._remove_protection_element()
            sheet_count += 1
            (f'Worksheet {sheet_count} unprotected...')

    def repackage(self):
        filepaths = _get_file_listing(self.unpacked_folderpath)
        with zipfile.ZipFile(self.zipped_filepath,'w') as repackaged_zip:
            for filepath in filepaths:
                rel_filepath = filepath.replace(self.unpacked_folderpath,'')
                repackaged_zip.write(filepath,arcname=rel_filepath)
            
        self._change_extension_to_original()
        print('File repackaged...')

    def cleanup(self):
        shutil.rmtree(self.unpacked_folderpath)
        print('Cleaning up temporary files...')

class WorkbookXML():
    def __init__(self,filepath):
        self.filepath = filepath
        self.protection_tag = WORKBOOK_TAG
       
    def _remove_protection_element(self):
        tree = etree.parse(self.filepath)
        root = tree.getroot()

        for element in root.iter():
            if self.protection_tag in element.tag:
                root.remove(element)

        tree.write(self.filepath, encoding='UTF-8', xml_declaration=True)

class WorksheetXML():
    def __init__(self,filepath):
        self.filepath = filepath
        self.protection_tag = WORKSHEET_TAG

    def _remove_protection_element(self):
        tree = etree.parse(self.filepath)
        root = tree.getroot()

        for element in root.iter():
            if self.protection_tag in element.tag:
                root.remove(element)

        tree.write(self.filepath, encoding='UTF-8', xml_declaration=True)

def _get_file_listing(folder):
        filepaths = []
        for root, folder, files in os.walk(folder): 
            for filename in files:
                filepath = os.path.join(root, filename) 
                filepaths.append(filepath)

        return filepaths
        
def Main():
    parser = argparse.ArgumentParser(description='Remove Workbook and Worksheet protection on Microsoft Excel files.')
    parser.add_argument('file', help='Filepath of the Excel file to be unlocked')
    group = parser.add_mutually_exclusive_group()
    group.add_argument('-ws', '--worksheet', action='store_true', 
                        help='unlocks the Worksheets only (leaves Workbook Protection intact)')
    group.add_argument('-wb', '--workbook', action='store_true',
                        help='unlocks the Workbook only (leaves Worksheet Protection intact)')
    parser.add_argument('--no_backup', action='store_true',
                        help='runs craXcel without making a backup of the original (use at own risk)')
    parser.add_argument('--debug', action='store_true',
                        help=f'retains the {TEMP_DIR} folder. Useful for dubugging exceptions')                        
    args = parser.parse_args()

    print('\nStarting craXcel...')
    print(f'Checking file {args.file}...')

    if os.path.isfile(args.file):
        cxl = ExcelFile(args.file)
        print('File accepted...')

        if args.no_backup == False:
            try:
                cxl.backup()
            except Exception as e:
                print(e)
                print("""
                Could not create backup, terminating program.
                You can try running the program with option --no_backup to circumvent this error (use this at your own risk!)
                """)
                sys.exit()

        try:
            cxl.unpackage()
        except Exception as e:
            print(e)
            print("""
            Could not unpack file, terminating program.
            The Excel format may not be supported. Check the list of supported Excel formats at the github repo.
            """)
            sys.exit()

        if args.worksheet:
            try:
                cxl.unprotect_worksheets()
            except Exception as e:
                print(e)
                if args.debug == False:
                    cxl.cleanup()
        elif args.workbook:
            try:
                cxl.unprotect_workbook()
            except Exception as e:
                print(e)
                if args.debug == False:
                    cxl.cleanup()
        else:
            try:
                cxl.unprotect_worksheets()
            except Exception as e:
                print(e)
            try:                
                cxl.unprotect_workbook()
            except Exception as e:
                print(e)                      

        cxl.repackage()

        if args.debug == False:
            cxl.cleanup()

    else:
        raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), args.file)

    print(f'\ncraXcel suXcessful')

if __name__ == '__main__':
    Main()