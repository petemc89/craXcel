import argparse
import errno
import os
import shutil
import sys
import uuid
import zipfile
from lxml import etree

APP_NAME = 'craXcel'
USER_HOME_DIR = os.path.expanduser('~')
APP_BASE_DIR = os.path.join(USER_HOME_DIR, APP_NAME)
APP_PROCESSING_DIR = os.path.join(APP_BASE_DIR, 'temp')

class FileInfo():
    def __init__(self, filepath):
        self.full_name = filepath
        self.name = os.path.basename(filepath)
        self.directory, self.extension = os.path.splitext(filepath)    

class ExcelFile():
    def __init__(self, args):
        self.file = FileInfo(args.filepath)
        self._args = args
        self._temp_processing_dir = os.path.join(APP_PROCESSING_DIR, str(uuid.uuid4()))
        self._xml_root_dir = 'xl'
        self._workbook_xml_filepath = os.path.join(self._temp_processing_dir, self._xml_root_dir, 'workbook.xml')
        self._worksheet_xml_dir = os.path.join(self._temp_processing_dir, self._xml_root_dir, 'worksheets')
        self._workbook_tag_names = ['fileSharing', 'workbookProtection']
        self._worksheet_tag_names = ['sheetProtection']

    def unlock(self):
        self._unpackage()

        if self._args.workbook:
            self._remove_workbook_protection()
        elif self._args.worksheet:
            self._remove_worksheet_protection()
        else:
            self._remove_workbook_protection()
            self._remove_worksheet_protection()
            
        self._repackage()

        if self._args.debug == False:
            self._cleanup()

    def _unpackage(self):
        zipfile.ZipFile(self.file.full_name,'r').extractall(self._temp_processing_dir)
        print('File unpacked...')

    def _remove_workbook_protection(self):
        self._remove_protection_element(self._workbook_xml_filepath, self._workbook_tag_names)
        print('Workbook protection removed...')

    def _remove_worksheet_protection(self):
        worksheet_xml_filepaths = self._get_file_listing(self._worksheet_xml_dir)

        for xml_filepath in worksheet_xml_filepaths:
            self._remove_protection_element(xml_filepath, self._worksheet_tag_names)
        print('Worksheet protection removed...')

    def _repackage(self):
        file_suffix = f'_{APP_NAME}{self.file.extension}'
        filename = self.file.name.replace(self.file.extension, file_suffix)
        unlocked_filepath = os.path.join(APP_BASE_DIR, filename)

        filepaths = self._get_file_listing(self._temp_processing_dir)
        with zipfile.ZipFile(unlocked_filepath,'w') as repackaged_zip:
            for filepath in filepaths:
                rel_filepath = filepath.replace(self._temp_processing_dir,'')
                repackaged_zip.write(filepath,arcname=rel_filepath)
            
        print('File repackaged...')

    def _cleanup(self):
        shutil.rmtree(self._temp_processing_dir)
        print('Cleaning up temporary files...')

    def _remove_protection_element(self, xml_filepath, tag_names_to_remove):
        tree = etree.parse(xml_filepath)
        root = tree.getroot()

        for element in root.iter():
            for tag_name in tag_names_to_remove:
                if tag_name in element.tag:
                    root.remove(element)

        tree.write(xml_filepath, encoding='UTF-8', xml_declaration=True)

    def _get_file_listing(self, directory):
            filepaths = []
            for root, folder, files in os.walk(directory): 
                for filename in files:
                    filepath = os.path.join(root, filename) 
                    filepaths.append(filepath)

            return filepaths
        
def Main():
    parser = argparse.ArgumentParser(description='Remove Workbook and Worksheet protection on Microsoft Excel files.')
    parser.add_argument('filepath', help='Filepath of the Excel file to be unlocked')
    group = parser.add_mutually_exclusive_group()
    group.add_argument('-ws', '--worksheet', action='store_true', 
                        help='unlocks the Worksheets only (leaves Workbook Protection intact)')
    group.add_argument('-wb', '--workbook', action='store_true',
                        help='unlocks the Workbook only (leaves Worksheet Protection intact)')
    parser.add_argument('--debug', action='store_true',
                        help=f'retains the temp folder. Useful for dubugging exceptions')                        
    args = parser.parse_args()

    print('\nStarting craXcel...')
    print(f'Checking file {args.filepath}...')

    if os.path.isfile(args.filepath):
        cxl = ExcelFile(args)
        print('File accepted...')

        try:
            cxl.unlock()
        except Exception as e:
            print(e)
            print("""
            Could not unpack file, terminating program.
            The Excel format may not be supported. Check the list of supported Excel formats at the github repo.
            """)
            sys.exit()

    else:
        raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), args.file)

    print(f'\ncraXcel suXcessful')

if __name__ == '__main__':
    Main()