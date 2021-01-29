"""
craXcel ("crack-cel") - removes password protection from Microsoft Office XML
based applications.
"""
import abc
import argparse
import os
import shutil
import uuid
import zipfile
import mmap
import re
from lxml import etree

APP_NAME = 'craXcel'

APP_ROOT_DIR = os.path.dirname(os.path.realpath(__file__))
APP_SAVE_DIR = os.path.join(APP_ROOT_DIR, 'unlocked')
APP_TEMP_DIR = os.path.join(APP_ROOT_DIR, 'temp')

MICROSOFT_EXCEL = 'MicrosoftExcel'
MICROSOFT_WORD = 'MicrosoftWord'
MICROSOFT_POWERPOINT = 'MicrosoftPowerpoint'

SUPPORTED_EXTENSIONS = {
    '.xlsx': MICROSOFT_EXCEL,
    '.xlsm': MICROSOFT_EXCEL,
    '.docx': MICROSOFT_WORD,
    '.docm': MICROSOFT_WORD,
    '.pptx': MICROSOFT_POWERPOINT,
    '.pptm': MICROSOFT_POWERPOINT
}

class FileInfo():
    """
    Class that encapsulates information related to a specified filepath.
    """

    def __init__(self, filepath):
        self.full_name = filepath
        self.name = os.path.basename(filepath)
        self.directory, self.extension = os.path.splitext(filepath)

class MicrosoftOfficeFile(metaclass=abc.ABCMeta):
    """
    Base class containing common logic for unlocking Microsoft Office XML 
    based applications.
    """

    def __init__(self, user_args, filepath, xml_root_dir_name):
        self._file = FileInfo(filepath)
        self._args = user_args

        # Creates a universally unique path in the app temp directory
        self._temp_processing_dir = os.path.join(APP_TEMP_DIR, str(uuid.uuid4()))

        # The root directory where XML files are stored when unpackaged
        self._xml_root_dir = os.path.join(self._temp_processing_dir, xml_root_dir_name)

        self._vba_filepath = os.path.join(self._xml_root_dir, 'vbaProject.bin')
    
    def unlock(self):
        """
        Unlocks the specified file according to arguments passed in by the user.
        """
        self._unpackage()

        self._remove_application_specific_protection()

        if self._args.vba:
            self._remove_vba_protection()
            
        self._repackage()

        if not self._args.debug:
            self._cleanup()

        print('Completed unlocking file!')

    def _unpackage(self):
        """
        Treats the target file as if it were a ZIP file and extracts the
        underlying XMLs.
        """
        zipfile.ZipFile(self._file.full_name,'r').extractall(self._temp_processing_dir)

        print('File unpacked...')

    def _repackage(self):
        """
        Takes the unpackaged XML files and repackages them into a ZIP file
        with the original file's extension restored. This makes the newly
        repackaged file openable by the original application.
        """
        file_suffix = f'_{APP_NAME}{self._file.extension}'
        filename = self._file.name.replace(self._file.extension, file_suffix)
        unlocked_filepath = os.path.join(APP_SAVE_DIR, filename)

        filepaths = self._get_file_listing(self._temp_processing_dir)
        with zipfile.ZipFile(unlocked_filepath,'w') as repackaged_zip:
            for filepath in filepaths:
                rel_filepath = filepath.replace(self._temp_processing_dir,'')
                repackaged_zip.write(filepath,arcname=rel_filepath)
            
        print('File repackaged...')

    def _cleanup(self):
        """
        Recursively deletes all files in the temporary processing directory.
        """
        shutil.rmtree(self._temp_processing_dir)

        print('Cleaning up temporary files...')

    def _get_file_listing(self, directory):
        """
        Retrieves a list of files from the specified directory.
        """
        filepaths = []
        for root, folder, files in os.walk(directory): 
            for filename in files:
                filepath = os.path.join(root, filename) 
                filepaths.append(filepath)

        return filepaths

    def _remove_protection_element(self, xml_filepath, tag_names_to_remove):
        """
        Reads through the XML in the specified filepath and removes the
        elements containing the specified tag names.
        """
        tree = etree.parse(xml_filepath)
        root = tree.getroot()

        for element in root.iter():
            for tag_name in tag_names_to_remove:
                if tag_name in element.tag:
                    root.remove(element)

        tree.write(xml_filepath, encoding='UTF-8', xml_declaration=True)

    def _remove_vba_protection(self):
        """
        Reads the file's underlying vbaProject.bin file in HEX form,
        replacing the string responsible for protecting the file with a
        password.        
        """
        if os.path.isfile(self._vba_filepath):
            with open(self._vba_filepath, 'r+b') as f:
                m = mmap.mmap(f.fileno(), 0)
                m[:] = re.sub(b'DPB', b'DPx', m[:])

            print('VBA protection removed...')

    @abc.abstractmethod
    def _remove_application_specific_protection(self):
        """
        Removes protection specific to the target application. Abstract method
        that requires implementation in all child classes.
        """

class MicrosoftExcel(MicrosoftOfficeFile):
    """
    Class encapsulating all specifc fields and logic required for the unlocking
    of Microsoft Excel XML based files.
    """

    def __init__(self, user_args, locked_filepath):
        super().__init__(user_args, locked_filepath, 'xl')
        self._workbook_xml_filepath = os.path.join(self._xml_root_dir, 'workbook.xml')
        self._worksheet_xml_dir = os.path.join(self._xml_root_dir, 'worksheets')
        self._workbook_tag_names = ['fileSharing', 'workbookProtection']
        self._worksheet_tag_names = ['sheetProtection']

    def _remove_application_specific_protection(self):
        if self._args.workbook:
            self._remove_workbook_protection()
        elif self._args.worksheet:
            self._remove_worksheet_protection()
        else:
            self._remove_workbook_protection()
            self._remove_worksheet_protection()

    def _remove_workbook_protection(self):
        """
        Takes the workbook XML and removes the protections within.
        """
        self._remove_protection_element(self._workbook_xml_filepath, self._workbook_tag_names)

        print('Workbook protection removed...')

    def _remove_worksheet_protection(self):
        """
        Iterates through the directory holding the worksheet XMLs and removes
        the protections in each file.
        """
        worksheet_xml_filepaths = self._get_file_listing(self._worksheet_xml_dir)

        for xml_filepath in worksheet_xml_filepaths:
            self._remove_protection_element(xml_filepath, self._worksheet_tag_names)

        print('Worksheet protection removed...')

class MicrosoftWord(MicrosoftOfficeFile):
    """
    Class encapsulating all specifc fields and logic required for the unlocking
    of Microsoft Word XML based files.
    """
    
    def __init__(self, user_args, locked_filepath):
        super().__init__(user_args, locked_filepath, 'word')
        self._document_xml_filepath = os.path.join(self._xml_root_dir, 'settings.xml')
        self._document_tag_names = ['writeProtection', 'documentProtection']

    def _remove_application_specific_protection(self):
        self._remove_protection_element(self._document_xml_filepath, self._document_tag_names)

        print('Document protection removed...')

class MicrosoftPowerpoint(MicrosoftOfficeFile):
    """
    Class encapsulating all specifc fields and logic required for the unlocking
    of Microsoft Powerpoint XML based files.
    """
    def __init__(self, user_args, locked_filepath):
        super().__init__(user_args, locked_filepath, 'ppt')
        self._presentation_xml_filepath = os.path.join(self._xml_root_dir, 'presentation.xml')
        self._presentation_tag_names = ['modifyVerifier']

    def _remove_application_specific_protection(self):
        self._remove_protection_element(self._presentation_xml_filepath, self._presentation_tag_names)
        print('Presentation protection removed...')   

def Main():
    """
    Main entry point of the application.
    """
    args = handle_args()

    print('\ncraXcel started')

    if args.list:
        print('\nList mode enabled')
        filepaths = read_list_of_filepaths(args.filepath)
        print(f'{len(filepaths)} files detected')
    else:
        filepaths = [args.filepath]

    files_unlocked = 0
    for locked_filepath in filepaths:
        print(f'\nChecking file {locked_filepath}...')

        if os.path.isfile(locked_filepath):
            file_info = FileInfo(locked_filepath)
            
            # Checks the extension of the file against the dictionary of
            # supported applications, returning the application name.
            try:
                detected_application = SUPPORTED_EXTENSIONS[file_info.extension]
            except:
                detected_application = 'unsupported'

            # Uses the deteted application to create the correct instance.
            if detected_application == MICROSOFT_EXCEL:
                cxl = MicrosoftExcel(args, locked_filepath)
            elif detected_application == MICROSOFT_WORD:
                cxl = MicrosoftWord(args, locked_filepath)
            elif detected_application == MICROSOFT_POWERPOINT:
                cxl = MicrosoftPowerpoint(args, locked_filepath)
            elif file_info.extension == '.txt':
                print('File rejected. Did you mean to use list mode? Try "python craxcel.py --help" for more info.')
                break
            else:
                print('File rejected. Unsupported file extension.')
                break

            print('File accepted...')            

            try:
                cxl.unlock()
                files_unlocked += 1           
            except Exception:
                print(f'An error occured while unlocking {locked_filepath}')

        else:
            print('File not found...')

    print(f'\nSummary: {files_unlocked}/{len(filepaths)} files unlocked')
    print('\ncraXcel finished')

def read_list_of_filepaths(list_filepath):
    """
    Reads a .txt file of line seperated filepaths and returns them as a list.
    """
    return [line.rstrip() for line in open(list_filepath, 'r')]

def handle_args():
    """
    Handles the command line arguments passed in by the user, returns them
    as an args object.
    """
    parser = argparse.ArgumentParser(description='Remove Workbook and Worksheet protection on Microsoft Excel files.')
    parser.add_argument('filepath', help='Target filepath')

    excel_group = parser.add_mutually_exclusive_group()
    excel_group.add_argument('-ws', '--worksheet', action='store_true', 
                        help='microsoft excel files: unlocks the Worksheets only (leaves Workbook Protection intact)')
    excel_group.add_argument('-wb', '--workbook', action='store_true',
                        help='microsoft excel files: unlocks the Workbook only (leaves Worksheet Protection intact)')
    
    parser.add_argument('-vba', '--vba', action='store_true',
                        help='removes projection from the VBA project of the file')

    parser.add_argument('--debug', action='store_true',
                        help='retains the temp folder. Useful for dubugging exceptions')
    parser.add_argument('--list', action='store_true',
                        help='unlock a list of files specified in a line-seperated .txt file')

    return parser.parse_args()

def create_directory_structure():
    """
    Creates the directory structure if it doesn't already exist.
    """
    if not os.path.exists(APP_SAVE_DIR):
        os.mkdir(APP_SAVE_DIR)

    if not os.path.exists(APP_TEMP_DIR):
        os.mkdir(APP_TEMP_DIR)

if __name__ == '__main__':
    create_directory_structure()
    Main()