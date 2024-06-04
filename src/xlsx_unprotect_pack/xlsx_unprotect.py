import os
import glob
import pathlib
import tkinter as tk
from tkinter import messagebox
import re
import shutil
import zipfile
# View
from . import file_control_frame as fc_vw


class XlsxUnProtect:
    TEMP_DIR = 'unprotect.temp'

    def __init__(self, file_path: str) -> None:
        self.init_instance_var(file_path=file_path)
        self.remove_protection_run()

    def init_instance_var(self, file_path: str) -> None:
        self.org_file_path = file_path
        self.org_dirname = ''
        self.org_basename = ''
        self.cpy_file_path = ''
        self.temp_path = ''
        self.org_file_extension = ''
        self.zip_path = ''
        self.new_zip_path = ''
        self.extracted_path = ''
        self.temp_dir = ''

    def remove_protection_run(self) -> None:
        if self._before_remove_protection():
            self._remove_protection()
            self._after_remove_protection()
            messagebox.showinfo(title='Remove protection', message='Succeed!')

    def _rename_to_zip(self) -> None:
        zip_path = re.sub(self.org_file_extension + '$',
                          '.zip', self.cpy_file_path)
        if not os.path.exists(zip_path):
            os.rename(self.cpy_file_path, zip_path)
        self.zip_path = zip_path

    def _extract_zip(self) -> None:
        extracted_path = os.path.join(self.temp_dir, 'extracted')
        if not os.path.exists(extracted_path):
            with zipfile.ZipFile(self.zip_path, 'r') as zip_ref:
                zip_ref.extractall(extracted_path)
        self.extracted_path = extracted_path

    def _remove_protection_from_sheets(self) -> None:
        if self.org_file_extension == '.xlsx':
            # get all sheets xml files to edit
            sheets_dir = os.path.join(self.extracted_path, 'xl', 'worksheets')
            sheets_glob = os.path.join(sheets_dir, '*.xml')
            sheets = glob.glob(sheets_glob)
            for sheet in sheets:
                with open(sheet, 'r+', encoding='utf-8') as f:
                    text = f.read()
                    # remove protection
                    text = re.sub('<sheetProtection.*?\/>', '', text)
                    f.seek(0)
                    f.write(text)
                    f.truncate()

    def _remove_protection_from_workbook(self) -> None:
        if self.org_file_extension == '.xlsx':
            # get workbook xml file to edit
            workbook = os.path.join(self.extracted_path, 'xl', 'workbook.xml')
            with open(workbook, '+r', encoding='utf-8') as f:
                text = f.read()
                # remove protection
                text = re.sub('<workbookProtection.*?\/>', '', text)
                f.seek(0)
                f.write(text)
                f.truncate()

    def _create_new_zip(self) -> None:
        new_zip_path = re.sub('.zip$', '_unprotected.zip', self.zip_path)
        prefix = re.sub('.zip$', '', new_zip_path)
        cwd = os.getcwd()
        with zipfile.ZipFile(new_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            info = zipfile.ZipInfo(prefix+'\\')
            zipf.writestr(info, '')
            os.chdir(self.extracted_path)
            for root, dirs, files in os.walk('./'):
                for d in dirs:
                    info = zipfile.ZipInfo(os.path.join(root, d)+'\\')
                    zipf.writestr(info, '')
                for file in files:
                    zipf.write(os.path.join(root, file))
        os.chdir(cwd)
        self.new_zip_path = new_zip_path

    def _turn_back_org_ext(self) -> None:
        new_file_path = re.sub('.zip$',
                               self.org_file_extension,
                               self.new_zip_path)
        if os.path.exists(self.new_zip_path):
            shutil.copy(self.new_zip_path, new_file_path)
        self.new_file_path = new_file_path

    def _remove_protection(self) -> None:
        self.org_file_extension = pathlib.Path(self.org_file_path).suffix
        self._rename_to_zip()
        self.temp_dir = os.path.dirname(self.zip_path)
        self._extract_zip()
        self._remove_protection_from_workbook()
        self._remove_protection_from_sheets()
        self._create_new_zip()
        self._turn_back_org_ext()
        shutil.copy(self.new_file_path, self.org_dirname)

    def _before_remove_protection(self) -> bool:
        if self.org_file_path is None:
            return False
        # create a temp path for working in
        self.org_basename = os.path.basename(self.org_file_path)
        self.org_dirname = os.path.dirname(self.org_file_path)
        self.temp_path = os.path.join(self.org_dirname,
                                      XlsxUnProtect.TEMP_DIR)
        if not os.path.exists(self.temp_path):
            os.makedirs(self.temp_path)
        # copy the file to destination dir
        dst_file = os.path.join(self.temp_path,
                                os.path.basename(self.org_file_path))
        if not os.path.exists(dst_file):
            shutil.copy(self.org_file_path, self.temp_path)
            self.cpy_file_path = os.path.join(self.temp_path,
                                              self.org_basename)
        return True

    def _after_remove_protection(self) -> None:
        # remove the temporary path
        if os.path.exists(self.temp_path) and os.path.isdir(self.temp_path):
            shutil.rmtree(self.temp_path)


class View(tk.Tk):
    TITLE = 'UnProtectExcel'
    GEOMETRY = '400x400'

    def __init__(self) -> None:
        super().__init__()
        self.title = self.TITLE
        # self.geometry = self.GEOMETRY
        self.init_instance_var()
        self.create_widget()
        self.mainloop()

    def init_instance_var(self) -> None:
        self.file_path = ''

    def create_widget(self) -> None:
        self.file_control_frame = fc_vw.FileControlFrame(master=self)
        self.file_control_frame\
            .unprotect_button.config(command=self.run_unprotect)

    def run_unprotect(self) -> None:
        XlsxUnProtect(file_path=self.file_control_frame.file_path.get())
