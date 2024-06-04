import tkinter as tk
from tkinter import ttk
from tkinter import filedialog


class FileControlFrame(ttk.Frame):
    PATH_LABEL_TEXT = 'File path'
    OPEN_BUTTON_TEXT = '...'
    OPEN_BUTTON_WIDTH = 2
    UNPROTECT_BUTTON_TEXT = 'UnProtect'
    PATH_ENTRY_WIDTH = 100

    def __init__(self, master) -> None:
        super().__init__(master=master)
        # Set instance var
        self._init_instance_var()
        # Create widgets
        self._create_widget()

    # Init Method
    def _init_instance_var(self) -> None:
        self.file_path = tk.StringVar()

    def _create_widget(self) -> None:
        path_label = ttk.Label(master=self,
                               text=FileControlFrame.PATH_LABEL_TEXT)
        open_button = ttk.Button(master=self,
                                 text=FileControlFrame.OPEN_BUTTON_TEXT,
                                 width=FileControlFrame.OPEN_BUTTON_WIDTH,
                                 command=lambda: self._open_file_dialog_command())
        self.unprotect_button = ttk.Button(master=self,
                                           text=FileControlFrame.UNPROTECT_BUTTON_TEXT)
        path_entry = ttk.Entry(master=self,
                               textvariable=self.file_path,
                               width=FileControlFrame.PATH_ENTRY_WIDTH)
        x_scrollbar = ttk.Scrollbar(master=self,
                                    orient=tk.HORIZONTAL,
                                    command=path_entry.xview)
        # Config
        path_entry.configure(xscrollcommand=x_scrollbar.set)
        # Pack
        path_label.grid(column=0, row=0)
        path_entry.grid(column=1, row=0)
        x_scrollbar.grid(column=1, row=1, sticky=tk.EW)
        open_button.grid(column=2, row=0)
        self.unprotect_button.grid(column=3, row=0)
        self.pack(anchor=tk.CENTER)

    # Button Command
    def _open_file_dialog_command(self,
                                  initial_dir: str = R'C:\Users\t.itihasi\Downloads\test',
                                  file_types: list = [('Excell ファイル', '*.xlsx')]) -> None:
        file_path = filedialog.askopenfilename(initialdir=initial_dir,
                                               filetypes=file_types)
        self.file_path.set(file_path)
