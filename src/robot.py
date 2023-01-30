import os
import pathlib
from dataclasses import dataclass
from time import sleep
from typing import List
import shutil
import psutil
import pywinauto
import requests
import win32com.client as win32
from pywinauto import Application
from pywinauto.timings import TimeoutError
from pywinauto.application import ProcessNotFoundError
from pywinauto.controls.hwndwrapper import InvalidWindowHandle
import openpyxl
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from colvir import Colvir
from data_structures import BranchInfo, Credentials, Process
from bot_notification import TelegramNotifier


@dataclass
class FilesInfo:
    path: str
    name: str
    full_path: str or None = None
    pid: int or None = None

    def __post_init__(self) -> None:
        self.full_path: str = os.path.join(self.path, self.name)
        pid: str = self.name[:self.name.find('_')]
        self.pid = int(pid)


class Robot:
    def __init__(self, credentials: Credentials, process: Process, notifier: TelegramNotifier, data: List[BranchInfo]) -> None:
        self.credentials: Credentials = credentials
        self.process: Process = process
        self.pids: List[int] = []
        self.restricted_pids: List[int] = []
        self.data: List[BranchInfo] = data

        self.notifier = notifier

        self.kill_colvirs()

        self.chunk_number: int = 5
        self.chunks: List[List[BranchInfo]] = [self.data[i:i + self.chunk_number] for i in
                                               range(0, len(self.data), self.chunk_number)]
        self.done_files: List[str] = []
        self.username: str = os.getlogin()
        self.counter: int = 0
        self.pids_number: int = 0

        self.rejected_data: List[BranchInfo] = []

        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel.DisplayAlerts = False

    def kill_colvirs(self) -> None:
        for proc in psutil.process_iter():
            if not any(process_name in proc.name() for process_name in [self.process.name, 'EXCEL']):
                continue
            try:
                p: psutil.Process = psutil.Process(proc.pid)
                p.terminate()
            except psutil.AccessDenied:
                if 'EXCEL' in proc.name():
                    continue
                self.restricted_pids.append(proc.pid)
                continue

    def kill_process(self, pid) -> None:
        self.pids.remove(pid)
        p: psutil.Process = psutil.Process(pid)
        p.terminate()

    def close_sessions(self) -> None:
        files_info: List[FilesInfo] = []
        for path, subdirs, files in os.walk(rf'C:\xls'):
            for name in files:
                if name in self.done_files or 'xlsx' in name:
                    continue
                files_info.append(FilesInfo(path=path, name=name))

        for file_info in files_info:
            path = file_info.path
            name = file_info.name
            full_path = file_info.full_path
            pid = file_info.pid
            if not os.path.exists(path=full_path) \
                    and os.path.getsize(filename=full_path) == 0:
                continue
            try:
                app: Application = Application(backend='win32').connect(process=pid)
                branch_info = self.find_branch_info(xls_path=path, xls_name=name)
                if branch_info not in self.rejected_data and self.is_errored(app=app):
                    self.rejected_data.append(branch_info)
                    continue
                if not any('Выбор отчета' in win.window_text() for win in app.windows()):
                    continue
                try:
                    os.rename(src=full_path, dst=full_path)
                except OSError:
                    continue
                if not self.is_correct_file(root=path, xls_file_path=name):
                    continue
                self.kill_process(pid=pid)
                self.counter += 1
                message = f'{self.counter}/{self.pids_number}\t{pid} was terminated'
                print(message)
                self.notifier.send_notification(message=message)
                self.done_files.append(name)
                self.convert_to_xlsb(xls_path=path, xls_name=name)
            except (ValueError, ProcessNotFoundError, InvalidWindowHandle):
                continue

    @staticmethod
    def is_errored(app):
        for win in app.windows():
            text = win.window_text()
            if text != 'Выбор отчета':
                continue
            win2 = app.window(handle=win.handle)
            for child in win2.iter_descendants():
                if 'Ошибка при обработке' in child.window_text():
                    return True
        return False

    def is_correct_file(self, root: str, xls_file_path: str) -> bool:
        xls_file_path = os.path.join(root, xls_file_path)
        shutil.copyfile(src=xls_file_path, dst=f'{xls_file_path}_copy.xls')
        xls_file_path = f'{xls_file_path}_copy.xls'
        xlsx_file_path = xls_file_path + 'x'

        if not os.path.exists(path=xlsx_file_path):
            wb = self.excel.Workbooks.Open(xls_file_path)
            wb.SaveAs(xlsx_file_path, FileFormat=51)
            wb.Close()

        workbook: Workbook = openpyxl.load_workbook(xlsx_file_path, data_only=True)
        sheet: Worksheet = workbook.active
        os.unlink(xlsx_file_path)
        os.unlink(xls_file_path)

        return next((True for row in sheet.iter_rows(max_row=50) for cell in row if cell.has_style), False)

    def find_branch_info(self, xls_path, xls_name):
        b_info: BranchInfo or None = None
        for branch_info in self.data:
            path, name = branch_info.save_path, branch_info.file_name
            if os.path.join(xls_path, xls_name[xls_name.find('_') + 1::]) != os.path.join(path, name):
                continue
            b_info = branch_info
        return b_info

    def convert_to_xlsb(self, xls_path: str, xls_name: str) -> None:
        b_info = self.find_branch_info(xls_path=xls_path, xls_name=xls_name)
        full_xls_path = os.path.join(xls_path, xls_name)
        if not b_info:
            print(f'Branch info not found for {full_xls_path}')
            return
        full_xlsb_path = os.path.join(b_info.final_save_path, b_info.final_name)

        try:
            wb = self.excel.Workbooks.Open(full_xls_path)
            wb.SaveAs(full_xlsb_path, 50)
            wb.Close()
            message = f'{full_xlsb_path} successfully converted'
            print(message)
            os.unlink(full_xls_path)
        except Exception as e:
            message = f'could not convert {xls_name}'
            self.rejected_data.append(b_info)
            print(str(e), message)
            pass

    def create_folder_structure(self):
        for branch_info in self.data:
            pathlib.Path(branch_info.final_save_path).mkdir(parents=True, exist_ok=True)

    def run(self) -> None:
        self.create_folder_structure()

        for i, chunk in enumerate(self.chunks):
            for j, branch_info in enumerate(chunk, start=(i * self.chunk_number)):
                message = f'{j + 1}/{len(self.data)} {branch_info}'
                print(message)
                colvir: Colvir = Colvir(
                    pids=self.pids,
                    restricted_pids=self.restricted_pids,
                    credentials=self.credentials,
                    process=self.process,
                    data=branch_info
                )
                colvir.open()
                self.notifier.send_notification(message=f'{j + 1}/{len(self.data)}')
                self.pids.append(colvir.pid)
                print(self.pids)
            self.notifier.send_notification(message=f'{len(self.pids)} processes are opened, starting to close sessions')
            self.counter = 0
            self.pids_number = len(self.pids)
            self.close_sessions()
        self.pids_number = len(self.pids)
        while self.pids:
            self.close_sessions()

        self.pids = []
        for i, branch_info in enumerate(self.rejected_data):
            message = f'{i + 1}/{len(self.rejected_data)} {branch_info}'
            print(message)
            colvir: Colvir = Colvir(
                pids=self.pids,
                restricted_pids=self.restricted_pids,
                credentials=self.credentials,
                process=self.process,
                data=branch_info,
            )
            colvir.open()
            self.notifier.send_notification(message=f'{i + 1}/{len(self.rejected_data)}')
            self.pids.append(colvir.pid)
        print(self.pids)
        self.notifier.send_notification(message=f'{len(self.pids)} processes are opened, starting to close sessions')
        self.pids_number = len(self.pids)
        self.counter = 0
        while self.pids:
            self.close_sessions()
        self.notifier.send_notification(message='Completed')
        self.excel.Quit()
