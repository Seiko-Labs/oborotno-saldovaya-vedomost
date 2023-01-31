import datetime
from time import sleep
from typing import List
import psutil
from pywinauto import Desktop, Application, WindowSpecification
from pywinauto.application import TimeoutError as AppTimeoutError
from pywinauto.base_wrapper import ElementNotEnabled, ElementNotVisible, InvalidElement
from pywinauto.controls.hwndwrapper import DialogWrapper
from pywinauto.findbestmatch import MatchError
from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError, WindowAmbiguousError, WindowNotFoundError
from pywinauto.timings import TimeoutError as TimingsTimeoutError
from data_structures import Credentials, Process, Date, BranchInfo
from utils import Utils


class Colvir:
    def __init__(self, pids: List[int], restricted_pids: List[int],
                 credentials: Credentials, process: Process, data: BranchInfo) -> None:
        self.credentials: Credentials = credentials
        self.process_name: str = process.name
        self.process_path: str = process.path

        self.pid: int or None = None
        self.pids: List[int] = pids
        self.restricted_pids: List[int] = restricted_pids

        self.desktop: Desktop = Desktop(backend='win32')
        self.app: Application or None = None

        self.branch_info: BranchInfo = data
        self.date: Date = Date(start=data.date_from, end=data.date_to)
        self.date_start: str = self.date.start.strftime('%d.%m.%y')
        self.date_end: str = self.date.end.strftime('%d.%m.%y')

        self._date: datetime.datetime = datetime.datetime.strptime(self.date_end, '%d.%m.%y')

        self.utils = Utils()

    def get_current_pid(self) -> int:
        res: int or None = None
        for proc in psutil.process_iter():
            if self.process_name in proc.name() \
                    and proc.pid not in self.pids \
                    and proc.pid not in self.restricted_pids:
                res = proc.pid
        return res

    def open(self) -> None:
        try:
            Application(backend='win32').start(cmd_line=self.process_path)
            self.login()
        except (ElementNotFoundError, TimingsTimeoutError) as e:
            self.retry()
            return
        self.pid: int = self.get_current_pid()
        self.app: Application = Application(backend='win32').connect(process=self.pid)
        self.confirm_warning()
        try:
            self.choose_mode()
        except ElementNotFoundError:
            self.retry()
            return
        try:
            self.run_action()
        except (ElementNotFoundError, TimingsTimeoutError, ElementNotEnabled, ElementAmbiguousError,
                ElementNotVisible, InvalidElement, WindowAmbiguousError, WindowNotFoundError,
                TimingsTimeoutError, MatchError, AppTimeoutError):
            self.retry()
            return
        print(self.pid)

    def run_action(self) -> None:
        mode = self.branch_info.mode
        action = self.branch_info.action

        filter_win: WindowSpecification = self.app.window(title='Фильтр')
        filter_win.wait(wait_for='exists', timeout=60)
        if mode == 'DD7':
            filter_win['Edit6'].wrapper_object().set_text(text='0114')
        elif mode == 'MCLIEN':
            filter_win['Edit2'].wrapper_object().set_text(text='720914400947')
        filter_win['OKButton'].wrapper_object().click()

        sleep(1)

        title: str = 'Субсчета ПС и лицевые счета клиентов' if mode == 'DD7' else 'Картотека физических и юридических лиц '
        main_win: WindowSpecification = self.app.window(title=title, found_index=0)
        main_win.wait(wait_for='exists', timeout=60)
        main_win.wrapper_object().send_keystrokes(keystrokes='{VK_F5}')

        self.prepare_for_export(file_name=self.branch_info.file_name, save_path=self.branch_info.save_path)

        settings_win: WindowSpecification = self.app.window(title='Параметры отчета ')
        settings_win.wait(wait_for='exists', timeout=60)

        if action == 'S_CLI_013':
            settings_win['Edit2'].set_text(text=self.date_end)
            settings_win['Edit4'].set_text(text=self.branch_info.branch)
        else:
            settings_win['Edit2'].wrapper_object().set_text(text=self.date_start)
            sleep(.1)
            settings_win['Edit4'].wrapper_object().set_text(text=self.date_end)
            sleep(.1)
            settings_win['Edit6'].wrapper_object().set_text(text=self.branch_info.branch)
            sleep(.1)
            if action in ['Z_160_GL_020', 'Z_160_GL_003']:
                edit_num: int = 5
                if action == 'Z_160_GL_020':
                    edit_num = 10 if (self.branch_info.branch != '00' or ',' not in self.branch_info.account) else 18
                settings_win[f'Edit{edit_num}'].wrapper_object().set_text(text=self.branch_info.account)
                sleep(.1)
                settings_win['СводныйCheckBox'].wrapper_object().click()
        sleep(.1)
        settings_win['OK'].wrapper_object().click()

    @staticmethod
    def is_active(app) -> bool:
        res: bool = False
        try:
            return app.active()
        except RuntimeError:
            return res

    def login(self) -> None:
        desktop: Desktop = Desktop(backend='win32')
        try:
            login_win = desktop.window(title='Вход в систему')
            login_win.wait(wait_for='exists', timeout=20)
            login_win['Edit2'].wrapper_object().set_text(text=self.credentials.usr)
            login_win['Edit'].wrapper_object().set_text(text=self.credentials.psw)
            login_win['OK'].wrapper_object().click()
        except ElementAmbiguousError:
            windows: List[DialogWrapper] = Desktop(backend='win32').windows()
            for win in windows:
                if 'Вход в систему' not in win.window_text():
                    continue
                self.utils.kill_process(pid=win.process_id())
            raise ElementNotFoundError

    def confirm_warning(self) -> None:
        try:
            self.app.backend.name = 'uia'
            self.app.Dialog.wait(wait_for='exists', timeout=60)
            self.app.Dialog['OK'].click()
        except (MatchError, ElementNotFoundError):
            self.app.backend.name = 'win32'
            self.retry()
            return
        self.app.backend.name = 'win32'

    def choose_mode(self) -> None:
        mode_win: WindowSpecification = self.app.window(title='Выбор режима')
        mode_win['Edit2'].wrapper_object().set_text(text=self.branch_info.mode)
        mode_win['Edit2'].wrapper_object().send_keystrokes(keystrokes='{ENTER}')
        print('successfully logged in')

    def prepare_for_export(self, file_name: str, save_path: str) -> None:
        file_name: str = f'{self.pid}_{file_name}'
        select_win: WindowSpecification = self.app.window(title='Выбор отчета')
        select_win.wait(wait_for='exists', timeout=60)

        select_win.wrapper_object().send_keystrokes(keystrokes='{VK_F9}')
        sleep(.2)

        filter_win: WindowSpecification = self.app.window(title='Фильтр')
        filter_win.wait(wait_for='exists', timeout=60)
        filter_win['Edit4'].wrapper_object().set_text(text=self.branch_info.action)
        filter_win['OK'].wrapper_object().click()

        sleep(.5)

        select_win['Предварительный просмотр'].wrapper_object().click()

        select_win['Экспорт в файл...'].wrapper_object().click()
        sleep(.1)

        file_win: WindowSpecification = self.app.window(title='Файл отчета ')
        file_win.wait(wait_for='exists', timeout=60)
        file_win['Edit4'].wrapper_object().set_text(text=file_name)
        sleep(.1)
        file_win['Edit2'].wrapper_object().set_text(text=save_path)
        sleep(.1)
        try:
            file_win['ComboBox'].wrapper_object().select(11)
        except (IndexError, ValueError):
            pass

        file_win['OK'].wrapper_object().click()

    def kill(self) -> None:
        try:
            p: psutil.Process = psutil.Process(pid=self.pid)
            p.terminate()
        except psutil.NoSuchProcess:
            self.pid: int = self.get_current_pid()
            p: psutil.Process = psutil.Process(self.pid)
            p.terminate()
        sleep(.5)

    def retry(self) -> None:
        self.kill()
        self.open()
