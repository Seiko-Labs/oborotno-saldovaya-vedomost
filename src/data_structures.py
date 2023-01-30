import datetime
from dataclasses import dataclass
from typing import Dict


@dataclass
class Credentials:
    usr: str
    psw: str


@dataclass
class Process:
    name: str
    path: str


@dataclass(init=True)
class Date:
    start: str or datetime.datetime
    end: str or datetime.datetime

    def __post_init__(self) -> None:
        self.start = datetime.date.fromisoformat(self.start)
        self.end = datetime.date.fromisoformat(self.end)


@dataclass
class ExcelInfo:
    path: str


@dataclass
class Dimension:
    x: int
    y: int


@dataclass
class BranchInfo:
    branch: str
    account: str = None
    account_name: str = None
    date_from: str = None
    date_to: str = None
    date_diff: int = None
    mode: str = None
    action: str = None
    file_name: str = None
    save_path: str = None
    final_name: str = None
    final_save_path: str = None

    @staticmethod
    def diff_month(start: datetime.datetime, end: datetime.datetime) -> int:
        return end.month - start.month + 1

    @staticmethod
    def get_quarter(d: datetime.datetime) -> int:
        return (d.month - 1) // 3 + 1

    @staticmethod
    def get_year_half(d: datetime.datetime) -> int:
        return (d.month - 1) // 6 + 1

    def __post_init__(self) -> None:
        if type(self.account) is int:
            self.account: str = str(self.account)

        if self.date_to is not None:
            date_to: datetime.datetime = datetime.datetime.strptime(self.date_to, '%Y-%m-%d')
            date_from: datetime.datetime = datetime.datetime.strptime(self.date_from, '%Y-%m-%d')
            self.date_diff: int = (date_to - date_from).days

        if not self.action:
            return

        self.mode: str = 'MCLIEN' if self.action in ['S_CLI_003', 'S_CLI_013'] else 'DD7'

        russian_months: Dict = {1: 'январь', 2: 'февраль', 3: 'март', 4: 'апрель', 5: 'май', 6: 'июнь', 7: 'июль', 8: 'август', 9: 'сентябрь', 10: 'октябрь', 11: 'ноябрь', 12: 'декабрь'}
        _date: datetime.datetime = datetime.datetime.strptime(self.date_to, '%Y-%m-%d')
        date_end: datetime.datetime = _date
        date_start: datetime.datetime = datetime.datetime.strptime(self.date_from, '%Y-%m-%d')
        date_end_str: str = date_end.strftime('%d.%m.%Y')

        month_year: str = f'{russian_months[_date.month]} {_date.year}'
        folder_names: Dict = {
            1: rf'За {month_year}',
            3: rf'За {self.get_quarter(date_start)} квартал {date_start.year}',
            6: rf'За {self.get_year_half(date_start)} полугодие {date_start.year}',
            9: rf'За 9 месяцев {date_start.year}',
            12: rf'За {date_start.year} год'
        }
        folder_name: str = folder_names[self.diff_month(start=date_start, end=date_end)]

        if self.action == 'Z_160_GL_020':
            self.save_path = rf'C:\xls\z_160_gl_020'
            self.save_path += rf'\{_date.year} год\{month_year}'

            if date_start == date_end:
                self.save_path += rf'\{date_end_str}'
            else:
                self.save_path += rf'\{folder_name}'

            if self.branch == '00':
                if date_start.day != date_end.day:
                    self.file_name = rf'{self.account_name}_{folder_name.lower()}.xls'
                else:
                    self.file_name = rf'{self.account_name}_{date_end_str}.xls'
            else:
                self.save_path += rf'\{self.account}'
                self.file_name = f'{self.account}_{self.branch}.xls'
        elif self.action == 'Z_160_GL_003':
            self.save_path = rf'C:\xls\z_160_gl_003'
            self.save_path += rf'\{date_start.year} год\{month_year}\Баланс ГК и обороты'

            if date_start.day != date_end.day:
                self.file_name = rf'{folder_name}.xls'
            else:
                self.file_name = rf'{date_end_str}.xls'
        elif self.action == 'S_CLI_003':
            self.save_path = rf'C:\xls\s_cli_003\{date_start.year}'
            if self.branch == '00':
                self.file_name = f'Ведомость коррекции Книги регистрации клиентов_' \
                        f'{russian_months[date_start.month]} {date_start.year}.xls'
            else:
                self.file_name = f'{self.branch}_{russian_months[date_start.month]} {date_start.year}.xls'
                self.save_path += rf'\{russian_months[date_start.month].capitalize()}'
        elif self.action == 'S_CLI_004':
            self.save_path = fr'C:\xls\s_cli_004'
            self.save_path += rf'\{date_start.year}\{russian_months[date_start.month].capitalize()} {date_start.year}'
            self.file_name = f'{self.branch}_{russian_months[date_start.month]} {date_start.year}.xls'
        elif self.action == 'S_CLI_013':
            self.save_path = rf'C:\xls\s_cli_013\{date_start.year}'
            self.save_path += rf'\Книга регистрации клиентов за {self.get_quarter(date_start)} квартал {date_start.year}'
            self.file_name = f'{self.branch}_{self.get_quarter(date_start)} кв {date_start.year}.xls'
        elif self.action == 'S_CLI_014':
            self.save_path = rf'C:\xls\s_cli_014\{date_start.year}'
            self.save_path += rf'\Книга регистрации счетов за {self.get_quarter(date_start)} квартал {date_start.year}'
            self.file_name = f'{self.branch}_{self.get_quarter(date_start)} кв {date_start.year}.xls'

        self.save_path = self.save_path.replace('31.12.2022', '30.12.2022')
        self.file_name = self.file_name.replace('31.12.2022', '30.12.2022')

        self.final_name = self.file_name.replace('xls', 'xlsb')
        self.final_save_path = self.get_final_save_path(save_path=self.save_path, action=self.action)

    @staticmethod
    def get_final_save_path(save_path, action):
        save_paths = {
            'Z_160_GL_020': r'finished_reports\ДБУ_Для Казначейства',
            'Z_160_GL_003': r'finished_reports\ДБУ_Для Казначейства',
            'S_CLI_003': r'finished_reports\УВГК_Книга регистрации\Ведомости\ведомости коррекции карточек клиентов',
            'S_CLI_004': r'finished_reports\УВГК_Книга регистрации\Ведомости\ведомости открытия, закрытия и кор-ии лицевых счетов',
            'S_CLI_013': r'finished_reports\УВГК_Книга регистрации\Книга регистрации\Книга регистрации клиентов',
            'S_CLI_014': r'finished_reports\УВГК_Книга регистрации\Книга регистрации\Книга регистрации лицевых счетов'
        }

        return save_path.replace(rf'xls\{action.lower()}', save_paths[action])
