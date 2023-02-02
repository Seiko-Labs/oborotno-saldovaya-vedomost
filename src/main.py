import datetime
import os
import platform
import warnings
from dataclasses import fields
from typing import List
import dotenv
import pandas
import psutil
import requests
from pandas.core.frame import DataFrame
from bot_notification import TelegramNotifier
from data_extractor import DataGetter
from data_structures import Credentials, Process, BranchInfo
from robot import Robot


def sort_list(data: List[BranchInfo]) -> List[BranchInfo]:
    z_160_gl_020 = []
    s_cli_003 = []
    z_160_gl_003 = []
    s_cli_004 = []
    s_cli_013 = []
    s_cli_014 = []

    for i, branch_info in enumerate(data):
        action = branch_info.action
        if action == 'Z_160_GL_020':
            z_160_gl_020.append(branch_info)
        elif action == 'S_CLI_003':
            s_cli_003.append(branch_info)
        elif action == 'Z_160_GL_003':
            z_160_gl_003.append(branch_info)
        elif action == 'S_CLI_004':
            s_cli_004.append(branch_info)
        elif action == 'S_CLI_013':
            s_cli_013.append(branch_info)
        else:
            s_cli_014.append(branch_info)

    z_160_gl_020.sort(key=lambda x: x.date_diff)
    s_cli_003.sort(key=lambda x: x.date_diff)
    z_160_gl_003.sort(key=lambda x: x.date_diff)
    s_cli_004.sort(key=lambda x: x.date_diff)
    s_cli_013.sort(key=lambda x: x.date_diff)
    s_cli_014.sort(key=lambda x: x.date_diff)

    return z_160_gl_020 + s_cli_003 + z_160_gl_003 + s_cli_004 + s_cli_013 + s_cli_014


def get_left_data(data: List[BranchInfo]) -> List[BranchInfo]:
    final_full_paths: List[str] = [os.path.join(x.final_save_path, x.final_name) for x in data]
    finished_full_paths: List[str] = []

    counter: int = 0
    for path, subdirs, names in os.walk(rf'C:\finished_reports'):
        for name in names:
            full_path: str = os.path.join(path, name)
            finished_full_paths.append(full_path)
            for i, compare_path in enumerate(final_full_paths):
                if full_path == compare_path:
                    continue
            else:
                counter += 1
    left_branch_infos: List[BranchInfo] = []
    for i, compare_path in enumerate(final_full_paths):
        if compare_path in finished_full_paths:
            continue
        counter += 1
        left_branch_infos.append(data[i])

    for branch_info in finished_full_paths:
        file_size = round(os.path.getsize(branch_info) / 1024, 2)
        if file_size > 20:
            continue
        dataframe: DataFrame = pandas.read_excel(branch_info, engine='pyxlsb')
        if len(dataframe) != 20:
            continue
        for i, compare_path in enumerate(final_full_paths):
            if compare_path != branch_info:
                continue
            counter += 1
            left_branch_infos.append(data[i])

    return left_branch_infos


def print_table(data: List[BranchInfo]) -> None:
    print('\t'.join([field.name for field in fields(data[0])]))
    for x in data:
        print('\t'.join([str(getattr(x, field.name)) for field in fields(x)]))


def kill_colvirs() -> None:
    for proc in psutil.process_iter():
        if not any(process_name in proc.name() for process_name in ['COLVIR', 'EXCEL']):
            continue
        try:
            p: psutil.Process = psutil.Process(proc.pid)
            p.terminate()
        except psutil.AccessDenied:
            continue


def main() -> None:
    warnings.simplefilter(action='ignore', category=UserWarning)
    dotenv.load_dotenv()

    colvir_usr, colvir_psw = os.getenv(f'COLVIR_USR'), os.getenv(f'COLVIR_PSW')
    process_name, process_path = 'COLVIR', os.getenv('COLVIR_PROCESS_PATH')

    data_getter = DataGetter()
    data: List[BranchInfo] = data_getter.info

    data.sort(key=lambda x: ['Z_160_GL_003', 'Z_160_GL_020', 'S_CLI_003', 'S_CLI_004', 'S_CLI_013', 'S_CLI_014'].index(x.action))
    data = [b for i, b in enumerate(data) if i % 2 == (0 if platform.node() == 'robot-7' else 1)]

    with requests.Session() as session:
        args = {
            'credentials': Credentials(usr=colvir_usr, psw=colvir_psw),
            'process': Process(name=process_name, path=process_path),
            'notifier': TelegramNotifier(chat_id=os.getenv(f'CHAT_ID'), session=session),
            'data': data
        }

        robot: Robot = Robot(**args)
        robot.run()

        full_path_data = [os.path.join(b.final_save_path, b.final_name) for b in data]

        finished_files = []
        for path, subdirs, names in os.walk(
                r'\\dbu-upload\c$\Users\bolatova.g\Desktop\робот январь 2023 02.02.23'):
            for name in names:
                full_path = os.path.join(path, name)
                finished_files.append(full_path)

        _data = [data[i] for i, path in enumerate(full_path_data) if path not in finished_files]

        args['data'] = _data

        if _data:
            robot = Robot(**args)
            robot.run()


if __name__ == '__main__':
    main()
