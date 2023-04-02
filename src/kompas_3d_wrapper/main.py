"""Main module."""
import logging
import shlex
import subprocess  # noqa: S404
import sys
import winreg
from os import path
from types import ModuleType

import psutil
import pythoncom
from win32com.client import Dispatch
from win32com.client import gencache
from pythoncom import com_error
from time import sleep


class k3d:
    """Класс для работы с КОМПАС-3D."""

    def __init__(self):
        """Инициализация класса."""
        self.start_kompas()

        sleep(5)  # Wait for Kompas to start

        self.constants = self.get_kompas_constants()
        self.module7, self.api7 = self.get_kompas_api7()
        self.application = self.api7.Application
        self.documents = self.application.Documents

    def open_document(self, document_path: str = None, create_new: bool = False):
        """Открывает документ КОМПАС-3D."""
        if document_path is not None:
            self.documents = self.application.Documents
            self.kompas_document = self.documents.Open(document_path, True, False)
        elif create_new:
            kompas_document = self.documents.Add(self.constants.ksDocumentDrawing, True)
        else:
            try:
                kompas_document = self.application.ActiveDocument
            except com_error as e:
                logging.exception(f"Не удалось получить активный документ: {e!r}")
                raise e

        kompas_document.Active = True
        self.document_2d = self.module7.IKompasDocument2D(kompas_document)

    def find_exe_by_file_extension(self, file_extension: str) -> str:
        """Находит исполняемый файл по расширению файла."""
        try:
            with winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT, file_extension, 0, winreg.KEY_READ
            ) as key:
                class_name, key_type = winreg.QueryValueEx(key, "")

                if key_type != winreg.REG_SZ:
                    raise FileNotFoundError(f"Нет ключа {file_extension} в реестре")

            with winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT,
                rf"{class_name}\shell\open\command",
                0,
                winreg.KEY_READ,
            ) as key:
                command_val, _ = winreg.QueryValueEx(key, "")
                exe_path = shlex.split(command_val)[0]
                return exe_path.strip('"')
        except FileNotFoundError as file_not_found_err:
            raise FileNotFoundError(
                f"Не удалось найти исполняемый файл по расширению файла {file_extension!r}"
            ) from file_not_found_err

    def get_pythonwin_path(self) -> str:
        """Возвращает путь к папке pythonwin КОМПАС-3D."""
        file_ext = ".cdm"
        py_scripter_path = self.find_exe_by_file_extension(file_ext)

        if "Python 3" not in py_scripter_path:
            raise FileNotFoundError("Библиотеки Python КОМПАС-3D не найдены")

        pythonwin_path = path.join(
            py_scripter_path.replace(path.basename(py_scripter_path), ""),
            "App\\Lib\\site-packages\\pythonwin",
        )

        if not path.exists(pythonwin_path):
            raise FileNotFoundError("папка pythonwin КОМПАС-3D не найдена")

        logging.debug(f"Pythonwin: {pythonwin_path}")

        return pythonwin_path

    def get_kompas_path(self) -> str:
        """Возвращает путь к исполняемому файлу КОМПАС-3D."""
        kompas_path = self.find_exe_by_file_extension(".cdw")

        if "KOMPAS-3D" not in kompas_path:
            raise FileNotFoundError("КОМПАС-3D с поддержкой макросов не установлен")

        logging.debug(f"Компас: {kompas_path}")

        return kompas_path

    def import_ldefin2d_mischelpers(
        self, kompas_pythonwin: str = None
    ) -> tuple[ModuleType, ModuleType]:
        """Импортирует модули LDefin2D и MiscellaneousHelpers из папки pythonwin КОМПАС-3D."""
        if kompas_pythonwin is None:
            kompas_pythonwin = self.get_pythonwin_path()
        # Проверяем, существует ли папка pythonwin
        if not path.exists(kompas_pythonwin):
            raise FileNotFoundError(
                f"Kompas pythonwin not found at {kompas_pythonwin!r}. "
                + "Please install Kompas with macro support."
            )

        sys.path.append(kompas_pythonwin)

        # Удаляем модули LDefin2D и MiscellaneousHelpers из sys.modules, если они там есть
        sys.modules.pop("MiscellaneousHelpers", None)
        sys.modules.pop("LDefin2D", None)

        try:
            # Импортируем модули LDefin2D и MiscellaneousHelpers как miscHelpers
            import LDefin2D  # pyright: reportMissingImports=false
            import MiscellaneousHelpers as miscHelpers
        except ImportError as import_err:
            # Если не удалось импортировать модули, то возвращаем исключение
            raise ImportError(f"Failed importing {import_err.name}") from ImportError
        finally:
            # Убеждаемся, что папка pythonwin удалена из sys.path
            sys.path.remove(kompas_pythonwin)

        return LDefin2D, miscHelpers

    def is_process_running(self, process_name: str) -> bool:
        """Проверяет, запущен ли процесс с именем process_name."""
        for proc in psutil.process_iter():
            try:
                if proc.name() == process_name:
                    logging.debug(f"Process {process_name} is running")
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        logging.debug(f"Process {process_name} is not running")
        return False

    def start_kompas(self, kompas_exe_path: str = None) -> bool:
        """Запускает КОМПАС-3D, если он ещё не запущен."""
        if kompas_exe_path is None:
            kompas_exe_path = self.get_kompas_path()
        if self.is_process_running(path.basename(kompas_exe_path)):
            return True

        try:
            subprocess.Popen(kompas_exe_path)  # noqa: S603
            logging.debug(f"Запустил компас из {kompas_exe_path!r}")
            return False
        except Exception as e:
            raise Exception(f"Ошибка запуска КОМПАС-3D: {e!r}") from e

    def get_kompas_constants(self) -> ModuleType:
        """Возвращает модуль constants КОМПАС-3D."""
        try:
            constants = gencache.EnsureModule(
                "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0
            ).constants
            return (
                constants  # Я не знаю как, но оно работает, даже если компас не запущен
            )
        except Exception as e:
            raise Exception("Не удалось получить константы КОМПАС-3D: " + str(e)) from e

    def get_kompas_constants_3d(self) -> ModuleType:
        """Возвращает модуль constants_3d КОМПАС-3D."""
        try:
            constants = gencache.EnsureModule(
                "{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0
            ).constants
            return constants
        except Exception as e:
            raise Exception(f"Не удалось получить константы 3d КОМПАС-3D: {e!r}") from e

    def get_kompas_api7(self) -> tuple[any, type]:
        """Получает COM API КОМПАС-3D версии 7."""
        try:
            module = gencache.EnsureModule(
                "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0
            )
            api = module.IKompasAPIObject(
                Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(
                    module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch
                )
            )
            return module, api
        except Exception as e:
            raise Exception(f"Failed to get Kompas COM API version 7: {e!r}") from e

    def get_kompas_api5(self) -> tuple[any, type]:
        """Получает COM API КОМПАС-3D версии 5."""
        try:
            module = gencache.EnsureModule(
                "{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0
            )
            api = module.KompasObject(
                Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(
                    module.KompasObject.CLSID, pythoncom.IID_IDispatch
                )
            )
            return module, api
        except Exception as e:
            raise Exception(
                "Failed to get Kompas COM API version 5: " + str(e)
            ) from Exception


if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)

    kompas = k3d()
