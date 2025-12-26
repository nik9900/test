import sys

from PyQt5.QtWidgets import QApplication

from macros.excel.work_exel import ExcelService
from macros.utils.loging import setup_logging
from macros.ui.ui_interface import App


def main() -> None:
    logger = setup_logging()
    excel_service = ExcelService()

    application = QApplication(sys.argv)
    window = App(excel_service=excel_service, logger=logger)
    window.show()
    sys.exit(application.exec_())


if __name__ == "__main__":
    main()