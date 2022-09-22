import logging
from datetime import datetime
from PreShin.config import CONFIG

class logger():
    def logs(self):
        date_s = (datetime.now().strftime('%Y-%m-%d-%H-%M-%S'))  # 파일 명에 사용 하기 위함
        logger = logging.getLogger()

        # logger의 level
        if CONFIG['log_level'] == 'info':
            logger.setLevel(logging.INFO)
        elif CONFIG['log_level'] == 'debug':
            logger.setLevel(logging.DEBUG)
        # formatter 지정
        formatter = logging.Formatter("%(asctime)s - [%(module)s - %(funcName)s] - %(levelname)s - %(message)s")

        # 파일 저장
        # console = logging.StreamHandler()
        file_handler = logging.FileHandler(filename=f"C:\woo_project\AI\log\{date_s}.log")

        # handler 출력 format 지정
        # console.setFormatter(formatter)
        file_handler.setFormatter(formatter)

        # logger에 handler 추가
        # logger.addHandler(console)
        logger.addHandler(file_handler)

        return logger


_logger = logger()
logger = _logger.logs()
