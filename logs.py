import logging

logger = logging.getLogger()
logger.setLevel(logging.INFO)
log_info = "relatório.log"
handler_info = logging.FileHandler(log_info, mode="a", encoding="utf-8")
handler_info.setLevel(logging.INFO)
formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s", datefmt="%Y/%m/%d %H:%M:%S")
handler_info.setFormatter(formatter)
logger.addHandler(handler_info)