import logging
import os
from app.utils.paths import get_log_path

_LOGGER_INITIALIZED = False


def setup_logger():
    global _LOGGER_INITIALIZED

    # =========================
    # GUARD: chỉ init 1 lần
    # =========================
    if _LOGGER_INITIALIZED:
        return logging.getLogger()

    log_dir = get_log_path()
    os.makedirs(log_dir, exist_ok=True)

    log_file = os.path.join(log_dir, "app.log")

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    # tránh duplicate handler
    if not logger.handlers:
        formatter = logging.Formatter(
            "%(asctime)s - %(levelname)s - %(name)s - %(message)s"
        )

        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setFormatter(formatter)

        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)

        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

    logger.info(f"Logger initialized: {log_file}")

    _LOGGER_INITIALIZED = True
    return logger