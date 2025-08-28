from pathlib import Path

import logging.handlers
import os
from dotenv import load_dotenv

load_dotenv()


class LogGenerator:
    """Classe responsável por gerar um logger para registro de logs em arquivo."""

    def __init__(self, log_folder="./data/logs", log_file_name="file_log.log"):
        """
        Inicializa uma instância do LogGenerator.

        Args:
            log_folder (str, optional): O diretório onde os arquivos de log serão armazenados.
                O padrão é './logs'.
            log_file_name (str, optional): O nome do arquivo de log.
                O padrão é 'file_log.log'.
        """
        self.log_folder = log_folder
        self.log_file_name = log_file_name
        self.logger = None

    def setup_logger(self) -> logging.Logger:
        """
        Configura o logger para armazenar logs em um arquivo rotativo diariamente.

        O logger será configurado com um manipulador TimedRotatingFileHandler, que permite a rotação
        diária do arquivo de log para manter um histórico de logs.

        Returns:
            logging.Logger: O objeto logger para registrar logs em um arquivo.

        """
        log_exec_dir = Path(os.getenv("LOGS_FOLDER", self.log_folder))
        log_exec_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_exec_dir.joinpath(self.log_file_name)

        handler = logging.handlers.TimedRotatingFileHandler(
            log_file, when="D", backupCount=30, utc=True
        )

        handler.setFormatter(
            logging.Formatter(
                "%(asctime)s [%(levelname)s] %(name)s - %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
            )
        )

        self.logger = logging.getLogger("logger")
        self.logger.setLevel(logging.INFO)
        self.logger.addHandler(handler)

        return self.logger