# config.py
import os
from dotenv import load_dotenv
from pathlib import Path

# Загружаем .env файл
env_path = Path('.') / '.env'
load_dotenv(dotenv_path=env_path)

class Config:
    """Класс для хранения конфигурации"""
    
    # KDV Аккаунт
    KDV_EMAIL = os.getenv('KDV_EMAIL')
    KDV_PASSWORD = os.getenv('KDV_PASSWORD')
    
    # Пути к файлам
    EXCEL_INPUT_PATH = os.getenv('EXCEL_INPUT_PATH', 'C:/TEMP/read.xlsx')
    EXCEL_OUTPUT_PATH = os.getenv('EXCEL_OUTPUT_PATH', 'C:/TEMP/read_new.xlsx')
    
    # Настройки драйвера
    HEADLESS_MODE = os.getenv('HEADLESS_MODE', 'False').lower() == 'true'
    MAX_PRODUCTS_TO_PROCESS = int(os.getenv('MAX_PRODUCTS_TO_PROCESS', 5))
    
    # Браузер
    USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36"
    
    
    @classmethod
    def validate(cls):
        """Проверяет наличие обязательных переменных"""
        required = ['KDV_EMAIL', 'KDV_PASSWORD']
        missing = [var for var in required if not getattr(cls, var)]
        
        if missing:
            raise ValueError(f"Отсутствуют обязательные переменные в .env: {', '.join(missing)}")
        
        return True

# Создаем глобальный объект конфигурации
config = Config()