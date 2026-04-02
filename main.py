# Entry point проекта
# Шаблон: читает .env и запускает "первый минимальный шаг"

from pathlib import Path
from dotenv import load_dotenv

def main() -> None:
    # Загружаем .env из корня проекта
    env_path = Path(__file__).resolve().parent / ".env"
    load_dotenv(dotenv_path=env_path)

    # TODO: здесь добавь реальный код проекта
    # Пример чтения переменной:
    import os
    host = os.getenv("HOST_CLICKHOUSE", "")
    print("HOST_CLICKHOUSE =", host if host else "(not set)")

    print("✅ Project started. Сейчас добавь твою логику в main().")

if __name__ == "__main__":
    main()