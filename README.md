# Outlook

Веб-панель для управления Outlook/Hotmail аккаунтами с поддержкой OAuth2 (Graph API), проверки валидности токенов, чтения входящих/спама и массовой рассылки.

## Стек технологий

- **Backend:** Python 3.10+, FastAPI, HTTPX, SQLite3
- **Frontend:** Vanilla HTML5, CSS3, JavaScript
- **API:** Microsoft Graph API / Outlook REST API

---

## Возможности

- 📥 **Массовый импорт:** умный парсинг строк любого формата (`email:password:client_id:refresh_token`, разделители `:`, `|`, `;` или tab).
- 🏷️ **Группировка:** распределение аккаунтов по категориям/папкам и быстрая фильтрация.
- 📬 **Чтение почты:** просмотр папок «Входящие» и «Нежелательная почта» прямо из интерфейса.
- ✉️ **Массовая рассылка:** отправка писем с выбранных аккаунтов в один клик.
- ⚡ **Быстрый чекер:** проверка валидности refresh-токенов при загрузке.
- 🌗 **Темная/светлая тема** и сохранение настроек интерфейса в LocalStorage.

---

## Установка и запуск

### 1. Клонирование репозитория
```bash
git clone [https://github.com/lemon143267/mail-manager.git](https://github.com/lemon143267/mail-manager.git)
cd mail-manager
2. Установка зависимостей
Рекомендуется использовать виртуальное окружение:

Bash
python -m venv venv
# Linux / macOS:
source venv/bin/activate
# Windows:
venv\Scripts\activate

pip install fastapi uvicorn httpx pydantic
3. Переменные окружения (опционально)
По умолчанию панель запускается с базовой HTTP-авторизацией admin:admin. Вы можете переопределить учетные данные и порт:

Bash
# Linux / macOS
export ADMIN_USERNAME="myuser"
export ADMIN_PASSWORD="strongpassword"
export PORT=8000

# Windows (PowerShell)
$env:ADMIN_USERNAME="myuser"
$env:ADMIN_PASSWORD="strongpassword"
4. Запуск приложения
Bash
python main.py
После запуска откройте браузер по адресу: http://localhost:8000

Автор
Разработчик: lemon143267

Профиль LOLZ: lolz.team/members/5957846/
