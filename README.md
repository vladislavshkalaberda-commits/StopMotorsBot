# Motor Search Bot

Телеграм-бот для поиска моторов по маркировке.

## Файлы
- `bot.py` — основной код бота
- `LISTE_100_-_31-12-2025.xlsx` — список моторов
- `Моторы_цены_диапазон.xlsx` — история цен

## Деплой на Railway

1. Создай репозиторий на GitHub и залей все файлы
2. На Railway: New Project → Deploy from GitHub repo
3. В настройках Railway добавь переменную окружения:
   - `BOT_TOKEN` = твой токен от @BotFather
4. Service → Settings → убедись что тип **Worker** (не Web)

## Локальный запуск
```bash
pip install -r requirements.txt
BOT_TOKEN=твой_токен python bot.py
```
