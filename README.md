# Sber-to-MicroBill

Парсер уведомлений СберБизнес о поступлениях → TSV-реестр для импорта в MicroBill.

## Возможности
- Подключение к почте по IMAP
- Поиск писем об операциях по счёту
- Парсинг HTML-писем (quoted-printable)
- Извлечение: дата, сумма, плательщик, ИНН, лицевой счёт, договор, УПД
- Пропуск списаний и эквайринга
- Вывод TSV-реестра

## Установка
```bash
git clone ...
cd sber-to-mikrobill
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
# отредактировать .env
