#!/usr/bin/env python3
"""
Парсер писем СберБизнес (HTML/quoted-printable) + выписка 1CClientBankExchange
-> TSV реестр для MicroBill.

Исправлено: конфликт имён переменных с модулем email.
"""

import re
import csv
import sys
import imaplib
import logging
import html as html_mod
from pathlib import Path
from datetime import datetime, timedelta
from email.header import decode_header
from email.message import Message  # Явный импорт класса
from typing import Optional, List, Dict
from dotenv import load_dotenv
import os
load_dotenv(dotenv_path=Path(__file__).parent / ".env")

# ======================== НАСТРОЙКИ (из .env) ========================
IMAP_SERVER = os.getenv("IMAP_SERVER", "imap.gmail.com")
IMAP_PORT = int(os.getenv("IMAP_PORT", "993"))
EMAIL_LOGIN = os.getenv("EMAIL_LOGIN", "")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "")
DELETE_AFTER_PROCESS = os.getenv("DELETE_AFTER_PROCESS", "False").lower() == "true"
MAILBOX = os.getenv("MAILBOX", "INBOX")
DOWNLOAD_DIR = Path(os.getenv("DOWNLOAD_DIR", "./sber_attachments"))
OUTPUT_TSV = os.getenv("OUTPUT_TSV", "microbill_reestr.tsv")

# ======================== ЛОГИРОВАНИЕ ========================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
log = logging.getLogger(__name__)

DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

# ================================================================
#  1. РАБОТА С ПОЧТОЙ (IMAP)
# ================================================================

def connect_to_mailbox() -> imaplib.IMAP4_SSL:
    """Подключается к почтовому ящику по IMAP."""
    log.info(f"Подключение к {IMAP_SERVER}:{IMAP_PORT}...")
    mail_conn = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail_conn.login(EMAIL_LOGIN, EMAIL_PASSWORD)
    log.info("Успешный вход в почтовый ящик.")
    return mail_conn


def search_sber_emails(mail_conn: imaplib.IMAP4_SSL, days_back: int = 7) -> List[bytes]:
    """Ищет письма-уведомления о ПОСТУПЛЕНИЯХ от СберБизнес."""
    mail_conn.select(MAILBOX)

    since_date = (datetime.now() - timedelta(days=days_back)).strftime("%d-%b-%Y")
    search_criteria = f'(SINCE "{since_date}")'

    log.info(f"Поиск писем с {since_date}...")
    status, messages = mail_conn.search(None, search_criteria)

    if status != "OK":
        log.error("Ошибка поиска писем")
        return []

    all_ids = messages[0].split()
    log.info(f"Всего писем за период: {len(all_ids)}")

    sber_ids = []

    for msg_id in all_ids:
        status, msg_data = mail_conn.fetch(msg_id, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT)])")
        if status != "OK":
            continue

        header_text = msg_data[0][1].decode('utf-8', errors='ignore')

        # Пропускаем списания сразу
        if 'списан' in header_text.lower():
            continue

        # Декодируем тему
        subject = ""
        from email.header import decode_header as dh
        import re as re_mod
        m = re_mod.search(r'Subject:\s*(.+)', header_text, re_mod.IGNORECASE | re_mod.DOTALL)
        if m:
            raw_subject = m.group(1).replace('\r\n', '').replace('\n', '').strip()
            try:
                parts = dh(raw_subject)
                subject = ''.join(
                    p.decode(c or 'utf-8') if isinstance(p, bytes) else str(p)
                    for p, c in parts
                )
            except:
                subject = raw_subject

        log.info(f"  Письмо #{msg_id.decode()}: тема='{subject[:80]}'")

        is_sber = 'sberbank' in header_text.lower()
        is_payment = (
            'операци' in subject.lower() or
            'поступили' in subject.lower()
        )

        if is_sber and is_payment:
            log.info(f"    -> ПОДХОДИТ")
            sber_ids.append(msg_id)

    log.info(f"Писем от Сбера о поступлениях: {len(sber_ids)}")
    return sber_ids

################

def fetch_email_message(mail_conn: imaplib.IMAP4_SSL, msg_id: bytes) -> Optional[Message]:
    """Загружает полное письмо по ID. Возвращает объект email.message.Message."""
    status, msg_data = mail_conn.fetch(msg_id, "(RFC822)")
    if status != "OK":
        log.error(f"Не удалось загрузить письмо #{msg_id.decode()}")
        return None
    
    # Используем импортированный класс Message
    return Message()


def fetch_email_bytes(mail_conn: imaplib.IMAP4_SSL, msg_id: bytes) -> Optional[bytes]:
    """Загружает полное письмо по ID. Возвращает сырые байты письма."""
    status, msg_data = mail_conn.fetch(msg_id, "(RFC822)")
    if status != "OK":
        log.error(f"Не удалось загрузить письмо #{msg_id.decode()}")
        return None
    return msg_data[0][1]


# ================================================================
#  2. ПАРСИНГ HTML-ПИСЬМА СБЕРА
# ================================================================

def decode_quoted_printable_html(part) -> str:
    """Декодирует quoted-printable HTML в обычный текст."""
    raw_bytes = part.get_payload(decode=True)
    charset = part.get_content_charset() or 'utf-8'
    return raw_bytes.decode(charset, errors='ignore')


def extract_text_from_html(html_content: str) -> str:
    """Убирает HTML-теги, оставляя текст."""
    text = html_content.replace('=\n', '')
    text = html_mod.unescape(text)
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def get_email_body_parts(email_msg: Message) -> List[str]:
    """Собирает все текстовые части письма."""
    parts = []
    if email_msg.is_multipart():
        for part in email_msg.walk():
            content_type = part.get_content_type()
            disposition = str(part.get("Content-Disposition", ""))
            if "attachment" in disposition:
                continue

            if content_type == "text/plain":
                charset = part.get_content_charset() or 'utf-8'
                text = part.get_payload(decode=True).decode(charset, errors='ignore')
                parts.append(text)
            elif content_type == "text/html":
                parts.append(decode_quoted_printable_html(part))
    else:
        parts.append(decode_quoted_printable_html(email_msg))
    return parts


def parse_sber_notification_from_html(html_content: str) -> Optional[Dict]:
    """Парсит HTML-уведомление Сбера о поступлении средств."""
    text = extract_text_from_html(html_content)
    result = {}

    log.info("--- Текст письма (начало) ---")
    log.info(text[:500])
    log.info("-----------------------------")

    # Сумма: "+700,00 RUB"
    sum_match = re.search(r'\+(\d[\d\s]*[.,]\d{2})\s*RUB', text)
    if sum_match:
        amount_str = sum_match.group(1).replace(' ', '').replace(',', '.')
        result['Сумма'] = float(amount_str)
        log.info(f"Сумма: {result['Сумма']} RUB")

    # Дата/время
    dt_match = re.search(r'на\s+(\d{2}\.\d{2}\.\d{4})\s+(\d{2}:\d{2}:\d{2})', text)
    if dt_match:
        result['ДатаПоступило'] = dt_match.group(1)
        result['ВремяПоступило'] = dt_match.group(2)
        log.info(f"Дата/время: {result['ДатаПоступило']} {result['ВремяПоступило']}")

    # Плательщик и реквизиты
    payer_match = re.search(
        r'Кто\s+отправитель\??\s*(.+?),\s*ИНН\s*\*(\d+).*?р/с\s*\*+(\d+).*?по\s+документу\s*№\s*(\d+)\s*от\s*(\d{2}\.\d{2}\.\d{4})',
        text,
        re.IGNORECASE | re.DOTALL
    )
    if payer_match:
        result['Плательщик'] = payer_match.group(1).strip()
        result['ИНН_маска'] = payer_match.group(2).strip()
        result['Счет_маска'] = payer_match.group(3).strip()
        result['НомерДокумента'] = payer_match.group(4).strip()
        result['ДатаДокумента'] = payer_match.group(5).strip()
        log.info(f"Плательщик: {result['Плательщик']}, ИНН *{result['ИНН_маска']}")

    # Полный ИНН
    inn_full = re.search(r'ИНН\s*(\d{10,12})', text)
    if inn_full:
        result['ИНН'] = inn_full.group(1)
    else:
        result['ИНН'] = result.get('ИНН_маска', '')

    # Назначение платежа
    purpose_match = re.search(
        r'Назначение\s+платежа:\s*(.+?)(?:\s*Пожалуйста, не отвечайте|\s*Скачать Приложение|\s*©)',
        text,
        re.DOTALL
    )
    if purpose_match:
        purpose = purpose_match.group(1).strip()
        purpose = re.sub(r'\s+', ' ', purpose)
        result['НазначениеПлатежа'] = purpose
        log.info(f"Назначение: {purpose[:120]}...")

    # Лицевой счёт
    if 'НазначениеПлатежа' in result:
        ls = extract_ls(result['НазначениеПлатежа'])
        if ls:
            result['ЛицевойСчет'] = ls
            log.info(f"Лицевой счёт: {ls}")

    # Баланс
    balance_match = re.search(
        r'Баланс\s+сч[её]та[^:]*:\s*([\d\s]+[.,]\d{2})\s*RUB',
        text
    )
    if balance_match:
        balance_str = balance_match.group(1).replace(' ', '').replace(',', '.')
        result['Баланс'] = float(balance_str)
        log.info(f"Баланс: {result['Баланс']} RUB")

    return result if result.get('Сумма') else None


def extract_ls(purpose: str) -> str:
    """Извлекает лицевой счёт из назначения платежа."""
    patterns = [
        r'(?:л/с|ЛС|лицевой\s*сч[её]т)\s*[:№]?\s*(\d[\d\-]*)',
        r'Договор\s+\d+[\s/]*ЛС\s*(\d[\d\-]*)',
    ]
    for pattern in patterns:
        match = re.search(pattern, purpose, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return ""


# ================================================================
#  3. ПАРСИНГ ВЛОЖЕНИЙ (1CClientBankExchange)
# ================================================================

def save_attachments(email_msg: Message, msg_id: str) -> List[Path]:
    """Сохраняет вложения письма на диск."""
    saved = []
    if not email_msg.is_multipart():
        return saved
    
    for part in email_msg.walk():
        disposition = str(part.get("Content-Disposition", ""))
        if "attachment" not in disposition:
            continue
        
        filename = part.get_filename()
        if not filename:
            continue
        
        # Декодируем имя файла
        decoded_parts = decode_header(filename)
        filename = ''
        for part_bytes, charset in decoded_parts:
            if isinstance(part_bytes, bytes):
                filename += part_bytes.decode(charset or 'utf-8', errors='ignore')
            else:
                filename += str(part_bytes)
        
        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        filepath = DOWNLOAD_DIR / f"{msg_id}_{filename}"
        
        with open(filepath, 'wb') as f:
            f.write(part.get_payload(decode=True))
        
        log.info(f"Сохранено вложение: {filepath}")
        saved.append(filepath)
    
    return saved


def parse_1c_bank_statement(filepath: str) -> List[Dict]:
    """Разбирает файл выписки 1CClientBankExchange."""
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    doc_pattern = re.compile(r'СекцияДокумент=.*?\n(.*?)КонецДокумента', re.DOTALL)
    operations = []

    for match in doc_pattern.finditer(content):
        block = match.group(1)
        op = {}
        for line in block.strip().split('\n'):
            if '=' in line:
                key, _, value = line.partition('=')
                op[key.strip()] = value.strip()

        if not op.get('Сумма') or not op.get('ДатаПоступило'):
            continue

        purpose = op.get('НазначениеПлатежа', '')
        ls = extract_ls(purpose)
        if ls:
            op['ЛицевойСчет'] = ls
        operations.append(op)

    return operations


def calculate_commission(purpose: str, amount: float) -> tuple:
    """Вычисляет комиссию из назначения платежа."""
    if not purpose:
        return amount, 0.0
    
    # Эквайринг
    summa = re.search(r'СУММА\s*([\d.]+)', purpose, re.IGNORECASE)
    usl = re.search(r'УСЛ\.БАНКА:?\s*([\d.]+)', purpose, re.IGNORECASE)
    if summa and usl:
        return float(summa.group(1)), float(usl.group(1))
    
    # Комиссия в тексте
    comm = re.search(r'(?:комис\.?|усл\.?банка:?)\s*([\d.]+)', purpose, re.IGNORECASE)
    if comm:
        c = float(comm.group(1))
        return amount + c, c
    
    return amount, 0.0


# ================================================================
#  4. ФОРМИРОВАНИЕ TSV ДЛЯ MICROBILL
# ================================================================
def generate_microbill_tsv(records: List[Dict], output_path: str) -> tuple[int, float]:
    """Формирует TSV-файл для импорта в MicroBill."""
    headers = [
        "Дата", "Сумма", "Комиссия", "Плательщик", "ИНН",
        "ЛицевойСчет", "Договор", "СчетУПД", "НазначениеПлатежа",
        "НомерПП", "Примечание"
    ]

    rows = 0
    with open(output_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f, delimiter='\t')
        writer.writerow(headers)
        total_sum = 0.0

        for rec in records:
            date_str = rec.get('ДатаПоступило', rec.get('ДатаДокумента', ''))
            amount = float(rec.get('Сумма', 0))
            purpose = rec.get('НазначениеПлатежа', '')
            payer = rec.get('Плательщик', '')
            inn = rec.get('ИНН', rec.get('ПлательщикИНН', ''))
            doc_num = rec.get('Номер', rec.get('НомерДокумента', ''))
            ls = rec.get('ЛицевойСчет', '')

            try:
                dt = datetime.strptime(date_str, '%d.%m.%Y')
                date_fmt = dt.strftime('%d.%m.%Y')
            except ValueError:
                date_fmt = date_str

            gross, commission = calculate_commission(purpose, amount)

            contract = ""
            ctr = re.search(r'(?:договор|дог\.)\s*(?:от|№)?\s*([\d\w\-/.]+)', purpose, re.IGNORECASE)
            if ctr:
                contract = ctr.group(1)

            invoice = ""
            inv = re.search(r'(?:УПД|сч[её]т|унив\.перед\.док\.)\s*[:№]?\s*(\d[\d\w\-/]*)', purpose, re.IGNORECASE)
            if inv:
                invoice = inv.group(1)

            note_parts = []
            if commission > 0:
                note_parts.append(f"Комиссия: {commission:.2f} руб.")
            if 'Баланс' in rec:
                note_parts.append(f"Баланс: {rec['Баланс']:.2f}")

            # Пропускаем сводные платежи эквайринга Сбера
            if 'ЗА Интернет;ПО ПЛАТЕЖАМ' in purpose and 'ЭЛ.РЕЕСТР' in purpose:
                continue
            # Пропускаем поступления от агрегаторов (Робокасса)
            if 'РОБОКАССА' in purpose.upper() or 'ROBOKASSA' in purpose.upper():
                continue

            writer.writerow([
                date_fmt,
                f"{gross:.2f}",
                f"{commission:.2f}",
                payer,
                inn,
                ls,
                contract,
                invoice,
                purpose,
                doc_num,
                "; ".join(note_parts)
            ])
            rows += 1
            total_sum += gross
    return rows, total_sum


# ================================================================
#  5. ГЛАВНАЯ ЛОГИКА
# ================================================================

def process_emails_and_generate_report(days_back: int = 7) -> List[Dict]:
    """Основной рабочий процесс."""
    all_records = []
    mail_conn = connect_to_mailbox()
    sber_ids = search_sber_emails(mail_conn, days_back)

    if not sber_ids:
        log.warning("Писем от Сбера о поступлениях не найдено.")
        mail_conn.logout()
        return all_records

    for msg_id in sber_ids:
        msg_id_str = msg_id.decode()
        log.info(f"\n{'='*60}")
        log.info(f"Обработка письма #{msg_id_str}")

        # Получаем сырые байты письма
        raw_email_bytes = fetch_email_bytes(mail_conn, msg_id)
        if not raw_email_bytes:
            continue

        # Парсим письмо из байтов
        import email.parser
        email_msg = email.parser.BytesParser().parsebytes(raw_email_bytes)

        # Получаем HTML-часть
        parts = get_email_body_parts(email_msg)
        html_body = next((p for p in parts if '<html' in p.lower()), parts[0] if parts else "")

        notification_data = parse_sber_notification_from_html(html_body)

        if notification_data:
            all_records.append(notification_data)
        else:
            log.warning("Не удалось извлечь данные из тела письма")

        # Вложения
        attachments = save_attachments(email_msg, msg_id_str)
        for fp in attachments:
            if fp.suffix.lower() in ['.txt']:
                ops = parse_1c_bank_statement(str(fp))
                log.info(f"Операций из вложения: {len(ops)}")
                all_records.extend(ops)

    mail_conn.logout()
    # Удаление обработанных писем (если включено)
    if DELETE_AFTER_PROCESS and sber_ids:
        mail_conn.select(MAILBOX)
        for msg_id in sber_ids:
            mail_conn.store(msg_id, '+FLAGS', '\\Deleted')
            log.info(f"Письмо #{msg_id.decode()} помечено на удаление")
        mail_conn.expunge()
        log.info("Обработанные письма удалены")
    # Удаляем дубли
    unique = []
    seen = set()
    for r in all_records:
        key = (r.get('Номер', r.get('НомерДокумента', '')), r.get('Сумма', 0))
        if key not in seen:
            seen.add(key)
            unique.append(r)

    log.info(f"Уникальных записей: {len(unique)}")
    return unique


def main():
    """Главная функция."""
    print("=" * 60)
    print("  Парсер уведомлений СберБизнес -> MicroBill")
    print("=" * 60)

    print(f"IMAP: {IMAP_SERVER}:{IMAP_PORT}")
    print(f"Логин: {EMAIL_LOGIN}")
    print(f"Пароль: {'***' if EMAIL_PASSWORD else 'НЕ ЗАДАН'}")
    print(f"Ящик: {MAILBOX}")
    print(f"Удалять письма: {DELETE_AFTER_PROCESS}")
    print(f"Папка вложений: {DOWNLOAD_DIR}")
    print(f"Выходной файл: {OUTPUT_TSV}")
    print("=" * 60)
    # Офлайн-режим
    if len(sys.argv) >= 2 and Path(sys.argv[1]).exists():
        input_file = sys.argv[1]
        print(f"Офлайн-режим: обработка файла '{input_file}'")
        records = parse_1c_bank_statement(input_file)
        out = sys.argv[2] if len(sys.argv) >= 3 else OUTPUT_TSV
        rows, total_exporeted = generate_microbill_tsv(records, out)
        print(f"Реестр сохранён: {out}")
        print(f"Строк данных: {rows}")
        total = sum(r.get('Сумма', 0) for r in records)
        print(f"Общая сумма: {total_exported:,.2f} руб.")
        return

    # Онлайн-режим
    try:
        records = process_emails_and_generate_report(days_back=7)
        if not records:
            print("Нет записей для экспорта.")
            return
        rows, total_exported = generate_microbill_tsv(records, OUTPUT_TSV)
        print(f"\nРеестр сохранён: {OUTPUT_TSV}")
        print(f"Строк данных: {rows}")
        total = sum(float(r.get('Сумма', 0)) for r in records)
        print(f"Общая сумма: {total_exported:,.2f} руб.")
    except Exception as e:
        log.error(f"Ошибка: {e}", exc_info=True)
        sys.exit(1)


if __name__ == '__main__':
    main()
