#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import json
import requests
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill

LOW_RATING_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
LOW_RATING_THRESHOLD = 3

# ---------- НАСТРОЙКИ ----------
BITRIX_URL = "https://24.uprav.ru/web_services/iblock/export.php"
STATIC_SECRET = "14b0fc053c141e47a5974b3859f5753f"

YANDEX_FILE_PATH = "/Bitrix_отчеты.xlsx"

def load_env_var(key):
    """Читает KEY=... из .env рядом со скриптом."""
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.exists(env_path):
        return None
    with open(env_path, encoding="utf-8") as f:
        for line in f:
            if line.strip().startswith(f"{key}="):
                return line.strip().split("=", 1)[1].strip()
    return None

BITRIX_WEBHOOK = (load_env_var("BITRIX_BASE") or "").rstrip("/") or None

# Маппинг полей для getReviews – DATE_CREATE теперь на первом месте
REVIEWS_FIELDS_MAP = {
    "DATE_CREATE": "Дата создания",
    "CLIENT": "Клиент",
    "EVENT": "Мероприятие",
    "MODULE": "Модуль",
    "DIRECTION": "Направление",
    "REVIEW": "Отзыв",
    "TEACHER": "Преподаватель",
    "RATING": "Рейтинг",
    "CLIENT_TYPE": "B2B/B2C"
}

# Маппинг полей для getAnkets – DATE_CREATE на первом месте
ANKETS_FIELDS_MAP = {
    "DATE_CREATE": "Дата создания",
    "EVENT": "Мероприятие",
    "REVIEW": "Отзыв",
    "EVENT_DATE_START": "Дата начала мероприятия",
    "EVENT_DATE_END": "Дата завершения мероприятия",
    "DIRECTION": "Направление мероприятия",
    "CLIENT": "Клиент",
    "TEACHERS": "Преподаватели",
    "RATING_ORG": "Оценка: Организация обучения",
    "COMMENT_ORG": "Комментарий: Организация обучения",
    "RATING_MATERIALS": "Оценка: Раздаточные материалы",
    "COMMENT_MATERIALS": "Комментарий: Раздаточные материалы",
    "RATING_ROOM": "Оценка: Аудитория и оборудование",
    "COMMENT_ROOM": "Комментарий: Аудитория и оборудование",
    "GOOD_EVENT": "Соответствие ожиданиям",
    "PUBLIC": "Согласие на размещение",
    "CLIENT_TYPE": "B2B/B2C"
}

# Колонки (1-based, в пределах ANKETS_FIELDS_MAP), для которых нужна подсветка низкой оценки
ANKETS_LOW_RATING_COLUMNS = {
    list(ANKETS_FIELDS_MAP).index("TEACHERS") + 1,
    list(ANKETS_FIELDS_MAP).index("RATING_ORG") + 1,
    list(ANKETS_FIELDS_MAP).index("RATING_MATERIALS") + 1,
    list(ANKETS_FIELDS_MAP).index("RATING_ROOM") + 1
}

# ---------- НАСТРОЙКИ TEACHBASE ----------
TEACHBASE_URL = "https://go.teachbase.ru"
TEACHBASE_COURSE_ID = 228785   # курс «Анкета обратной связи»
TEACHBASE_POLL_ID = 33018      # опрос внутри курса (общий для всех потоков)

# Вопросы анкеты Teachbase: question_id -> ключ
TB_QUESTIONS = {
    95420: "REVIEW",          # мнение о программе (свободный текст)
    95425: "RATING_ORG",      # качество организации, 1-5
    95426: "RATING_MATERIALS",# раздаточные материалы, 1-5
    95427: "RATING_PLATFORM", # удобство платформы, 1-5
    118019: "RATING_TEACHERS",# экспертность преподавателей, 1-5
    95422: "GOOD_EVENT",      # соответствие ожиданиям
    95423: "PUBLIC"           # согласие на размещение
}

TB_FIELDS_MAP = {
    "DATE_CREATE": "Дата заполнения",
    "EVENT": "Мероприятие",
    "CLIENT": "Слушатель",
    "EMAIL": "Email",
    "REVIEW": "Мнение о программе",
    "RATING_ORG": "Оценка: Организация",
    "RATING_MATERIALS": "Оценка: Материалы",
    "RATING_PLATFORM": "Оценка: Платформа",
    "RATING_TEACHERS": "Оценка: Преподаватели",
    "GOOD_EVENT": "Соответствие ожиданиям",
    "PUBLIC": "Согласие на размещение"
}

TB_RATING_KEYS = ("RATING_ORG", "RATING_MATERIALS", "RATING_PLATFORM", "RATING_TEACHERS")

# ---------- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ----------
def load_yandex_token():
    token = load_env_var("YANDEX_TOKEN")
    if not token:
        raise Exception("YANDEX_TOKEN не найден в .env")
    return token

def flatten_dict(d, parent=''):
    items = []
    for k, v in d.items():
        key = f"{parent}[{k}]" if parent else k
        if isinstance(v, dict):
            items.extend(flatten_dict(v, key))
        else:
            items.append((key, v))
    return items

def extract_text(value):
    if isinstance(value, dict) and 'TEXT' in value:
        return value['TEXT']
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False)
    return value if value is not None else ""

def is_current_month(date_str):
    if not date_str:
        return False
    try:
        if ' ' in date_str:
            dt = datetime.strptime(date_str, "%d.%m.%Y %H:%M:%S")
        else:
            dt = datetime.strptime(date_str, "%d.%m.%Y")
    except:
        return False
    now = datetime.now()
    return dt.year == now.year and dt.month == now.month

def get_bitrix_data(action, page_size=200):
    data = {
        "SORT": {"DATE_CREATE": "DESC"},
        "FILTER": {},
        "nav": {"nPageSize": str(page_size), "iNumPage": "1"}
    }
    post_fields = [('secret', STATIC_SECRET), ('action', action)] + flatten_dict(data, 'data')
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    try:
        resp = requests.post(BITRIX_URL, data=post_fields, headers=headers, verify=False, timeout=60)
        resp.raise_for_status()
        result = resp.json()
        if result.get("errors"):
            print(f"Ошибка Bitrix ({action}): {result['errors']}")
            return []
        raw_items = result.get("data", [])
        filtered = [item for item in raw_items if is_current_month(item.get("DATE_CREATE", ""))]
        print(f"  Загружено {len(raw_items)}, отфильтровано за текущий месяц: {len(filtered)}")
        return filtered
    except Exception as e:
        print(f"Ошибка при запросе {action}: {e}")
        return []

def fetch_contacts_company_map(contact_ids):
    """Батч-запрос COMPANY_ID контактов через crm.contact.list. Возвращает {contact_id: company_id}."""
    contact_ids = sorted({str(cid) for cid in contact_ids if cid})
    if not contact_ids or not BITRIX_WEBHOOK:
        return {}
    company_by_contact = {}
    try:
        for i in range(0, len(contact_ids), 50):
            chunk = contact_ids[i:i + 50]
            resp = requests.post(
                f"{BITRIX_WEBHOOK}/crm.contact.list.json",
                json={"filter": {"ID": chunk}, "select": ["ID", "COMPANY_ID"]},
                verify=False, timeout=30
            )
            resp.raise_for_status()
            for row in resp.json().get("result", []):
                company_by_contact[row["ID"]] = row.get("COMPANY_ID", "")
    except Exception as e:
        print(f"Ошибка при запросе crm.contact.list: {e}")
    return company_by_contact

def detect_b2b(company_id):
    return "B2B" if company_id and str(company_id) != "0" else "B2C"

# ---------- TEACHBASE ----------
def get_teachbase_token():
    client_id = load_env_var("TEACHBASE_CLIENT_ID")
    client_secret = load_env_var("TEACHBASE_CLIENT_SECRET")
    if not client_id or not client_secret:
        raise Exception("TEACHBASE_CLIENT_ID/TEACHBASE_CLIENT_SECRET не найдены в .env")
    resp = requests.post(f"{TEACHBASE_URL}/oauth/token", params={
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "client_credentials",
        "scope": "public"
    }, verify=False, timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]

def tb_get(token, path, params=None):
    resp = requests.get(f"{TEACHBASE_URL}/endpoint/v1{path}",
                        headers={"Authorization": f"Bearer {token}"},
                        params=params or {}, verify=False, timeout=30)
    resp.raise_for_status()
    return resp.json()

def get_teachbase_ankets():
    """Возвращает (headers, rows, ids, low_rating_cells) для листа анкет Teachbase за текущий месяц."""
    import time as _time
    headers = list(TB_FIELDS_MAP.values())
    try:
        token = get_teachbase_token()
    except Exception as e:
        print(f"Ошибка авторизации Teachbase: {e}")
        return headers, [], [], []

    # 1. Статистика по курсу-анкете. begin_at фильтрует по updated_at записи статистики,
    # а заполнение анкеты сразу завершает курс и обновляет статистику — поэтому окна
    # с 1-го числа текущего месяца достаточно, чтобы поймать всех заполнивших в этом месяце.
    begin_at = datetime.now().strftime("%Y-%m-01T00:00:00+03:00")
    stats = []
    page = 1
    try:
        while True:
            chunk = tb_get(token, f"/courses/{TEACHBASE_COURSE_ID}/course_stats",
                           {"per_page": 500, "page": page, "begin_at": begin_at})
            if not chunk:
                break
            stats.extend(chunk)
            page += 1
            _time.sleep(0.34)
    except Exception as e:
        print(f"Ошибка запроса course_stats Teachbase: {e}")
        return headers, [], [], []

    # Не фильтруем по флагу completed: у части заполнивших он False при progress 100%.
    # Истина — сам факт ответа на опрос (проверяется ниже по created_at ответа).
    by_user = {s["user_id"]: s for s in stats}
    session_ids = sorted({s["course_session_id"] for s in stats})
    print(f"  Teachbase: записей статистики {len(stats)}, потоков {len(session_ids)}")

    # 2. Ответы на опрос по каждому потоку, фильтр по месяцу заполнения
    month_prefix = datetime.now().strftime("%Y-%m")
    rows, ids, low_rating_cells = [], [], []
    for sid in session_ids:
        try:
            answers = tb_get(token, f"/course_sessions/{sid}/polls/{TEACHBASE_POLL_ID}/answers",
                             {"per_page": 1000})
        except Exception as e:
            print(f"  Teachbase: поток {sid} пропущен ({e})")
            continue
        _time.sleep(0.34)
        for ans in answers:
            if not ans.get("created_at", "").startswith(month_prefix):
                continue
            u = by_user.get(ans["user_id"], {})
            a = {}
            for item in ans.get("answers", []):
                key = TB_QUESTIONS.get(item.get("question_id"))
                if key:
                    a[key] = item["answer"][0] if item.get("answer") else ""
            dt = ans["created_at"][:19].replace("T", " ")
            rows.append([
                dt,
                u.get("course_session_name", ""),
                u.get("username", ""),
                u.get("user_email", ""),
                a.get("REVIEW", ""),
                a.get("RATING_ORG", ""),
                a.get("RATING_MATERIALS", ""),
                a.get("RATING_PLATFORM", ""),
                a.get("RATING_TEACHERS", ""),
                a.get("GOOD_EVENT", "").strip(),
                a.get("PUBLIC", "")
            ])
            ids.append(ans["id"])
            row_low = set()
            for key in TB_RATING_KEYS:
                if is_low_rating(a.get(key, "")):
                    row_low.add(list(TB_FIELDS_MAP).index(key) + 1)
            low_rating_cells.append(row_low)

    # свежие внизу — сортируем по дате заполнения по возрастанию
    order = sorted(range(len(rows)), key=lambda i: rows[i][0])
    rows = [rows[i] for i in order]
    ids = [ids[i] for i in order]
    low_rating_cells = [low_rating_cells[i] for i in order]
    return headers, rows, ids, low_rating_cells

def parse_reviews_data(raw_items):
    headers = list(REVIEWS_FIELDS_MAP.values())
    rows = []
    ids = []
    for item in raw_items:
        date_create = item.get("DATE_CREATE", "")
        client_data = item.get("PROPERTY_CLIENT_DATA", {})
        client_name = client_data.get("FULL_NAME", item.get("PROPERTY_CLIENT_VALUE", ""))
        event_name = item.get("PROPERTY_EVENT_NAME", "")
        module_name = item.get("PROPERTY_MODULE_NAME", "")
        direction = item.get("PROPERTY_DIRECTION_VALUE", "")
        review_raw = item.get("PROPERTY_REVIEW_VALUE", "")
        review_text = extract_text(review_raw)
        teacher = item.get("PROPERTY_TEACHER_NAME", "")
        rating = item.get("PROPERTY_RATING_VALUE", "")
        rows.append([
            date_create,
            client_name,
            event_name,
            module_name,
            direction,
            review_text,
            teacher,
            rating,
            client_data.get("ID", "")  # временно: ID контакта, заменяется на B2B/B2C в main()
        ])
        ids.append(item.get("ID", ""))
    return headers, rows, ids

def format_teachers(teachers):
    lines = []
    has_low_rating = False
    for t in teachers or []:
        rating = t.get("rating", "")
        line = f"{t.get('teacher', '')} — {rating}"
        comment = t.get("comment", "")
        if comment:
            line += f" ({comment})"
        lines.append(line)
        try:
            if rating != "" and float(rating) <= LOW_RATING_THRESHOLD:
                has_low_rating = True
        except ValueError:
            pass
    return "\n".join(lines), has_low_rating

def is_low_rating(rating):
    try:
        return rating != "" and float(rating) <= LOW_RATING_THRESHOLD
    except ValueError:
        return False

def parse_ankets_data(raw_items):
    headers = list(ANKETS_FIELDS_MAP.values())
    rows = []
    ids = []
    low_rating_cells = []  # список множеств номеров колонок (1-based) с низкой оценкой, по строке
    for item in raw_items:
        date_create = item.get("DATE_CREATE", "")
        event_name = item.get("PROPERTY_EVENT_REGISTER_NAME", "")
        review_raw = item.get("PROPERTY_REVIEW_VALUE", "")
        review_text = extract_text(review_raw)
        date_start = item.get("PROPERTY_EVENT_DATE_START_VALUE", "")
        date_end = item.get("PROPERTY_EVENT_DATE_END_VALUE", "")
        direction = item.get("PROPERTY_EVENT_DIRECTION_VALUE", "")
        client_data = item.get("PROPERTY_CLIENT_DATA", {})
        client_name = client_data.get("FULL_NAME", "")

        teachers_text, teachers_low = format_teachers(item.get("PROPERTY_RATING_TEACHER_VALUE"))

        quality = item.get("PROPERTY_RATING_QUALITY_VALUE") or {}
        rating_org = quality.get("rating7", "")
        comment_org = quality.get("comment7", "")
        rating_materials = quality.get("rating3", "")
        comment_materials = quality.get("comment3", "")
        rating_room = quality.get("rating4", "")
        comment_room = quality.get("comment4", "")

        good_event = item.get("PROPERTY_GOOD_EVENT_VALUE", "")
        public = item.get("PROPERTY_PUBLIC_VALUE", "")

        rows.append([
            date_create,
            event_name,
            review_text,
            date_start,
            date_end,
            direction,
            client_name,
            teachers_text,
            rating_org,
            comment_org,
            rating_materials,
            comment_materials,
            rating_room,
            comment_room,
            good_event,
            public,
            client_data.get("ID", "")  # временно: ID контакта, заменяется на B2B/B2C в main()
        ])

        row_low_cols = set()
        teachers_col = list(ANKETS_FIELDS_MAP).index("TEACHERS") + 1
        if teachers_low:
            row_low_cols.add(teachers_col)
        if is_low_rating(rating_org):
            row_low_cols.add(list(ANKETS_FIELDS_MAP).index("RATING_ORG") + 1)
        if is_low_rating(rating_materials):
            row_low_cols.add(list(ANKETS_FIELDS_MAP).index("RATING_MATERIALS") + 1)
        if is_low_rating(rating_room):
            row_low_cols.add(list(ANKETS_FIELDS_MAP).index("RATING_ROOM") + 1)
        low_rating_cells.append(row_low_cols)
        ids.append(item.get("ID", ""))

    return headers, rows, ids, low_rating_cells

def get_upload_link(token, disk_path):
    url = "https://cloud-api.yandex.net/v1/disk/resources/upload"
    headers = {"Authorization": f"OAuth {token}"}
    params = {"path": disk_path, "overwrite": "true"}
    resp = requests.get(url, headers=headers, params=params)
    resp.raise_for_status()
    return resp.json()["href"]

def upload_file(token, disk_path, file_buffer):
    upload_url = get_upload_link(token, disk_path)
    resp = requests.put(upload_url, data=file_buffer)
    resp.raise_for_status()
    print(f"Файл {disk_path} успешно загружен")

def download_existing_workbook(token, file_path):
    """Скачивает существующий файл с Яндекс.Диска. Если файла нет – возвращает None."""
    url = "https://cloud-api.yandex.net/v1/disk/resources/download"
    headers = {"Authorization": f"OAuth {token}"}
    params = {"path": file_path}
    try:
        resp = requests.get(url, headers=headers, params=params)
        if resp.status_code == 404:
            return None
        resp.raise_for_status()
        download_url = resp.json()["href"]
        file_resp = requests.get(download_url)
        file_resp.raise_for_status()
        wb = openpyxl.load_workbook(BytesIO(file_resp.content))
        print(f"Существующий файл {file_path} загружен")
        return wb
    except Exception as e:
        print(f"Не удалось загрузить существующий файл: {e}")
        return None

def update_or_create_sheet(wb, sheet_name, headers, rows, ids, low_rating_cells=None):
    """Создаёт лист (если его нет) и дописывает только новые строки (по ID), не трогая уже существующие."""
    id_col = len(headers) + 1
    id_col_letter = get_column_letter(id_col)

    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        existing_ids = {
            str(row[0].value) for row in ws.iter_rows(min_row=2, min_col=id_col, max_col=id_col)
            if row[0].value not in (None, "")
        }
        next_row = ws.max_row + 1
    else:
        ws = wb.create_sheet(title=sheet_name)
        for col_idx, header in enumerate(headers, 1):
            ws.cell(row=1, column=col_idx, value=header)
        ws.cell(row=1, column=id_col, value="ID")
        ws.column_dimensions[id_col_letter].hidden = True
        existing_ids = set()
        next_row = 2

    appended = 0
    for i, (row_data, row_id) in enumerate(zip(rows, ids)):
        if str(row_id) in existing_ids:
            continue
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=next_row, column=col_idx, value=value)
            if low_rating_cells and col_idx in low_rating_cells[i]:
                cell.fill = LOW_RATING_FILL
        ws.cell(row=next_row, column=id_col, value=row_id)
        next_row += 1
        appended += 1

    # Автоширина (кроме скрытой колонки ID)
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter == id_col_letter:
            continue
        max_len = 0
        for cell in col:
            try:
                if len(str(cell.value)) > max_len:
                    max_len = len(str(cell.value))
            except:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 2, 50)
    print(f"Лист '{sheet_name}': добавлено новых строк {appended} (пропущено уже существующих: {len(rows) - appended})")

def get_month_name_ru():
    months = {
        1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
        5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
        9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
    }
    return months[datetime.now().month]

# ---------- ОСНОВНАЯ ФУНКЦИЯ ----------
def main():
    try:
        yandex_token = load_yandex_token()
    except Exception as e:
        print(f"Ошибка загрузки токена: {e}")
        return

    # 1. Получаем данные из Bitrix (только за текущий месяц)
    print("Запрос getReviews...")
    reviews_raw = get_bitrix_data("getReviews")
    print(f"Отзывов за текущий месяц: {len(reviews_raw)}")

    print("Запрос getAnkets...")
    ankets_raw = get_bitrix_data("getAnkets")
    print(f"Анкет за текущий месяц: {len(ankets_raw)}")

    # Bitrix отдаёт данные от новых к старым, а строки в таблице дописываются
    # в конец без сдвига существующих — переворачиваем, чтобы новые строки
    # добавлялись в хронологическом порядке (старые сверху).
    reviews_raw = list(reversed(reviews_raw))
    ankets_raw = list(reversed(ankets_raw))

    # 2. Преобразуем в строки и заголовки
    rev_headers, rev_rows, rev_ids = parse_reviews_data(reviews_raw)
    ank_headers, ank_rows, ank_ids, ank_low_rating_cells = parse_ankets_data(ankets_raw)

    # 2.1 Определяем B2B/B2C по COMPANY_ID контакта (последний элемент каждой строки — временно ID контакта)
    all_contact_ids = [row[-1] for row in rev_rows] + [row[-1] for row in ank_rows]
    company_by_contact = fetch_contacts_company_map(all_contact_ids)
    for row in rev_rows:
        row[-1] = detect_b2b(company_by_contact.get(str(row[-1]), ""))
    for row in ank_rows:
        row[-1] = detect_b2b(company_by_contact.get(str(row[-1]), ""))

    # 2.2 Анкеты Teachbase за текущий месяц
    print("Запрос анкет Teachbase...")
    tb_headers, tb_rows, tb_ids, tb_low_rating_cells = get_teachbase_ankets()
    print(f"Анкет Teachbase за текущий месяц: {len(tb_rows)}")

    # 3. Загружаем существующую книгу (если есть)
    wb = download_existing_workbook(yandex_token, YANDEX_FILE_PATH)
    if wb is None:
        wb = openpyxl.Workbook()
        # Удаляем стандартный лист, который создаётся по умолчанию
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])
        print("Создана новая книга")

    # 4. Добавляем/обновляем листы текущего месяца
    month_ru = get_month_name_ru()
    sheet_reviews = f"Отз поМод {month_ru}"
    sheet_ankets = f"анкеты {month_ru}"
    sheet_tb = f"анкеты ТБ {month_ru}"

    update_or_create_sheet(wb, sheet_reviews, rev_headers, rev_rows, rev_ids)
    update_or_create_sheet(wb, sheet_ankets, ank_headers, ank_rows, ank_ids, ank_low_rating_cells)
    update_or_create_sheet(wb, sheet_tb, tb_headers, tb_rows, tb_ids, tb_low_rating_cells)

    # 5. Сохраняем и загружаем на Яндекс.Диск
    with BytesIO() as buf:
        wb.save(buf)
        buf.seek(0)
        try:
            upload_file(yandex_token, YANDEX_FILE_PATH, buf)
            print(f"Готово. Файл {YANDEX_FILE_PATH} обновлён. Листы: {sheet_reviews}, {sheet_ankets}, {sheet_tb}")
            print(f"Записано отзывов: {len(rev_rows)}, анкет: {len(ank_rows)}, анкет ТБ: {len(tb_rows)}")
        except Exception as e:
            print(f"Ошибка загрузки на Яндекс.Диск: {e}")

if __name__ == "__main__":
    main()