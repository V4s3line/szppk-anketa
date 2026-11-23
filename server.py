"""Онлайн-анкетирование пассажиров АО «СЗППК».

Запуск:  ADMIN_TOKEN=секрет python server.py   (порт по умолчанию 8090, PORT=...)
Анкета:  http://<сервер>:8090/
Админка: http://<сервер>:8090/admin?token=секрет
Выгрузка в формат листа «База оценок» (вставлять с ячейки B19):
         /admin/export.xlsx?token=...&year=2027&q=1

Зависимости: только стандартная библиотека; для .xlsx — openpyxl (без него доступен CSV).
"""
import csv, hashlib, io, json, os, sqlite3, time
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

HERE = os.path.dirname(os.path.abspath(__file__))
DB = os.environ.get('ANKETA_DB', os.path.join(HERE, 'anketa.sqlite3'))
TOKEN = os.environ.get('ADMIN_TOKEN', '')
PORT = int(os.environ.get('PORT', '8090'))
SALT = os.environ.get('IP_SALT', 'szppk')
MIN_SECONDS = 60          # быстрее минуты анкету честно не заполнить
MAX_PER_IP_HOUR = 5
# Плановые значения индекса по кварталам (правьте при необходимости)
PLAN = {1: 85.6, 2: 85.8, 3: 85.5, 4: 85.5}

# Заголовки столбцов листа «База оценок» (индекс 0 = столбец A)
HEADERS = {
    1: 'Филиал/ ДЗО/ ППК', 2: 'Структурное подразделение/ Филиал ДЗО/ Филиал ППК', 3: 'Дорога',
    4: 'Название вокзального комплекса/ остановочного пункта', 5: 'Номер поезда', 6: 'Номер вагона',
    7: 'Класс обслуживания', 8: 'Тип вагона', 9: 'Станция отправления', 10: 'Станция назначения',
    11: 'Цель поездки', 12: 'Частота поездок', 13: 'Пол респондента', 14: 'Возраст респондента',
    15: 'Способ оформления билета', 16: 'Время ожидания в очереди в кассу', 17: 'Клиентоориентированность кассира',
    18: 'Доступность изложения кассиром информации', 19: 'Абонементный билет/ транспортная карта (касса)',
    20: 'Сложный маршрут с пересадками (касса)', 21: 'Удобство поиска информации', 22: 'Быстрота оформления и оплаты',
    23: 'Исправность работы', 24: 'Абонементный билет/ транспортная карта (онлайн)', 25: 'Сложный маршрут с пересадками (онлайн)',
    52: 'Исправность работы оборудования', 53: 'Чистота', 54: 'Комфорт посадочного места', 55: 'Места для вещей и багажа',
    56: 'Температурный режим', 57: 'Уровень шума', 58: 'Освещенность', 59: 'Понятность объявлений',
    60: 'Чувство личной безопасности', 61: 'Сохранность личных вещей', 62: 'Ассортимент услуг',
    63: 'Зарядка мобильных устройств', 64: 'Санитарная комната (туалет)', 72: 'Питание на пассажирском месте',
    73: 'Сопутствующие товары и сувениры', 74: 'Беспроводной доступ в интернет', 75: 'Клиентоориентированность представителя перевозчика',
    76: 'Удобство времени отправления', 77: 'Удобство времени прибытия', 78: 'Продолжительность поездки',
    80: 'Стоимость проезда', 82: 'Стоимость сопутствующих товаров', 83: 'Обращение: на сайте', 84: 'Обращение: в мобильном приложении',
    85: 'Обращение: виртуальная приемная', 86: 'Обращение: ЦПК', 87: 'Обращение: книга жалоб', 88: 'С чем было связано обращение',
    89: 'ЦПК: клиентоориентированность', 90: 'ЦПК: компетентность', 91: 'ЦПК: оперативность', 92: 'ЦПК: полезность информации',
    93: 'ЦПК: решен ли вопрос', 94: 'РЖД-Бонус', 95: 'Приоритет: расписание', 96: 'Приоритет: безопасность',
    97: 'Приоритет: стоимость', 98: 'Приоритет: комфорт вагона', 99: 'Приоритет: вокзалы/ОП', 100: 'Приоритет: доп. услуги',
    101: 'Приоритет: информация', 102: 'Приоритет: оформление билета', 103: 'Приоритет: обслуживание',
    104: 'Изменение качества за год', 105: 'Изменить: расписание', 106: 'Изменить: безопасность', 107: 'Изменить: стоимость доп. услуг',
    108: 'Изменить: комфорт вагона', 109: 'Изменить: инфраструктура', 110: 'Изменить: доп. услуги', 111: 'Изменить: информация',
    112: 'Изменить: оформление билета', 113: 'Изменить: обслуживание', 114: 'Изменить: другое', 115: 'NPS',
}
LAST_COL = 115
RATING_COLS = set(range(16, 26)) | set(range(52, 65)) | {72, 73, 74, 75, 76, 77, 78, 80, 82} | set(range(83, 88)) | set(range(89, 93))
TEXT_COLS = {5, 6, 7, 9, 10, 11, 12, 13, 14, 15, 88, 93, 104, 114}
# В индекс ППК не входят каналы обратной связи (83–87) и ЦПК (89–92): в «Отчете» их вес = 0 при фильтре «ППК»
INDEX_COLS = RATING_COLS - set(range(83, 88)) - set(range(89, 93))
RANK_COLS = set(range(95, 104))
BIN_COLS = set(range(105, 114))


def db():
    con = sqlite3.connect(DB)
    con.execute('CREATE TABLE IF NOT EXISTS answers(id INTEGER PRIMARY KEY, ts REAL, year INT, q INT, iph TEXT, secs INT, row TEXT)')
    return con


def clean(row):
    """Проверка и нормализация присланной строки; возвращает dict{col:int->value}."""
    out = {}
    for k, v in row.items():
        c = int(k)
        if c in RATING_COLS:
            if v not in (1, 2, 3, 4, 5):
                raise ValueError(f'оценка {c}')
            out[c] = v
        elif c in RANK_COLS:
            if v not in range(1, 10):
                raise ValueError(f'ранг {c}')
            out[c] = v
        elif c in BIN_COLS:
            out[c] = 1 if v == 1 else 0
        elif c == 115:
            if v not in range(1, 11):
                raise ValueError('NPS')
            out[c] = v
        elif c in TEXT_COLS:
            if isinstance(v, (int, float)) and c in (5, 6, 114):
                out[c] = int(v)
            else:
                s = str(v).strip()[:300]
                if s:
                    out[c] = s
    if sorted(out.get(c) for c in RANK_COLS if c in out) != list(range(1, 10)):
        raise ValueError('приоритеты должны быть 1..9 без повторов')
    picked = sum(out.get(c, 0) for c in BIN_COLS) + (1 if out.get(114) not in (None, 0) else 0)
    if not 1 <= picked <= 3:
        raise ValueError('нужно выбрать 1–3 пункта')
    out[1], out[2] = 'ППК', 'СЗППК'
    return out


def index_of(rows):
    """Индекс удовлетворенности по формуле листа «Отчет»: средняя всех оценок 1–5 → 100-балльная шкала,
    вес каждого параметра = доля его оценок (эквивалентно Σ L_блока × X_блока)."""
    # Точно как ветка «ППК» формулы L37: + H271*V271 (книга жалоб) в числителе, но её оценки не входят в знаменатель
    s = n = 0
    for r in rows:
        for c in INDEX_COLS:
            v = r.get(c)
            if isinstance(v, int):
                s += (v - 1) * 25
                n += 1
        v = r.get(87)
        if isinstance(v, int):
            s += (v - 1) * 25
    return (s / n if n else None), n


def nps_of(rows):
    v = [r[115] for r in rows if 115 in r]
    if not v:
        return None
    return 100 * (sum(x >= 9 for x in v) - sum(x <= 6 for x in v)) / len(v)


def load(year=None, q=None):
    con = db()
    sql, args = 'SELECT ts,year,q,row FROM answers WHERE 1=1', []
    if year:
        sql += ' AND year=?'; args.append(int(year))
    if q:
        sql += ' AND q=?'; args.append(int(q))
    res = []
    for ts, y, qq, row in con.execute(sql + ' ORDER BY id', args):
        d = {int(k): v for k, v in json.loads(row).items()}
        res.append((ts, y, qq, d))
    con.close()
    return res


class H(BaseHTTPRequestHandler):
    server_version = 'anketa'

    def _send(self, code, body, ctype='text/plain; charset=utf-8', extra=None):
        if isinstance(body, str):
            body = body.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('X-Content-Type-Options', 'nosniff')
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _ip(self):
        return self.headers.get('X-Real-IP') or self.client_address[0]

    def _authed(self, qs):
        return TOKEN and qs.get('token', [''])[0] == TOKEN

    def log_message(self, fmt, *a):
        pass

    def do_GET(self):
        u = urlparse(self.path)
        qs = parse_qs(u.query)
        if u.path in ('/', '/index.html'):
            with open(os.path.join(HERE, 'index.html'), 'rb') as f:
                return self._send(200, f.read(), 'text/html; charset=utf-8')
        if u.path.startswith('/admin'):
            if not self._authed(qs):
                return self._send(403, 'Нет доступа')
            if u.path == '/admin':
                return self._send(200, self.admin_page(qs), 'text/html; charset=utf-8')
            if u.path in ('/admin/export.xlsx', '/admin/export.csv'):
                return self.export(qs, u.path.endswith('.xlsx'))
        self._send(404, 'Не найдено')

    def do_POST(self):
        if urlparse(self.path).path != '/api/submit':
            return self._send(404, 'Не найдено')
        try:
            n = int(self.headers.get('Content-Length', '0'))
            if n > 50_000:
                return self._send(413, 'Слишком большой запрос')
            data = json.loads(self.rfile.read(n))
            if data.get('hp'):
                return self._send(200, 'ok')          # бот заполнил скрытое поле — молча игнорируем
            secs = int(data.get('secs') or 0)
            row = clean(data['row'])
        except Exception as e:
            return self._send(400, f'Некорректные данные: {e}')
        iph = hashlib.sha256((SALT + self._ip()).encode()).hexdigest()[:16]
        now = time.time()
        con = db()
        recent = con.execute('SELECT COUNT(*) FROM answers WHERE iph=? AND ts>?', (iph, now - 3600)).fetchone()[0]
        if recent >= MAX_PER_IP_HOUR:
            con.close()
            return self._send(429, 'Слишком много анкет с одного адреса')
        dt = datetime.now()
        con.execute('INSERT INTO answers(ts,year,q,iph,secs,row) VALUES(?,?,?,?,?,?)',
                    (now, dt.year, (dt.month - 1) // 3 + 1, iph, secs, json.dumps(row, ensure_ascii=False)))
        con.commit()
        con.close()
        self._send(200, 'ok')

    def admin_page(self, qs):
        tok = qs['token'][0]
        all_rows = load()
        groups = {}
        for ts, y, q, r in all_rows:
            groups.setdefault((y, q), []).append((ts, r))
        con = db()
        fast_map = dict(((y, q), n) for y, q, n in con.execute(
            'SELECT year,q,COUNT(*) FROM answers WHERE secs<? GROUP BY year,q', (MIN_SECONDS,)))
        con.close()
        trs = []
        for (y, q) in sorted(groups):
            items = groups[(y, q)]
            rows = [r for _, r in items]
            fast = fast_map.get((y, q), 0)
            idx, nr = index_of(rows)
            nps = nps_of(rows)
            plan = PLAN.get(q)
            diff = f'{idx - plan:+.2f}' if idx is not None and plan else '—'
            trs.append(f'<tr><td>{y}</td><td>{q}</td><td>{len(rows)}{f" (быстрых &lt;1 мин: {fast})" if fast else ""}</td><td>{nr}</td>'
                       f'<td><b>{idx:.2f}</b></td><td>{plan}</td><td>{diff}</td><td>{nps:.0f}</td>'
                       f'<td><a href="/admin/export.xlsx?token={tok}&year={y}&q={q}">xlsx</a> · '
                       f'<a href="/admin/export.csv?token={tok}&year={y}&q={q}">csv</a></td></tr>'
                       if idx is not None else f'<tr><td>{y}</td><td>{q}</td><td>{len(rows)}</td><td colspan=6>нет оценок</td></tr>')
        body = ''.join(trs) or '<tr><td colspan=9>Анкет пока нет</td></tr>'
        return f'''<!doctype html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Анкеты СЗППК</title><style>body{{font:15px/1.4 Segoe UI,Arial,sans-serif;margin:16px;background:#fff;color:#1c2230}}
table{{border-collapse:collapse;width:100%;max-width:900px}}td,th{{border:1px solid #d9dde5;padding:6px 8px;text-align:center}}
th{{background:#f1f3f6}}</style></head><body><h2>Онлайн-анкетирование: сводка</h2>
<p>Индекс считается по формуле листа «Отчет» (средняя всех оценок 1–5 в 100-балльной шкале, вес — число оценок). Онлайн-анкеты — отдельный метод: в первый год их учитывают раздельно с традиционными (п. 4.3 Методики).</p>
<table><tr><th>Год</th><th>Квартал</th><th>Анкет</th><th>Оценок</th><th>Индекс</th><th>План</th><th>Откл.</th><th>NPS</th><th>Выгрузка</th></tr>{body}</table>
<p><a href="/admin/export.xlsx?token={tok}">Выгрузить всё (xlsx)</a></p></body></html>'''

    def export(self, qs, as_xlsx):
        year = qs.get('year', [None])[0]
        q = qs.get('q', [None])[0]
        rows = load(year, q)
        cols = list(range(1, LAST_COL + 1))
        head = [HEADERS.get(c, '') for c in cols] + ['Дата заполнения', 'Метод']
        data = [[r.get(c) for c in cols] + [datetime.fromtimestamp(ts).strftime('%d.%m.%Y %H:%M'), 'онлайн']
                for ts, _, _, r in rows]
        name = f'anketa_{year or "all"}_{"q" + q if q else "all"}'
        if as_xlsx:
            try:
                import openpyxl
            except ImportError:
                return self._send(500, 'openpyxl не установлен — используйте export.csv')
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = 'База оценок (онлайн)'
            ws.append(head)
            for d in data:
                ws.append(d)
            ws.freeze_panes = 'A2'
            buf = io.BytesIO()
            wb.save(buf)
            return self._send(200, buf.getvalue(),
                              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                              {'Content-Disposition': f'attachment; filename="{name}.xlsx"'})
        buf = io.StringIO()
        w = csv.writer(buf, delimiter=';')
        w.writerow(head)
        w.writerows(data)
        return self._send(200, '﻿' + buf.getvalue(), 'text/csv; charset=utf-8',
                          {'Content-Disposition': f'attachment; filename="{name}.csv"'})


if __name__ == '__main__':
    if not TOKEN:
        print('ВНИМАНИЕ: ADMIN_TOKEN не задан — админка и выгрузка отключены.')
    db().close()
    print(f'Анкета: http://127.0.0.1:{PORT}/')
    ThreadingHTTPServer(('0.0.0.0', PORT), H).serve_forever()
