"""wallet-diary 的算数脚本：一句话记账背后的所有数，只用 Python 标准库。

所有金额、余额、合计、id 都在这里算，Agent 只读本脚本输出的 JSON，不心算。

Author: Griffin-CJH (https://github.com/Griffin-CJH)
"""

__author__ = "Griffin-CJH"
__author_url__ = "https://github.com/Griffin-CJH"
__license__ = "MIT"

import argparse
import calendar
import csv
import datetime as dt
import json
import os
import re
import sys
import webbrowser
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path

LEDGER_COLS = ["id", "date", "time", "amount", "account", "category", "item", "raw_text"]
ACCOUNT_COLS = ["account", "currency", "opening_balance", "opening_date"]
CATEGORY_COLS = ["category", "direction", "definition"]
RECURRING_COLS = ["item", "amount", "account", "category", "cycle", "day", "active", "last_logged"]

# 内置类别由 skill 定义，不写进 categories.csv。校准的差额可正可负，所以是「双向」
BUILTIN_CATEGORIES = {"转出": "支出", "转入": "收入", "校准": "双向"}
TRANSFER = ("转出", "转入")
UNASSIGNED = "（未指定账户）"

# 全部按「钱花在什么用途上」分，不按付钱方式分（没有「订阅」「网购」这类）。
# 定义写成「包含什么；不包含的归哪」，让一句话能稳定归类；会员、软件归哪类不在这里替用户定
DEFAULT_CATEGORIES = [
    ("餐饮", "支出", "现成能吃的：堂食、外卖、咖啡奶茶、零食饮料；请别人吃→人情；买生的回家做→买菜"),
    ("买菜", "支出", "回家做饭用的生鲜、粮油调料、水果；超市混合小票按大头归"),
    ("交通", "支出", "公交地铁、打车、火车机票、加油停车、车辆保养和保险"),
    ("住房", "支出", "房租、房贷、物业费：住下来就得付、自己控制不了的；水电网费话费→生活缴费"),
    ("生活缴费", "支出", "水电燃气、宽带、话费：用多少自己能控制的日常账单"),
    ("日用", "支出", "用完要再买的消耗品：洗护清洁、纸巾、化妆品、理发；耐用的→购物"),
    ("购物", "支出", "能用很久的东西：衣服鞋包、数码家电、家具；消耗品→日用"),
    ("娱乐", "支出", "为了开心花的：电影演出、游戏、运动健身、旅行玩乐"),
    ("医疗", "支出", "看病、买药、体检、牙科"),
    ("学习", "支出", "为了变强花的：书、课程、考试"),
    ("人情", "支出", "为关系花的：请客、送礼、发红包、给家人的钱；AA 自己那份→餐饮"),
    ("其他支出", "支出", "以上都放不进的：保险、罚款、手续费、捐款等"),
    ("工资", "收入", "主业收入：工资、奖金、补贴"),
    ("副业", "收入", "主业以外靠干活挣的：兼职、稿费、接单、平台分成"),
    ("投资收入", "收入", "利息、分红、理财和基金收益"),
    ("其他收入", "收入", "收到的红包礼金、退税、中奖等；退款不算收入，冲回原来那笔支出"),
]
DEFAULT_SETTINGS = {"default_currency": "CNY", "show_balance": True, "celebrate_income": True}

DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
TIME_RE = re.compile(r"^([01]\d|2[0-3]):[0-5]\d$")


class WalletError(Exception):
    """用户看得懂的错误：main() 把它写到 stderr，退出码 1。"""


def pointer_file():
    """记着数据夹在哪的指针文件；测试用环境变量换掉它，不碰用户家目录。"""
    override = os.environ.get("WALLET_DIARY_POINTER")
    return Path(override) if override else Path.home() / ".wallet-diary.json"


def money(s):
    try:
        d = Decimal(str(s).strip().replace(",", ""))
    except InvalidOperation:
        raise WalletError(f"金额不是数字：{s!r}")
    if not d.is_finite():
        raise WalletError(f"金额不是数字：{s!r}")
    return d


def num(d):
    """Decimal → JSON 里的数：整数给 int，其余给 float。"""
    return int(d) if d == d.to_integral_value() else float(d)


def parse_date(s):
    s = (s or "").strip()
    if not DATE_RE.match(s):
        raise WalletError(f"日期要写成 YYYY-MM-DD：{s!r}")
    try:
        return dt.date.fromisoformat(s)
    except ValueError:
        raise WalletError(f"没有这一天：{s}")


def read_table(path, required):
    """读一张 CSV，返回 (表头, 行)。多余的列原样保留；格子数对不上就报错，不静默丢行。"""
    path = Path(path)
    if not path.exists():
        raise WalletError(f"找不到 {path}")
    try:
        with open(path, encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            header = list(reader.fieldnames or [])
            missing = [c for c in required if c not in header]
            if missing:
                raise WalletError(f"{path.name} 缺列：{'、'.join(missing)}")
            rows = []
            for row in reader:
                if None in row:
                    raise WalletError(f"{path.name} 第 {reader.line_num} 行的格子比表头多（原话里有逗号却没加引号？）")
                if any(v is None for v in row.values()):
                    raise WalletError(f"{path.name} 第 {reader.line_num} 行的格子比表头少")
                if not any(v.strip() for v in row.values()):
                    continue
                rows.append(row)
    except UnicodeDecodeError:
        raise WalletError(f"{path.name} 不是 UTF-8 编码：在 Excel 里另存为「CSV UTF-8」格式再试")
    return header, rows


def write_table(path, header, rows):
    """先写临时文件再替换，写到一半出错也不会留下半张表。"""
    path = Path(path)
    tmp = path.with_name(path.name + ".tmp")
    try:
        with open(tmp, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=header, lineterminator="\n")
            writer.writeheader()
            writer.writerows(rows)
        tmp.replace(path)
    except OSError as e:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                # tmp 删除失败不重要；我们关心的错误是原来的 OSError
                pass
        raise WalletError(f"写不进 {path.name}：Excel 或 OneDrive 可能在占用这个文件，先关掉它再试") from e


def init_data(data_dir, currency, set_pointer=True):
    """建数据夹。已有的文件一个都不覆盖，所以对已有数据夹再跑一次 = 只更新指针。"""
    data_dir = Path(data_dir).resolve()
    data_dir.mkdir(parents=True, exist_ok=True)
    (data_dir / "chapters").mkdir(exist_ok=True)
    tables = {
        "ledger.csv": (LEDGER_COLS, []),
        "accounts.csv": (ACCOUNT_COLS, []),
        "categories.csv": (CATEGORY_COLS, [dict(zip(CATEGORY_COLS, c)) for c in DEFAULT_CATEGORIES]),
        "recurring.csv": (RECURRING_COLS, []),
    }
    created = []
    for name, (cols, rows) in tables.items():
        if not (data_dir / name).exists():
            write_table(data_dir / name, cols, rows)
            created.append(name)
    settings = data_dir / "settings.json"
    if not settings.exists():
        settings.write_text(
            json.dumps({**DEFAULT_SETTINGS, "default_currency": currency}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        created.append("settings.json")
    if set_pointer:
        pointer_file().write_text(json.dumps({"data_dir": str(data_dir)}, ensure_ascii=False), encoding="utf-8")
    return {"data_dir": str(data_dir), "created": created}


def resolve_data_dir(arg):
    if arg:
        return Path(arg).resolve()
    p = pointer_file()
    if not p.exists():
        raise WalletError("还没有数据文件夹：先运行 init --data <路径>")
    try:
        data_dir = Path(json.loads(p.read_text(encoding="utf-8"))["data_dir"])
    except (json.JSONDecodeError, KeyError, TypeError) as e:
        raise WalletError(f"指针文件 {p} 破损或格式错误，删掉它重新 init") from e
    if not data_dir.exists():
        raise WalletError(f"指针指向的数据文件夹不存在了：{data_dir}")
    return data_dir


class Book:
    """一次性读进整个数据夹。改完调 save_* 写回。"""

    def __init__(self, data_dir):
        self.dir = Path(data_dir)
        self.ledger_header, self.ledger = read_table(self.dir / "ledger.csv", LEDGER_COLS)
        _, self.accounts_rows = read_table(self.dir / "accounts.csv", ACCOUNT_COLS)
        self.accounts = {a["account"]: a for a in self.accounts_rows}
        _, self.category_rows = read_table(self.dir / "categories.csv", CATEGORY_COLS)
        self.categories = {c["category"]: c["direction"] for c in self.category_rows}
        self.recurring_header, self.recurring = read_table(self.dir / "recurring.csv", RECURRING_COLS)
        try:
            settings_file = self.dir / "settings.json"
            if not settings_file.exists():
                raise WalletError("settings.json 找不到，数据文件夹可能已损坏")
            settings = json.loads(settings_file.read_text(encoding="utf-8"))
        except json.JSONDecodeError as e:
            raise WalletError("settings.json 格式错误（不是合法 JSON）") from e
        except OSError as e:
            raise WalletError("settings.json 读不了（权限或文件问题）") from e
        self.settings = {**DEFAULT_SETTINGS, **settings}

    def direction(self, category):
        return BUILTIN_CATEGORIES.get(category) or self.categories.get(category)

    def currency(self, account):
        """账户的币种；账户留空用默认币种；不认识的账户返回 None（check 会报）。"""
        if not account or account == UNASSIGNED:
            return self.settings["default_currency"]
        acc = self.accounts.get(account)
        return acc["currency"] if acc else None

    def save_ledger(self):
        write_table(self.dir / "ledger.csv", self.ledger_header, self.ledger)

    def save_recurring(self):
        write_table(self.dir / "recurring.csv", self.recurring_header, self.recurring)


def account_list(book):
    names = list(book.accounts)
    if any(not r["account"] for r in book.ledger):
        names.append(UNASSIGNED)
    return names


def opening(book, account):
    """期初 = opening_date 当天开始时的余额。没写账户的流水从 0 起算、不设期初日。"""
    if account == UNASSIGNED:
        return Decimal(0), dt.date.min
    a = book.accounts[account]
    return money(a["opening_balance"] or "0"), parse_date(a["opening_date"])


def balance_at(book, account, day):
    """该账户在 day 当天结束时的余额：期初日及以后的流水往上加，更早的往回倒推。"""
    base, start = opening(book, account)
    key = "" if account == UNASSIGNED else account
    total = base
    for r in book.ledger:
        if r["account"] != key:
            continue
        d = parse_date(r["date"])
        if start <= d <= day:
            total += money(r["amount"])
        elif day < d < start:
            total -= money(r["amount"])
    return total


def balances_view(book, day, names=None):
    names = names if names is not None else account_list(book)
    return [
        {"account": n, "currency": book.currency(n), "balance": num(balance_at(book, n, day))}
        for n in names
    ]


def parse_cycle_day(cycle, day):
    day = (day or "").strip()
    try:
        if cycle == "月" and 1 <= int(day) <= 31:
            return int(day)
        if cycle == "周" and 1 <= int(day) <= 7:
            return int(day)
        if cycle == "年":
            m, d = (int(x) for x in day.split("-"))
            dt.date(2024, m, d)  # 2024 是闰年，02-29 也合法
            return (m, d)
    except ValueError:
        pass
    raise WalletError(f"cycle「{cycle}」配 day「{day}」不合法：月填 1–31，周填 1–7（周一=1），年填 MM-DD")


def row_problems(book, r):
    errors, warnings = [], []
    tag = f"id {r['id']}（{r['date']} {r['item']}）"
    try:
        parse_date(r["date"])
    except WalletError as e:
        errors.append(f"{tag}：{e}")
    if r["time"].strip() and not TIME_RE.match(r["time"].strip()):
        errors.append(f"{tag}：时间要写成 HH:MM")
    try:
        amt = money(r["amount"])
    except WalletError as e:
        errors.append(f"{tag}：{e}")
        amt = None
    if r["account"] and r["account"] not in book.accounts:
        errors.append(f"{tag}：账户「{r['account']}」不在 accounts.csv 里")
    if r["category"] in TRANSFER and not r["account"]:
        errors.append(f"{tag}：转账必须写明账户")
    direction = book.direction(r["category"])
    if direction is None:
        errors.append(f"{tag}：类别「{r['category']}」不在 categories.csv 里，也不是内置类别")
    if amt is not None and direction:
        if amt == 0:
            warnings.append(f"{tag}：金额是 0")
        elif direction == "支出" and amt > 0:
            warnings.append(f"{tag}：类别「{r['category']}」是支出，金额却是正数——漏了负号，还是退款？")
        elif direction == "收入" and amt < 0:
            warnings.append(f"{tag}：类别「{r['category']}」是收入，金额却是负数——多了负号？")
    return errors, warnings


def transfer_problems(book, rows):
    """同一组转账（同 date、同 raw_text）要有转出也有转入；同币种时两边要正负相抵。"""
    ids = "、".join(r["id"] for r in rows)
    has_out = any(r["category"] == "转出" for r in rows)
    has_in = any(r["category"] == "转入" for r in rows)
    if not (has_out and has_in):
        return [f"转账 id {ids}（{rows[0]['date']}）只有{'转出' if has_out else '转入'}，缺另一半"]
    if len({book.currency(r["account"]) for r in rows}) == 1:
        try:
            total = sum((money(r["amount"]) for r in rows), Decimal(0))
        except WalletError:
            return []  # 金额本身的错 row_problems 已经报了
        if total != 0:
            return [f"转账 id {ids}（{rows[0]['date']}）两边不相抵，差 {num(total)}"]
    return []


def table_problems(book):
    errors, warnings = [], []
    seen = set()
    for a in book.accounts_rows:
        tag = f"账户「{a['account']}」"
        if not a["account"].strip():
            errors.append("accounts.csv 有一行没写账户名")
        if a["account"] in seen:
            errors.append(f"{tag}重复")
        seen.add(a["account"])
        if not a["currency"].strip():
            errors.append(f"{tag}没写币种")
        try:
            money(a["opening_balance"] or "0")
        except WalletError:
            errors.append(f"{tag}的期初余额不是数字")
        try:
            parse_date(a["opening_date"])
        except WalletError:
            errors.append(f"{tag}的期初日期要写成 YYYY-MM-DD")
    seen = set()
    for c in book.category_rows:
        tag = f"类别「{c['category']}」"
        if c["category"] in seen:
            errors.append(f"{tag}重复")
        seen.add(c["category"])
        if c["category"] in BUILTIN_CATEGORIES:
            errors.append(f"{tag}是内置类别，不要写进 categories.csv")
        if c["direction"] not in ("支出", "收入"):
            errors.append(f"{tag}的方向只能是 支出 / 收入")
    seen = set()
    for rec in book.recurring:
        tag = f"定期项「{rec['item']}」"
        if rec["item"] in seen:
            errors.append(f"{tag}重复")
        seen.add(rec["item"])
        try:
            amt = money(rec["amount"])
        except WalletError:
            errors.append(f"{tag}的金额不是数字")
            amt = None
        if rec["account"] and rec["account"] not in book.accounts:
            errors.append(f"{tag}的账户「{rec['account']}」不在 accounts.csv 里")
        direction = book.direction(rec["category"])
        if direction is None:
            errors.append(f"{tag}的类别「{rec['category']}」不存在")
        elif amt is not None and ((direction == "支出" and amt > 0) or (direction == "收入" and amt < 0)):
            warnings.append(f"{tag}的金额正负号和类别「{rec['category']}」的方向对不上")
        try:
            parse_cycle_day(rec["cycle"], rec["day"])
        except WalletError as e:
            errors.append(f"{tag}：{e}")
        if rec["active"].strip().upper() not in ("Y", "N"):
            errors.append(f"{tag}的 active 只能填 Y / N")
        if rec["last_logged"].strip():
            try:
                parse_date(rec["last_logged"])
            except WalletError:
                errors.append(f"{tag}的 last_logged 要写成 YYYY-MM-DD")
    return errors, warnings


def check(book, ids=None):
    """ids=None 全表体检；给 id 集合就只查这些行和它们所在的转账组。id 重复永远全表查。"""
    errors, warnings = [], []
    seen = set()
    for r in book.ledger:
        rid = r["id"].strip()
        if not rid.isdigit():
            errors.append(f"id {rid!r} 不是正整数（{r['date']} {r['item']}）")
        elif rid in seen:
            errors.append(f"id {rid} 重复")
        seen.add(rid)
    for r in book.ledger:
        if ids is None or r["id"] in ids:
            e, w = row_problems(book, r)
            errors += e
            warnings += w
    groups = defaultdict(list)
    for r in book.ledger:
        if r["category"] in TRANSFER:
            groups[(r["date"], r["raw_text"])].append(r)
    for rows in groups.values():
        if ids is None or any(r["id"] in ids for r in rows):
            warnings += transfer_problems(book, rows)
    if ids is None:
        e, w = table_problems(book)
        errors += e
        warnings += w
    return {"errors": errors, "warnings": warnings}


def occurrences(rec, start, end):
    """定期项在 [start, end] 内的到期日。月底不够的日子（31 号遇到 2 月）落到当月最后一天。"""
    spec = parse_cycle_day(rec["cycle"], rec["day"])
    out = []
    if rec["cycle"] == "月":
        y, m = start.year, start.month
        while True:
            d = dt.date(y, m, min(spec, calendar.monthrange(y, m)[1]))
            if d > end:
                break
            if d >= start:
                out.append(d)
            y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    elif rec["cycle"] == "年":
        mm, dd = spec
        for y in range(start.year, end.year + 1):
            d = dt.date(y, mm, min(dd, calendar.monthrange(y, mm)[1]))
            if start <= d <= end:
                out.append(d)
    else:  # 周
        d = start + dt.timedelta(days=(spec - 1 - start.weekday()) % 7)
        while d <= end:
            out.append(d)
            d += dt.timedelta(days=7)
    return out


def rec_view(book, rec, day):
    return {
        "date": day.isoformat(),
        "item": rec["item"],
        "amount": num(money(rec["amount"])),
        "currency": book.currency(rec["account"]),
        "account": rec["account"],
        "category": rec["category"],
    }


def is_active(rec):
    return rec["active"].strip().upper() == "Y"


def due_items(book, today):
    """到期还没记的定期收支。从没记过的只列最近一次，免得一次问出一长串。"""
    out = []
    for rec in book.recurring:
        dates = pending_dates_for_item(book, rec["item"], today)
        out += [rec_view(book, rec, d) for d in dates]
    return sorted(out, key=lambda x: x["date"])


def upcoming(book, start, end):
    """排程上的到期日。账户期初日之前的期数不列：那笔钱已经算在期初余额里，不在账内。"""
    out = []
    for rec in book.recurring:
        if not is_active(rec):
            continue
        acc = book.accounts.get(rec["account"])
        since = parse_date(acc["opening_date"]) if acc else start
        out += [rec_view(book, rec, d) for d in occurrences(rec, max(start, since), end)]
    return sorted(out, key=lambda x: x["date"])


ADD_REQUIRED = ("date", "amount", "category", "item", "raw_text")
COUNT_MILESTONES = (100, 500, 1000, 2000, 5000)
DAY_MILESTONES = ((30, "记账满一个月"), (60, "记账满两个月"), (100, "记账满 100 天"),
                  (200, "记账满 200 天"), (365, "记账满一年"))


def row_view(book, r):
    return {**r, "currency": book.currency(r["account"])}


def sort_key(r):
    return (r["date"], r["time"], int(r["id"]) if r["id"].isdigit() else 0)


def next_id(book):
    """删掉的 id 不复用：deleted.csv 里的也算进去，免得找回时和新记的撞号。"""
    rows = list(book.ledger)
    if (book.dir / "deleted.csv").exists():
        rows += read_table(book.dir / "deleted.csv", ["id"])[1]
    return max((int(r["id"]) for r in rows if r["id"].strip().isdigit()), default=0) + 1


def find_row(book, rid):
    for r in book.ledger:
        if r["id"] == str(rid):
            return r
    raise WalletError(f"找不到 id {rid}")


def touched_accounts(*rows):
    return sorted({r["account"] or UNASSIGNED for r in rows})


def transfer_partners(book, rows):
    """和这些行同属一组转账（同 date、同 raw_text）、现在还在账本里的行的 id。"""
    keys = {(r["date"], r["raw_text"]) for r in rows if r["category"] in TRANSFER}
    return {r["id"] for r in book.ledger if r["category"] in TRANSFER and (r["date"], r["raw_text"]) in keys}


def milestones(old_rows, new_rows):
    out = []
    before, after = len(old_rows), len(old_rows) + len(new_rows)
    out += [f"第 {n} 笔" for n in COUNT_MILESTONES if before < n <= after]
    if old_rows:
        dates = [parse_date(r["date"]) for r in old_rows]
        first, last_before = min(dates), max(dates)
        last_after = max([last_before] + [parse_date(r["date"]) for r in new_rows])
        out += [label for n, label in DAY_MILESTONES
                if (last_before - first).days < n <= (last_after - first).days]
    return out


def add_rows(book, new_rows, today):
    """写入一笔或多笔。任何一笔有硬错误就一笔都不写；正负号这类疑问写入后放进 warnings。"""
    if not new_rows:
        raise WalletError("没有要写入的记录")
    allowed = set(book.ledger_header) - {"id"}
    nid = next_id(book)
    prepared, errors = [], []
    for i, raw in enumerate(new_rows, 1):
        if not isinstance(raw, dict):
            errors.append(f"第 {i} 笔不是 JSON 对象（或合法的字典）")
            continue
        unknown = set(raw) - allowed
        if unknown:
            errors.append(f"第 {i} 笔有不认识的字段：{'、'.join(sorted(unknown))}")
            continue
        raw = {k: "" if v is None else v for k, v in raw.items()}  # JSON null 当空串
        missing = [k for k in ADD_REQUIRED if not str(raw.get(k, "")).strip()]
        if missing:
            errors.append(f"第 {i} 笔缺：{'、'.join(missing)}")
            continue
        r = {col: "" for col in book.ledger_header}
        r.update({k: str(v) if k == "raw_text" else str(v).strip() for k, v in raw.items()})
        r["id"] = str(nid)
        nid += 1
        e, _ = row_problems(book, r)
        errors += e
        if not e:
            d = money(r["amount"])
            r["amount"] = format(d, "f")
        prepared.append(r)
    if errors:
        raise WalletError("没写入任何一笔：\n" + "\n".join(errors))
    old = list(book.ledger)
    book.ledger.extend(prepared)
    book.save_ledger()
    report = check(book, {r["id"] for r in prepared})
    return {
        "added": [row_view(book, r) for r in prepared],
        "balances": balances_view(book, today, touched_accounts(*prepared)),
        "errors": report["errors"],
        "warnings": report["warnings"],
        "milestones": milestones(old, prepared),
        "due": due_items(book, today),
    }


RAW_SEP = " | "


def edit_row(book, rid, changes, today, append_raw=None):
    """append_raw：用户后来补的话（澄清、改口），由脚本接在原话后面，Agent 不用重写整串原话。"""
    r = find_row(book, rid)
    bad = set(changes) - (set(book.ledger_header) - {"id"})
    if bad:
        raise WalletError(f"不能改的字段：{'、'.join(sorted(bad))}")
    if append_raw is not None and "raw_text" in changes:
        raise WalletError("--append-raw 和 --set raw_text= 只能用一个")
    if not changes and not append_raw:
        raise WalletError("没有要改的：用 --set 字段=值 或 --append-raw")
    before = dict(r)
    changes = {k: "" if v is None else v for k, v in changes.items()}  # None 当空串
    if append_raw:
        changes["raw_text"] = r["raw_text"] + RAW_SEP + append_raw
    r.update({k: v if k == "raw_text" else v.strip() for k, v in changes.items()})
    errors, _ = row_problems(book, r)
    if errors:
        r.clear()
        r.update(before)
        raise WalletError("没改：\n" + "\n".join(errors))
    d = money(r["amount"])
    r["amount"] = format(d, "f")
    book.save_ledger()
    ids = {r["id"]} | transfer_partners(book, [before, r])
    return {
        "before": row_view(book, before),
        "after": row_view(book, r),
        "balances": balances_view(book, today, touched_accounts(before, r)),
        "warnings": check(book, ids)["warnings"],
    }


def keep_deleted(book, rows, today):
    """删掉的行原样追加进 deleted.csv，删错了能抄回来。只有这里写它，别的命令都不读。"""
    path = book.dir / "deleted.csv"
    header, kept = (read_table(path, LEDGER_COLS) if path.exists() else (book.ledger_header + ["deleted_at"], []))
    header += [c for c in book.ledger_header + ["deleted_at"] if c not in header]
    kept = [{**{c: "" for c in header}, **k} for k in kept]
    kept += [{**{c: "" for c in header}, **r, "deleted_at": today.isoformat()} for r in rows]
    write_table(path, header, kept)


def delete_rows(book, ids, today):
    rows = [find_row(book, i) for i in ids]
    gone = {r["id"] for r in rows}
    keep_deleted(book, rows, today)  # 先留底再删：留底失败就什么都不删
    book.ledger = [r for r in book.ledger if r["id"] not in gone]
    book.save_ledger()
    partners = transfer_partners(book, rows)
    return {
        "kept_in": "deleted.csv",
        "deleted": [row_view(book, r) for r in rows],
        "balances": balances_view(book, today, [a for a in touched_accounts(*rows) if a in account_list(book)]),
        "warnings": check(book, partners)["warnings"] if partners else [],
    }


def calibrate(book, account, target, raw_text, day):
    if account not in book.accounts:
        raise WalletError(f"账户「{account}」不在 accounts.csv 里")
    opening_day = parse_date(book.accounts[account]["opening_date"])
    if day < opening_day:
        raise WalletError(
            f"校准日期 {day.isoformat()} 早于账户「{account}」的期初日期 {opening_day.isoformat()}："
            "期初余额已经包含了那之前的所有流水，不用也不能在那之前校准")
    current = balance_at(book, account, day)
    diff = money(target) - current
    out = {"account": account, "currency": book.currency(account),
           "before": num(current), "target": num(money(target)), "diff": num(diff), "added": None}
    if diff != 0:
        out.update(add_rows(book, [{"date": day.isoformat(), "amount": str(diff), "account": account,
                                    "category": "校准", "item": "余额校准", "raw_text": raw_text}], day))
    return out


def pending_dates_for_item(book, item, today):
    """计算某一定期项的待处理期数（自 last_logged 之后到当前日期的到期日）。"""
    recs = [r for r in book.recurring if r["item"] == item]
    if len(recs) != 1:
        return []
    rec = recs[0]
    if not is_active(rec):
        return []
    if rec["last_logged"].strip():
        return occurrences(rec, parse_date(rec["last_logged"]) + dt.timedelta(days=1), today)
    else:
        return occurrences(rec, today - dt.timedelta(days=366), today)[-1:]


def _require_earliest_pending(book, rec, day, today):
    """检查某定期项的给定日期是否是最早待处理期数。若不是，抛 WalletError。"""
    pending = pending_dates_for_item(book, rec["item"], today)
    if not pending:
        raise WalletError(f"定期项「{rec['item']}」没有待处理的期数")
    earliest = min(pending)
    if day != earliest:
        raise WalletError(f"定期项「{rec['item']}」先处理更早的期数：{earliest.isoformat()}")


def log_recurring(book, item, day, today):
    recs = [r for r in book.recurring if r["item"] == item]
    if len(recs) != 1:
        raise WalletError(f"recurring.csv 里找不到唯一一项「{item}」")
    rec = recs[0]
    _require_earliest_pending(book, rec, day, today)
    result = add_rows(book, [{
        "date": day.isoformat(), "amount": rec["amount"], "account": rec["account"],
        "category": rec["category"], "item": rec["item"],
        "raw_text": f"定期收支「{item}」{day.isoformat()} 到期，用户确认后记入",
    }], today)
    new_id = result["added"][0]["id"]
    rec["last_logged"] = day.isoformat()
    try:
        book.save_recurring()
    except WalletError as e:
        raise WalletError(
            f"账本行 id {new_id} 已经写入，但 recurring.csv 没更新成功（last_logged 没变）：{e}\n"
            f"不要重新 log-recurring（会重复记账）：手动把「{item}」的 last_logged 改成 {day.isoformat()}"
        ) from e
    result["due"] = due_items(book, today)
    return result


def skip_recurring(book, item, day, today):
    """跳过定期项的一期，不记账本行。"""
    recs = [r for r in book.recurring if r["item"] == item]
    if len(recs) != 1:
        raise WalletError(f"recurring.csv 里找不到唯一一项「{item}」")
    rec = recs[0]
    _require_earliest_pending(book, rec, day, today)
    rec["last_logged"] = day.isoformat()
    book.save_recurring()
    return {
        "skipped": {"item": item, "date": day.isoformat()},
        "due": due_items(book, today),
    }


def prev_period(start, end):
    """上期：整月 → 上一个整月；月初到某天 → 上月同期；周一到周日 → 上一个整周；其他 → 紧挨着的等长窗口。
    由脚本按日期推，不交给 LLM：「回顾九月」要比八月，不是比 08-02~08-31。"""
    if start.day == 1 and (start.year, start.month) == (end.year, end.month):
        py, pm = (start.year - 1, 12) if start.month == 1 else (start.year, start.month - 1)
        last = calendar.monthrange(py, pm)[1]
        if end.day == calendar.monthrange(end.year, end.month)[1]:
            return dt.date(py, pm, 1), dt.date(py, pm, last), "month"
        return dt.date(py, pm, 1), dt.date(py, pm, min(end.day, last)), "month_to_date"
    days = (end - start).days + 1
    prev_end = start - dt.timedelta(days=1)
    kind = "week" if days == 7 and start.weekday() == 0 else "window"
    return prev_end - dt.timedelta(days=days - 1), prev_end, kind


def tracked_since(book):
    """从哪天开始记账：最早的期初日或最早的流水（补记的更早流水也算）。"""
    starts = [parse_date(a["opening_date"]) for a in book.accounts_rows] + [parse_date(r["date"]) for r in book.ledger]
    return min(starts) if starts else None


def row_kind(book, category):
    """一行算支出、收入、转账还是校准。页面只认这个标签，不自己判断类别方向。"""
    if category in TRANSFER:
        return "transfer"
    if category == "校准":
        return "calibration"
    return "expense" if book.direction(category) == "支出" else "income"


def tagged_row(book, r, recurring_items):
    """给页面和编剧的流水行：金额转成数，打上 kind、「是不是定期项」（item 和 recurring.csv 同名，停用的也算）和 weekday（0 = 周一）。"""
    return {**r, "amount": num(money(r["amount"])), "currency": book.currency(r["account"]),
            "kind": row_kind(book, r["category"]), "recurring": r["item"] in recurring_items,
            "weekday": parse_date(r["date"]).weekday()}


def rows_between(book, start, end):
    return [r for r in book.ledger if start <= parse_date(r["date"]) <= end]


def summarize(book, rows):
    """支出、收入按类别方向算（退款冲减支出），转账不计，校准单列。"""
    acc = defaultdict(lambda: {"expense": Decimal(0), "income": Decimal(0), "calibration": Decimal(0)})
    for r in rows:
        cur, amt, cat = book.currency(r["account"]), money(r["amount"]), r["category"]
        if cat in TRANSFER:
            continue
        if cat == "校准":
            acc[cur]["calibration"] += amt
        elif book.direction(cat) == "支出":
            acc[cur]["expense"] -= amt
        else:
            acc[cur]["income"] += amt
    return {
        cur: {"expense": num(v["expense"]), "income": num(v["income"]),
              "net": num(v["income"] - v["expense"]), "calibration": num(v["calibration"])}
        for cur, v in sorted(acc.items())
    }


def category_totals(book, rows, prev_rows):
    def tally(rs):
        amounts, counts = defaultdict(Decimal), defaultdict(int)
        for r in rs:
            if r["category"] in BUILTIN_CATEGORIES:
                continue
            key = (r["category"], book.currency(r["account"]))
            amounts[key] += money(r["amount"])
            counts[key] += 1
        return amounts, counts

    now, counts = tally(rows)
    before, _ = tally(prev_rows)
    keys = sorted(set(now) | set(before), key=lambda k: (-abs(now.get(k, Decimal(0))), k[0], k[1]))
    return [
        {"category": c, "direction": book.direction(c), "currency": cur,
         "amount": num(now.get((c, cur), Decimal(0))), "count": counts.get((c, cur), 0),
         "prev_amount": num(before.get((c, cur), Decimal(0)))}
        for c, cur in keys
    ]


def spending_patterns(book, rows):
    """支出按星期、按钟点分布。只算用户自己的支出类别；没写时间的单独合计，并数一下几笔写了时间。
    定期项不算：房租落在哪天是排程定的，不是习惯，算进来会撑出假峰（「周二花得最多」其实是 1 号交房租）。"""
    recurring_items = {r["item"] for r in book.recurring}
    weekday = defaultdict(lambda: [Decimal(0)] * 7)
    hour = defaultdict(lambda: [Decimal(0)] * 24)
    untimed = defaultdict(Decimal)
    timed = defaultdict(lambda: {"rows": 0, "timed": 0})
    for r in rows:
        if row_kind(book, r["category"]) != "expense" or r["item"] in recurring_items:
            continue
        cur, spent = book.currency(r["account"]), -money(r["amount"])
        weekday[cur][parse_date(r["date"]).weekday()] += spent
        timed[cur]["rows"] += 1
        if r["time"].strip():
            hour[cur][int(r["time"].strip()[:2])] += spent
            timed[cur]["timed"] += 1
        else:
            untimed[cur] += spent
    return weekday, hour, untimed, timed


NOTE_KEYS = ("from", "to", "generated", "title", "opening", "closing")


def recent_chapters(book, start, limit=3):
    """最近几章的记录（build 时写在章节旁边的 .json），新的在前：只取结束日早于本期起点的，同一范围生成过几次只算最后一次。
    第一章是「上一章」，编剧接它收尾点名的事；几章一起给，是为了这次的标题、开场白、收尾别和最近几次撞句式。
    记录坏了就跳过，不为它报错；旧记录没有 title / opening，给 null。"""
    latest = {}
    for p in (book.dir / "chapters").glob("*.json"):
        try:
            note = json.loads(p.read_text(encoding="utf-8"))
            frm, to = parse_date(note["from"]), parse_date(note["to"])
        except (OSError, ValueError, KeyError, TypeError, WalletError):
            continue
        if to >= start:
            continue
        n = int(p.stem.rsplit("_", 1)[1]) if p.stem.count("_") == 2 else 1   # 2025-03-01_2025-03-31_2 → 2
        key = (str(note.get("generated", "")), n)
        if (frm, to) not in latest or key > latest[(frm, to)][0]:
            latest[(frm, to)] = (key, {k: note.get(k) for k in NOTE_KEYS})
    ordered = sorted(latest.items(), key=lambda kv: (kv[0][1], kv[1][0]), reverse=True)
    return [v[1] for _, v in ordered[:limit]]


def stats(book, start, end, include_rows=False, today=None):
    """today 只影响 upcoming_30d：回顾已经过完的时段时，从今天往后列，不把已经过去的日子算成「接下来」。"""
    if end < start:
        raise WalletError("--from 不能晚于 --to")
    days = (end - start).days + 1
    prev_start, prev_end, kind = prev_period(start, end)
    rows, prev_rows = rows_between(book, start, end), rows_between(book, prev_start, prev_end)
    before = start - dt.timedelta(days=1)  # 起点余额取本期前一天：上期不一定紧挨着本期
    accounts, change = [], defaultdict(Decimal)
    for name in account_list(book):
        s, e = balance_at(book, name, before), balance_at(book, name, end)
        cur = book.currency(name)
        change[cur] += e - s
        # 没写账户的流水从 0 起算，起止余额没有意义，只给净变化
        shown = (None, None) if name == UNASSIGNED else (num(s), num(e))
        accounts.append({"account": name, "currency": cur, "start": shown[0], "end": shown[1], "change": num(e - s)})
    totals, prev_totals = summarize(book, rows), summarize(book, prev_rows)
    weekday, hour, untimed, timed = spending_patterns(book, rows)
    # 所有按币种分的对象用同一组键：页面和编剧不用判断「这个币种上期有没有」
    curs = sorted(set(change) | set(totals) | set(prev_totals) | set(weekday))
    zero = {"expense": 0, "income": 0, "net": 0, "calibration": 0}
    since = tracked_since(book)
    covered = max(0, (end - max(start, since)).days + 1) if since else 0
    ahead = max(end + dt.timedelta(days=1), today or start)
    out = {
        "range": {"from": start.isoformat(), "to": end.isoformat(), "days": days},
        "coverage": {"tracked_since": since.isoformat() if since else None, "days": covered},
        "prev_range": {"from": prev_start.isoformat(), "to": prev_end.isoformat(), "kind": kind},
        "row_count": len(rows),
        "accounts": accounts,
        "wallet_change": {c: num(change[c]) for c in curs},
        "totals": {c: totals.get(c, dict(zero)) for c in curs},
        "prev_totals": {c: prev_totals.get(c, dict(zero)) for c in curs},
        "categories": category_totals(book, rows, prev_rows),
        "weekday_expense": {c: [num(x) for x in weekday[c]] for c in curs},
        "hour_expense": {c: [num(x) for x in hour[c]] for c in curs},
        "hour_coverage": {c: dict(timed[c]) for c in curs},
        "untimed_expense": {c: num(untimed[c]) for c in curs},
        "recurring_in_range": upcoming(book, start, end),
        "upcoming_30d": upcoming(book, ahead, ahead + dt.timedelta(days=29)),
        "recent_chapters": recent_chapters(book, start),
    }
    if include_rows:
        items = {r["item"] for r in book.recurring}
        out["rows"] = [tagged_row(book, r, items) for r in sorted(rows, key=sort_key)]
    return out


def daily_balance(book, start, end):
    """每个账户每天结束时的余额。先算 start 前一天的余额，再逐日加当天的流水；
    结果和逐日调 balance_at 一致（期初日前后都成立），但只扫一遍流水。"""
    n = (end - start).days + 1
    out = []
    for name in account_list(book):
        key = "" if name == UNASSIGNED else name
        delta = defaultdict(Decimal)
        for r in book.ledger:
            if r["account"] == key:
                delta[parse_date(r["date"])] += money(r["amount"])
        bal, values = balance_at(book, name, start - dt.timedelta(days=1)), []
        for i in range(n):
            bal += delta[start + dt.timedelta(days=i)]
            values.append(num(bal))
        out.append({"account": name, "currency": book.currency(name), "values": values})
    return {"from": start.isoformat(), "to": end.isoformat(), "accounts": out}


def chapter_payload(book, start, end, today=None):
    """一章要注入页面的数据：本期 stats、全部流水（打好标签，看板可以选任何日期）、按账户的每日余额。
    每日余额从记账第一天的前一天画起：那天是「还没有任何流水」时的余额，任意日期段的起点都查得到。"""
    first = min(tracked_since(book) or start, start) - dt.timedelta(days=1)
    last = max([end] + [parse_date(r["date"]) for r in book.ledger])
    items = {r["item"] for r in book.recurring}
    return {
        "stats": stats(book, start, end, today=today),
        "ledger": [tagged_row(book, r, items) for r in sorted(book.ledger, key=sort_key)],
        "daily_balance": daily_balance(book, first, last),
        "default_currency": book.settings["default_currency"],
    }


DATA_SLOT = "/*__WALLET_DATA__*/null"
TECHNIQUES = ("转折", "人物", "场景", "反差")
COUNT_PHRASE = re.compile(r"(第\s*)?(?<![\d.])(\d+)\s*(?:次|顿|笔|回|趟|杯|单)")   # 带「第」的是序数，命中了也跳过


def script_problems(book, script, start, end):
    """剧本的硬伤：会让页面坏掉或说谎的，逐条报。措辞好不好不归这里管。"""
    if not isinstance(script, dict):
        return ["剧本要是一个 JSON 对象"]
    filled = lambda s: isinstance(s, str) and s.strip()
    out = [f"缺「{k}」或不是文字" for k in ("title", "closing") if not filled(script.get(k))]
    op = script.get("opening")
    if not (filled(op) or isinstance(op, list) and 1 <= len(op) <= 4 and all(map(filled, op))):
        out.insert(1, "opening 要是一段话或 1–4 条短句")   # 封面上：一段话，或一条一行的短句
    findings = script.get("findings", [])
    if not isinstance(findings, list):
        return out + ["findings 要是数组"]
    if len(findings) > 3:
        out.append(f"发现最多 3 个，现在有 {len(findings)} 个")
    ids = {r["id"] for r in book.ledger}
    for n, f in enumerate(findings, 1):
        if not isinstance(f, dict):
            out.append(f"第 {n} 个发现不是对象")
            continue
        for k in ("title", "claim"):
            if not isinstance(f.get(k), str) or not f[k].strip():
                out.append(f"第 {n} 个发现缺「{k}」")
        if f.get("technique") not in TECHNIQUES:
            out.append(f"第 {n} 个发现的 technique 要是 {' / '.join(TECHNIQUES)} 之一")
        fids = f.get("ids")
        if not isinstance(fids, list) or not fids:
            out.append(f"第 {n} 个发现没有 ids：每个发现都要有撑得住它的那几笔账")
        else:
            missing = [str(i) for i in fids if str(i) not in ids]
            if missing:
                out.append(f"第 {n} 个发现引用了不存在的 id：{', '.join(missing)}")
            # 发现里说的次数只能是 ids 里除去退款的笔数：退款（支出类别的正数行）放进 ids 是为了合计对得上，不算一次。
            # 只核阿拉伯数字；「十几顿」「好几次」不核，「第 2 次」是序数不核。量词只收数「笔」的：「6 天」是天数，不在里面
            by_id = {r["id"]: r for r in book.ledger}
            times = sum(1 for i in fids if str(i) in by_id and not (
                row_kind(book, by_id[str(i)]["category"]) == "expense" and money(by_id[str(i)]["amount"]) > 0))
            for k in ("claim", "body"):
                for m in COUNT_PHRASE.finditer(f.get(k) or ""):
                    if not m.group(1) and int(m.group(2)) != times:
                        out.append(f"第 {n} 个发现的 {k} 写的是「{m.group(0)}」这个次数，ids 里是 {times} 笔（退款不算）：要么改次数，要么改 ids")
    return out


def build(book, start, end, script, template, today):
    problems = script_problems(book, script, start, end)
    if problems:
        raise WalletError("剧本有问题，改好再 build：\n" + "\n".join(problems))
    for f in script.get("findings", []):
        f["ids"] = [str(i) for i in f["ids"]]
    try:
        html = Path(template).read_text(encoding="utf-8")
    except OSError as e:
        raise WalletError(f"读不到模板 {template}：{e}")
    if html.count(DATA_SLOT) != 1:
        raise WalletError("template.html 里的数据插槽不是正好一个，模板坏了")
    payload = {**chapter_payload(book, start, end, today), "script": script, "generated": today.isoformat()}
    # 原话里可能有 </script>：把 < 全部写成 \u003c，JSON 读回来不变，页面也不会被提前截断
    data = json.dumps(payload, ensure_ascii=False).replace("<", "\\u003c")
    folder = book.dir / "chapters"
    folder.mkdir(exist_ok=True)
    stem = f"{start.isoformat()}_{end.isoformat()}"
    path, n = folder / f"{stem}.html", 2
    while path.exists():  # 章节是快照：同一范围再生成一次不覆盖
        path, n = folder / f"{stem}_{n}.html", n + 1
    path.write_bytes(html.replace(DATA_SLOT, data).encode("utf-8"))
    # 旁边留一份小记录：之后的编剧从 stats 的 recent_chapters 读到这章写了什么、收尾点名了什么
    note = {"from": start.isoformat(), "to": end.isoformat(), "generated": today.isoformat(),
            "title": script["title"], "opening": script["opening"], "closing": script["closing"]}
    path.with_suffix(".json").write_bytes(json.dumps(note, ensure_ascii=False, indent=2).encode("utf-8"))
    return {"chapter": str(path), "findings": len(script.get("findings", [])), "ledger_rows": len(payload["ledger"])}


def opt_date(s):
    return parse_date(s) if s else dt.date.today()


def split_ids(s):
    return [x.strip() for x in s.split(",") if x.strip()]


def load_clean(args):
    """坏数据上算出的数不可信：全表有硬错误就拒绝执行，让用户先修。"""
    book = Book(resolve_data_dir(args.data))
    errors = check(book)["errors"]
    if errors:
        raise WalletError("账本有错误，先修好再算：\n" + "\n".join(errors))
    return book


def cmd_init(a):
    if not a.data:
        raise WalletError("init 要用 --data 指定数据文件夹")
    return init_data(a.data, a.currency), 0


def cmd_balance(a):
    book, day = load_clean(a), opt_date(a.date)
    if a.account and a.account not in book.accounts:
        raise WalletError(f"账户「{a.account}」不在 accounts.csv 里")
    names = [a.account] if a.account else account_list(book)
    by_currency = defaultdict(Decimal)
    for n in names:
        if n != UNASSIGNED:  # 未指定账户从 0 累加，不是余额，不进币种合计（自己那行在 accounts 里）
            by_currency[book.currency(n)] += balance_at(book, n, day)
    return {"date": day.isoformat(), "accounts": balances_view(book, day, names),
            "by_currency": {c: num(v) for c, v in sorted(by_currency.items())}}, 0


def cmd_add(a):
    book = load_clean(a)
    if a.file:
        try:
            text = Path(a.file).read_text(encoding="utf-8-sig")
        except FileNotFoundError:
            raise WalletError(f"读不到文件：{a.file}")
        except UnicodeDecodeError:
            raise WalletError(
                f"{a.file} 不是 UTF-8 编码：PowerShell 的 Out-File 和 > 默认写 UTF-16。"
                "请用文件写入工具，或加 -Encoding utf8 再试")
        except OSError as e:
            raise WalletError(f"读不到文件 {a.file}：{e}")
    else:
        text = sys.stdin.read()
    try:
        rows = json.loads(text)
    except json.JSONDecodeError as e:
        raise WalletError(f"add 读到的不是合法 JSON：{e}")
    result = add_rows(book, rows if isinstance(rows, list) else [rows], dt.date.today())
    if a.file:
        # 写入成功才删：旧批次不可能被重放；失败时留着，方便改完重试
        try:
            Path(a.file).unlink()
            result["input_file_removed"] = True
        except OSError:
            result["input_file_removed"] = False
    return result, 0


def cmd_edit(a):
    changes = {}
    for pair in a.set or []:
        if "=" not in pair:
            raise WalletError(f"--set 要写成 字段=值：{pair}")
        k, v = pair.split("=", 1)
        changes[k.strip()] = v
    return edit_row(load_clean(a), a.id, changes, dt.date.today(), a.append_raw), 0


def cmd_delete(a):
    return delete_rows(load_clean(a), split_ids(a.ids), dt.date.today()), 0


def cmd_calibrate(a):
    return calibrate(load_clean(a), a.account, a.balance, a.raw_text, opt_date(a.date)), 0


def cmd_check(a):
    book = Book(resolve_data_dir(a.data))
    report = check(book, set(split_ids(a.ids)) if a.ids else None)
    return report, 1 if report["errors"] else 0


def cmd_due(a):
    return {"due": due_items(load_clean(a), opt_date(a.date))}, 0


def cmd_log_recurring(a):
    return log_recurring(load_clean(a), a.item, parse_date(a.date), dt.date.today()), 0


def cmd_skip_recurring(a):
    return skip_recurring(load_clean(a), a.item, parse_date(a.date), dt.date.today()), 0


def cmd_stats(a):
    end = opt_date(a.to)
    start = parse_date(a.frm) if a.frm else end - dt.timedelta(days=29)
    return stats(load_clean(a), start, end, a.rows, today=dt.date.today()), 0


def cmd_build(a):
    end = opt_date(a.to)
    start = parse_date(a.frm) if a.frm else end - dt.timedelta(days=29)
    try:
        script = json.loads(Path(a.script).read_text(encoding="utf-8-sig"))
    except FileNotFoundError:
        raise WalletError(f"读不到剧本文件：{a.script}")
    except UnicodeDecodeError:
        raise WalletError(f"{a.script} 不是 UTF-8 编码：请用文件写入工具，或加 -Encoding utf8 再试")
    except json.JSONDecodeError as e:
        raise WalletError(f"剧本不是合法 JSON：{e}")
    result = build(load_clean(a), start, end, script, Path(__file__).with_name("template.html"), dt.date.today())
    if a.open:
        result["opened"] = webbrowser.open(Path(result["chapter"]).resolve().as_uri())
    return result, 0


def build_parser():
    p = argparse.ArgumentParser(prog="wallet.py", description="钱包日记的算数脚本，输出 JSON")
    p.add_argument("--data", help="数据文件夹（不写就读 ~/.wallet-diary.json）")
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("init", help="建数据文件夹")
    s.add_argument("--currency", default="CNY")
    s.set_defaults(func=cmd_init)

    s = sub.add_parser("balance", help="余额")
    s.add_argument("--account")
    s.add_argument("--date")
    s.set_defaults(func=cmd_balance)

    s = sub.add_parser("add", help="记一笔或多笔（JSON，从 stdin 或 --file 读）")
    s.add_argument("--file", help="JSON 文件；写入成功后脚本会删掉它")
    s.set_defaults(func=cmd_add)

    s = sub.add_parser("edit", help="按 id 改一笔")
    s.add_argument("--id", required=True)
    s.add_argument("--set", action="append", metavar="字段=值")
    s.add_argument("--append-raw", help="用户补的话，用 ' | ' 接在原话后面")
    s.set_defaults(func=cmd_edit)

    s = sub.add_parser("delete", help="按 id 删（删掉的行留底在 deleted.csv）")
    s.add_argument("--ids", required=True)
    s.set_defaults(func=cmd_delete)

    s = sub.add_parser("calibrate", help="按用户报的余额校准")
    s.add_argument("--account", required=True)
    s.add_argument("--balance", required=True)
    s.add_argument("--raw-text", required=True)
    s.add_argument("--date")
    s.set_defaults(func=cmd_calibrate)

    s = sub.add_parser("check", help="体检")
    s.add_argument("--ids")
    s.set_defaults(func=cmd_check)

    s = sub.add_parser("due", help="到期没记的定期收支")
    s.add_argument("--date")
    s.set_defaults(func=cmd_due)

    s = sub.add_parser("log-recurring", help="用户确认后记一笔定期收支")
    s.add_argument("--item", required=True)
    s.add_argument("--date", required=True)
    s.set_defaults(func=cmd_log_recurring)

    s = sub.add_parser("skip-recurring", help="这一期没扣 / 不记，跳过")
    s.add_argument("--item", required=True)
    s.add_argument("--date", required=True)
    s.set_defaults(func=cmd_skip_recurring)

    s = sub.add_parser("stats", help="一段时间的汇总")
    s.add_argument("--from", dest="frm")
    s.add_argument("--to")
    s.add_argument("--rows", action="store_true", help="带上范围内的明细")
    s.set_defaults(func=cmd_stats)

    s = sub.add_parser("build", help="生成一章日记（HTML），写进数据文件夹的 chapters/")
    s.add_argument("--from", dest="frm")
    s.add_argument("--to")
    s.add_argument("--script", required=True, help="剧本 JSON 文件")
    s.add_argument("--open", action="store_true", help="生成后用默认浏览器打开")
    s.set_defaults(func=cmd_build)
    return p


def main(argv=None):
    # 中文 Windows 的控制台默认不是 utf-8：三条流都要改，只改 stdout 的话报错会乱码
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        stream.reconfigure(encoding="utf-8")
    args = build_parser().parse_args(argv)
    try:
        result, code = args.func(args)
    except WalletError as e:
        print(f"错误：{e}", file=sys.stderr)
        return 1
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return code


if __name__ == "__main__":
    sys.exit(main())
