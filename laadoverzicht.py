#!/usr/bin/env python3
"""Maandelijks Zaptec-laadoverzicht (xlsx) maken en mailen naar de werkgever."""
import argparse
import logging
import os
import smtplib
import sys
from datetime import date, datetime, timedelta, timezone
from decimal import ROUND_HALF_UP, Decimal
from email.message import EmailMessage
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

BASE_DIR = Path(__file__).resolve().parent
TOKEN_URL = "https://api.zaptec.com/oauth/token"
HISTORY_URL = "https://api.zaptec.com/api/chargehistory"
PRICES_URL = "https://api.energyzero.nl/v1/energyprices"
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
TZ = ZoneInfo("Europe/Amsterdam")
PAGE_SIZE = 100
TIMEOUT = 30
MONTHS_NL = [
    "januari", "februari", "maart", "april", "mei", "juni", "juli",
    "augustus", "september", "oktober", "november", "december",
]
XLSX_MIME = ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")

log = logging.getLogger("laadoverzicht")


class ConfigError(Exception):
    pass


class PriceError(Exception):
    """Marktprijzen ontbreken of zijn onbruikbaar; de run faalt liever dan te gokken."""


def previous_month(today=None):
    today = today or date.today()
    return (today.year - 1, 12) if today.month == 1 else (today.year, today.month - 1)


def parse_month(value):
    try:
        d = datetime.strptime(value, "%Y-%m")
    except ValueError:
        raise argparse.ArgumentTypeError(f"ongeldige maand '{value}', gebruik YYYY-MM")
    return d.year, d.month


def month_label(year, month):
    return f"{MONTHS_NL[month - 1]} {year}"


def month_bounds(year, month):
    """Begin (inclusief) en einde (exclusief) van de kalendermaand in Europe/Amsterdam, als UTC."""
    start = datetime(year, month, 1, tzinfo=TZ)
    end = datetime(year + (month == 12), month % 12 + 1, 1, tzinfo=TZ)
    return start.astimezone(timezone.utc), end.astimezone(timezone.utc)


def subject_for(owner, year, month):
    return f"declaratie {owner} laadpaal {month_label(year, month)}"


def load_config(env=None):
    if env is None:
        load_dotenv(BASE_DIR / ".env")
        env = os.environ
    get = lambda k, d="": (env.get(k) or d).strip()
    cfg = {
        "zaptec_username": get("ZAPTEC_USERNAME"),
        "zaptec_password": get("ZAPTEC_PASSWORD"),
        "charger_id": get("ZAPTEC_CHARGER_ID"),
        "gmail_user": get("GMAIL_USER"),
        "gmail_password": get("GMAIL_APP_PASSWORD"),
        "mail_to": get("MAIL_TO"),
        "mail_to_name": get("MAIL_TO_NAME"),
        "owner": get("OWNER_NAME"),
        "output_dir": Path(get("OUTPUT_DIR", str(BASE_DIR / "output"))),
    }
    cfg["mail_cc"] = get("MAIL_CC") or cfg["gmail_user"]
    try:
        cfg["rate"] = Decimal(get("RATE", "0.27505").replace(",", "."))
    except Exception:
        raise ConfigError(f"RATE is geen geldig getal: {get('RATE')!r}")
    if cfg["rate"] < 0:
        raise ConfigError("RATE mag niet negatief zijn")
    cfg["rate_mode"] = get("RATE_MODE", "fixed").lower()
    if cfg["rate_mode"] not in ("fixed", "dynamic"):
        raise ConfigError(f"RATE_MODE moet 'fixed' of 'dynamic' zijn: {get('RATE_MODE')!r}")
    # Energiebelasting elektriciteit 2026, schijf 1 (t/m 10.000 kWh): EUR 0,1108481 incl. btw
    # = EUR 0,09161 excl. btw per kWh (bron: Belastingdienst, tarieven 2026). Pas jaarlijks aan.
    for key, name, default in (("dynamic_markup", "DYNAMIC_MARKUP", "0"),
                               ("energy_tax", "ENERGY_TAX", "0.09161"),
                               ("vat", "VAT", "0.21")):
        try:
            cfg[key] = Decimal(get(name, default).replace(",", "."))
        except Exception:
            raise ConfigError(f"{name} is geen geldig getal: {get(name)!r}")
    if cfg["energy_tax"] < 0 or cfg["vat"] < 0:
        raise ConfigError("ENERGY_TAX en VAT mogen niet negatief zijn")
    cfg["log_file"] = Path(get("LOG_FILE")) if get("LOG_FILE") else cfg["output_dir"] / "laadoverzicht.log"
    return cfg


def validate(cfg, dry_run):
    required = ["zaptec_username", "zaptec_password", "owner"]
    if not dry_run:
        required += ["gmail_user", "gmail_password", "mail_to"]
    names = {"zaptec_username": "ZAPTEC_USERNAME", "zaptec_password": "ZAPTEC_PASSWORD", "owner": "OWNER_NAME",
             "gmail_user": "GMAIL_USER", "gmail_password": "GMAIL_APP_PASSWORD", "mail_to": "MAIL_TO"}
    missing = [names[k] for k in required if not cfg[k]]
    if missing:
        raise ConfigError("ontbrekende instellingen: " + ", ".join(missing))


def get_token(cfg):
    r = requests.post(
        TOKEN_URL,
        data={"grant_type": "password", "username": cfg["zaptec_username"],
              "password": cfg["zaptec_password"]},
        timeout=TIMEOUT,
    )
    r.raise_for_status()
    return r.json()["access_token"]


def parse_dt(value):
    """Zaptec geeft UTC-tijden, soms zonder offset; naive waarden zijn UTC."""
    dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt


def parse_details(raw):
    """EnergyDetails als oplopende lijst (tijdstip UTC, cumulatieve kWh); onleesbare punten vervallen."""
    points = []
    for p in raw or []:
        ts = p.get("Timestamp") or p.get("timestamp")
        en = p.get("Energy") if p.get("Energy") is not None else p.get("energy")
        if ts is None or en is None:
            continue
        points.append((parse_dt(ts), Decimal(str(en))))
    return sorted(points, key=lambda p: p[0])


def fetch_sessions(cfg, year, month, token=None, details=False):
    """Alle laadsessies die in de maand zijn gestart, oplopend gesorteerd."""
    start, end = month_bounds(year, month)
    token = token or get_token(cfg)
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    fmt = "%Y-%m-%dT%H:%M:%SZ"
    params = {"From": start.strftime(fmt), "To": end.strftime(fmt), "PageSize": PAGE_SIZE,
              "SortDescending": "false"}
    if details:
        params["DetailLevel"] = 1  # EnergyDetails: cumulatieve kWh-meetpunten per sessie
    if cfg["charger_id"]:
        params["ChargerId"] = cfg["charger_id"]
    sessions, seen, page = [], set(), 0
    while True:
        r = requests.get(HISTORY_URL, headers=headers, timeout=TIMEOUT,
                         params={**params, "PageIndex": page})
        r.raise_for_status()
        data = r.json().get("Data") or []
        for item in data:
            if item.get("Id") in seen:
                continue
            seen.add(item.get("Id"))
            s = parse_dt(item["StartDateTime"])
            if not start <= s < end:
                continue
            e = parse_dt(item["EndDateTime"]) if item.get("EndDateTime") else s
            sessions.append({
                "start": s.astimezone(TZ),
                "end": e.astimezone(TZ),
                "kwh": Decimal(str(item.get("Energy") or 0)),
                "charger": item.get("DeviceName") or item.get("ChargerId") or "",
                "details": parse_details(item.get("EnergyDetails")) if details else [],
            })
        if len(data) < PAGE_SIZE:
            break
        page += 1
    sessions.sort(key=lambda s: s["start"])
    return sessions


def money(value):
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def priced(sessions, rate):
    """Kosten per sessie op centen afgerond; totaal is de som van de getoonde regels."""
    rows = [dict(s, cost=money(s["kwh"] * rate)) for s in sessions]
    return rows, sum((r["kwh"] for r in rows), Decimal(0)), sum((r["cost"] for r in rows), Decimal(0))


HOUR = timedelta(hours=1)


def fetch_prices(sessions, year, month):
    """Day-ahead marktprijzen (EUR/kWh excl. btw) per uur als {begin uur (UTC): prijs}.

    Bron: EnergyZero (https://api.energyzero.nl/v1/energyprices), gratis en zonder API-sleutel.
    De API levert uurprijzen (interval=4); sinds de overgang van EPEX naar kwartierprijzen
    (okt 2025) zijn dat gemiddelden van de vier kwartieren. Kwartierprijzen biedt de API niet.
    """
    start, end = month_bounds(year, month)
    # Een sessie kan na middernacht van de maand doorlopen; haal die uren ook op.
    end = max([end] + [s["end"].astimezone(timezone.utc) for s in sessions]) + HOUR
    fmt = "%Y-%m-%dT%H:%M:%S.000Z"
    r = requests.get(PRICES_URL, timeout=TIMEOUT, params={
        "fromDate": start.strftime(fmt), "tillDate": (end - timedelta(milliseconds=1)).strftime(fmt),
        "interval": 4, "usageType": 1, "inclBtw": "false"})
    r.raise_for_status()
    prices = {}
    for p in r.json().get("Prices") or []:
        slot = parse_dt(p["readingDate"])
        if slot != slot.replace(minute=0, second=0, microsecond=0):
            raise PriceError(f"prijsbron levert geen uurprijzen (tijdstip {p['readingDate']})")
        prices[slot] = Decimal(str(p["price"]))
    return prices


def consumption_blocks(session):
    """Verbruik als [(begin, einde, kWh)]: per Zaptec-meetinterval, anders gelijkmatig over de sessie.

    EnergyDetails zijn cumulatief; het verschil tussen twee punten is het verbruik van dat interval.
    De intervallen worden geschaald naar de sessie-kWh zodat kosten en kWh-kolom kloppen.
    """
    start, end, kwh = session["start"].astimezone(timezone.utc), session["end"].astimezone(timezone.utc), session["kwh"]
    pts = session.get("details") or []
    deltas = [(a[0], b[0], b[1] - a[1]) for a, b in zip(pts, pts[1:])]
    if deltas and all(d[2] >= 0 for d in deltas) and sum(d[2] for d in deltas) > 0:
        total = sum(d[2] for d in deltas)
        return [(a, b, d * kwh / total) for a, b, d in deltas]
    if kwh > 0:
        log.warning("sessie %s: geen bruikbaar verbruik per interval van Zaptec, kWh gelijkmatig verdeeld over de duur",
                    session["start"].strftime("%d-%m-%Y %H:%M"))
    return [(start, end, kwh)]


def block_cost(begin, end, kwh, prices, cfg):
    """Kosten (excl. afronding) van een blok; verdeelt over uurgrenzen naar tijdsduur."""
    if kwh == 0:
        return Decimal(0)
    seconds = (end - begin).total_seconds()
    cost, t = Decimal(0), begin
    while True:
        slot = t.replace(minute=0, second=0, microsecond=0)
        nxt = min(slot + HOUR, end) if seconds > 0 else end
        share = Decimal(str((nxt - t).total_seconds() / seconds)) if seconds > 0 else Decimal(1)
        if slot not in prices:
            raise PriceError(f"geen marktprijs voor {slot.astimezone(TZ):%d-%m-%Y %H:%M} (uur van sessie met "
                             f"{kwh:.3f} kWh); gebruik RATE_MODE=fixed of probeer later opnieuw")
        cost += kwh * share * dynamic_rate(prices[slot], cfg)
        if nxt >= end:
            return cost
        t = nxt


def dynamic_rate(market, cfg):
    """EUR/kWh incl. btw: (marktprijs + opslag + energiebelasting) x (1 + btw)."""
    return (market + cfg["dynamic_markup"] + cfg["energy_tax"]) * (1 + cfg["vat"])


def priced_dynamic(sessions, prices, cfg):
    """Per sessie kosten op centen; tarief is het gemiddelde EUR/kWh (kosten / kWh)."""
    rows = []
    for s in sessions:
        cost = money(sum((block_cost(a, b, k, prices, cfg) for a, b, k in consumption_blocks(s)), Decimal(0)))
        rows.append(dict(s, cost=cost, rate=cost / s["kwh"] if s["kwh"] else Decimal(0)))
    return rows, sum((r["kwh"] for r in rows), Decimal(0)), sum((r["cost"] for r in rows), Decimal(0))


def build_xlsx(cfg, year, month, sessions, path, prices=None):
    dynamic = cfg["rate_mode"] == "dynamic"
    if dynamic:
        rows, total_kwh, total_cost = priced_dynamic(sessions, prices, cfg)
    else:
        rows, total_kwh, total_cost = priced(sessions, cfg["rate"])
    charger = cfg["charger_id"] or ", ".join(sorted({s["charger"] for s in rows if s["charger"]})) or "Zaptec Go2"
    wb = Workbook()
    ws = wb.active
    ws.title = "Laadsessies"
    bold = Font(bold=True)
    ws["A1"] = f"Laadoverzicht {month_label(year, month)}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"], ws["B2"] = "Naam", cfg["owner"]
    ws["A3"], ws["B3"] = "Maand", month_label(year, month)
    ws["A4"], ws["B4"] = "Laadpaal", charger
    for r in (2, 3, 4):
        ws[f"A{r}"].font = bold
    header_row = 6
    for col, name in enumerate(["Datum", "Starttijd", "Eindtijd", "kWh", "Dynamisch tarief (gem. EUR/kWh)" if dynamic else "Tarief (EUR/kWh)", "Kosten"], 1):
        c = ws.cell(header_row, col, name)
        c.font = bold
        c.alignment = Alignment(horizontal="center")
    r = header_row
    for s in rows:
        r += 1
        ws.cell(r, 1, s["start"].date()).number_format = "dd-mm-yyyy"
        ws.cell(r, 2, s["start"].time().replace(microsecond=0, tzinfo=None)).number_format = "hh:mm"
        ws.cell(r, 3, s["end"].time().replace(microsecond=0, tzinfo=None)).number_format = "hh:mm"
        ws.cell(r, 4, float(s["kwh"])).number_format = "0.00"
        ws.cell(r, 5, float(s["rate"] if dynamic else cfg["rate"])).number_format = '"€" 0.00000'
        ws.cell(r, 6, float(s["cost"])).number_format = '"€" #,##0.00'
    r += 1
    ws.cell(r, 1, "Totaal").font = bold
    kwh = ws.cell(r, 4, float(total_kwh))
    kwh.number_format, kwh.font = "0.00", bold
    tot = ws.cell(r, 6, float(total_cost))
    tot.number_format, tot.font = '"€" #,##0.00', bold
    for col, width in zip("ABCDEF", (14, 11, 11, 10, 32 if dynamic else 18, 12)):
        ws.column_dimensions[col].width = width
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return total_kwh, total_cost


def euro_nl(value):
    s = f"{value:,.2f}"
    return "€ " + s.replace(",", "X").replace(".", ",").replace("X", ".")


def send_mail(cfg, year, month, xlsx_path, total_cost, total_kwh, count):
    subject = subject_for(cfg["owner"], year, month)
    msg = EmailMessage()
    msg["From"] = cfg["gmail_user"]
    msg["To"] = cfg["mail_to"]
    msg["Cc"] = cfg["mail_cc"]
    msg["Subject"] = subject
    greeting = f"Beste {cfg['mail_to_name']}," if cfg.get("mail_to_name") else "Beste,"
    msg.set_content(
        f"{greeting}\n\nIn de bijlage het laadoverzicht van {month_label(year, month)}.\n\n"
        f"Maand: {month_label(year, month)}\n"
        f"Aantal laadsessies: {count}\n"
        f"Totaal: {total_kwh:.2f} kWh\n"
        f"Totaalbedrag: {euro_nl(total_cost)}\n\n"
        f"Met vriendelijke groet,\n{cfg['owner']}\n"
    )
    msg.add_attachment(xlsx_path.read_bytes(), maintype=XLSX_MIME[0], subtype=XLSX_MIME[1],
                       filename=xlsx_path.name)
    recipients = [a.strip() for a in f"{cfg['mail_to']},{cfg['mail_cc']}".split(",") if a.strip()]
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=TIMEOUT) as smtp:
        smtp.starttls()
        smtp.login(cfg["gmail_user"], cfg["gmail_password"])
        smtp.send_message(msg, to_addrs=recipients)


def run(cfg, year, month, dry_run=False, token=None):
    validate(cfg, dry_run)
    dynamic = cfg["rate_mode"] == "dynamic"
    sessions = fetch_sessions(cfg, year, month, token, details=dynamic)
    log.info("%d laadsessies in %s", len(sessions), month_label(year, month))
    prices = fetch_prices(sessions, year, month) if dynamic and sessions else None
    path = cfg["output_dir"] / (subject_for(cfg["owner"], year, month) + ".xlsx")
    total_kwh, total_cost = build_xlsx(cfg, year, month, sessions, path, prices)
    log.info("xlsx geschreven: %s (totaal %s)", path, euro_nl(total_cost))
    if dry_run:
        log.info("dry-run: geen mail verstuurd")
    else:
        send_mail(cfg, year, month, path, total_cost, total_kwh, len(sessions))
        log.info("mail verstuurd naar %s (cc %s)", cfg["mail_to"], cfg["mail_cc"])
    return path, total_kwh, total_cost


def setup_logging(log_file):
    handlers = [logging.StreamHandler()]
    try:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))
    except OSError as exc:
        print(f"kan logbestand {log_file} niet openen: {exc}", file=sys.stderr)
    logging.basicConfig(level=logging.INFO, handlers=handlers, force=True,
                        format="%(asctime)s %(levelname)s %(message)s")


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--month", type=parse_month, help="maand als YYYY-MM (standaard: vorige maand)")
    p.add_argument("--dry-run", action="store_true", help="maak de xlsx maar verstuur geen mail")
    args = p.parse_args(argv)
    try:
        cfg = load_config()
        setup_logging(cfg["log_file"])
        year, month = args.month or previous_month()
        run(cfg, year, month, args.dry_run)
    except ConfigError as exc:
        print(f"configuratiefout: {exc}", file=sys.stderr)
        return 2
    except Exception:
        log.exception("laadoverzicht mislukt")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
