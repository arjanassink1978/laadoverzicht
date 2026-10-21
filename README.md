# laadoverzicht

Maandelijks Zaptec Go2 laadoverzicht (xlsx) automatisch mailen naar de werkgever. Draait op de Raspberry Pi.

Op de 1e van de maand haalt het script via de Zaptec Cloud API alle laadsessies van de vorige kalendermaand op,
zet ze in een xlsx (datum, start, eind, kWh, tarief, kosten, totaalregel) en mailt die via Gmail naar de werkgever,
met een cc naar jezelf.

- Onderwerp: `declaratie <OWNER_NAME> laadpaal <maand jaar>`, bijv. `declaratie Arjan laadpaal oktober 2026`
- Bijlage en opgeslagen bestand: onderwerp + `.xlsx`, bijv. `declaratie Arjan laadpaal oktober 2026.xlsx`
- Eén tarief (`RATE`, standaard EUR 0,27505/kWh incl. btw) voor alle sessies; geen dal-/nachttarief
- Tijden staan in Europe/Amsterdam; de maandgrenzen ook

## Installatie op de Pi

```bash
sudo mkdir -p /opt/laadoverzicht /mnt/docker-data/laadoverzicht
sudo chown -R $USER /opt/laadoverzicht /mnt/docker-data/laadoverzicht
git clone <repo-url> /opt/laadoverzicht
cd /opt/laadoverzicht
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
cp .env.example .env
chmod 600 .env
nano .env        # vul de gegevens in, zie .env.example
```

Zet `OUTPUT_DIR` en `LOG_FILE` op de SSD (`/mnt/docker-data/laadoverzicht/...`), zodat de SD-kaart niet slijt.

Test eerst zonder te mailen (maakt wel de xlsx):

```bash
.venv/bin/python laadoverzicht.py --month 2026-10 --dry-run
```

Daarna een echte run voor een maand naar keuze:

```bash
.venv/bin/python laadoverzicht.py --month 2026-10
```

Opties: `--month YYYY-MM` (standaard vorige maand), `--dry-run` (xlsx maken, niet mailen).

## Gmail app-wachtwoord aanmaken

1. Zet 2-stapsverificatie aan op https://myaccount.google.com/security.
2. Ga naar https://myaccount.google.com/apppasswords (zoek anders naar "App-wachtwoorden" in je Google-account).
3. Geef een naam op, bijv. `laadoverzicht`, en kies **Maken**.
4. Kopieer het wachtwoord van 16 tekens naar `GMAIL_APP_PASSWORD` in `.env` (spaties mogen weg).
   `GMAIL_USER` is je Gmail-adres.

## Cron: elke 1e van de maand om 07:00

`crontab -e`:

```cron
0 7 1 * * /opt/laadoverzicht/.venv/bin/python /opt/laadoverzicht/laadoverzicht.py >> /mnt/docker-data/laadoverzicht/cron.log 2>&1
```

Alle paden zijn absoluut; het script leest `.env` uit de map van het script, dus de werkmap van cron doet er niet toe.
Het script logt zelf naar `LOG_FILE` (standaard `OUTPUT_DIR/laadoverzicht.log`) en geeft een exitcode ongelijk aan 0 bij een fout
(2 = ontbrekende of ongeldige instellingen, 1 = andere fout). Controleer na de eerste 1e of de mail aankwam.

## Instellingen

Uit `.env` naast het script, of uit environment-variabelen (die winnen van `.env`). Zie `.env.example` voor uitleg per veld:
`ZAPTEC_USERNAME`, `ZAPTEC_PASSWORD`, `ZAPTEC_CHARGER_ID` (optioneel), `GMAIL_USER`, `GMAIL_APP_PASSWORD`,
`MAIL_TO` (verplicht, behalve bij `--dry-run`), `MAIL_CC` (leeg = `GMAIL_USER`), `RATE`, `OWNER_NAME`, `OUTPUT_DIR`, `LOG_FILE`.

`.env` staat in `.gitignore`; commit nooit echte geheimen.

## Tests

```bash
.venv/bin/pip install -r requirements-dev.txt
.venv/bin/python -m pytest
```

De tests mocken de Zaptec API en SMTP (maandgrenzen, kosten en totaal, cc-standaard, onderwerp en bijlagenaam, dry-run).
