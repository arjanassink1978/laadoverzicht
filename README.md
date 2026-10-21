# laadoverzicht

Maandelijks een laadoverzicht (xlsx) van je Zaptec-laadpaal automatisch mailen, bijvoorbeeld om laadkosten bij je werkgever te declareren.

Op de 1e van de maand haalt het script via de Zaptec Cloud API alle laadsessies van de vorige kalendermaand op,
zet ze in een xlsx (datum, start, eind, kWh, tarief, kosten, totaalregel) en mailt die via Gmail naar een ontvanger naar keuze,
met een cc naar jezelf.

## Wat heb je nodig

- Een Zaptec-laadpaal met een account op https://portal.zaptec.com (getest met de Zaptec Go2; andere Zaptec-palen in dezelfde portal werken in principe ook)
- Een machine die op de geplande tijd aan staat, met Linux, macOS of iets anders met Python 3 en cron (of een andere scheduler, zoals launchd, systemd-timers of Taakplanner). Een Raspberry Pi is een prima voorbeeld, maar niet vereist.
- Een Gmail-account met een app-wachtwoord om mee te mailen (zie hieronder)

## Wat het doet

- Onderwerp: `declaratie <OWNER_NAME> laadpaal <maand jaar>`, bijv. `declaratie Jan laadpaal oktober 2026`
- Bijlage en opgeslagen bestand: onderwerp + `.xlsx`, bijv. `declaratie Jan laadpaal oktober 2026.xlsx`
- Eén vast tarief (`RATE`) voor alle sessies, of een dynamisch uurtarief (`RATE_MODE=dynamic`); geen dal-/nachttarief
- Tijden staan in Europe/Amsterdam; de maandgrenzen ook

### Tarief

Er zijn twee modi, in te stellen met `RATE_MODE`.

**Vast tarief (`RATE_MODE=fixed`, standaard).** Het tarief is je eigen keuze: neem het all-in kWh-tarief incl. btw uit je energiecontract.
De standaardwaarde (EUR 0,27505/kWh) is slechts een voorbeeld en zal niet bij jouw contract passen.
Een bestaande `.env` zonder de nieuwe velden werkt ongewijzigd door.

**Dynamisch uurtarief (`RATE_MODE=dynamic`).** Voor contracten met een prijs per uur (Tibber, Frank Energie, ANWB, Zonneplan en dergelijke).
Het script haalt de day-ahead marktprijzen van de maand op bij [EnergyZero](https://api.energyzero.nl/v1/energyprices) (gratis, geen API-sleutel)
en rekent per uur waarin geladen is:

```
prijs per kWh = (marktprijs + DYNAMIC_MARKUP + ENERGY_TAX) x (1 + VAT)
```

- `DYNAMIC_MARKUP`: opslag of inkoopvergoeding van je leverancier per kWh excl. btw (standaard 0). Zoek die op in je contract.
- `ENERGY_TAX`: energiebelasting op elektriciteit per kWh excl. btw. De standaard is het tarief van 2026 voor schijf 1 (t/m 10.000 kWh): EUR 0,1108481 incl. btw = EUR 0,09161 excl. btw (bron: Belastingdienst). Controleer dit elk jaar en pas het zo nodig aan.
- `VAT`: btw als fractie, standaard `0.21`.
- De marktprijzen van EnergyZero zijn uurprijzen. Sinds de overstap van EPEX op kwartierprijzen (oktober 2025) zijn dat gemiddelden van de vier kwartieren; kwartierprijzen biedt de API niet.
- Verbruik per tijdsblok: het script vraagt bij Zaptec de meetpunten per sessie op (`DetailLevel=1`, `EnergyDetails`) en rekent per meetinterval, verdeeld over de uren waarin dat interval valt. Levert Zaptec voor een sessie geen bruikbare meetpunten, dan wordt de kWh gelijkmatig over de duur van de sessie verdeeld en komt er een waarschuwing in het log.
- Ontbreekt de marktprijs voor een uur waarin geladen is, dan stopt de run met een foutmelding (exitcode 1) en wordt er niets gemaild. Er wordt niets geschat.
- In de xlsx toont de kolom tarief het gemiddelde EUR/kWh van de sessie (kosten / kWh) en heet de kop "Dynamisch tarief (gem. EUR/kWh)". De totaalregel blijft hetzelfde.

Voorbeeld `.env`-fragment:

```
RATE_MODE=dynamic
DYNAMIC_MARKUP=0.0185
```

## Installatie

```bash
git clone https://github.com/arjanassink1978/laadoverzicht.git <pad-naar-laadoverzicht>
cd <pad-naar-laadoverzicht>
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
cp .env.example .env
chmod 600 .env
$EDITOR .env     # vul de gegevens in, zie de tabel hieronder
```

## Configuratie

Instellingen komen uit `.env` naast het script, of uit environment-variabelen (die winnen van `.env`).
`.env` staat in `.gitignore`; commit nooit echte geheimen.

| Variabele | Verplicht | Standaard | Uitleg |
|---|---|---|---|
| `ZAPTEC_USERNAME` | ja | | Gebruikersnaam van je Zaptec-portaalaccount |
| `ZAPTEC_PASSWORD` | ja | | Wachtwoord van je Zaptec-portaalaccount |
| `ZAPTEC_CHARGER_ID` | nee | alle laadpalen van het account | ID van één specifieke laadpaal |
| `GMAIL_USER` | ja | | Je Gmail-adres (afzender) |
| `GMAIL_APP_PASSWORD` | ja | | Gmail-app-wachtwoord van 16 tekens, niet je gewone wachtwoord |
| `MAIL_TO` | ja, behalve bij `--dry-run` | | E-mailadres van de ontvanger |
| `MAIL_TO_NAME` | nee | leeg | Naam voor de aanhef: `Piet` geeft `Beste Piet,`; leeg geeft `Beste,` |
| `MAIL_CC` | nee | `GMAIL_USER` | cc-adres |
| `RATE_MODE` | nee | `fixed` | `fixed` (vast tarief `RATE`) of `dynamic` (dynamisch uurtarief), zie [Tarief](#tarief) |
| `RATE` | nee | `0.27505` | Vast tarief in EUR per kWh incl. btw; alleen bij `RATE_MODE=fixed` |
| `DYNAMIC_MARKUP` | nee | `0` | Alleen bij `dynamic`: opslag/inkoopvergoeding leverancier in EUR per kWh excl. btw |
| `ENERGY_TAX` | nee | `0.09161` | Alleen bij `dynamic`: energiebelasting in EUR per kWh excl. btw (tarief 2026, schijf 1) |
| `VAT` | nee | `0.21` | Alleen bij `dynamic`: btw als fractie |
| `OWNER_NAME` | ja, ook bij `--dry-run` | | Jouw naam in onderwerp, bestandsnaam en kop. Geen standaardwaarde; zonder dit veld stopt het script met een foutmelding. |
| `OUTPUT_DIR` | nee | map `output` naast het script | Map voor de xlsx; gebruik een absoluut pad |
| `LOG_FILE` | nee | `OUTPUT_DIR/laadoverzicht.log` | Logbestand; gebruik een absoluut pad |

## Gebruik

Test eerst zonder te mailen (maakt wel de xlsx):

```bash
.venv/bin/python laadoverzicht.py --month 2026-10 --dry-run
```

Daarna een echte run voor een maand naar keuze:

```bash
.venv/bin/python laadoverzicht.py --month 2026-10
```

Opties:

- `--month YYYY-MM`: de maand om te verwerken (standaard de vorige maand)
- `--dry-run`: xlsx maken, niet mailen

Het script geeft een exitcode ongelijk aan 0 bij een fout (2 = ontbrekende of ongeldige instellingen, 1 = andere fout).

## Gmail app-wachtwoord aanmaken

1. Zet 2-stapsverificatie aan op https://myaccount.google.com/security.
2. Ga naar https://myaccount.google.com/apppasswords (zoek anders naar "App-wachtwoorden" in je Google-account).
3. Geef een naam op, bijv. `laadoverzicht`, en kies **Maken**.
4. Kopieer het wachtwoord van 16 tekens naar `GMAIL_APP_PASSWORD` in `.env` (spaties mogen weg).
   `GMAIL_USER` is je Gmail-adres.

## Automatisch draaien: elke 1e van de maand

Met cron (`crontab -e`), elke 1e om 08:00:

```cron
0 8 1 * * <pad-naar-laadoverzicht>/.venv/bin/python <pad-naar-laadoverzicht>/laadoverzicht.py >> <pad-naar-laadoverzicht>/cron.log 2>&1
```

Gebruik absolute paden; het script leest `.env` uit de map van het script, dus de werkmap van de scheduler doet er niet toe.
Het script logt zelf naar `LOG_FILE`. Gebruik je een andere scheduler, laat die dan hetzelfde commando op de 1e van de maand draaien.
Controleer na de eerste 1e of de mail aankwam.

## Tests

```bash
.venv/bin/pip install -r requirements-dev.txt
.venv/bin/python -m pytest
```

De tests mocken de Zaptec API en SMTP (maandgrenzen, kosten en totaal, cc-standaard, onderwerp en bijlagenaam, dry-run) en de EnergyZero-prijs-API (fixed ongewijzigd, berekening per interval, gelijkmatige verdeling als terugval, ontbrekende prijzen, btw en energiebelasting).

## Licentie

MIT, zie [LICENSE](LICENSE).
