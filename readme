# UA Seasonal Listing Builder

Aplikace vytváří **kompletní zákaznický listing aktivního sezónního sortimentu** — nejen položek, které USDistribution již předobjednal v OOB.

## Co vstupní soubory znamenají

1. **EMEA Line List** + **Licensed List** jsou výchozí seznam produktů, které jsou v dané sezóně aktivní a mohou být objednatelné.
2. **Change Log** aktualizuje tento rozsah:
   - `DROP` vyřadí colorway z širokého aktivního portfolia;
   - `ADD` jej přidá, ale pouze tehdy, pokud jsou v Material Data Reportu dostupné konkrétní EANy a velikosti;
   - datumové změny přepisují `Shipment Start Date` a `Launch Date`.
3. **Material Data Report** rozpadá aktivní colorwaye na konkrétní EANy, velikosti a produktová / logistická data.
4. **OOB** není filtr rozsahu. Označuje již objednané položky a chrání jednotlivé potvrzené EANy, které v aktivní kombinaci Line List + Master Data chybějí.
5. **Referenční listing / muster** určuje výstupní layout a je referencí pro `Size UA → Size EUR` a chybějící `Size Scale`.

## Přesná logika rozsahu

```text
Finální listing
= aktivní Line List + Licensed List
  − aktuální DROP z Change Logu
  + aktuální ADD z Change Logu, pokud mají EANy v Masterdatech
  + jednotlivé OOB EANy mimo tento rozsah (potvrzené výjimky)
```

Proto:

- produkt v **Line Listu**, ale ne v OOB → **je v listingu**, protože jej můžete později doobjednat;
- produkt pouze v **Masterdatech**, mimo Line List a Change Log ADD → **není v listingu**;
- EAN v **OOB**, který je ve starším Change Logu `DROP` → **zůstává v listingu**, ale je označen jako konflikt;
- `ADD` bez EANů v Masterdatech → **nezařadí se**, ale objeví se v auditu.

## Výstupy

- `FW26_active_listing_data.xlsx` — zákaznický listing přesně ve sloupcové struktuře Musteru.
- `FW26_active_listing_audit.xlsx`
  - **Summary** — počty a použitá pravidla;
  - **Exceptions** — konflikty, chybějící data a manuální kontroly;
  - **Scope** — každý výsledný EAN s informací, zda vznikl z aktivního Line Listu, Change Log ADD, nebo jako potvrzená OOB výjimka.

## Lokální spuštění

```bash
pip install -r requirements.txt
streamlit run app.py
```

Pro každou sezónu stačí nahrát nových pět exportů. Názvy souborů nejsou v aplikaci natvrdo.
