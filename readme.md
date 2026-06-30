# UA Seasonal Listing Builder

Aplikace vytváří jeden zákaznický listing kompletního **aktivního sezónního portfolia**. Standardní UA produkty a licencované produkty Centric Brands jsou v ní vědomě rozděleny do dvou datových proudů, protože obecný UA Material Data Report není pro Centric autoritativní.

## Vstupy

### Standardní UA portfolio

1. **OOB** – pouze potvrzuje objednané standardní UA EANy. Neomezuje šíři sezonního listingu.
2. **UA Material Data Report** – EANová, produktová a logistická master data standardních UA produktů.
3. **EMEA Line List** včetně Licensed Listu – určuje, které standardní UA colorwaye jsou v sezoně aktivní.
4. **EMEA Line List Change Log** – aplikuje ADD/DROP a změny termínů na standardní UA portfolio.
5. **Referenční listing / muster** – cílové sloupce a formát exportu; reference pro Size EUR a Size Scale.

### Centric Brands – licencované produkty

6. **FW26 Master_Data_Underwear_In Line.xlsx**
   - používá se `In_line Underwear_Undershirts` a `Boys_Underwear`;
   - list `MFO_Underwear_Undershirts` se vždy ignoruje.
7. **FW26 Master_Data_Outerwear_In Line.xlsx** – vše s `Range Segment = IN-LINE`.
8. **FW26 Master_Data_Sportswear_In Line.xlsx** – vše s `Range Segment = IN-LINE`.

## Pravidla rozsahu

```text
Finální listing
= aktivní standardní UA portfolio z Line Listu + Change Logu
+ OOB potvrzené UA EANové výjimky
+ Centric In-line underwear / boys underwear / kids outerwear / kids sportswear
− Centric MFO
```

Centric produkty nejsou závislé na OOB ani na UA Material Data Reportu. Pokud generic UA Material Data obsahuje stejný Centric article, Centric řádek jej nahradí, i když generic UA report uvádí jiný EAN, COO, FEDAS nebo materiál.

## GHL (Global Hero Look)

- Pro standardní UA položky se `GHL` přebírá z pole **Hero Look Name** v EMEA Line Listu.
- Je-li pro konkrétní colorway `Hero Look Name` neprázdný, výsledná hodnota je `ANO`; je-li prázdný nebo pro položku neexistuje Line List řádek, hodnota je `NE`.
- FW26 hodnoty Hero Look Name začínají Q3 nebo Q4. Skript však záměrně testuje neprázdnost pole, aby zůstal platný i při případné budoucí změně pojmenování.
- Audit na listu **Scope** obsahuje výsledný příznak `GHL` i původní `Hero Look Name`.

## Kódy Centric

- Pokud má Centric vyplněný `UA Full Article Code`, použije se tento standardní UA kód, například `6011474-001`.
- Pokud jej nemá, ale má `UA Style Code` + `UA Colour Code`, vytvoří se standardní UA article, například `1383915-001`.
- Pokud je v obou polích `N/A`, parser použije Centric kód. Například:

```text
Raw Material Number: 25UJFJM07F-001-JPC
Style:               25UJFJM07F
Article:             25UJFJM07F-001
SKU / usdis:         25UJFJM07F-001-6-7YR
```

Přípona jako `-JPC`, `-EPC` nebo `-PC` se do zákaznického Article/SKU nepřepisuje; zůstává v auditu jako **Source Material Number**.

## Důležité kontroly

- Nenumerický nebo chybějící EAN/UPC, například `tbc`, se nezapíše do zákaznického listingu a je v auditu jako `Centric row without valid EAN`.
- U Centric outerwear a sportswear jsou v aktuálním FW26 souboru prázdné hodnoty `Country of Origin for PO`; výstup je nechá prázdné a audit je označí jako `Centric COO missing`.
- Centric zdroj nedává oddělené Size EUR pro věkové velikosti, proto se u nich bezpečně zachovává zdrojový label, například `6-7YR`, namísto neověřeného převodu na 122/128.
- Centric zdroj nedává oficiální Size Scale. Aplikace ji transparentně odvodí z velikostí daného stylu a označí v auditu.

## Výstupy

- `FW26_active_listing_data.xlsx` – zákaznický listing ve struktuře Musteru.
- `FW26_active_listing_audit.xlsx`
  - **Summary** – počty podle zdrojů a použitých pravidel;
  - **Exceptions** – chybějící EANy, COO, Size EUR, konflikty a další kontroly;
  - **Scope** – každý výsledný EAN a jeho původ, včetně Centric raw Material Number.

## Lokální spuštění

```bash
pip install -r requirements.txt
streamlit run app.py
```

Pro každou sezónu nahrajete osm aktuálních souborů. Názvy souborů nejsou v aplikaci natvrdo, rozhodující je struktura jednotlivých exportů.
