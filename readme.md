# UA Seasonal Listing Builder

Aplikace vytváří jeden zákaznický listing kompletního **aktivního sezónního portfolia**. Standardní UA produkty a licencované produkty Centric Brands jsou v ní vědomě rozděleny do dvou datových proudů, protože obecný UA Material Data Report není pro Centric autoritativní.

## Vstupy

### Standardní UA portfolio

1. **OOB** – potvrzuje objednané standardní UA EANy. Neomezuje běžnou šíři sezonního listingu, ale je povinným potvrzením pro položky s posledním stavem `ADD` v Change Logu.
2. **UA Material Data Report** – EANová, produktová a logistická master data standardních UA produktů.
3. **EMEA Line List** – určuje, které standardní UA colorwaye jsou v sezoně aktivní. Záložka `Licensed List` se nepoužívá ani nevyžaduje; licencované produkty se berou výhradně z Centric Brands masterdat.
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
= aktivní standardní UA portfolio pouze ze záložky Line List
+ Change Log ADD pouze tehdy, když je stejný style/colorway v OOB
+ OOB potvrzené UA EANové výjimky
+ Centric In-line underwear / boys underwear / kids outerwear / kids sportswear
− Change Log ADD bez potvrzení v OOB
− Centric MFO
```

Centric produkty nejsou závislé na OOB ani na UA Material Data Reportu. Záložka `Licensed List` se záměrně nečte, i když je v exportu přítomná. Pokud generic UA Material Data obsahuje stejný Centric article, Centric řádek jej nahradí, i když generic UA report uvádí jiný EAN, COO, FEDAS nebo materiál.

### Change Log ADD a OOB

- Je-li u standardního UA style/colorway v Change Logu poslední stav `ADD`, zařadí se do zákaznického listingu pouze tehdy, když se tentýž `Article Generic` nachází v OOB.
- Platí to i pro ADD, který už je současně viditelný v aktuálním Line Listu. OOB je zde pojistka, že nejde o předčasně založenou nebo nekompletní položku (například bez HTS dat).
- Audit uvádí vyřazené položky jako `Change Log ADD excluded (not in OOB)`.
- Toto pravidlo se netýká Centric In-line produktů, protože ty v OOB z principu nejsou.


## Brand

- Původní cílový sloupec `DTC exclusive` je v exportu nahrazen sloupcem **`Brand`** ve stejném pořadí a se stejným formátováním.
- Standardní UA položky mají hodnotu `Under Armour`.
- Licencované položky z Centric Brands — underwear, boys underwear, dětský outerwear a dětský sportswear — mají hodnotu `Centric Brand`.
- Aplikace přijme jak nový Muster se sloupcem `Brand`, tak starší Muster se sloupcem `DTC exclusive`; výsledný export vždy použije `Brand`.
- U licencovaných Centric položek se hodnota `KHQ Branded` ve sloupci **Product Division** automaticky normalizuje na `Apparel`. Ostatní division hodnoty Centric zůstávají beze změny.


## Story Tier

- `Storytier` z EMEA Line Listu se do zákaznického listingu ani do auditu nepřenáší.
- Proto není tento sloupec v Line Listu vyžadován a lze použít novější exporty, které jej již neobsahují.

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


## Podbarvení prvních velikostí

- V zákaznickém exportu je šedě podbarven pouze první řádek každého `Article` (tedy první velikost daného stylu a barvy).
- Další velikosti stejného `Article` zůstávají bez šedého podbarvení.
- Barvy a styly se řadí podle `Style`, `Article` a velikosti; šedý řádek proto vždy funguje jako vizuální začátek skupiny variant.

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

## Rozměry a objem

- `Article Lenght`, `Article Width`, `Article Height` jsou vždy v centimetrech a `Dimensions (cm)` je z nich vždy sestaven ve stejném pořadí **L × W × H**. Zákazník tak nikdy nedostane dva rozporné centimetrové údaje.
- Pokud Line List obsahuje u produktu text `Dimensions When Full` s označenými palci (`L`, `W`, `H`), tento údaj je pro zákaznický listing prioritní. Aplikace jej přepíše do všech tří centimetrových sloupců, převede do cm a sjednotí i `Dimensions (inch)` do pořadí **L × W × H**.
- Pokud takový údaj v Line Listu není, ponechají se centimetry z UA Material Data a `Dimensions (cm)` se vytvoří přímo z nich.
- `Volume` obsahuje pouze hodnotu v litrech, například `12 L` nebo `23 L`; z katalogového textu se odstraňuje objem v cubic inches, materiálové složení i `Imported`.

## Compatibility update: worksheet names

The builder now identifies the OOB, UA Material Data Report and reference Muster
by their column headers, not only by a generic worksheet label such as `Sheet1`.
It therefore supports updated exports where the working sheet has been renamed
(for example from `Sheet1` to `Sheet2`).
