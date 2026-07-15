# UA Seasonal Listing Builder

Aplikace vytváří jeden zákaznický listing kompletního **aktivního sezónního portfolia**. Standardní UA produkty a licencované produkty Centric Brands jsou v ní vědomě rozděleny do dvou datových proudů, protože obecný UA Material Data Report není pro Centric autoritativní.

## Vstupy

### Standardní UA portfolio

1. **OOB** – potvrzuje objednané standardní UA EANy. Neomezuje běžnou šíři sezonního listingu, ale je povinným potvrzením pro položky s posledním stavem `ADD` v Change Logu.
2. **UA Material Data Report** – kompletní EANová, produktová a logistická master data standardních UA produktů. Builder z něj používá pouze řádky `Masterproductline = Inline` a `Article Status Desc = Active`; MFO, Licensed / Non-Uniform, Promo, Deplete a Discontinued řádky se ignorují.
3. **EMEA Line List** – určuje, které standardní UA colorwaye jsou v sezoně aktivní. Záložka `Licensed List` se nepoužívá ani nevyžaduje; licencované produkty se berou výhradně z Centric Brands masterdat.
4. **EMEA Line List Change Log** – aplikuje ADD/DROP a změny termínů na standardní UA portfolio.
5. **Referenční listing / muster** – cílové sloupce a formát exportu; reference pro Size EUR a Size Scale.

### Centric Brands – licencované produkty

6. **FW26_Master data_centric_BMMcr.xlsx** nebo jiný sloučený Centric listingový soubor.
   - obsahuje underwear, boys underwear, kids outerwear a kids sportswear v jednom souboru;
   - nahrává se pouze tento jeden Centric soubor místo původních tří;
   - očekává se listingová struktura se sloupci `Style`, `Article`, `Size UA`, `EAN`, `Product Division`, `Brand` atd.

## Pravidla rozsahu

```text
Finální listing
= aktivní standardní UA portfolio pouze ze záložky Line List
+ UA Material Data pouze v rozsahu Inline + Active
+ Change Log ADD pouze tehdy, když je stejný style/colorway v OOB
+ OOB potvrzené UA EANové výjimky
+ sloučený Centric soubor pro underwear / boys underwear / kids outerwear / kids sportswear
− Change Log ADD bez potvrzení v OOB
− UA Material Data řádky mimo Inline / Active
```

Centric produkty nejsou závislé na OOB ani na UA Material Data Reportu. Záložka `Licensed List` se záměrně nečte, i když je v exportu přítomná. Pokud generic UA Material Data obsahuje stejný Centric article, Centric řádek jej nahradí, i když generic UA report uvádí jiný EAN, COO, FEDAS nebo materiál. Od této verze se Centric nahrává jako jeden konsolidovaný listingový soubor.

### Change Log ADD a OOB

- Je-li u standardního UA style/colorway v Change Logu poslední stav `ADD`, zařadí se do zákaznického listingu pouze tehdy, když se tentýž `Article Generic` nachází v OOB.
- Platí to i pro ADD, který už je současně viditelný v aktuálním Line Listu. OOB je zde pojistka, že nejde o předčasně založenou nebo nekompletní položku (například bez HTS dat).
- Audit uvádí vyřazené položky jako `Change Log ADD excluded (not in OOB)`.
- Toto pravidlo se netýká Centric In-line produktů, protože ty v OOB z principu nejsou.


## UA Material Data Report – filtr Inline / Active

Kompletní Material Data Report může obsahovat také MFO, licencované non-uniform položky, promo artikly a položky mimo aktivní prodejní status. Builder proto před jakýmkoliv párováním s Line Listem a Change Logem používá pouze:

- `Masterproductline = Inline`;
- `Article Status Desc = Active`.

Všechno ostatní se ze standardního UA proudu vyřadí ještě před rozšířením colorwayů na EANy. Toto pravidlo pomáhá hlavně u Change Logu, kde se mohou objevit `ADD` MFO položky, které do zákaznického zalistování nepatří. Počty vyřazených řádků jsou uvedené v auditním listu **Summary**.

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

Sloučený Centric soubor už používá zákaznické listingové kódy ve sloupcích `Style`, `Article`, `SKU` a `usdis`. Builder je proto pouze převezme a neprovádí další skládání z raw Centric material number. `Source Material Number` v auditu se vyplní z `SKU` / `usdis`, pokud je k dispozici.

## Důležité kontroly

- Nenumerický nebo chybějící EAN/UPC, například `tbc`, se nezapíše do zákaznického listingu a je v auditu jako `Centric row without valid EAN`.
- Pokud má sloučený Centric soubor prázdné `COO COUNTRY`, výstup jej nechá prázdný a audit jej označí jako `Centric COO missing`.
- Centric soubor už obsahuje `Size EUR` i `Size Scale`; builder tyto hodnoty převezme, pokud jsou vyplněné.


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

Pro každou sezónu nahrajete šest aktuálních souborů: pět standardních UA podkladů a jeden sloučený Centric soubor. Názvy souborů nejsou v aplikaci natvrdo, rozhodující je struktura jednotlivých exportů.

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
