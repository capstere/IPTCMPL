# QDIP v2 — Arkitektur-blueprint & härdningsplan

**Skapad:** 2026-07-22
**Syfte:** Dokumentera det verifierade nuläget, fastställa den korrekta fiskalkalendern, och lägga en fasad plan för att göra QDIP robust, förenklat, dokumenterat och framtidssäkert — **utan att kasta bort fungerande arbete.**

---

## 0. Huvudbeslut: Refaktorisering, inte omskrivning

Jag mätte hela ditt projekt innan jag skrev detta. Slutsatsen är entydig:

| Vad jag mätte | Värde | Bedömning |
|---|---|---|
| Measures | 80, konsekvent namngivna | Behåll |
| Relationer | 19, ren stjärnstruktur mot Dim_Date/Danaher Calendar | Behåll |
| Presentation (sida 20) | Polerad, mötesklar (bekräftat i skärmdump) | Behåll |
| OneLake + SharePoint composite | Korrekt separerat | Behåll |
| **Skräptabeller** | 8 auto-date (LocalDateTable/DateTableTemplate) | **Städa** |
| **Fiskalkalender** | Beräknad med formler → fel årsgränser | **Ersätt med lookup-tabell** |
| **`_OL`-suffix** | ~22 measures, kvarleva från migrering | **Byt namn försiktigt** |
| **Datakälla** | Hårt kopplad till OneLake/SharePoint | **Abstrahera för Snowflake-framtid** |
| **Power App-inmatning** | Fungerar men känns plottrig | **Förenkla struktur** |
| **Dokumentation** | Saknas | **Skapa** |

**Att bygga om från noll skulle kasta bort hundratals timmar och återinföra alla buggar vi redan fixat.** Det du behöver är en riktad härdning av de 6 punkterna i fetstil. Den här planen gör exakt det.

---

## 1. DEN KORREKTA FISKALKALENDERN (kärnfyndet)

Din nya `Fiscal Calendar.xlsx` är korrekt och sträcker sig **2015–2042** på dagsnivå. Jag verifierade den mot skärmdumpen: fiskala juni 2026 = veckorna W22–W26, exakt som presentationen visar. ✅

### Verifierad struktur (Danaher 4-4-5, Sat–Fri veckor)

**Fiskalår 2026:** 1 jan (tors) → 31 dec (tors), 365 dagar
**Veckor:** Lördag → fredag. W01 är partiell (1–2 jan = 2 dagar), W53 är partiell (26–31 dec = 6 dagar)
**Kvartalsmönster:** 4-4-5 veckor

| Månad | Namn | Startar | Slutar | Veckor | Dagar |
|---|---|---|---|---|---|
| M01 | January | 2026-01-01 (tors) | 2026-01-23 | 4* | 23 |
| M02 | February | 2026-01-24 (lör) | 2026-02-20 | 4 | 28 |
| M03 | March | 2026-02-21 (lör) | 2026-03-27 | 5 | 35 |
| M04 | April | 2026-03-28 (lör) | 2026-04-24 | 4 | 28 |
| M05 | May | 2026-04-25 (lör) | 2026-05-22 | 4 | 28 |
| M06 | June | 2026-05-23 (lör) | 2026-06-26 | 5 | 35 |
| M07 | July | 2026-06-27 (lör) | 2026-07-24 | 4 | 28 |
| M08 | August | 2026-07-25 (lör) | 2026-08-21 | 4 | 28 |
| M09 | September | 2026-08-22 (lör) | 2026-09-25 | 5 | 35 |
| M10 | October | 2026-09-26 (lör) | 2026-10-23 | 4 | 28 |
| M11 | November | 2026-10-24 (lör) | 2026-11-20 | 4 | 28 |
| M12 | December | 2026-11-21 (lör) | 2026-12-31 | 6* | 41 |

*M01 och M12 är partiella pga att fiskalåret pinnas till kalenderår (1 jan–31 dec).

### ⚠️ Var den gamla logiken var fel

Din ursprungliga specifikation sa: *"December 2026 är en lång fiskalmånad och slutar 2027-01-01."*

**Det stämmer inte.** Den korrekta filen visar:
- FY2026 M12 slutar **2026-12-31** (tors), W53
- **2027-01-01** (fre) tillhör **FY2027 M01 W01** — inte december 2026

December **är** lång (6 veckor / 41 dagar), men den slutar 31 dec, inte 1 jan. Den gamla formeln som beräknade månadsgränser fick detta fel. Det är precis den typ av fel som uppstår när man *beräknar* fiskallogik istället för att *slå upp* den.

### Den robusta principen: lookup, inte beräkning

**Gammalt (fel):** `Map_FiscalMonths` + DAX/M-formler som beräknade var månader börjar/slutar → fragilt, fel vid årsgränser, ej framtidssäkert.

**Nytt (rätt):** Ladda `Fiscal Calendar.xlsx` som den auktoritativa `Dim_Date`. Varje datum har redan sin fiskala månad/vecka/kvartal förberäknad och verifierad t.o.m. 2042. Ingen formel kan bli fel, för det finns ingen formel — bara en uppslagstabell.

### Ny Dim_Date — kolumndesign

Ladda dimDate-fliken och forma till dessa kolumner (Power Query):

| Kolumn | Källa i xlsx | Exempel | Syfte |
|---|---|---|---|
| `Date` | Reference Date | 2026-05-23 | Nyckel, relationer |
| `FiscalYear` | Year Number | 2026 | Årsaxel |
| `FiscalQuarter` | Quarter Id | Q2 | Kvartalsaxel |
| `FiscalYearQuarter` | Year Quarter | 2026-Q2 | Sorterbar kvartal |
| `FiscalMonthNum` | Month Id → 4→"04" | 06 | Sortering |
| `FiscalMonthName` | härledd från Month Id | June | Slicer-etikett |
| `FiscalYearMonth` | Year Month | 2026-M06 | Sorterbar månad, unik |
| `FiscalWeekNum` | Week Id → W22→22 | 22 | Veckoetikett i kalender |
| `FiscalYearWeek` | Year Week | 2026-W22 | Sorterbar vecka, unik |
| `DayOfWeekName` | härledd | Saturday | Kalenderkolumner |
| `DayOfWeekSort` | härledd (Sat=1..Fri=7) | 1 | Kalenderkolumn-ordning |
| `IsWeekend` | härledd (Sat/Sun) | true | Skuggning helger |
| `FiscalWeekInMonth` | härledd (rank vecka inom månad) | 1 | Kalenderrad-ordning |
| `Today_Offset` | härledd = DATEDIFF(Date, TODAY) | -3 | Recent/rollover-logik |

**Sorteringar (Sort by column):** `FiscalMonthName` → sortera på `FiscalMonthNum`; `FiscalYearMonth`/`FiscalYearWeek` sorterar naturligt som text.

**Kritiskt:** Denna enda tabell ersätter både gamla `Dim_Date` **och** `Map_FiscalMonths`. En sanningskälla.

---

## 2. DATAKÄLLA-ABSTRAKTION (Snowflake-framtiden)

Du sa: Production Order-datan i SharePoint försvinner och ersätts av Snowflake/SQL Server. Vi förbereder för det **nu** så bytet blir smärtfritt senare.

### Nuläge
- **Produktionsdata** (Historical/Active Production Orders, Danaher Calendar, PO Products, Modules, Packing): DirectQuery → AS "Solna Quality report" (OneLake)
- **QDIP-input** (EHS/Quality/Delivery/Actions Daily): Import från SharePoint-listor

### Strategin: kanoniskt schemakontrakt + staginglager

Måtten ska **aldrig** behöva ändras när källan byts. Det uppnås med ett tydligt lager mellan källa och semantik.

```
KÄLLA (byts ut)          STAGING (kanoniskt)         SEMANTIK (rörs ej)
─────────────────        ────────────────────        ──────────────────
OneLake AS  ──┐
              ├──►  stg_HistoricalProductionOrders ──►  measures
Snowflake   ──┘      (samma kolumnnamn alltid)         (D_Lagging_*, KPI_*)
(framtid)
```

**Kanoniskt kontrakt** — de enda kolumner måtten beror på (dokumenteras och fryses):

`stg_HistoricalProductionOrders` måste alltid leverera:
- `Order#`, `IPT/FQC date`, `Lead time`, `All Orders`, `Quality under 20h`, `Material`, `SAP Batch#`

`stg_ActiveProductionOrders` måste alltid leverera:
- `Order#`, `IPT LT`, `Filter IPT/FQC`, `Material`, `Work Center`, `ROBAL - Actual end date/ time`, `Packing info`, `IPT - Test results`, `IPT/Resampling date`, `IPT - Compiling of data`

**Så här fungerar bytet senare:**
1. Idag: staging-lagret pekar på OneLake-tabellerna (bara en `Source =`-rad i Power Query, eller en tunn passthrough)
2. När Snowflake kommer: ändra **enbart** `Source`-steget + kolumnmappning i staging. Kanoniska kolumnnamn behålls.
3. Måtten, relationerna, visualerna: **noll ändringar.**

**Rekommendation:** Använd en Power Query-parameter `pSource` ("OneLake" / "Snowflake") och en `Source`-switch i varje staging-query. Då är hela bytet en parameterändring + verifiering.

**Notering om DirectQuery:** OneLake-tabellerna är idag DirectQuery mot AS, vilket är svårare att lägga staging framför. Del av härdningen är att bestämma: fortsätter vi DirectQuery (då blir Snowflake också DirectQuery/dual), eller går vi mot Import? Det avgör vi i Fas 3 — men kontraktet ovan gäller oavsett.

---

## 3. POWER APP — förenklad inmatningslogik

### Problemet med nuvarande design
DailyScreen samlar Leading (idag) + Lagging (3 senaste dagar) på samma yta med 65+ kontroller synliga. Två separata Save-knappar. Plottrigt, lätt att göra fel.

### Ny princip: "en dag, en uppgift, ett spar"

**Kärnidé:** Användaren väljer *vilken dag* de fyller i, och ser bara *den dagens* formulär — rent, fokuserat, ett Save.

```
┌─────────────────────────────────────┐
│  QDIP Input                          │
│  ┌────────────┐  Välj dag:           │
│  │ ● Today    │  [Today ▼]           │  ← dagväljare, inte två zoner
│  │ ○ Yesterday│                      │
│  │ ○ 2 d ago  │                      │
│  └────────────┘                      │
│                                      │
│  ┌─ EHS ──────┐ ┌─ Quality ┐ ┌─ D ─┐ │  ← 3 statuskort
│  │ ✓ 28/28    │ │ ○ ej inl│ │● 2  │ │     (klick öppnar modal)
│  └────────────┘ └──────────┘ └─────┘ │
│                                      │
│  [ Spara dagen ]                     │  ← ETT spar för vald dag
└─────────────────────────────────────┘
```

**Vinster:**
- En dagväljare ersätter "Leading-zon + Lagging-zon"-uppdelningen. Samma UI oavsett vilken dag.
- Tre statuskort visar direkt vad som är ifyllt/saknas för vald dag.
- Klick på kort → fokuserad modal (EHS / Quality / Delivery), en sak i taget.
- **Delivery-causes blir obligatoriska på Lagging** (dag i det förflutna med overdue lots) — inbyggt i modalens validering.
- Ett enda "Spara dagen" — ingen förvirring om två knappar.

**Robusthet:**
- All Patch-logik återanvänds (den är redan verifierad och korrekt)
- ClearCollect-mönstret för delegationssäkerhet behålls
- Modalerna vi redan designat i tidigare plan blir byggstenarna

Detaljerad byggguide kommer i Fas 5 — men detta är riktningen.

---

## 4. HÄRDNINGSPLAN — 6 faser

Varje fas är oberoende körbar. Vi gör en i taget och verifierar innan nästa. **Fet = obligatorisk nu. Kursiv = kan vänta.**

### **Fas 1 — Modellhygien (låg risk, hög vinst)** ~30 min
1. Stäng av **Auto date/time** (File → Options → Data Load) → tar bort alla 8 skräptabeller
2. Verifiera att inga visualer använde auto-hierarkier (de använder Dim_Date, så säkert)
3. Ta bort/arkivera "Page 1" (9 visualer — verifiera först att det är en gammal yta)
4. Spara som ny version `QDIP_v2.0`

**Varför först:** Ren modell gör allt annat lättare att resonera om. Noll risk för affärslogik.

### **Fas 2 — Korrekt Dim_Date (kärnan i framtidssäkring)** ~90 min
1. Ladda `Fiscal Calendar.xlsx` dimDate-flik via Power Query
2. Forma till kolumndesignen i sektion 1
3. Bygg om relationerna: alla QDIP-tabeller + Danaher Calendar-brygga mot nya `Dim_Date[Date]`
4. Ta bort gamla `Map_FiscalMonths`
5. Verifiera fiskala juni 2026 = W22–W26 mot skärmdumpen
6. Uppdatera slicern på sida 20 till nya `FiscalMonthName`

**Varför nu:** Detta är den enskilt viktigaste framtidssäkringen. Löser årsgränsbuggen permanent.

### *Fas 3 — Datakälla-abstraktion (framtidssäkring för Snowflake)* ~60 min
1. Bestäm DirectQuery vs Import för produktionstabellerna
2. Skapa staging-queries med kanoniskt schema (sektion 2)
3. Lägg `pSource`-parameter
4. Peka måtten mot staging (om namnbyte krävs)
5. Dokumentera kontraktet

**Varför kan vänta:** SharePoint-produktionsdatan finns kvar ett tag. Men gör detta innan Snowflake-migreringen faktiskt sker.

### **Fas 4 — Measure-städning & namngivning** ~45 min
1. Ta bort `_OL`-suffix på ~22 Delivery-measures (försiktigt, verifiera visual-referenser)
2. Ta bort ev. döda measures
3. Standardisera formatsträngar (procent, heltal)
4. Verifiera alla visualer efter namnbyte

**Varför nu:** Gör modellen underhållbar. Måste göras försiktigt så inget bryts.

### **Fas 5 — Power App förenkling** ~2-3 tim
1. Bygg dagväljar-mönstret (sektion 3)
2. Tre statuskort + modaler
3. Obligatoriska Delivery-causes på Lagging
4. Ett "Spara dagen"
5. Behåll all verifierad Patch-logik

**Varför:** Den robustare, enklare inmatningen du efterfrågade.

### **Fas 6 — Dokumentation (det som gör det "framtidssäkert")** ~60 min
1. **Data dictionary:** alla tabeller, kolumner, källor, kanoniskt kontrakt
2. **Measure-katalog:** de 80 måtten grupperade med syfte
3. **Arkitekturöversikt:** dataflöde källa→staging→semantik→visual
4. **Fiskalkalender-spec:** strukturen i sektion 1
5. **Runbook:** hur man byter datakälla, lägger till en månad, felsöker

**Varför sist:** Dokumenterar det färdiga, städade tillståndet — inte ett rörligt mål.

---

## 5. Vad som INTE ändras (skyddat)

- Presentationssidan 20:s layout och design
- Detaljsidorna 21/22/23 och tooltips 24/25
- Måttlogiken (bara namn städas, inte beräkningar)
- Den verifierade Patch-logiken i Power App
- SharePoint-listornas kolumnstruktur
- OneLake-relationerna (bara ev. staging framför)

---

## 6. Rekommenderad startpunkt

**Fas 1 + Fas 2** ger dig 80% av robusthetsvinsten:
- Ren modell (Fas 1)
- Korrekt, framtidssäker fiskalkalender som lookup (Fas 2)

Det löser den akuta fiskalbuggen och rensar skräpet. Resten (abstraktion, namngivning, app, dokumentation) bygger vi ovanpå en solid grund.

**Säg "Kör Fas 1"** så börjar vi med modellhygienen — låg risk, snabb vinst, och en ren bas för allt annat.
