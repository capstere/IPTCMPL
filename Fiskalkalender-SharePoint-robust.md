# Fiskalkalender i SharePoint — robust & framtidssäker placering

**Kärnfråga:** Hur placerar jag fiskalkalendern i SharePoint på det mest robusta, framtidssäkra sättet — med tanke på att både Power BI och Power App ska kunna använda den, och att vi enkelt ska kunna byta mål senare?

---

## Rekommendation: SharePoint-**Lista**, inte Excel-fil

Jag rekommenderar bestämt en **SharePoint-lista** (`Dim_FiscalCalendar`), inte en Excel-fil i ett dokumentbibliotek. Här är varför — och varför det spelar roll just för dig:

| Kriterium | Excel-fil i bibliotek | **SharePoint-lista** |
|---|---|---|
| Typade kolumner | Excel gissar typer, kan bli fel | ✅ Native Date/Number/Text |
| Parsning | `Excel.Workbook` kan brista (fliknamn, headers) | ✅ Ingen filparsning |
| Power App kan läsa direkt | ❌ Klumpigt via connector | ✅ Native datakälla |
| Power BI-refresh i Service | Kräver ofta gateway/filsökväg | ✅ Molnbaserad, ingen gateway |
| Risk för korruption | Någon öppnar & råkar ändra struktur | ✅ Kolumnstruktur låst |
| Versionshantering | Filversioner | ✅ Listversioner per rad |
| Konsekvens med din data | — | ✅ Samma typ som QDIP_-listorna |

**Avgörande för dig:** Du sa att Power Appen måste kunna använda den och att målet ska vara lätt att byta. En lista är **native läsbar av både Power BI och Power App samtidigt** — en sanningskälla, konsumerad av båda. En Excel-fil är det inte. Dessutom har du redan alla QDIP-data som listor, så detta är konsekvent.

---

## Steg 1 — Filen är redan förberedd

Jag har genererat **`Dim_FiscalCalendar_SharePoint.xlsx`** åt dig:
- **6 940 rader** (fiskalår 2024–2042 — operativ historik + framtid)
- **Rena kolumnnamn utan mellanslag** (undviker `_x0020_`-helvetet du sett tidigare i t.ex. `Work_x0020_Center`)
- **Rätt typer** förberedda: RefDate som datum, siffror som siffror
- **Formaterad som Excel-tabell** (krävs för "From Excel"-import)
- Verifierad: June 2026 = veckorna 22–26 ✅

### Kolumndesign (medvetet valda namn)

| Kolumn | Typ | Exempel | Syfte |
|---|---|---|---|
| `Title` | Text | "2026-05-23" | SharePoint kräver Title; datumsträng = unik nyckel |
| `RefDate` | Date | 2026-05-23 | **Relationsnyckeln** mot QDIP-data |
| `FiscalYear` | Number | 2026 | Årsslicer |
| `FiscalQuarter` | Text | "Q2" | Kvartalsgruppering |
| `FiscalMonthNum` | Number | 6 | Sortering av månadsnamn |
| `FiscalMonthName` | Text | "June" | Slicer-etikett |
| `FiscalWeekNum` | Number | 22 | Kalenderns radetikett |
| `FiscalYearMonth` | Text | "2026-M06" | Entydig månad över år |
| `FiscalYearWeek` | Text | "2026-W22" | Entydig vecka över år |

**Varför inga mellanslag i namnen:** SharePoint kodar mellanslag till `_x0020_` i det interna namnet, vilket gör DAX/Power Query/Power App fult och felbenäget. Rena namn nu = robust för alltid.

**Varför siffror som Number (inte text):** `FiscalWeekNum` och `FiscalMonthNum` som riktiga tal sorteras och filtreras korrekt utan konvertering i både PBI och Power App.

---

## Steg 2 — Skapa listan från Excel

1. Gå till din QDIP SharePoint-site (samma site som QDIP_EHS_Daily osv.)
2. **New → List → From Excel**
3. **Upload file** → välj `Dim_FiscalCalendar_SharePoint.xlsx`
4. SharePoint läser tabellen `FiscalCalendar` och visar en förhandsvisning
5. **Kontrollera kolumntyperna** i förhandsvisningen:
   - `RefDate` → **Date and time** (viktig! om den föreslås som text, ändra till Date)
   - `FiscalYear`, `FiscalMonthNum`, `FiscalWeekNum` → **Number**
   - Övriga → **Single line of text**
6. Namnge listan: **`Dim_FiscalCalendar`**
7. **Create**

SharePoint skapar listan med alla 6 940 rader och rätt typer.

---

## Steg 3 — Indexera RefDate (prestanda >5000 rader)

Listan har fler än 5 000 rader, så SharePoint kräver ett index för snabb åtkomst. Power BI läser ändå alla rader via API, men indexet gör listan robust.

1. Öppna listan → **Settings (kugghjul) → List settings**
2. Under **Columns**, klicka **Indexed columns**
3. **Create a new index** → Primary column: **RefDate** → **Create**

**Notering:** Detta påverkar bara SharePoint-UI:t. Power BI och Power App läser oberoende av 5000-tröskeln. Men indexet är god praxis.

---

## Steg 4 — Framtidssäker anslutning i Power BI (bytbart mål)

Här är nyckeln till "lätt att byta mål". Vi använder **Power Query-parametrar** så att en källändring är en parameterändring — inte en query-redigering.

### 4.1 Skapa en delad site-parameter

Alla dina SharePoint-källor (QDIP-listorna + fiskalkalendern) ligger på samma site. Lägg site-URL:en i **en** parameter:

1. **Transform data → Manage Parameters → New Parameter**
2. Name: **`pSharePointSite`**
3. Type: **Text**
4. Current Value: `https://<dittföretag>.sharepoint.com/sites/<dinsite>`

**Vinst:** Om siten någonsin flyttar/byter namn ändrar du **en** parameter, och alla queries följer med.

### 4.2 Peka Dim_Date mot listan via parametern

Ersätt Dim_Date-queryn (från Fas 2) med denna **enklare** version — nu behövs nästan ingen transformation eftersom listan redan har rena, typade kolumner:

```m
let
    // Källa via delad parameter (byt mål = ändra pSharePointSite)
    Source = SharePoint.Tables(pSharePointSite, [Implementation="2.0", ViewMode="All"]),
    ListData = Source{[Title="Dim_FiscalCalendar"]}[Items],

    // Välj och byt namn till modellens kanoniska kolumnnamn
    Selected = Table.SelectColumns(ListData, {
        "RefDate", "FiscalYear", "FiscalQuarter", "FiscalMonthNum",
        "FiscalMonthName", "FiscalWeekNum", "FiscalYearMonth", "FiscalYearWeek"
    }),

    Renamed = Table.RenameColumns(Selected, {
        {"RefDate", "Date"},
        {"FiscalWeekNum", "FiscalWeek2026"}   // behåll namnet tills Fas 4-namnbyte
    }),

    // Härled hjälpkolumner som modellen använder
    AddDOW = Table.AddColumn(Renamed, "_DOW", each Date.DayOfWeek([Date], Day.Saturday), Int64.Type),
    AddWSort = Table.AddColumn(AddDOW, "WeekdaySortOrder", each [_DOW] + 1, Int64.Type),
    AddDName = Table.AddColumn(AddWSort, "DayOfWeekName", each Date.ToText(Date.From([Date]), [Format="dddd", Culture="en-US"]), type text),
    AddWknd = Table.AddColumn(AddDName, "IsWeekend", each ([_DOW] = 0 or [_DOW] = 1), type logical),
    AddDoM = Table.AddColumn(AddWknd, "DayOfMonth", each Date.Day([Date]), Int64.Type),

    // Typa Date som datetime (matchar relationer)
    Typed = Table.TransformColumnTypes(AddDoM, {{"Date", type datetime}, {"FiscalMonthName", type text}, {"FiscalYear", type text}, {"FiscalMonthNum", type text}}),

    Final = Table.SelectColumns(Typed, {
        "Date", "DayOfMonth", "DayOfWeekName", "WeekdaySortOrder", "IsWeekend",
        "FiscalWeek2026", "FiscalMonthNum", "FiscalMonthName", "FiscalYear",
        "FiscalYearMonth", "FiscalYearWeek", "FiscalQuarter"
    })
in
    Final
```

**Notera:** `FiscalYear`/`FiscalMonthNum` typas om till text för att exakt matcha nuvarande modell (där de var strängar). Om du hellre vill ha dem som tal — säg till, men då måste ev. sorteringar verifieras.

### 4.3 Framtiden: byte till Snowflake

När produktions-SharePoint blir Snowflake ändrar du bara **`Source`-raden** i respektive query (och lägger ev. en `pSource`-parameter enligt blueprintens Fas 3). Kolumnnamnen ut ur queryn är oförändrade → measures och visualer rörs aldrig. Fiskalkalendern kan ligga kvar som SharePoint-lista oberoende av var produktionsdatan bor.

---

## Steg 5 — Framtidssäker användning i Power App (bytbart mål)

Du bad mig tänka på detta för Power Appen. Här är principen vi bygger in **nu** så det blir enkelt senare:

### Abstrahera varje datakälla bakom en collection i App.OnStart

Istället för att referera `Dim_FiscalCalendar` (eller QDIP-listorna) direkt utspritt i appen, ladda dem **en gång** i `App.OnStart` till collections, och referera collections överallt:

```
// App.OnStart — en enda plats att byta mål
ClearCollect(colFiscal, Dim_FiscalCalendar);
ClearCollect(colEHS, QDIP_EHS_Daily);
ClearCollect(colQuality, QDIP_Quality_Daily);
ClearCollect(colDelivery, QDIP_Delivery_Daily);
ClearCollect(colActions, QDIP_Actions);
```

**Vinsten är dubbel:**
1. **Bytbart mål:** Om en lista byter namn/källa ändrar du **en rad** i OnStart. Resten av appen (som refererar `colFiscal` etc.) rörs inte.
2. **Delegationssäkert:** Samma mönster vi redan använt för `colActions` — filtrering sker lokalt, inga Delegation Warnings.

**Konkret nytta av colFiscal i appen:** När användaren väljer en dag att fylla i kan appen slå upp fiskalmånad/vecka direkt från `colFiscal` (t.ex. visa "You are entering: June 2026, Week 22") — samma sanningskälla som Power BI, garanterat konsekvent.

Detta bygger vi in när vi kommer till Power App-förenklingen (Fas 5).

---

## Sammanfattning — varför detta är robust och framtidssäkert

1. **En sanningskälla** för fiskallogik, konsumerad av både Power BI och Power App
2. **Native typer** — ingen filparsning, inga `_x0020_`-namn
3. **Molnrefresh** utan gateway
4. **Bytbart mål via parametrar** (Power BI) och **collections** (Power App) — källbyte = en rad att ändra
5. **Oberoende av produktionsdatans framtid** — fiskalkalendern påverkas inte av Snowflake-migreringen
6. **Verifierad korrekt** t.o.m. 2042

---

## Din åtgärdslista

- [ ] Ladda ner `Dim_FiscalCalendar_SharePoint.xlsx` (bifogad)
- [ ] Skapa SharePoint-lista `Dim_FiscalCalendar` via New → List → From Excel
- [ ] Verifiera kolumntyper (RefDate = Date, siffror = Number)
- [ ] Indexera RefDate
- [ ] Skapa parameter `pSharePointSite` i Power BI
- [ ] Ersätt Dim_Date-queryn med list-versionen (steg 4.2)
- [ ] Verifiera June 2026 = veckorna 22–26
- [ ] (Fas 5) Bygg in collection-abstraktionen i Power App

---

När listan är skapad och Power BI läser från den, säg **"Klar"** så fortsätter vi. Vill du att jag genererar en motsvarande **ren importfil för något annat** (eller justerar årsintervallet 2024–2042) — säg bara till.
