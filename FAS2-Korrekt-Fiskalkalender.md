# FAS 2 — Korrekt fiskalkalender som lookup-tabell

**Mål:** Ersätta den felaktiga, beräknade fiskallogiken med en korrekt uppslagstabell från `Fiscal Calendar.xlsx` (verifierad 2015–2042). Löser årsgränsbuggen permanent.

**Verifierat innan start:**
- Måtten rör bara `Dim_Date[Date]` (11×) och `Dim_Date[IsWeekend]` (3×)
- Visualerna rör: `Date`, `DayOfMonth`, `DayOfWeekName`, `FiscalWeek2026`, `FiscalMonthName`, `FiscalMonthNum`, `FiscalYear`
- `Map_FiscalMonths` används **inte** av något mått eller visual → säkert att ta bort
- Gamla trasiga kolumner (`FiscalMonthStart/End`, `IsPreviousCalendarMonth`) används **inte** av något → försvinner
- **Roten till buggen:** Map_FiscalMonths hade `January 2026 → startar 2025-12-27` (fel). Korrekt: 2026-01-01.

**Strategin:** Behåll exakt samma kolumnnamn måtten/visualerna redan använder. Byt bara datakällan bakom, så datan blir korrekt. **Noll ändringar i measures och visualer.**

---

## Steg 2.1 — Placera fiskalkalender-filen

Filen måste ligga där Power BI når den — både i Desktop nu och i Service senare.

**Alternativ A (snabbast, för att testa nu):**
Lägg `Fiscal Calendar.xlsx` i en stabil lokal mapp, t.ex. `C:\QDIP\Fiscal Calendar.xlsx`. Fungerar direkt i Desktop.

**Alternativ B (produktion, rekommenderas innan publicering):**
Ladda upp `Fiscal Calendar.xlsx` till samma SharePoint-site som dina QDIP-listor (dokumentbibliotek). Då fungerar schemalagd refresh i Service utan gateway.

**För Fas 2:** Använd Alternativ A för att få det verifierat nu. Byt till B när du publicerar (bara en `Source`-rad att ändra).

---

## Steg 2.2 — Ersätt Dim_Date:s Power Query

1. **Home → Transform data** (öppnar Power Query Editor)
2. I vänsterpanelen **Queries**, klicka på **Dim_Date**
3. **Home → Advanced Editor**
4. **Markera allt** (Ctrl+A) och **radera**
5. **Klistra in** hela skriptet nedan
6. **Ändra sökvägen** på rad `Source` till din faktiska sökväg (Alternativ A eller B)
7. Klicka **Done**

```m
let
    // ===== KÄLLA (ändra bara denna rad vid byte lokal <-> SharePoint) =====
    Source = Excel.Workbook(File.Contents("C:\QDIP\Fiscal Calendar.xlsx"), null, true),

    dimDate_Sheet = Source{[Item="dimDate", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(dimDate_Sheet, [PromoteAllScalars=true]),

    // Typa källkolumnerna
    Typed = Table.TransformColumnTypes(Promoted, {
        {"Reference Date", type date},
        {"Year Number", Int64.Type},
        {"Quarter Id", type text},
        {"Month Id", type text},
        {"Week Id", type text},
        {"Year Month", type text},
        {"Year Week", type text}
    }),

    // Rensa bort ev. tomma rader
    Cleaned = Table.SelectRows(Typed, each [Reference Date] <> null),

    // ===== Bygg kanoniska kolumner (samma namn som nuvarande modell) =====

    // Date (dateTime) — nyckel för alla relationer
    AddDate = Table.AddColumn(Cleaned, "Date", each DateTime.From([Reference Date]), type datetime),

    // Veckodag: Sat=0..Fri=6 (helper)
    AddDOW = Table.AddColumn(AddDate, "_DOW", each Date.DayOfWeek([Reference Date], Day.Saturday), Int64.Type),

    // WeekdaySortOrder: Sat=1..Fri=7 (för att sortera DayOfWeekName)
    AddWSort = Table.AddColumn(AddDOW, "WeekdaySortOrder", each [_DOW] + 1, Int64.Type),

    // DayOfWeekName: "Saturday".."Friday"
    AddDName = Table.AddColumn(AddWSort, "DayOfWeekName", each
        Date.ToText([Reference Date], [Format="dddd", Culture="en-US"]), type text),

    // IsWeekend: Sat eller Sun (används av measures)
    AddWknd = Table.AddColumn(AddDName, "IsWeekend", each ([_DOW] = 0 or [_DOW] = 1), type logical),

    // DayOfMonth: dagsiffran som visas i kalendercellen
    AddDoM = Table.AddColumn(AddWknd, "DayOfMonth", each Date.Day([Reference Date]), Int64.Type),

    // FiscalWeek2026: "W22" -> 22 (kalenderns radetikett)
    AddFWeek = Table.AddColumn(AddDoM, "FiscalWeek2026", each
        Number.FromText(Text.Middle([Week Id], 1)), Int64.Type),

    // FiscalMonthNum: "M06" -> "06" (sträng, för sortering)
    AddFMNum = Table.AddColumn(AddFWeek, "FiscalMonthNum", each Text.End([Month Id], 2), type text),

    // FiscalMonthName: M01->January .. M12->December
    AddFMName = Table.AddColumn(AddFMNum, "FiscalMonthName", each
        {"January","February","March","April","May","June",
         "July","August","September","October","November","December"}
        {Number.FromText(Text.End([Month Id], 2)) - 1}, type text),

    // FiscalYear: "2026" (sträng, matchar nuvarande modell)
    AddFYear = Table.AddColumn(AddFMName, "FiscalYear", each Text.From([Year Number]), type text),

    // ===== Bonus: framtidssäkra, entydiga kolumner (bryter inget) =====
    AddFYM = Table.AddColumn(AddFYear, "FiscalYearMonth", each [Year Month], type text),      // "2026-M06"
    AddFYW = Table.AddColumn(AddFYM, "FiscalYearWeek", each [Year Week], type text),          // "2026-W22"
    AddFQ  = Table.AddColumn(AddFYW, "FiscalQuarter", each [Quarter Id], type text),          // "Q2"

    // Välj slutliga kolumner i ordning
    Final = Table.SelectColumns(AddFQ, {
        "Date", "DayOfMonth", "DayOfWeekName", "WeekdaySortOrder", "IsWeekend",
        "FiscalWeek2026", "FiscalMonthNum", "FiscalMonthName", "FiscalYear",
        "FiscalYearMonth", "FiscalYearWeek", "FiscalQuarter"
    })
in
    Final
```

**Vad skriptet gör:** Läser dimDate-fliken, härleder alla kolumner måtten/visualerna behöver — med **identiska namn och datatyper** — men nu från den korrekta filen. Plus tre bonuskolumner för framtida entydighet.

---

## Steg 2.3 — Applicera och sätt sorteringar

1. **Home → Close & Apply** (laddar in den nya datan)
2. Vänta tills det laddat (nu 2015–2042, ~10 000 rader — fortfarande litet)

Nu sätter vi sorteringskolumner (dessa är modell-egenskaper som måste sättas om):

3. Gå till **Model view** (eller Data view)
4. Klicka på kolumnen **DayOfWeekName** i Dim_Date
5. I menyfliken **Column tools → Sort by column → WeekdaySortOrder**
6. Klicka på kolumnen **FiscalMonthName**
7. **Column tools → Sort by column → FiscalMonthNum**

**Effekt:** Veckodagarna visas Lördag→Fredag i kalendern, och månaderna sorteras Jan→Dec i slicern (inte alfabetiskt).

---

## Steg 2.4 — Hantera år i slicern (viktigt)

**Problemet:** Gamla Dim_Date täckte bara 2026, så slicern visade bara 2026-månader. Nu täcker den 2015–2042, så `FiscalMonthName` = "June" skulle matcha alla år.

**Lösningen — lägg till en FiscalYear-slicer bredvid månadsslicern:**

1. Gå till **sida 20 (Presentation)**
2. Klicka på tom yta → **Insert → Slicer** (eller kopiera befintlig månadsslicer)
3. Sätt fältet till **Dim_Date[FiscalYear]**
4. Placera den bredvid månadsdropdownen i headern
5. Format → Slicer settings → **Single select** = On
6. Sätt default: klicka **2026** i slicern (den kommer ihåg valet)

**Alternativ (en enda dropdown istället för två):** Byt månadsslicerns fält från `FiscalMonthName` till `FiscalYearMonth`. Då visar den "2026-M06" — entydigt men mindre snyggt. Rekommendation: håll två slicers (År + Månad), det matchar mötesflödet bäst.

**Verifiera:** Välj FiscalYear = 2026, FiscalMonthName = June → kalendrarna ska visa veckorna 22–26.

---

## Steg 2.5 — Ta bort Map_FiscalMonths

Nu när Dim_Date är självständig och korrekt behövs inte den gamla hårdkodade tabellen.

1. **Model view** → högerklicka på tabellen **Map_FiscalMonths**
2. **Delete from model**
3. Bekräfta

Relationerna `Dim_Date.FiscalMonthName → Map_FiscalMonths` och range-relationerna (FiscalMonthStart/End) försvinner automatiskt. Eftersom inget mått/visual använde Map_FiscalMonths bryts ingenting.

**Om Power BI varnar** att Dim_Date-queryn refererar Map_FiscalMonths: det ska inte hända längre eftersom vi ersatte hela Dim_Date-queryn i steg 2.2 (den refererar inte längre Map_FiscalMonths). Om varning ändå kommer — verifiera att du klistrade in hela det nya skriptet.

---

## Steg 2.6 — Verifiera relationerna

1. **Model view**
2. Kontrollera att dessa relationer finns kvar (de bygger på `Date`, som vi behöll):
   - `QDIP_EHS_Daily[InputDate] → Dim_Date[Date]`
   - `QDIP_Quality_Daily[InputDate] → Dim_Date[Date]`
   - `QDIP_Delivery_Daily[InputDate] → Dim_Date[Date]`
   - `QDIP_Actions[DateIdentified] → Dim_Date[Date]`

**Om någon relation är borta** (kan hända om Power BI tappar den vid query-byte): dra `[InputDate]`/`[DateIdentified]` till `Dim_Date[Date]` för att återskapa. Single direction, many-to-one.

**Notera:** OneLake-relationen (`Historical Production Orders[IPT/FQC date] → Danaher Calendar[Date]`) påverkas INTE — den går mot Danaher Calendar, inte Dim_Date. Delivery-måtten använder fortfarande TREATAS från Dim_Date[Date] till Danaher Calendar[Date], vilket fungerar oförändrat.

---

## Steg 2.7 — Slutverifiering mot skärmdumpen

Detta är det avgörande testet att fiskallogiken nu är korrekt:

1. **Sida 20**, sätt slicers: FiscalYear = **2026**, Månad = **June**
2. Kolla EHS/Q/D-kalendrarna. Radetiketterna (veckonummer) ska vara: **22, 23, 24, 25, 26**
3. Första raden (vecka 22) ska börja på **lördag 23 maj**
4. Sista raden (vecka 26) ska sluta på **fredag 26 juni**

Testa även årsgränsen (det som var trasigt):
5. Sätt FiscalYear = **2026**, Månad = **December**
6. December ska visa **6 veckor** (W48–W53), sista dagen **31 dec** (inte 1 jan)
7. Sätt FiscalYear = **2027**, Månad = **January** → 1 jan 2027 ska tillhöra denna månad

Om dessa stämmer är fiskalkalendern korrekt och framtidssäker. ✅

---

## Steg 2.8 — Spara

**File → Save** (behåll QDIP_v2.0, eller Save As `QDIP_v2.1` om du vill ha en milstolpe).

---

## Checklista

- [ ] 2.1 Fiscal Calendar.xlsx placerad (lokal för nu)
- [ ] 2.2 Dim_Date Power Query ersatt med nya skriptet (sökväg justerad)
- [ ] 2.3 Close & Apply OK + sort-by satt (DayOfWeekName→WeekdaySortOrder, FiscalMonthName→FiscalMonthNum)
- [ ] 2.4 FiscalYear-slicer tillagd, default 2026
- [ ] 2.5 Map_FiscalMonths borttagen
- [ ] 2.6 Fyra QDIP→Dim_Date-relationer verifierade
- [ ] 2.7 June 2026 = veckor 22–26 ✓, December slutar 31 dec ✓, jan 2027 korrekt ✓
- [ ] 2.8 Sparad

---

## Om något går fel

**Kalendern blir tom efter bytet**
→ Relationerna QDIP→Dim_Date tappades. Återskapa dem (steg 2.6).

**Veckodagarna visas i fel ordning (Sön först)**
→ Sort-by på DayOfWeekName saknas. Sätt Column tools → Sort by column → WeekdaySortOrder.

**Månaderna sorteras alfabetiskt i slicern (April, August, December...)**
→ Sort-by på FiscalMonthName saknas. Sätt → FiscalMonthNum.

**Slicern "June" visar data från flera år**
→ Lägg till FiscalYear-slicern (steg 2.4) och välj 2026.

**"June" visar veckorna 21–25 istället för 22–26**
→ Fel Week Id-parsning. Verifiera att `FiscalWeek2026`-kolumnen läser rätt (öppna Data view, filtrera June 2026, kolla att veckonummer = 22–26).

**Power Query-fel: "dimDate not found"**
→ Fliknamnet måste vara exakt `dimDate`. Kolla `Source{[Item="dimDate",Kind="Sheet"]}` matchar filens flik.

---

När allt är verifierat, säg **"Klar"** så går vi vidare till Fas 3 (datakälla-abstraktion för Snowflake-framtiden) — eller så hoppar vi till Fas 4/5/6 om du hellre vill prioritera measure-städning, Power App-förenkling eller dokumentation först.
