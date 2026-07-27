# Finance

Budžeti, unit economics, scenariji, cash-flow modeli, pricing inputi, hardverska marža i finansijska kontrola kompanije.

Osetljivi finansijski podaci, bankarski podaci i ugovori ne čuvaju se u javnom repozitorijumu; ovde ostaju strukture, metodologija i anonimizovani modeli.

## Dokumenti

| Dokument | Verzija / datum | Sadržaj |
|---|---|---|
| `AgriX_Finansijski_model.xlsx` | v1 · 26.07.2026. | Listovi `Pretpostavke`, `Prihod`, `Kapacitet`, `CashFlow`, `Scenariji`. Cene u EUR, godišnje, bez PDV-a. |

Pravila korišćenja modela:

- ulazi se menjaju **isključivo** na listu `Pretpostavke`; ostalo su formule;
- model ne prognozira — pokazuje posledice datih pretpostavki po prihod i po vreme osnivača;
- **troškovi (odeljak F) su prazni**; dok se ne popune, redovi neto rezultata i cash-flow-a nemaju smisla;
- izvor cena su odluke 339, 341, 349–358; iznosi moraju biti identični sa `docs/Sales/AgriX_Cenovnik_2027.pdf`;
- scenariji rasta: odluka 375, izabran scenario C. **Neusklađenost:** odluka 375 daje raspon 12–15 novih / 15–18 ukupno, a model koristi tačku 14 novih / 17 ukupno uz kapacitetnu kolonu od 18. Treba odlučiti šta je merodavno i uskladiti sa `02_STRATEGY.md` §9 i `04_MARKET.md` §9.1;
- okidač za zapošljavanje druge tehnički ovlašćene osobe: BC3 i 336 (najkasnije pri 15–20 firmi), a red `Status` na listu `Kapacitet` može ga pomeriti ranije;
- unit economics je fino podešavanje posle određenih cena, ne preduslov za ponudu (odluka 408).
