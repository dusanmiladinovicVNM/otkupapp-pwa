# Isplata gotovinom — PDV nadoknada nikad gotovinom

> Odluke operatera 05.11.2026. Važi za **svaku** isplatu kooperantu koja razdužuje
> otkupni blok ili dokument doplate (`DOPLATA.md`). Ko piše novac:
> `docs/DOMEN/WHO_WRITES.md` (`tblNovac`).

## Odluke

| ID | Odluka |
|---|---|
| **ISP-ODL-01** | **PDV nadoknada se nikad ne isplaćuje gotovinom.** Gotovina je novac otkupnog mesta kooperantu (`KesOtkupacKoop`, ekran F5 „iz avansa otkupnog mesta"); banka (virman) sme da plati sve, i nadoknadu. Avans kooperanta je uvek virman (`modNovac.JeAvansKooperanta`), pa primena avansa nije gotovina. |
| **ISP-ODL-02** | **Granica:** zbir svih gotovinskih isplata na jednom otkupnom bloku (ili dokumentu doplate) ne sme preći **osnovicu** = zaokruženo(bruto / (1 + stopa)), isto pravilo kao na papiru (otkupni list, papir doplate, specifikacija); stopa iz podešavanja `PDV_NADOKNADA_STOPA` (podrazumevano 8%). Primer: blok 108 = osnovica 100 + nadoknada 8 → gotovinom najviše 100 ukupno; ostatak (najmanje nadoknada) samo banka. |
| **ISP-ODL-03** | **Zatečeni podaci** (gotovina već veća od osnovice, pre ovog pravila) se ne menjaju; Oporavak / integritet ih prijavljuje kao nalaz **GOT1**. |

## Gde živi pravilo

- Jedna implementacija, dva pozivaoca (obrazac `IsplataBlokProblem`,
  `.claude/rules/testovi.md` §5): `modNovac.GotovinaNadoknadaProblem` zove ekran
  F5 (`modNovacUnos.IsplataValidiraj`, poruka uz polje) i pisac
  (`modDokumenta.SaveOMUlaz_TX`, greška po imenu) — samo za gotovinu.
- Osnovica: `modNovac.OsnovicaBezNadoknade` — isti broj na papiru i u kapiji.
- Signal: `modIntegritet` → `GOT1` iz `modNovac.GotovinaPrekoOsnoviceRedovi`.
- Prikaz (zahtev operatera 05.11.2026): uz „Keš iz OM avansa" ploča na F5 menja
  natpis u „KEŠOM NAJVIŠE" i pokazuje granicu gotovine za izabran blok, dokument
  doplate ili specifikaciju (`modNovac.GotovinaNajviseNaIzboru`, isto pravilo kao
  kapija). Operater granicu vidi pre klika. Test 347, sabotaža
  `isp-ploca-bez-osnovice`.
- Gotovina po specifikaciji doplate (DOP-ODL-28, rez B) prolazi istu kapiju po
  dokumentu.

## Test i dokaz

- `modTest` **341** `T_Gotovina_NadoknadaNikadGotovinom`: kapija na bloku i na
  dokumentu doplate (osnovica prolazi, cent preko ne), pisac nad snimkom novca
  (gotovina 400, virman 30 se ne broji, cent preko dozvoljenog odbijen iako je
  bruto ostatak veći, tačno do osnovice prolazi), ekran F5 sa istom porukom uz
  polje iznosa, i zatečena gotovina preko osnovice kao nalaz `GOT1`.
- Sabotaže (`tools/sabotaza.py`): `isp-gotovina-bez-kapije-pisca`,
  `isp-gotovina-bez-kapije-ekrana`, `isp-gotovina-broji-virman`, `isp-got1-nema`.
  Uz njih je re-sidrena `dop-papir-nadoknada-od-bruto` (papir sada računa osnovicu
  kroz `OsnovicaBezNadoknade`).
- Excel (05.11.2026): pun prolaz 12 suita ZELENO (RunAllTests 341/0, BFP
  3014/0); dokaz `isp-` + `dop-papir-nadoknada-od-bruto` **5/5 DOKAZANO**, izvor
  pre/posle identičan. Ručni Compile PASS. Ručno na F5 (05.11.2026): blok
  3/260726 sa ostatkom 627 i „Keš iz OM avansa" odbija 627 porukom „gotovinom
  na ovaj dokument još najviše 580,56" (627 / 1,08). **Zatvoreno.**
