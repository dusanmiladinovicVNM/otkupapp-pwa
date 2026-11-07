# Provera unosa stare sveske

Alat **samo čita** staru svesku i prijavljuje greške koje je stari unos ostavio u
podacima. Ne menja ništa: SHA256 sveske pre i posle čitanja mora biti isti.

Iz root-a repoa, u Git Bash-u:

```bash
python tools/migracija/provera_unosa.py "C:/Users/Dusan/Desktop/AgriX - C002/Venivo/Original kraj sezone pred konacno stanje/AgriX_C002_2.28.4_2026-10-16.xlsm"
```

Izveštaj ide u folder pored sveske: `Migracija_pregled/<ime>_provera_unosa.xlsx`
(nikad u repo — nosi stvarne podatke). Listovi:

- **Sazetak** — koliko nalaza po kodu i ozbiljnosti;
- **Nalazi** — jedan red po nalazu: ozbiljnost, kod, tabela, ID, opis, predlog ispravke.

| Ozbiljnost | Značenje |
|---|---|
| `GRESKA` | podaci sami dokazuju grešku |
| `SUMNJA` | potpis greške, ali postoji i legitimno objašnjenje — operater proverava |
| `INFO` | skreće pažnju; nije greška sama po sebi |

## Odakle klase grešaka

Svaka klasa je nađena čitanjem koda koji je C002 koristio (tagovi `vba-v2.2.3`
… `vba-v2.28.4`): polje koje posle snimanja ostaje popunjeno, predlog koji
zastari kad se promeni vozač ili stanica, veza koju pisac nije proveravao, storno
koji nije kaskadirao. Za svaku je zapisan potpis u podacima i legitimni slučajevi
koji liče na grešku, pa ih je nezavisan skeptik proverio ponovnim čitanjem koda.

Polazni slučaj: prijemnica nasledi broj zbirne od prethodnog unosa
(`btnUnosPrij_Click` ne prazni `txtBrojZbirnePrij`, `vba-v2.11.0`), pa jedna
zbirna ima dve prijemnice, a prava nijednu.

## Šta se proverava

Kod nalaza je ime klase. Kad tabela postoji a kolona koju provera traži fali,
stoji jedan `INFO <KOD>_PRESKOCENO` sa imenom kolone; tabela koje nema se tiho
preskače.

| Modul | Kodovi |
|---|---|
| `prijemnica` | `PRJ_DVE_NA_ZBIRNOJ_ISTA_STANICA` (varijanta b), `PRJ_ZBIRNA_DRUGE_STANICE` (varijanta a), `PRJ_NASLEDJENA_ZBIRNA`, `ZBR_ZIVA_BEZ_PRIJEMNICE`, `PRJ_ZASTAREO_VOZAC`, `PRJ_KUPAC_RAZLICIT_OD_ZBIRNE`, `PRJ_DATUM_ZASTAREO_ILI_PRE_ZBIRNE`, `PRJ_ROBA_RAZLICITA_OD_ZBIRNE`, `PRJ_KLASA_RAZLICITA_OD_ZBIRNE`, `PRJ_ZBIRNA_NE_POSTOJI_ILI_STORNIRANA`, `PRJ_AKTIVNA_NA_STORNIRANOJ_ZBIRNOJ`, `PRJ_BROJ_ZBIRNE_NENORMALIZOVAN`, `PRJ_DUPLI_BROJ_PRIJEMNICE`, `PRJ_CENA_NULA`, `PRJ_KG_GAJBE_VAN_OPSEGA`, `PRJ_GAJBE_NULA_ILI_BEZ_TIPA`, `PRJ_PREVEZANA_BEZ_PROVERE`, `PRJ_ISPRAVKA_NA_DRUGOJ_ZBIRNOJ`, `PRJ_AUTO_LANAC_NEPOTPUN`, `PRIJEMNICA_POKRIVA_DVE_TURE`, `TESTNI_REDOVI_U_PRODUKCIJI` |
| `otpremnica_zbirna` | `OTP_VOZAC_NIJE_PAR_STANICE`, `TURA_PO_VOZACU_IZ_OTKUPA`, `OTP_BROJ_DATUM_NESLAGANJE`, `OTP_BROJ_TUDJA_STANICA`, `OTP_DUPLI_BROJ_AKTIVAN`, `ZBR_DUPLI_BROJ_AKTIVAN`, `ZBR_VISI_OTP_STORNIRANE`, `ZBR_AKTIVNA_STORNIRANE_OTPREMNICE`, `ISPRAVKA_ZAVRSENA_BEZ_NOVE_ZBIRNE`, `OTP_CEKA_NOVU_ZBIRNU`, `OTP_AKTIVNA_NA_STORNIRANOJ_ZBIRNOJ`, `OTP_BEZ_ZBIRNE_MALINA`, `HLADNJACA_LANAC_NEPOTPUN`, `ZBR_TURA_BEZ_PRIJEMNICE`, `ZBR_CLAN_VAN_1_1`, `ZBR_ISTI_BROJ_STORNIRANA_I_AKTIVNA` |
| `otkup` | `OTK_ZASTARELI_KONTEKST_DATUM`, `OTK_VEZA_NESAGLASNA_SA_OTP`, `OTK_PANEL_PREGAZIO_VEZU`, `OTK_AKTIVAN_BEZ_ZIVE_OTPREMNICE`, `SYNC_AUTO_OTPREMNICA_ZA_NEVEZANE`, `OTK_BROJ_ZBIRNE_ZAOSTAO`, `OTK_VOZAC_ZAOSTAO_ILI_LOSE_IZVUCEN`, `VOZAC_HLADNJACA_LITERAL`, `OTK_KOOPERANT_ZAOSTAO`, `OTK_KOOPERANT_NEPOSTOJECI_ID`, `KOOP_DUPLIKAT_AUTOKREIRANJEM`, `OTK_CENA_TIP_ZAOSTALI`, `ZAOSTALA_CENA_POSLE_DESELEKCIJE`, `OTK_BROJ_VAN_NIZA_ILI_PRAZAN`, `PWA_STANICA_I_DATUM`, `OTK_DUPLI_UNOS_BLOKA`, `DUPLI_LANAC_HLADNJACE_POSLE_GUARD2`, `OTK_KULTURAID_PO_VRSTI`, `OTK_GAJBE_KG_NEPROVERENO` |
| `novac` | `NOV_DUPLA_ISPLATA_RUCNO_I_IZVOD`, `NOV_DUPLI_RUCNI_UNOS`, `NOV_PREPLATA_OTKUPA`, `NOV_ISPLATA_OSIROCENA_STORNOM`, `NOV_ISPLATA_ODVEZANA_OD_STORNIRANOG_OTKUPA`, `OTK_ISPLACENO_POSLE_STORNA_NOVCA`, `NOV_STORNO_NOVCA_BEZ_LANCA`, `NOV_AVANS_RASPODELA_RAZDVOJENA_STORNOM`, `NOV_VEZA_NA_TUDJI_ILI_STORNIRAN_DOKUMENT`, `NOV_UPLATA_KUPCA_NA_STORNIRANU_ILI_PREPLACENU_FAKTURU`, `NOV_UPLATA_KUPCA_ODVEZANA_STORNOM_FAKTURE`, `FAK_UPLATA_NA_STORNIRANU_FAKTURU`, `NOV_IZVOD_RUCNO_KUPAC_OM_POGRESNO`, `NOV_IZVOD_POGRESAN_KOOPERANT`, `NOV_TIP_ISPLATE_PO_PODRAZUMEVANOJ`, `NOV_ZASTAREO_DATUM_FORME`, `NOV_IZVOD_DATUM_LOKAL`, `IZVOD_VAN_KNJIGE`, `IZVOD_BEZ_POZIVA_ODE_U_AVANS`, `IZVOD_DEDUPE_NEPROVERLJIV`, `REVERS_STORNIRAN_NOVAC_OSTAO`, `UPLATA_OM_NEVIDLJIVA_U_SALDU`, `POZICIONI_POMAK_TBLNOVAC` |
| `amb_pal_fak` | `AMB_AKTIVNA_NOGA_STORNIRANOG_DOKUMENTA`, `AMB_REVERS_VOZAC_NIJE_PAR_STANICE`, `AMB_REVERS_STANICA_ZASTARELA`, `AMB_REVERS_TIP_ZASTAREO`, `AMB_REVERS_DATUM_ZASTAREO`, `AMB_DATUM_NASLEDJEN_IZ_ISPRAVKE`, `AMB_BROJ_REVERSA_NE_PRATI_DATUM`, `AMB_OMULAZ_ZABORAVLJEN_SMER`, `AMB_BEZ_STORNA_KUPCI_OTPREMNICA_OMULAZ`, `AMB_KOMPENZACIJA_BEZ_STORNA`, `KUPAC_ZASTAREO_NA_KUPCI_OTPREMNICI_I_UPLATI`, `PAL_STAVKA_NA_STORNIRANOJ_PRIJEMNICI`, `PAL_BROJ_ZBIRNE_ZASTAREO`, `PAL_GAJBE_KG_NESAGLASNI`, `PAL_IDENTITET_RAZLICIT_OD_PRIJEMNICE`, `PAL_PREVEZIVANJE_PO_BROJU_TUDJI_KUPAC`, `PAL_DETACH_PO_BROJU_TUDJA_PRIJEMNICA`, `PAL_ZBIRNA_PREVEZANA_PO_BROJU`, `PAL_MESANI_KUPCI`, `FAK_*` (C002 nema faktura — tiho nad praznim tabelama) |
| `storno` | `OTP_AKTIVNA_SA_STORNIRANIM_BLOKOM`, `KOLATERALNI_STORNO_PO_BROJU`, `PRJ_STORNO_PO_BROJU_TUDJI_KUPAC`, `OTK_STORNO_PO_BROJU_TUDJI_BLOK`, `ZBR_PONISTENJE_OBARA_TUDJU_PRIJEMNICU`, `AUTO_USKRSNUCE_POSLE_STORNA`, `DOK_DELIMICNO_STORNIRAN`, `STORNO_KONTEKST_NEDOVRSEN`, `OTK_BROJZBIRNE_NA_STORNIRANOJ_ZBIRNOJ`, `FAK_OSIROCENO_ZASTARELO` |

Ista stvar koju vide dva modula (zbirna bez prijemnice: `ZBR_ZIVA_BEZ_PRIJEMNICE` i
`ZBR_TURA_BEZ_PRIJEMNICE`) prijavljuje se jednom, uz napomenu ko je još prijavio
(`pokretac.ISTO`).

**Nije kalibrisano na pravim podacima.** Pragovi (dani, kg, procenti) su izvedeni
iz koda i proverani na izmišljenim sveskama: čista sveska (i velika, 900 lanaca)
daje nula nalaza, a velika sveska sa podmetnutih 6 oblika greške iz operaterovog
ručnog spiska daje tačno te nalaze, sa tačnim predlogom.

## Veze u staroj svesci

Stara sveska nema tabele članstva ni `GeneracijaID`; veze su kolone na samim
dokumentima:

| Veza | Kako |
|---|---|
| otkup → otpremnica | `tblOtkup.OtpremnicaID` (po ID-ju) |
| otpremnica → zbirna | `tblOtpremnica.BrojZbirne` (po broju) |
| prijemnica → zbirna | `tblPrijemnica.BrojZbirne` (po broju) |
| novac → otkup | `tblNovac.OtkupID` |
| ambalaža → dokument | `DokumentTip` + `DokumentID` (reversi: broj dokumenta) |
| paletna stavka → prijemnica | `tblPaletaStavka.PrijemnicaID` (+ broj prijemnice i zbirne) |
| faktura stavka → prijemnica | `tblFakturaStavke.PrijemnicaID` |

Brojevi: otpremnica `x/ddmmyy[-n]` (x = broj stanice), zbirna `S` + broj
otpremnice, prijemnica `1/ddmmyy[-n]` (niz po kupcu i danu), otkup `x/ddmmyy[-n]`.

## Kod

| Fajl | Šta |
|---|---|
| `../provera_unosa.py` | komanda |
| `pokretac.py` | učita svesku, pusti module, skupi nalaze; pad modula je nalaz `MODUL_PAO` |
| `zajednicko.py` | `Kontekst` (tabele, indeksi), čitanje vrednosti, `parse_broj`, `nalaz` |
| `fixture.py` | čista izmišljena stara sveska — nad njom nijedna provera ne sme ništa da prijavi |
| `prijemnica.py`, `otpremnica_zbirna.py`, `otkup.py`, `novac.py`, `amb_pal_fak.py`, `storno.py` | provere po oblasti, svaka sa svojim `test_*.py` |

Testovi idu uz `python tools/migracija/migracija.py --self-test`.
