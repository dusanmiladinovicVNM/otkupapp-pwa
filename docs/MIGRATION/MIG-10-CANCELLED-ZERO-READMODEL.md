# MIG-10 — storno otkupa sa nultom kolicinom (read-only kompatibilnost)

**Status:** PATCH NA IZDVOJENOJ GRANI; VBA COMPILE/UAT NA STVARNOM EXCELU NISU IZVRSENI; PRODUKCIJA NO-GO.

## Dokaz i granica
- Referentni data-transfer: MIG-09, 47 kanonskih tabela, 9.610 redova u 25 migriranih tabela, VBA binary preserved.
- MIG-10.0: jedna istorijska anomalija od 1.702 otkupna lista (source ordinal 123).
- MIG-10.1: dokument je STORNIRAN, kolicina 0, cena pozitivna, ambalaza pozitivna, original == staging.
- Stari OtpremnicaID/BrojZbirne su istorijske reference, NE dokaz aktivnog kanonskog clanstva.
- Nema promene istorijskih podataka, staginga, ambalazne knjige, isplata ili VBA writer-a.

## Minimalni domain rez
1. `modOtkup.StavkeOtkupaRedovi`: ista stroga validacija za sve osim `Kolicina=0` uz `tblOtkup.Stornirano=Da`, i dalje sa postojecim stvarnim zaglavljem, pozitivnom cenom, nenegativnom ambalazom, jedinstvenim OtkupID-om i bez dupliranja klase. Granica je READ, ne CREATE; ne skida status storna.
2. `modOtkup.VrednostOtkupa` i `CreateOtkup_TX`: nepromenjeni, aktivna komanda ostaje stroga. Na storniranom istorijskom dokumentu `VrednostOtkupa` jos uvek odbija nulu; taj API nije istorijski pregled.
3. `modOtkup.KgOtkupaZaDan` + `modOtkupUI.CountForDate(..., samoAktivni=True)`: KPI i njegov broj dokumenata ukljucuju samo aktivni otkup. Ostale vrste dokumenata zadrzavaju raniju semantiku broja.
4. `modDokumenta.AktivnoOtpClanstvoPoKanonu`: istorijske veze uz stornirane otpremnice ostaju u tabeli, ali aktivna otpremnica koja sadrzi storniran otkup pada po imenu. Bez tihog racunanja nula kg ili pozitivnih istorijskih gajbi kao aktivnog fizickog izvora.
5. Izvestaji i otvorene obaveze vec filtriraju stornirana zaglavlja: `modNovac.GetOpenOtkupi`, `modIzvestaj.ReportSaldoOM`, `ReportKarticaKooperanta`, `ReportOtkupListe`, `ReportZbirniOM`, `modStammdatenSync.OtkupZaglavljaZaIzvoz`. `modPrint.SpecifikacijaBlokovaRedovi` ima izricito odbijanje storniranog izvora aktivne otpremnice.

## Test i verifikacija
- `modBusinessFlowProTests.Test_MIG10_CancelledZeroQuantity` se izvrsava kroz `RunBusinessFlowProSuite` iskljucivo na namenskoj QA svesci, NIKADA na originalnoj klijentskoj svesci. Fixture je transakciono snapshotovan i rollback-uje promene.
- Pokriva aktivan otkup sa nultom kolicinom (1906), storniran sa nultom kolicinom i pozitivnom cenom (read PASS), storniran sa negativnom kolicinom (1906), storniran sa nultom cenom (1906), istorijski bulk agregat, otvorene isplate, aktivni KPI, istorijsku listu i strogu `VrednostOtkupa` (1904).
- Vec postojece regresije za dupli OtkupID, nedostajuce stavke i regularan CREATE zahtevaju ponovno izvrsavanje.
- Potreban dodatni UAT na izolovanoj KOPIJI MIG-09: otvaranje `Otkazane`, kontrola istorijskog reda; aktivna lista, izvestaji, placanja, specifikacija otpremnice; broj dokumenata nepromenjen; original/staging hash i sve migrirane vrednosti nepromenjene.
- Poseban MIG-11 gate: proveriti `OtpremnicaIzvori`/`ZbirnaIzvori` i istorijske veze, bez izmisljanja ili brisanja odnosa. MIG-12 AMB-10 i MIG-13 finansije ostaju BLOKIRANI do svojih domain gate-ova.

## GO/NO-GO
- Kod na grani je samo kandidat; testovi nisu izvrseni u ovoj izradi.
- **Nema GO** pre VBA Compile, Save, Close, Reopen, lokalni smoke svih navedenih ruta i kontrola efekata pre/posle na lokalnoj kopiji.
- U migrirani `.xlsm` se ne upisuje direktno: koristiti postojeci `modVbaTools.ImportAllVBA` samo na QA kopiji i ponoviti Compile/Save/Reopen.
