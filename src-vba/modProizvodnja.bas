Attribute VB_Name = "modProizvodnja"
Option Explicit

' ============================================================
' modProizvodnja -- Prerada 2.0, proizvodno jezgro (Faza A).
' Spec: docs/PRERADA_2_MODEL_I_PLAN.md.
'
' Faza A nosi:
'  - seed maticnih podataka (tipovi procesa, proizvodi iz tblKulture i
'    tblVrstaGotovihProizvoda) -- idempotentno, po prirodnom kljucu;
'  - MATERIJALIZACIJU legacy prerada u lager jedinice: svaka tblPrerada
'    (i stornirana -- LJ nasledjuje Stornirano, da istorijske utovarne i
'    fakturne stavke uvek imaju jedinicu) dobija tacno jednu LJ sa
'    IzvorTip=PRERADA, IzvorID=PreradaID; obrnuti pokazivac ide u
'    tblPrerada.LagerJedinicaID;
'  - backfill LagerJedinicaID na utovarnim/fakturnim stavkama iz mape
'    PreradaID -> LJ (deterministicki 1:1, zato SME u self-heal);
'  - RaspolozivoPoJedinici: JEDNA mapa raspolozivog po jedinici (fizicko
'    minus utovareno; ulazi procesa dolaze u Fazi B2, blokade u D).
'    Nema KgTrenutno kesa -- saldo se izvodi pri citanju, kao ambalaza.
'  - LjOznaka / LjRokTrajanja -- prikaz i rok (SNAPSHOT DatumIsteka).
'
' Writer-i sarze (OtvoriSarzu_TX / ZavrsiSarzu_TX) dolaze u Fazi B2.
'
' Fajl mora ostati 100% ASCII.
' ============================================================

Private Const SRC_MOD As String = "modProizvodnja"

' ============================================================
' SEED MATICNIH PODATAKA
' ============================================================

' Tipovi procesa: idempotentno po Sifra. Vraca broj dodatih redova.
' Redosled procesa se NE hardkoduje -- jedina kapija toka je opciona
' DozvoljenaUlaznaForma (prazno = sve forme).
Public Function SeedTipoviProcesa() As Long
    Const SRC As String = "modProizvodnja.SeedTipoviProcesa"
    If GetTable(TBL_TIPOVI_PROCESA) Is Nothing Then Exit Function

    Dim postoji As Object: Set postoji = CreateObject("Scripting.Dictionary")
    postoji.CompareMode = vbTextCompare
    Dim d As Variant, i As Long, cS As Long
    d = GetTableData(TBL_TIPOVI_PROCESA)
    If IsArray(d) Then
        cS = RequireColumnIndex(TBL_TIPOVI_PROCESA, COL_TPR_SIFRA, SRC)
        For i = 1 To UBound(d, 1)
            postoji(Trim$(CStr(nz(d(i, cS))))) = True
        Next i
    End If

    Dim sveze As String, smrz As String
    sveze = PRZ_FORMA_SVEZE
    smrz = PRZ_FORMA_SMRZNUTO
    Dim seed As Variant
    seed = Array( _
        Array("PRANJE", "Pranje", "Ne", "Ne", sveze, ""), _
        Array("SORTIRANJE", "Sortiranje", "Da", "Da", sveze & ";" & smrz, ""), _
        Array("KALIBRACIJA", "Kalibracija", "Da", "Da", sveze & ";" & smrz, ""), _
        Array("PREBIRANJE", "Prebiranje", "Da", "Da", sveze & ";" & smrz, ""), _
        Array("ZAMRZAVANJE", "Zamrzavanje", "Da", "Da", sveze, _
              "VREME_ULAZ;VREME_IZLAZ;TEMP_ROBE_ULAZ;TEMP_ROBE_IZLAZ;CILJNA_TEMP"), _
        Array("IZBIJANJE_KOSTICE", "Izbijanje ko" & ChrW(353) & "tice", "Da", "Da", sveze & ";" & smrz, ""), _
        Array("PAKOVANJE", "Pakovanje", "Da", "Ne", smrz & ";" & PRZ_FORMA_BULK, ""), _
        Array("PREPAKIVANJE", "Prepakivanje", "Ne", "Ne", "", ""), _
        Array("PASIRANJE", "Pasiranje", "Da", "Da", sveze & ";" & smrz, "BRIX"), _
        Array("BLOK", "Blok", "Da", "Da", smrz & ";" & PRZ_FORMA_PIRE & ";" & PRZ_FORMA_BULK, ""), _
        Array("ODMRZAVANJE", "Odmrzavanje", "Da", "Ne", smrz, ""), _
        Array("PRERADA_LEGACY", "Prerada (legacy)", "Da", "Ne", "", ""))

    Dim r As Variant, n As Long
    For Each r In seed
        If Not postoji.Exists(CStr(r(0))) Then
            PrzAppendRow TBL_TIPOVI_PROCESA, _
                Array(COL_TPR_SIFRA, COL_TPR_NAZIV, COL_TPR_MENJA_PROIZVOD, _
                      COL_TPR_ZAHTEVA_OPREMU, COL_TPR_ULAZNA_FORMA, _
                      COL_TPR_OBAVEZNI_PARAM, COL_TPR_AKTIVAN), _
                Array(r(0), r(1), r(2), r(3), r(4), r(5), STATUS_AKTIVAN)
            postoji(CStr(r(0))) = True
            n = n + 1
        End If
    Next r
    SeedTipoviProcesa = n
End Function

' Proizvodi: po jedan SVEZE proizvod za svaku VrstaVoca iz tblKulture
' (IzvorTip=KULTURA) i po jedan prodajni za svaki TipGotovogProizvoda iz
' tblVrstaGotovihProizvoda (IzvorTip=VGP). Idempotentno po IzvorTip+IzvorKljuc.
' RokMeseci se NE kopira: rok se racuna kroz modUtovar.RokIstekaZaTip samo
' pri nastanku jedinice (snapshot DatumIsteka).
Public Function SeedProizvodi() As Long
    Const SRC As String = "modProizvodnja.SeedProizvodi"
    If GetTable(TBL_PROIZVODI) Is Nothing Then Exit Function

    Dim kult As Object: Set kult = ProizvodiPoIzvoru(PRZ_IZVOR_KULTURA)
    Dim vgp As Object: Set vgp = ProizvodiPoIzvoru(PRZ_IZVOR_VGP)
    Dim n As Long, d As Variant, i As Long, k As String

    ' --- sveze voce iz tblKulture (distinct VrstaVoca) ---
    If Not GetTable(TBL_KULTURE) Is Nothing Then
        d = GetTableData(TBL_KULTURE)
        If IsArray(d) Then
            Dim cV As Long
            cV = GetColumnIndex(TBL_KULTURE, "VrstaVoca")
            If cV > 0 Then
                For i = 1 To UBound(d, 1)
                    k = Trim$(CStr(nz(d(i, cV))))
                    If Len(k) > 0 Then
                        If Not kult.Exists(k) Then
                            PrzAppendRow TBL_PROIZVODI, _
                                Array(COL_PRZ_ID, COL_PRZ_VRSTA, COL_PRZ_NAZIV, COL_PRZ_FORMA, _
                                      COL_PRZ_PRODAJNI, COL_PRZ_IZVOR_TIP, COL_PRZ_IZVOR_KLJUC, _
                                      COL_PRZ_AKTIVAN), _
                                Array(GetNextID(TBL_PROIZVODI, COL_PRZ_ID, PRZ_ID_PREFIKS), _
                                      k, k & ", sve" & ChrW(382) & "e", PRZ_FORMA_SVEZE, _
                                      "Ne", PRZ_IZVOR_KULTURA, k, STATUS_AKTIVAN)
                            kult(k) = "seed"
                            n = n + 1
                        End If
                    End If
                Next i
            End If
        End If
    End If

    ' --- gotovi proizvodi iz tblVrstaGotovihProizvoda ---
    If Not GetTable(TBL_VRSTA_GP) Is Nothing Then
        d = GetTableData(TBL_VRSTA_GP)
        If IsArray(d) Then
            Dim cT As Long, cA As Long, akt As String
            cT = GetColumnIndex(TBL_VRSTA_GP, COL_VGP_TIP)
            cA = GetColumnIndex(TBL_VRSTA_GP, "Aktivan")
            If cT > 0 Then
                For i = 1 To UBound(d, 1)
                    k = Trim$(CStr(nz(d(i, cT))))
                    If Len(k) > 0 Then
                        If Not vgp.Exists(k) Then
                            akt = STATUS_AKTIVAN
                            If cA > 0 Then
                                If Len(Trim$(CStr(nz(d(i, cA))))) > 0 Then akt = Trim$(CStr(nz(d(i, cA))))
                            End If
                            PrzAppendRow TBL_PROIZVODI, _
                                Array(COL_PRZ_ID, COL_PRZ_VRSTA, COL_PRZ_NAZIV, COL_PRZ_FORMA, _
                                      COL_PRZ_PRODAJNI, COL_PRZ_IZVOR_TIP, COL_PRZ_IZVOR_KLJUC, _
                                      COL_PRZ_AKTIVAN), _
                                Array(GetNextID(TBL_PROIZVODI, COL_PRZ_ID, PRZ_ID_PREFIKS), _
                                      "", k, PRZ_FORMA_SMRZNUTO, _
                                      "Da", PRZ_IZVOR_VGP, k, akt)
                            vgp(k) = "seed"
                            n = n + 1
                        End If
                    End If
                Next i
            End If
        End If
    End If
    SeedProizvodi = n
End Function

' Oba seed-a; vraca kratak sazetak za log.
Public Function SeedProizvodnjaMaticni() As String
    Dim t As Long, p As Long
    t = SeedTipoviProcesa()
    p = SeedProizvodi()
    SeedProizvodnjaMaticni = "tipova procesa +" & CStr(t) & ", proizvoda +" & CStr(p)
End Function

' Mapa IzvorKljuc -> ProizvodID za dati IzvorTip (prvi pojav pobedjuje).
Public Function ProizvodiPoIzvoru(ByVal izvorTip As String) As Object
    Const SRC As String = "modProizvodnja.ProizvodiPoIzvoru"
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set ProizvodiPoIzvoru = d
    If GetTable(TBL_PROIZVODI) Is Nothing Then Exit Function

    Dim data As Variant, i As Long, cId As Long, cT As Long, cK As Long, k As String
    data = GetTableData(TBL_PROIZVODI)
    If Not IsArray(data) Then Exit Function
    cId = RequireColumnIndex(TBL_PROIZVODI, COL_PRZ_ID, SRC)
    cT = RequireColumnIndex(TBL_PROIZVODI, COL_PRZ_IZVOR_TIP, SRC)
    cK = RequireColumnIndex(TBL_PROIZVODI, COL_PRZ_IZVOR_KLJUC, SRC)
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(CStr(nz(data(i, cT)))), izvorTip, vbTextCompare) = 0 Then
            k = Trim$(CStr(nz(data(i, cK))))
            If Len(k) > 0 Then
                If Not d.Exists(k) Then d.Add k, Trim$(CStr(nz(data(i, cId))))
            End If
        End If
    Next i
End Function

' ProizvodID za tip gotovog proizvoda (legacy naziv), "" kad ga nema.
Public Function ProizvodIDZaTipGP(ByVal tipGP As String) As String
    Dim m As Object: Set m = ProizvodiPoIzvoru(PRZ_IZVOR_VGP)
    If m.Exists(Trim$(tipGP)) Then ProizvodIDZaTipGP = CStr(m(Trim$(tipGP)))
End Function

' ============================================================
' LAGER JEDINICE -- materijalizacija legacy prerada
' ============================================================

' Mapa IzvorID -> LagerJedinicaID za dati IzvorTip. Uzima SVE redove (i
' stornirane): postojanje jedinice je pitanje identiteta, ne stanja.
' Dupli IzvorID daje "" (P4 ga prijavljuje; ovde se nista ne bira naslepo).
Public Function LjPoIzvoru(ByVal izvorTip As String) As Object
    Const SRC As String = "modProizvodnja.LjPoIzvoru"
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set LjPoIzvoru = d
    If GetTable(TBL_LAGER_JEDINICE) Is Nothing Then Exit Function

    Dim data As Variant, i As Long, cId As Long, cT As Long, cIz As Long, k As String
    data = GetTableData(TBL_LAGER_JEDINICE)
    If Not IsArray(data) Then Exit Function
    cId = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_ID, SRC)
    cT = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_IZVOR_TIP, SRC)
    cIz = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_IZVOR_ID, SRC)
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(CStr(nz(data(i, cT)))), izvorTip, vbTextCompare) = 0 Then
            k = Trim$(CStr(nz(data(i, cIz))))
            If Len(k) > 0 Then
                If d.Exists(k) Then
                    d(k) = ""
                Else
                    d.Add k, Trim$(CStr(nz(data(i, cId))))
                End If
            End If
        End If
    Next i
End Function

' Jedinstvena stanica-hladnjaca (tblStanice.JeHladnjaca=Da), inace "".
' Vise hladnjaca = ne bira se naslepo; P6 prijavljuje jedinicu bez stanice.
Public Function HladnjacaStanicaID() As String
    Dim d As Variant, i As Long, cId As Long, cH As Long, n As Long, id As String
    If GetTable(TBL_STANICE) Is Nothing Then Exit Function
    cId = GetColumnIndex(TBL_STANICE, COL_STA_ID)
    cH = GetColumnIndex(TBL_STANICE, COL_STA_JE_HLADNJACA)
    If cId = 0 Or cH = 0 Then Exit Function
    d = GetTableData(TBL_STANICE)
    If Not IsArray(d) Then Exit Function
    For i = 1 To UBound(d, 1)
        If UCase$(Trim$(CStr(nz(d(i, cH))))) = "DA" Then
            n = n + 1
            id = Trim$(CStr(nz(d(i, cId))))
        End If
    Next i
    If n = 1 Then HladnjacaStanicaID = id
End Function

' Materijalizuje SVE prerade bez lager jedinice. Idempotentno: prerada koja
' vec ima LJ dobija samo obrnuti pokazivac ako mu fali. Dupli PreradaID se
' preskace (korupcija -- P4). Vraca broj NOVIH jedinica.
Public Function MaterijalizujLegacyPrerade() As Long
    Const SRC As String = "modProizvodnja.MaterijalizujLegacyPrerade"
    If GetTable(TBL_LAGER_JEDINICE) Is Nothing Then Exit Function
    If GetTable(TBL_PRERADA) Is Nothing Then Exit Function

    Dim d As Variant
    d = GetTableData(TBL_PRERADA)
    If Not IsArray(d) Then Exit Function

    Dim cId As Long, cLj As Long, i As Long, pid As String, n As Long
    cId = RequireColumnIndex(TBL_PRERADA, COL_PRE_ID, SRC)
    cLj = GetColumnIndex(TBL_PRERADA, COL_PRE_LJ_ID)

    ' Dupli PreradaID = korupcija: takva prerada ne dobija jedinicu.
    Dim brojac As Object: Set brojac = CreateObject("Scripting.Dictionary")
    brojac.CompareMode = vbTextCompare
    For i = 1 To UBound(d, 1)
        pid = Trim$(CStr(nz(d(i, cId))))
        If Len(pid) > 0 Then
            If brojac.Exists(pid) Then brojac(pid) = CLng(brojac(pid)) + 1 Else brojac.Add pid, 1
        End If
    Next i

    Dim mapa As Object: Set mapa = LjPoIzvoru(LJ_IZVOR_PRERADA)
    Dim prz As Object: Set prz = ProizvodiPoIzvoru(PRZ_IZVOR_VGP)
    Dim stanica As String: stanica = HladnjacaStanicaID()
    Dim ljID As String

    For i = 1 To UBound(d, 1)
        pid = Trim$(CStr(nz(d(i, cId))))
        If Len(pid) > 0 Then
            If CLng(brojac(pid)) > 1 Then
                LogWarn SRC, "Dupli PreradaID " & pid & " -- jedinica se ne materijalizuje (P4)."
            ElseIf mapa.Exists(pid) Then
                If cLj > 0 And Len(CStr(mapa(pid))) > 0 Then
                    If Len(Trim$(CStr(nz(d(i, cLj))))) = 0 Then
                        RequireUpdateCell TBL_PRERADA, i, COL_PRE_LJ_ID, CStr(mapa(pid)), SRC
                    End If
                End If
            Else
                ljID = LjIzPreradeReda(d, i, prz, stanica, SRC)
                mapa.Add pid, ljID
                If cLj > 0 Then RequireUpdateCell TBL_PRERADA, i, COL_PRE_LJ_ID, ljID, SRC
                n = n + 1
            End If
        End If
    Next i
    MaterijalizujLegacyPrerade = n
End Function

' Materijalizuje JEDNU preradu (zove SavePrerada_TX u istoj transakciji).
' Prerada mora postojati tacno jednom. Vraca LagerJedinicaID (postojeci ili nov).
Public Function MaterijalizujPreradu(ByVal preradaID As String, _
                                     ByVal SRC As String) As String
    Dim hits As Collection
    Set hits = FindRows(TBL_PRERADA, COL_PRE_ID, Trim$(preradaID))
    If hits.count <> 1 Then
        Err.Raise vbObjectError + 7430, SRC, _
                  "Prerada " & preradaID & " nije nadjena tacno jednom (" & _
                  CStr(hits.count) & ") -- lager jedinica se ne pravi."
    End If
    Dim mapa As Object: Set mapa = LjPoIzvoru(LJ_IZVOR_PRERADA)
    Dim ljID As String
    If mapa.Exists(Trim$(preradaID)) Then
        If Len(CStr(mapa(Trim$(preradaID)))) = 0 Then
            Err.Raise vbObjectError + 7431, SRC, _
                      "Prerada " & preradaID & " vec ima VISE lager jedinica (P4)."
        End If
        ljID = CStr(mapa(Trim$(preradaID)))
    Else
        Dim d As Variant
        d = GetTableData(TBL_PRERADA)
        ljID = LjIzPreradeReda(d, CLng(hits(1)), ProizvodiPoIzvoru(PRZ_IZVOR_VGP), _
                               HladnjacaStanicaID(), SRC)
    End If
    If GetColumnIndex(TBL_PRERADA, COL_PRE_LJ_ID) > 0 Then
        RequireUpdateCell TBL_PRERADA, CLng(hits(1)), COL_PRE_LJ_ID, ljID, SRC
    End If
    MaterijalizujPreradu = ljID
End Function

' Jezgro materijalizacije: LJ red iz reda tblPrerada. Kolone koje sveska
' nema (pre nadogradnje) ostaju prazne, ne obaraju upis.
Private Function LjIzPreradeReda(ByRef d As Variant, ByVal r As Long, _
                                 ByVal prz As Object, ByVal stanica As String, _
                                 ByVal SRC As String) As String
    Dim cId As Long, cBroj As Long, cGod As Long, cDat As Long, cNeto As Long, cSt As Long
    cId = RequireColumnIndex(TBL_PRERADA, COL_PRE_ID, SRC)
    cBroj = RequireColumnIndex(TBL_PRERADA, COL_PRE_BROJ, SRC)
    cGod = RequireColumnIndex(TBL_PRERADA, COL_PRE_GODINA, SRC)
    cDat = RequireColumnIndex(TBL_PRERADA, COL_PRE_DATUM, SRC)
    cNeto = RequireColumnIndex(TBL_PRERADA, COL_PRE_NETO_IZLAZ, SRC)
    cSt = RequireColumnIndex(TBL_PRERADA, COL_STORNIRANO, SRC)

    Dim tipGP As String, pid As String, proizvod As String
    pid = Trim$(CStr(nz(d(r, cId))))
    tipGP = Trim$(CStr(nz(PrzCell(d, r, GetColumnIndex(TBL_PRERADA, COL_PRE_TIP_GP)))))
    If prz.Exists(tipGP) Then proizvod = CStr(prz(tipGP))

    Dim datum As Variant, rok As Variant
    datum = Empty
    If IsDate(d(r, cDat)) Then datum = CDate(d(r, cDat))
    rok = PrzCell(d, r, GetColumnIndex(TBL_PRERADA, COL_PRE_ROK))
    If Not IsDate(rok) Then
        ' Stari lot bez snapshota: isti fallback koji koristi stampa
        ' utovarne liste (RokIstekaZaTip po TEKUCEM pravilu), jednom, sad.
        rok = Empty
        If IsDate(datum) Then rok = modUtovar.RokIstekaZaTip(tipGP, CDate(datum))
    Else
        rok = CDate(rok)
    End If

    Dim storno As String
    storno = ""
    If UCase$(Trim$(CStr(nz(d(r, cSt))))) = "DA" Then storno = "Da"

    Dim ljID As String
    ljID = GetNextID(TBL_LAGER_JEDINICE, COL_LJ_ID, LJ_ID_PREFIKS)
    PrzAppendRow TBL_LAGER_JEDINICE, _
        Array(COL_LJ_ID, COL_LJ_BROJ, COL_LJ_GODINA, COL_LJ_TIP, COL_LJ_PROIZVOD, _
              COL_LJ_KLASA, COL_LJ_KALIBRACIJA, COL_LJ_KG_POCETNO, COL_LJ_LOT, _
              COL_LJ_TIP_KUTIJE, COL_LJ_KUTIJE, COL_LJ_TIP_KESE, COL_LJ_KESE, _
              COL_LJ_TEZINA_PALETE, COL_LJ_BRUTO, COL_LJ_DATUM, COL_LJ_ROK, _
              COL_LJ_STANICA, COL_LJ_IZVOR_TIP, COL_LJ_IZVOR_ID, COL_LJ_NAPOMENA, _
              COL_STORNIRANO), _
        Array(ljID, NzL(d(r, cBroj)), NzL(d(r, cGod)), LJ_TIP_PALETA, proizvod, _
              "", "", NzD(d(r, cNeto)), "", _
              PrzTxt(d, r, COL_PRE_TIP_KUTIJE), PrzNum(d, r, COL_PRE_KUTIJE), _
              PrzTxt(d, r, COL_PRE_TIP_KESE), PrzNum(d, r, COL_PRE_KESE), _
              PrzNum(d, r, COL_PRE_TEZINA_PALETE), PrzNum(d, r, COL_PRE_BRUTO), _
              datum, rok, _
              stanica, LJ_IZVOR_PRERADA, pid, "", _
              storno)
    LjIzPreradeReda = ljID
End Function

' Backfill LagerJedinicaID na utovarnim i fakturnim stavkama iz mape
' PreradaID -> LJ. Nista se ne izmislja (mapa je 1:1 iz materijalizacije);
' stavka bez PreradaID ostaje bez LJ i P5 je prijavljuje. Vraca broj upisa.
Public Function BackfillLjNaStavkama() As Long
    Dim n As Long
    n = BackfillLjKolona(TBL_UTOVAR_STAVKE, COL_UTS_PRERADA_ID, COL_UTS_LJ_ID)
    n = n + BackfillLjKolona(TBL_FAKTURA_STAVKE, COL_FS_PRERADA_ID, COL_FS_LJ_ID)
    BackfillLjNaStavkama = n
End Function

Private Function BackfillLjKolona(ByVal tbl As String, ByVal colPre As String, _
                                  ByVal colLj As String) As Long
    Const SRC As String = "modProizvodnja.BackfillLjKolona"
    If GetTable(tbl) Is Nothing Then Exit Function
    Dim cPre As Long, cLj As Long
    cPre = GetColumnIndex(tbl, colPre)
    cLj = GetColumnIndex(tbl, colLj)
    If cPre = 0 Or cLj = 0 Then Exit Function

    Dim d As Variant, i As Long, pre As String, n As Long
    d = GetTableData(tbl)
    If Not IsArray(d) Then Exit Function
    Dim mapa As Object: Set mapa = LjPoIzvoru(LJ_IZVOR_PRERADA)
    For i = 1 To UBound(d, 1)
        pre = Trim$(CStr(nz(d(i, cPre))))
        If Len(pre) > 0 And Len(Trim$(CStr(nz(d(i, cLj))))) = 0 Then
            If mapa.Exists(pre) Then
                If Len(CStr(mapa(pre))) > 0 Then
                    RequireUpdateCell tbl, i, colLj, CStr(mapa(pre)), SRC
                    n = n + 1
                End If
            End If
        End If
    Next i
    BackfillLjKolona = n
End Function

' ============================================================
' RASPOLOZIVO -- jedna mapa za mrezu, writer i storno
' ============================================================

' LagerJedinicaID -> raspolozivo kg, nad NESTORNIRANIM jedinicama.
'   raspolozivo = fizicko - utovareno   (Faza A)
'                - ulazi procesa        (Faza B2)
'                - blokirano            (Faza D)
' Fizicko: KgPocetno; za IzvorTip=PALETA ziva tblPaleta.NetoKg (header
' palete sme da se menja dok paleta nije usla u proces). Utovareno: za
' legacy lot preko UtovarenoPoPreradi(IzvorID) -- do B1 utovar zna samo
' PreradaID. Ostecene stavke blokiraju prodaju u writer-u (#248
' RequireStavkeKonzistentne); B1 tu kapiju prevodi na LJ kljuc.
Public Function RaspolozivoPoJedinici() As Object
    Const SRC As String = "modProizvodnja.RaspolozivoPoJedinici"
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set RaspolozivoPoJedinici = d
    If GetTable(TBL_LAGER_JEDINICE) Is Nothing Then Exit Function

    Dim lj As Variant
    lj = GetTableData(TBL_LAGER_JEDINICE)
    If Not IsArray(lj) Then Exit Function
    lj = ExcludeStornirano(lj, TBL_LAGER_JEDINICE)
    If Not IsArray(lj) Then Exit Function

    Dim cId As Long, cKg As Long, cT As Long, cIz As Long
    cId = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_ID, SRC)
    cKg = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_KG_POCETNO, SRC)
    cT = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_IZVOR_TIP, SRC)
    cIz = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_IZVOR_ID, SRC)

    Dim utov As Object: Set utov = modUtovar.UtovarenoPoPreradi()
    Dim palNeto As Object: Set palNeto = PaletaNetoMapa()

    Dim i As Long, id As String, tip As String, izvor As String
    Dim fiz As Double, pot As Double
    For i = 1 To UBound(lj, 1)
        id = Trim$(CStr(nz(lj(i, cId))))
        If Len(id) > 0 Then
            tip = UCase$(Trim$(CStr(nz(lj(i, cT)))))
            izvor = Trim$(CStr(nz(lj(i, cIz))))
            fiz = NzD(lj(i, cKg))
            If tip = LJ_IZVOR_PALETA And palNeto.Exists(izvor) Then fiz = CDbl(palNeto(izvor))
            pot = 0#
            If tip = LJ_IZVOR_PRERADA And utov.Exists(izvor) Then pot = CDbl(utov(izvor))
            If d.Exists(id) Then
                d(id) = -1#           ' dupli LagerJedinicaID: nikad raspoloziva
            Else
                d.Add id, fiz - pot
            End If
        End If
    Next i
End Function

Public Function RaspolozivoKg(ByVal ljID As String) As Double
    Dim m As Object: Set m = RaspolozivoPoJedinici()
    If m.Exists(Trim$(ljID)) Then RaspolozivoKg = CDbl(m(Trim$(ljID)))
End Function

' PaletaID -> NetoKg (sve palete; stornirana paleta ce od B2 nositi i
' storniranu LJ, pa je ovde ne treba posebno filtrirati).
Private Function PaletaNetoMapa() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set PaletaNetoMapa = d
    If GetTable(TBL_PALETA) Is Nothing Then Exit Function
    Dim data As Variant, i As Long, cId As Long, cN As Long, k As String
    data = GetTableData(TBL_PALETA)
    If Not IsArray(data) Then Exit Function
    cId = GetColumnIndex(TBL_PALETA, COL_PAL_ID)
    cN = GetColumnIndex(TBL_PALETA, COL_PAL_NETO)
    If cId = 0 Or cN = 0 Then Exit Function
    For i = 1 To UBound(data, 1)
        k = Trim$(CStr(nz(data(i, cId))))
        If Len(k) > 0 Then
            If Not d.Exists(k) Then d.Add k, NzD(data(i, cN))
        End If
    Next i
End Function

' ============================================================
' PRIKAZ I ROK
' ============================================================

' Oznaka jedinice za operatera: "PRE 51/2026" (legacy lot), "PAL 31/2026"
' (paleta), "LJ 12/2026" (rodjena u procesu). Nepoznata jedinica vraca ID.
Public Function LjOznaka(ByVal ljID As String) As String
    Dim hits As Collection, r As Long, d As Variant
    LjOznaka = Trim$(ljID)
    If GetTable(TBL_LAGER_JEDINICE) Is Nothing Then Exit Function
    Set hits = FindRows(TBL_LAGER_JEDINICE, COL_LJ_ID, Trim$(ljID))
    If hits.count <> 1 Then Exit Function
    r = CLng(hits(1))
    d = GetTableData(TBL_LAGER_JEDINICE)
    Dim broj As Long, god As Long, pref As String
    broj = NzL(PrzCell(d, r, GetColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_BROJ)))
    god = NzL(PrzCell(d, r, GetColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_GODINA)))
    If broj = 0 Then Exit Function
    Select Case UCase$(Trim$(CStr(nz(PrzCell(d, r, GetColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_IZVOR_TIP))))))
        Case LJ_IZVOR_PRERADA: pref = "PRE"
        Case LJ_IZVOR_PALETA:  pref = "PAL"
        Case Else:             pref = "LJ"
    End Select
    LjOznaka = pref & " " & CStr(broj) & "/" & CStr(god)
End Function

' Rok trajanja jedinice: SNAPSHOT DatumIsteka. Stara legacy jedinica bez
' snapshota dobija fallback po tekucem pravilu (isti kao stampa utovarne
' liste); jedinica bez datuma nastanka daje Empty. Nikad se ne racuna iz
' tekuceg pravila kad snapshot postoji.
Public Function LjRokTrajanja(ByVal ljID As String) As Variant
    LjRokTrajanja = Empty
    If GetTable(TBL_LAGER_JEDINICE) Is Nothing Then Exit Function
    Dim hits As Collection, r As Long, d As Variant
    Set hits = FindRows(TBL_LAGER_JEDINICE, COL_LJ_ID, Trim$(ljID))
    If hits.count <> 1 Then Exit Function
    r = CLng(hits(1))
    d = GetTableData(TBL_LAGER_JEDINICE)
    Dim rok As Variant
    rok = PrzCell(d, r, GetColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_ROK))
    If IsDate(rok) Then
        LjRokTrajanja = CDate(rok)
        Exit Function
    End If
    If UCase$(Trim$(CStr(nz(PrzCell(d, r, GetColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_IZVOR_TIP)))))) _
       <> LJ_IZVOR_PRERADA Then Exit Function
    Dim datum As Variant, tipGP As String
    datum = PrzCell(d, r, GetColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_DATUM))
    If Not IsDate(datum) Then Exit Function
    tipGP = Trim$(CStr(nz(LookupValue(TBL_PRERADA, COL_PRE_ID, _
                          Trim$(CStr(nz(PrzCell(d, r, GetColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_IZVOR_ID))))), _
                          COL_PRE_TIP_GP))))
    LjRokTrajanja = modUtovar.RokIstekaZaTip(tipGP, CDate(datum))
End Function

' ============================================================
' POMOCNE
' ============================================================

' Append po IMENU kolone (mirror PalAppendRow): kolona koje nema se preskace,
' pa je upis bezbedan pod schema drift-om. Pada glasno ako append ne uspe.
Private Sub PrzAppendRow(ByVal tblName As String, ByVal cols As Variant, ByVal vals As Variant)
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then
        Err.Raise vbObjectError + 7432, "modProizvodnja.PrzAppendRow", "Nema tabele: " & tblName
    End If
    Dim n As Long: n = lo.ListColumns.count
    Dim rowData() As Variant
    ReDim rowData(0 To n - 1)
    Dim i As Long, idx As Long
    For i = LBound(cols) To UBound(cols)
        idx = GetColumnIndex(tblName, CStr(cols(i)))
        If idx >= 1 And idx <= n Then rowData(idx - 1) = vals(i)
    Next i
    If AppendRow(tblName, rowData) = 0 Then
        Err.Raise vbObjectError + 7433, "modProizvodnja.PrzAppendRow", _
                  "AppendRow nije uspeo za tabelu: " & tblName
    End If
End Sub

' Celija ili Empty kad kolona ne postoji (c = 0) ili je van opsega.
Private Function PrzCell(ByRef d As Variant, ByVal r As Long, ByVal c As Long) As Variant
    PrzCell = Empty
    If c <= 0 Then Exit Function
    If Not IsArray(d) Then Exit Function
    If r < 1 Or r > UBound(d, 1) Then Exit Function
    If c > UBound(d, 2) Then Exit Function
    PrzCell = d(r, c)
End Function

Private Function PrzTxt(ByRef d As Variant, ByVal r As Long, ByVal colName As String) As String
    PrzTxt = Trim$(CStr(nz(PrzCell(d, r, GetColumnIndex(TBL_PRERADA, colName)))))
End Function

' Broj ili Empty: prazno ostaje prazno (ne izmislja se nula).
Private Function PrzNum(ByRef d As Variant, ByVal r As Long, ByVal colName As String) As Variant
    Dim v As Variant
    v = PrzCell(d, r, GetColumnIndex(TBL_PRERADA, colName))
    If IsNumeric(v) And Len(Trim$(CStr(nz(v)))) > 0 Then
        PrzNum = CDbl(v)
    Else
        PrzNum = Empty
    End If
End Function
