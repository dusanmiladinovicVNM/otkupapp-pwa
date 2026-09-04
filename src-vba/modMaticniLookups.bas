Attribute VB_Name = "modMaticniLookups"
'Attribute VB_Name = "modMaticniLookups"
Option Explicit

' ============================================================
' modMaticniLookups - jedinstveni (data-driven) meni "Maticni podaci"
'
' Ceo meni frmMaticniPodaci se gradi iz JEDNE registracije sekcija
' (MaticniSekcijeGrupisano). Za svaku sekciju se dinamicki kreira dugme
' (Controls.Add) pa se njegov klik hvata preko clsLookupMenuBtn
' (WithEvents). Tako:
'   - frmMaticniPodaci.frx se NE dira,
'   - sve sekcije (postojece + nove) idu kroz isti mehanizam,
'   - dodavanje nove sekcije = jedan red u MaticniSekcije + Case u
'     frmStammdaten.
'
' Postojeca staticna dugmad na formi se sakrivaju (ostaju u .frx kao
' fallback ako dinamicka izgradnja ne uspe).
'
' Otvaranje sekcije ide kroz frmMaticniPodaci.OpenSekcija (koji vec
' ispravno upravlja m_IsOpeningChild flagom, da se meni ne zatvori).
' ============================================================

Private mWrappers As Collection   ' clsLookupMenuBtn instance (drzi WithEvents zivim)
Private mBtns As Collection       ' MSForms.CommandButton kontrole (za reset/hover)
Private mHoverNm As String        ' poslednje hover-ovano dugme (anti-flicker)

Private Const STATIC_BTNS As String = _
    "btnKooperanti;btnStanice;btnKupci;btnVozaci;btnArtikli;btnParcele"

' Registracija svih sekcija: Array(Naziv u meniju, Tag za frmStammdaten).
' Redosled ovde = redosled u meniju.
Public Function MaticniSekcije() As Variant
    ' Ravna lista (zadrzana radi kompatibilnosti) izvedena iz grupisane
    ' registracije MaticniSekcijeGrupisano - jedan izvor istine.
    Dim groups As Variant
    groups = MaticniSekcijeGrupisano()

    Dim out As Collection
    Set out = New Collection

    Dim gi As Long, ii As Long
    Dim grp As Variant, items As Variant
    For gi = LBound(groups) To UBound(groups)
        grp = groups(gi)
        items = grp(1)
        For ii = LBound(items) To UBound(items)
            out.Add items(ii)
        Next ii
    Next gi

    Dim a() As Variant, k As Long
    ReDim a(0 To out.count - 1)
    For k = 1 To out.count
        a(k - 1) = out(k)
    Next k
    MaticniSekcije = a
End Function

' Grupisana registracija sekcija menija "Maticni podaci".
' Vraca Array(Array(GrupaNaziv, Array(Array(Caption, Tag), ...)), ...).
' Redosled grupa i stavki = redosled u meniju. Pakovanje (ambalaza, palete,
' kutije, kese) je svoja grupa - vizuelno podredjena osnovnim sifarnicima,
' a ne ravnopravno sa Kooperanti/Kupci/Stanice.
' Dodavanje sekcije = jedan red u odgovarajucoj grupi + Case u frmStammdaten.
Public Function MaticniSekcijeGrupisano() As Variant
    MaticniSekcijeGrupisano = Array( _
        Array(ChrW(352) & "ifarnici", Array( _
            Array("Kooperanti", "Kooperanti"), _
            Array("Stanice", "Stanice"), _
            Array("Kupci", "Kupci"), _
            Array("Vozaci", "Vozaci"), _
            Array("Parcele", "Parcele"))), _
        Array("Proizvodi i cene", Array( _
            Array("Artikli", "Artikli"), _
            Array("Kulture", "Kulture"), _
            Array("Cenovnik", "Cenovnik"), _
            Array("Vrsta got. proizvoda", "VrstaGP"))), _
        Array("Ambala" & ChrW(382) & "a i pakovanje", Array( _
            Array("Ambala" & ChrW(382) & "a", "TipAmbalaze"), _
            Array("Palete", "TipPalete"), _
            Array("Kutije", "Kutije"), _
            Array("Kese", "Kese"))), _
        Array("Sistem", Array( _
            Array(Poruka("MATICNI_MSG_PODESAVANJA"), "Pode" & ChrW(353) & "avanja"), _
            Array("Admin", "Admin"), _
            Array("Korisnici", "Korisnici"))))
End Function


' Otpusti module-level reference (clsLookupMenuBtn WithEvents + dugmad) pre
' self-update importa (zivi event-sink obara CodeModule edit). Idempotentno.
Public Sub MaticniMenu_Release()
    On Error Resume Next
    Set mWrappers = Nothing
    Set mBtns = Nothing
End Sub
