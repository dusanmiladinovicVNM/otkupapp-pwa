Attribute VB_Name = "modSEFMapper"
Option Explicit

Public Function BuildSEFInvoiceDto(ByVal fakturaID As String) As clsSEFInvoiceSnapshot
    
    Dim dto As clsSEFInvoiceSnapshot
    Set dto = New clsSEFInvoiceSnapshot
    
    Dim fakture As Variant
    Dim stavke As Variant
    
    Dim colFakturaID As Long
    Dim colKupacID As Long
    Dim colBroj As Long
    Dim colDatum As Long
    Dim colIznos As Long
    
    Dim kupacID As String
    Dim i As Long
    Dim found As Boolean
    
    ' ========================
    ' tblFakture
    ' ========================
    
    fakture = GetTableData(TBL_FAKTURE)
    
    If IsEmpty(fakture) Then
        Err.Raise ERR_SEF_VALIDATION, "BuildSEFInvoiceDto", "TBL_FAKTURE is empty."
    End If
    
    colFakturaID = GetColumnIndex(TBL_FAKTURE, "FakturaID")
    colKupacID = GetColumnIndex(TBL_FAKTURE, "KupacID")
    colBroj = GetColumnIndex(TBL_FAKTURE, "BrojFakture")
    colDatum = GetColumnIndex(TBL_FAKTURE, "Datum")
    colIznos = GetColumnIndex(TBL_FAKTURE, "Iznos")
    
    If colFakturaID = 0 Or colKupacID = 0 Or colBroj = 0 Or colDatum = 0 Or colIznos = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "BuildSEFInvoiceDto", _
            "Required columns missing in tblFakture."
    End If
    
    For i = 1 To UBound(fakture, 1)
        If CStr(fakture(i, colFakturaID)) = fakturaID Then
            found = True
            
            dto.fakturaID = fakturaID
            dto.InvoiceNumber = CStr(fakture(i, colBroj))
            dto.InvoiceDate = CDate(fakture(i, colDatum))
            dto.TotalNet = CDbl(fakture(i, colIznos))
            dto.TotalVat = Round(dto.TotalNet * 0.1, 2)
            dto.TotalGross = Round(dto.TotalNet + dto.TotalVat, 2)
            
            kupacID = CStr(fakture(i, colKupacID))
            Exit For
        End If
    Next i
    
    If Not found Then
        Err.Raise ERR_SEF_VALIDATION, "BuildSEFInvoiceDto", _
            "Faktura not found: " & fakturaID
    End If
    
    ' ========================
    ' tblKupci
    ' ========================
    
    dto.BuyerID = kupacID
    dto.BuyerName = CStr(LookupValue("tblKupci", "KupacID", kupacID, "Naziv"))
    dto.BuyerPIB = CStr(LookupValue("tblKupci", "KupacID", kupacID, "PIB"))
    
    If Len(Trim$(dto.BuyerName)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "BuildSEFInvoiceDto", "Buyer name missing."
    End If
    
    If Len(Trim$(dto.BuyerPIB)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "BuildSEFInvoiceDto", "Buyer PIB missing."
    End If
    
    ' ========================
    ' Seller from tblConfig
    ' ========================
    
    dto.SellerName = GetConfigValue("SELLER_NAME")
    dto.SellerPIB = GetConfigValue("SELLER_PIB")
    
    If Len(Trim$(dto.SellerName)) = 0 Then
        Err.Raise ERR_SEF_CONFIG, "BuildSEFInvoiceDto", "SELLER_NAME missing in tblConfig."
    End If
    
    If Len(Trim$(dto.SellerPIB)) = 0 Then
        Err.Raise ERR_SEF_CONFIG, "BuildSEFInvoiceDto", "SELLER_PIB missing in tblConfig."
    End If
    
    dto.CurrencyCode = "RSD"
    dto.versionNo = GetNextSEFVersionNo(fakturaID)
    
    ' ========================
    ' Lines from tblFakturaStavke + tblPrijemnica
    ' ========================
    
    Set dto.Lines = New Collection
    
    stavke = GetTableData("tblFakturaStavke")
    
    If IsEmpty(stavke) Then
        Err.Raise ERR_SEF_VALIDATION, "BuildSEFInvoiceDto", "tblFakturaStavke is empty."
    End If
    
    Dim colStFakturaID As Long
    Dim colPrijemnicaID As Long
    Dim colKolicina As Long
    Dim colCena As Long
    Dim colKlasa As Long
    Dim colBrojPrijemnice As Long
    
    colStFakturaID = GetColumnIndex("tblFakturaStavke", "FakturaID")
    colPrijemnicaID = GetColumnIndex("tblFakturaStavke", "PrijemnicaID")
    colKolicina = GetColumnIndex("tblFakturaStavke", "Kolicina")
    colCena = GetColumnIndex("tblFakturaStavke", "Cena")
    colKlasa = GetColumnIndex("tblFakturaStavke", "Klasa")
    colBrojPrijemnice = GetColumnIndex("tblFakturaStavke", "BrojPrijemnice")
    
    If colStFakturaID = 0 Or colPrijemnicaID = 0 Or colKolicina = 0 Or colCena = 0 Or colKlasa = 0 Or colBrojPrijemnice = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "BuildSEFInvoiceDto", _
            "Required columns missing in tblFakturaStavke."
    End If
    
    Dim line As clsSEFLine
    Dim prijemnicaID As String
    Dim brojPrijemnice As String
    Dim vrstaVoca As String
    Dim sortaVoca As String
    Dim opis As String

    Dim lineNet As Double
    Dim lineVat As Double
    Dim lineGross As Double

    Dim calcTotalNet As Double
    Dim calcTotalVat As Double
    Dim calcTotalGross As Double

    calcTotalNet = 0
    calcTotalVat = 0
    calcTotalGross = 0

    For i = 1 To UBound(stavke, 1)
    
        If CStr(stavke(i, colStFakturaID)) = fakturaID Then
        
            prijemnicaID = CStr(stavke(i, colPrijemnicaID))
            brojPrijemnice = CStr(stavke(i, colBrojPrijemnice))
        
            vrstaVoca = CStr(LookupValue("tblPrijemnica", "PrijemnicaID", prijemnicaID, "VrstaVoca"))
            sortaVoca = CStr(LookupValue("tblPrijemnica", "PrijemnicaID", prijemnicaID, "SortaVoca"))
        
            opis = Trim$(vrstaVoca & " " & sortaVoca)
            opis = Trim$(opis & " po prijemnici " & brojPrijemnice)
        
            lineNet = CDbl(stavke(i, colKolicina)) * CDbl(stavke(i, colCena))
            lineVat = Round(lineNet * 0.1, 2)
            lineGross = Round(lineNet + lineVat, 2)
        
            Set line = New clsSEFLine
        
            line.prijemnicaID = prijemnicaID
            line.brojPrijemnice = brojPrijemnice
            line.Naziv = opis
            line.kolicina = CDbl(stavke(i, colKolicina))
            line.cena = CDbl(stavke(i, colCena))
            line.klasa = CStr(stavke(i, colKlasa))
            line.Neto = lineNet
            line.PDV = lineVat
            line.iznos = lineGross
            
            dto.Lines.Add line
        
            calcTotalNet = calcTotalNet + lineNet
            calcTotalVat = calcTotalVat + lineVat
            calcTotalGross = calcTotalGross + lineGross
        
        End If
    
    Next i

If dto.Lines.count = 0 Then
    Err.Raise ERR_SEF_VALIDATION, "BuildSEFInvoiceDto", _
        "Invoice has no lines."
End If
    If dto.Lines.count = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "BuildSEFInvoiceDto", _
        "Invoice has no lines."
    End If

    ' Optionaler Abgleich Kopf vs. Zeilen
    If Round(calcTotalNet, 2) <> Round(dto.TotalNet, 2) Then
    ' erstmal kein harter Fehler, nur vorbereitet
    ' sp‰ter kann das zu Err.Raise werden
    End If

    If dto.Lines.count = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "BuildSEFInvoiceDto", _
            "Invoice has no lines."
    End If
    
    ' optional sanity check
    If Round(dto.TotalNet + dto.TotalVat, 2) <> Round(dto.TotalGross, 2) Then
        ' Noch kein harter Fehler, aber fachlich auff‰llig
        ' Kann sp‰ter zu Err.Raise werden, wenn du streng sein willst
    End If
    
    Set BuildSEFInvoiceDto = dto

End Function

Public Function SerializeSEFRequest(ByVal dto As clsSEFInvoiceSnapshot) As String
    
    Dim sb As String
    Dim i As Long
    Dim ln As clsSEFLine
    
    If dto Is Nothing Then
        Err.Raise ERR_SEF_VALIDATION, "SerializeSEFRequest", "DTO is Nothing."
    End If
    
    sb = ""
    sb = sb & "{"
    sb = sb & """FakturaID"":" & JsonString(dto.fakturaID) & ","
    sb = sb & """VersionNo"":" & CStr(dto.versionNo) & ","
    sb = sb & """InvoiceNumber"":" & JsonString(dto.InvoiceNumber) & ","
    sb = sb & """InvoiceDate"":" & JsonString(Format$(dto.InvoiceDate, "yyyy-mm-dd")) & ","
    sb = sb & """CurrencyCode"":" & JsonString(dto.CurrencyCode) & ","
    
    sb = sb & """Buyer"":{"
    sb = sb & """BuyerID"":" & JsonString(dto.BuyerID) & ","
    sb = sb & """Name"":" & JsonString(dto.BuyerName) & ","
    sb = sb & """PIB"":" & JsonString(dto.BuyerPIB)
    sb = sb & "},"
    
    sb = sb & """Seller"":{"
    sb = sb & """Name"":" & JsonString(dto.SellerName) & ","
    sb = sb & """PIB"":" & JsonString(dto.SellerPIB)
    sb = sb & "},"
    
    sb = sb & """Totals"":{"
    sb = sb & """Net"":" & JsonNumber(dto.TotalNet) & ","
    sb = sb & """VAT"":" & JsonNumber(dto.TotalVat) & ","
    sb = sb & """Gross"":" & JsonNumber(dto.TotalGross)
    sb = sb & "},"
    
    sb = sb & """Lines"":["
    
    For i = 1 To dto.Lines.count
        Set ln = dto.Lines(i)
        
        If i > 1 Then sb = sb & ","
        
        sb = sb & "{"
        sb = sb & """PrijemnicaID"":" & JsonString(ln.prijemnicaID) & ","
        sb = sb & """BrojPrijemnice"":" & JsonString(ln.brojPrijemnice) & ","
        sb = sb & """Naziv"":" & JsonString(ln.Naziv) & ","
        sb = sb & """Kolicina"":" & JsonNumber(ln.kolicina) & ","
        sb = sb & """Cena"":" & JsonNumber(ln.cena) & ","
        sb = sb & """Klasa"":" & JsonString(ln.klasa) & ","
        sb = sb & """Neto"":" & JsonNumber(ln.Neto) & ","
        sb = sb & """PDV"":" & JsonNumber(ln.PDV) & ","
        sb = sb & """Iznos"":" & JsonNumber(ln.iznos)
        sb = sb & "}"
    Next i
    
    sb = sb & "]"
    sb = sb & "}"
    
    SerializeSEFRequest = sb

End Function

Private Function JsonString(ByVal s As String) As String
    Dim t As String
    
    t = s
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, "\n")
    t = Replace(t, vbCr, "\n")
    t = Replace(t, vbLf, "\n")
    
    JsonString = """" & t & """"
End Function

Private Function JsonNumber(ByVal n As Double) As String
    Dim s As String
    
    s = Format$(Round(n, 2), "0.00")
    s = Replace(s, ",", ".")
    
    JsonNumber = s
End Function


Public Function SerializeUBLInvoice(ByVal dto As clsSEFInvoiceSnapshot) As String
    
    Dim xml As String
    Dim i As Long
    Dim ln As clsSEFLine
    
    Dim sellerMaticni As String
    Dim sellerStreet As String
    Dim sellerCity As String
    Dim sellerPostalCode As String
    Dim sellerCountryCode As String
    Dim sellerAccount As String
    Dim sellerEmail As String
    
    Dim paymentMeansCode As String
    Dim paymentDueDays As Long
    Dim noteText As String
    
    Dim buyerMaticni As String
    Dim buyerStreet As String
    Dim buyerCity As String
    Dim buyerPostalCode As String
    Dim buyerCountryCode As String
    Dim buyerEmail As String
    
    Dim dueDate As Date
    Dim deliveryDate As Date
    
    If dto Is Nothing Then
        Err.Raise ERR_SEF_VALIDATION, "SerializeUBLInvoiceV2", "DTO is Nothing."
    End If
    
    sellerMaticni = GetConfigValue("SELLER_MATICNI_BROJ")
    sellerStreet = GetConfigValue("SELLER_STREET")
    sellerCity = GetConfigValue("SELLER_CITY")
    sellerPostalCode = GetConfigValue("SELLER_POSTAL_CODE")
    sellerCountryCode = GetConfigValue("SELLER_COUNTRY_CODE")
    sellerAccount = GetConfigValue("SELLER_ACCOUNT")
    sellerEmail = GetConfigValue("SELLER_EMAIL")
    
    paymentMeansCode = GetConfigValue("SEF_PAYMENT_MEANS_CODE")
    noteText = GetConfigValue("SEF_NOTE_DEFAULT")
    
    If Len(Trim$(sellerCountryCode)) = 0 Then sellerCountryCode = "RS"
    If Len(Trim$(paymentMeansCode)) = 0 Then paymentMeansCode = "30"
    If Len(Trim$(noteText)) = 0 Then noteText = "Otkupljena roba prema prijemnici"
    
    If Len(Trim$(GetConfigValue("SEF_PAYMENT_DUE_DAYS"))) = 0 Then
        paymentDueDays = 15
    Else
        paymentDueDays = CLng(GetConfigValue("SEF_PAYMENT_DUE_DAYS"))
    End If
    
    dueDate = DateAdd("d", paymentDueDays, dto.InvoiceDate)
    
    deliveryDate = GetInvoiceDeliveryDate(dto.fakturaID, dto.InvoiceDate)
    
    buyerMaticni = NzStr(LookupValue("tblKupci", "KupacID", dto.BuyerID, "MaticniBroj"))
    buyerStreet = NzStr(LookupValue("tblKupci", "KupacID", dto.BuyerID, "Ulica"))
    buyerCity = NzStr(LookupValue("tblKupci", "KupacID", dto.BuyerID, "Mesto"))
    buyerPostalCode = NzStr(LookupValue("tblKupci", "KupacID", dto.BuyerID, "PostanskiBroj"))
    buyerCountryCode = NzStr(LookupValue("tblKupci", "KupacID", dto.BuyerID, "Drzava"))
    buyerEmail = NzStr(LookupValue("tblKupci", "KupacID", dto.BuyerID, "Email"))
    
    If Len(Trim$(buyerCountryCode)) = 0 Then buyerCountryCode = "RS"
    
    Dim taxPercent As Double
    Dim taxCategoryID As String
    Dim descriptionCode As String

    taxPercent = GetDefaultTaxPercent()
    taxCategoryID = GetDefaultTaxCategoryID()
    descriptionCode = GetDefaultInvoicePeriodDescriptionCode()
    
    xml = ""
    xml = xml & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    xml = xml & "<Invoice " & _
                "xmlns:cec=""urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2"" " & _
                "xmlns:cac=""urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"" " & _
                "xmlns:cbc=""urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"" " & _
                "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
                "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
                "xmlns:sbt=""http://mfin.gov.rs/srbdt/srbdtext"" " & _
                "xmlns=""urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"">" & vbCrLf
    
    ' ========================
    ' Header
    ' ========================
    xml = xml & "  <cbc:CustomizationID>urn:cen.eu:en16931:2017#compliant#urn:mfin.gov.rs:srbdt:2022</cbc:CustomizationID>" & vbCrLf
    xml = xml & "  <cbc:ID>" & XmlEscape(dto.InvoiceNumber) & "</cbc:ID>" & vbCrLf

    Dim forceTodayIssueDate As String
    forceTodayIssueDate = UCase$(GetConfigValue("SEF_FORCE_TODAY_ISSUE_DATE"))

    If forceTodayIssueDate = "DA" Then
        xml = xml & "  <cbc:IssueDate>" & Format$(Date, "yyyy-mm-dd") & "</cbc:IssueDate>" & vbCrLf
    Else
        xml = xml & "  <cbc:IssueDate>" & Format$(dto.InvoiceDate, "yyyy-mm-dd") & "</cbc:IssueDate>" & vbCrLf
    End If

    xml = xml & "  <cbc:DueDate>" & Format$(dueDate, "yyyy-mm-dd") & "</cbc:DueDate>" & vbCrLf
    xml = xml & "  <cbc:InvoiceTypeCode>380</cbc:InvoiceTypeCode>" & vbCrLf
    xml = xml & "  <cbc:Note>" & XmlEscape(noteText) & "</cbc:Note>" & vbCrLf
    xml = xml & "  <cbc:DocumentCurrencyCode>" & XmlEscape(dto.CurrencyCode) & "</cbc:DocumentCurrencyCode>" & vbCrLf
    If Len(Trim$(descriptionCode)) > 0 Then
        xml = xml & "  <cac:InvoicePeriod>" & vbCrLf
        xml = xml & "    <cbc:DescriptionCode>" & XmlEscape(descriptionCode) & "</cbc:DescriptionCode>" & vbCrLf
        xml = xml & "  </cac:InvoicePeriod>" & vbCrLf
    End If
    
    ' ========================
    ' Supplier
    ' ========================
    xml = xml & "  <cac:AccountingSupplierParty>" & vbCrLf
    xml = xml & "    <cac:Party>" & vbCrLf
    xml = xml & "      <cbc:EndpointID schemeID=""9948"">" & XmlEscape(dto.SellerPIB) & "</cbc:EndpointID>" & vbCrLf
    
    xml = xml & "      <cac:PartyName>" & vbCrLf
    xml = xml & "        <cbc:Name>" & XmlEscape(dto.SellerName) & "</cbc:Name>" & vbCrLf
    xml = xml & "      </cac:PartyName>" & vbCrLf
    
    xml = xml & "      <cac:PostalAddress>" & vbCrLf
    xml = xml & "        <cbc:StreetName>" & XmlEscape(sellerStreet) & "</cbc:StreetName>" & vbCrLf
    xml = xml & "        <cbc:CityName>" & XmlEscape(sellerCity) & "</cbc:CityName>" & vbCrLf
    xml = xml & "        <cbc:PostalZone>" & XmlEscape(sellerPostalCode) & "</cbc:PostalZone>" & vbCrLf
    xml = xml & "        <cac:Country><cbc:IdentificationCode>" & XmlEscape(sellerCountryCode) & "</cbc:IdentificationCode></cac:Country>" & vbCrLf
    xml = xml & "      </cac:PostalAddress>" & vbCrLf
    
    xml = xml & "      <cac:PartyTaxScheme>" & vbCrLf
    xml = xml & "        <cbc:CompanyID>RS" & XmlEscape(dto.SellerPIB) & "</cbc:CompanyID>" & vbCrLf
    xml = xml & "        <cac:TaxScheme><cbc:ID>VAT</cbc:ID></cac:TaxScheme>" & vbCrLf
    xml = xml & "      </cac:PartyTaxScheme>" & vbCrLf
    
    xml = xml & "      <cac:PartyLegalEntity>" & vbCrLf
    xml = xml & "        <cbc:RegistrationName>" & XmlEscape(dto.SellerName) & "</cbc:RegistrationName>" & vbCrLf
    xml = xml & "        <cbc:CompanyID>" & XmlEscape(sellerMaticni) & "</cbc:CompanyID>" & vbCrLf
    xml = xml & "      </cac:PartyLegalEntity>" & vbCrLf
    
    If Len(Trim$(sellerEmail)) > 0 Then
        xml = xml & "      <cac:Contact>" & vbCrLf
        xml = xml & "        <cbc:ElectronicMail>" & XmlEscape(sellerEmail) & "</cbc:ElectronicMail>" & vbCrLf
        xml = xml & "      </cac:Contact>" & vbCrLf
    End If
    
    xml = xml & "    </cac:Party>" & vbCrLf
    xml = xml & "  </cac:AccountingSupplierParty>" & vbCrLf
    
    ' ========================
    ' Customer
    ' ========================
    xml = xml & "  <cac:AccountingCustomerParty>" & vbCrLf
    xml = xml & "    <cac:Party>" & vbCrLf
    xml = xml & "      <cbc:EndpointID schemeID=""9948"">" & XmlEscape(dto.BuyerPIB) & "</cbc:EndpointID>" & vbCrLf
    
    xml = xml & "      <cac:PartyName>" & vbCrLf
    xml = xml & "        <cbc:Name>" & XmlEscape(dto.BuyerName) & "</cbc:Name>" & vbCrLf
    xml = xml & "      </cac:PartyName>" & vbCrLf
    
    xml = xml & "      <cac:PostalAddress>" & vbCrLf
    xml = xml & "        <cbc:StreetName>" & XmlEscape(buyerStreet) & "</cbc:StreetName>" & vbCrLf
    xml = xml & "        <cbc:CityName>" & XmlEscape(buyerCity) & "</cbc:CityName>" & vbCrLf
    
    If Len(Trim$(buyerPostalCode)) > 0 Then
        xml = xml & "        <cbc:PostalZone>" & XmlEscape(buyerPostalCode) & "</cbc:PostalZone>" & vbCrLf
    End If
    
    xml = xml & "        <cac:Country><cbc:IdentificationCode>" & XmlEscape(buyerCountryCode) & "</cbc:IdentificationCode></cac:Country>" & vbCrLf
    xml = xml & "      </cac:PostalAddress>" & vbCrLf
    
    xml = xml & "      <cac:PartyTaxScheme>" & vbCrLf
    xml = xml & "        <cbc:CompanyID>RS" & XmlEscape(dto.BuyerPIB) & "</cbc:CompanyID>" & vbCrLf
    xml = xml & "        <cac:TaxScheme><cbc:ID>VAT</cbc:ID></cac:TaxScheme>" & vbCrLf
    xml = xml & "      </cac:PartyTaxScheme>" & vbCrLf
    
    xml = xml & "      <cac:PartyLegalEntity>" & vbCrLf
    xml = xml & "        <cbc:RegistrationName>" & XmlEscape(dto.BuyerName) & "</cbc:RegistrationName>" & vbCrLf
    xml = xml & "        <cbc:CompanyID>" & XmlEscape(buyerMaticni) & "</cbc:CompanyID>" & vbCrLf
    xml = xml & "      </cac:PartyLegalEntity>" & vbCrLf
    
    If Len(Trim$(buyerEmail)) > 0 Then
        xml = xml & "      <cac:Contact>" & vbCrLf
        xml = xml & "        <cbc:ElectronicMail>" & XmlEscape(buyerEmail) & "</cbc:ElectronicMail>" & vbCrLf
        xml = xml & "      </cac:Contact>" & vbCrLf
    End If
    
    xml = xml & "    </cac:Party>" & vbCrLf
    xml = xml & "  </cac:AccountingCustomerParty>" & vbCrLf
    
    ' ========================
    ' Delivery
    ' ========================
    xml = xml & "  <cac:Delivery>" & vbCrLf
    xml = xml & "    <cbc:ActualDeliveryDate>" & Format$(deliveryDate, "yyyy-mm-dd") & "</cbc:ActualDeliveryDate>" & vbCrLf
    xml = xml & "  </cac:Delivery>" & vbCrLf
    
    ' ========================
    ' PaymentMeans
    ' ========================
    xml = xml & "  <cac:PaymentMeans>" & vbCrLf
    xml = xml & "    <cbc:PaymentMeansCode>" & XmlEscape(paymentMeansCode) & "</cbc:PaymentMeansCode>" & vbCrLf
    xml = xml & "    <cbc:PaymentID>" & XmlEscape(dto.InvoiceNumber) & "</cbc:PaymentID>" & vbCrLf
    xml = xml & "    <cac:PayeeFinancialAccount>" & vbCrLf
    xml = xml & "      <cbc:ID>" & XmlEscape(sellerAccount) & "</cbc:ID>" & vbCrLf
    xml = xml & "    </cac:PayeeFinancialAccount>" & vbCrLf
    xml = xml & "  </cac:PaymentMeans>" & vbCrLf
    
    ' ========================
    ' Tax total
    ' ========================
    xml = xml & "  <cac:TaxTotal>" & vbCrLf
    xml = xml & "    <cbc:TaxAmount currencyID=""" & XmlEscape(dto.CurrencyCode) & """>" & XmlAmount(dto.TotalVat) & "</cbc:TaxAmount>" & vbCrLf
    xml = xml & "    <cac:TaxSubtotal>" & vbCrLf
    xml = xml & "      <cbc:TaxableAmount currencyID=""" & XmlEscape(dto.CurrencyCode) & """>" & XmlAmount(dto.TotalNet) & "</cbc:TaxableAmount>" & vbCrLf
    xml = xml & "      <cbc:TaxAmount currencyID=""" & XmlEscape(dto.CurrencyCode) & """>" & XmlAmount(dto.TotalVat) & "</cbc:TaxAmount>" & vbCrLf
    xml = xml & "      <cac:TaxCategory>" & vbCrLf
    xml = xml & "        <cbc:ID>" & XmlEscape(taxCategoryID) & "</cbc:ID>" & vbCrLf
    xml = xml & "        <cbc:Percent>" & Replace(CStr(taxPercent), ",", ".") & "</cbc:Percent>" & vbCrLf
    xml = xml & "        <cac:TaxScheme><cbc:ID>VAT</cbc:ID></cac:TaxScheme>" & vbCrLf
    xml = xml & "      </cac:TaxCategory>" & vbCrLf
    xml = xml & "    </cac:TaxSubtotal>" & vbCrLf
    xml = xml & "  </cac:TaxTotal>" & vbCrLf
    
    ' ========================
    ' Monetary total
    ' ========================
    xml = xml & "  <cac:LegalMonetaryTotal>" & vbCrLf
    xml = xml & "    <cbc:LineExtensionAmount currencyID=""" & XmlEscape(dto.CurrencyCode) & """>" & XmlAmount(dto.TotalNet) & "</cbc:LineExtensionAmount>" & vbCrLf
    xml = xml & "    <cbc:TaxExclusiveAmount currencyID=""" & XmlEscape(dto.CurrencyCode) & """>" & XmlAmount(dto.TotalNet) & "</cbc:TaxExclusiveAmount>" & vbCrLf
    xml = xml & "    <cbc:TaxInclusiveAmount currencyID=""" & XmlEscape(dto.CurrencyCode) & """>" & XmlAmount(dto.TotalGross) & "</cbc:TaxInclusiveAmount>" & vbCrLf
    xml = xml & "    <cbc:AllowanceTotalAmount currencyID=""" & XmlEscape(dto.CurrencyCode) & """>0.00</cbc:AllowanceTotalAmount>" & vbCrLf
    xml = xml & "    <cbc:PrepaidAmount currencyID=""" & XmlEscape(dto.CurrencyCode) & """>0.00</cbc:PrepaidAmount>" & vbCrLf
    xml = xml & "    <cbc:PayableRoundingAmount currencyID=""" & XmlEscape(dto.CurrencyCode) & """>0.00</cbc:PayableRoundingAmount>" & vbCrLf
    xml = xml & "    <cbc:PayableAmount currencyID=""" & XmlEscape(dto.CurrencyCode) & """>" & XmlAmount(dto.TotalGross) & "</cbc:PayableAmount>" & vbCrLf
    xml = xml & "  </cac:LegalMonetaryTotal>" & vbCrLf
    
    ' ========================
    ' Invoice lines
    ' ========================
    For i = 1 To dto.Lines.count
        
        Set ln = dto.Lines(i)
        
        xml = xml & "  <cac:InvoiceLine>" & vbCrLf
        xml = xml & "    <cbc:ID>" & CStr(i) & "</cbc:ID>" & vbCrLf
        xml = xml & "    <cbc:InvoicedQuantity unitCode=""KGM"">" & XmlAmount(ln.kolicina) & "</cbc:InvoicedQuantity>" & vbCrLf
        xml = xml & "    <cbc:LineExtensionAmount currencyID=""" & XmlEscape(dto.CurrencyCode) & """>" & XmlAmount(ln.Neto) & "</cbc:LineExtensionAmount>" & vbCrLf
        
        xml = xml & "    <cac:Item>" & vbCrLf
        xml = xml & "      <cbc:Name>" & XmlEscape(ln.Naziv) & "</cbc:Name>" & vbCrLf
        xml = xml & "      <cac:SellersItemIdentification>" & vbCrLf
        xml = xml & "        <cbc:ID>" & XmlEscape(ln.prijemnicaID) & "</cbc:ID>" & vbCrLf
        xml = xml & "      </cac:SellersItemIdentification>" & vbCrLf
        xml = xml & "      <cac:ClassifiedTaxCategory>" & vbCrLf
        xml = xml & "        <cbc:ID>" & XmlEscape(taxCategoryID) & "</cbc:ID>" & vbCrLf
        xml = xml & "        <cbc:Percent>" & Replace(CStr(taxPercent), ",", ".") & "</cbc:Percent>" & vbCrLf
        xml = xml & "        <cac:TaxScheme><cbc:ID>VAT</cbc:ID></cac:TaxScheme>" & vbCrLf
        xml = xml & "      </cac:ClassifiedTaxCategory>" & vbCrLf
        xml = xml & "    </cac:Item>" & vbCrLf
        
        xml = xml & "    <cac:Price>" & vbCrLf
        xml = xml & "      <cbc:PriceAmount currencyID=""" & XmlEscape(dto.CurrencyCode) & """>" & XmlAmount(ln.cena) & "</cbc:PriceAmount>" & vbCrLf
        xml = xml & "    </cac:Price>" & vbCrLf
        
        xml = xml & "  </cac:InvoiceLine>" & vbCrLf
        
    Next i
    
    xml = xml & "</Invoice>"
    
    SerializeUBLInvoice = xml

End Function

Private Function XmlEscape(ByVal s As String) As String
    
    Dim t As String
    
    t = s
    t = Replace(t, "&", "&amp;")
    t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;")
    t = Replace(t, """", "&quot;")
    t = Replace(t, "'", "&apos;")
    
    XmlEscape = t

End Function

Private Function XmlAmount(ByVal n As Double) As String
    
    Dim s As String
    
    s = Format$(Round(n, 2), "0.00")
    s = Replace(s, ",", ".")
    
    XmlAmount = s

End Function

Private Function NzStr(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then
        NzStr = ""
    Else
        NzStr = Trim$(CStr(v))
    End If
End Function

Private Function GetInvoiceDeliveryDate(ByVal fakturaID As String, ByVal fallbackDate As Date) As Date
    
    Dim stavke As Variant
    Dim colFakturaID As Long
    Dim colPrijemnicaID As Long
    Dim i As Long
    Dim prijemnicaID As String
    Dim v As Variant
    
    stavke = GetTableData("tblFakturaStavke")
    
    If IsEmpty(stavke) Then
        GetInvoiceDeliveryDate = fallbackDate
        Exit Function
    End If
    
    colFakturaID = GetColumnIndex("tblFakturaStavke", "FakturaID")
    colPrijemnicaID = GetColumnIndex("tblFakturaStavke", "PrijemnicaID")
    
    For i = 1 To UBound(stavke, 1)
        If CStr(stavke(i, colFakturaID)) = fakturaID Then
            prijemnicaID = CStr(stavke(i, colPrijemnicaID))
            v = LookupValue("tblPrijemnica", "PrijemnicaID", prijemnicaID, "Datum")
            
            If IsDate(v) Then
                GetInvoiceDeliveryDate = CDate(v)
            Else
                GetInvoiceDeliveryDate = fallbackDate
            End If
            Exit Function
        End If
    Next i
    
    GetInvoiceDeliveryDate = fallbackDate

End Function

Public Function ComputePayloadHash(ByVal payload As String) As String

    Dim i As Long
    Dim h As Double
    Dim c As Long
    
    h = 5381
    
    For i = 1 To Len(payload)
        
        c = Asc(Mid$(payload, i, 1))
        
        ' Rolling hash ohne Bitoperatoren
        h = (h * 33#) + c
        
        ' manuelles Modulo, damit h nie zu groﬂ wird
        h = h - Int(h / 2147483647#) * 2147483647#
        
    Next i
    
    ComputePayloadHash = Hex$(CLng(h))

End Function



Public Sub Test_BuildSEFInvoiceDto()

    On Error GoTo EH
    
    Dim dto As clsSEFInvoiceSnapshot
    Dim ln As clsSEFLine
    
    Set dto = BuildSEFInvoiceDto("FAK-00001")
    
    Debug.Print "DTO OK"
    Debug.Print "InvoiceNumber: "; dto.InvoiceNumber
    Debug.Print "BuyerName: "; dto.BuyerName
    Debug.Print "SellerName: "; dto.SellerName
    Debug.Print "TotalNet: "; dto.TotalNet
    Debug.Print "TotalVat: "; dto.TotalVat
    Debug.Print "TotalGross: "; dto.TotalGross
    Debug.Print "Lines.Count: "; dto.Lines.count
    
    If dto.Lines.count = 0 Then
        Debug.Print "NO LINES"
        Exit Sub
    End If
    
    Debug.Print "Type of dto.Lines(1): "; TypeName(dto.Lines.Item(1))
    
    Set ln = dto.Lines.Item(1)
    
    Debug.Print "LINE OK"
    Debug.Print "Naziv: "; ln.Naziv
    Debug.Print "Neto: "; ln.Neto
    Debug.Print "PDV: "; ln.PDV
    Debug.Print "Iznos: "; ln.iznos
    
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.Description
End Sub

Public Sub Test_SerializeSEFRequest()

    On Error GoTo EH
    
    Dim dto As clsSEFInvoiceSnapshot
    Dim payload As String
    
    Set dto = BuildSEFInvoiceDto("FAK-00001")
    payload = SerializeSEFRequest(dto)
    
    Debug.Print payload
    
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.Description

End Sub

Public Sub Test_PayloadHash()

    Dim dto As clsSEFInvoiceSnapshot
    Dim payload As String
    Dim hash As String
    
    Set dto = BuildSEFInvoiceDto("FAK-00001")
    
    payload = SerializeSEFRequest(dto)
    
    hash = ComputePayloadHash(payload)
    
    Debug.Print hash

End Sub

Public Sub Test_SerializeUBLInvoice()

    On Error GoTo EH
    
    Dim dto As clsSEFInvoiceSnapshot
    Dim xml As String
    
    Set dto = BuildSEFInvoiceDto("FAK-00001")
    xml = SerializeUBLInvoice(dto)
    
    Debug.Print xml
    
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.Description

End Sub
