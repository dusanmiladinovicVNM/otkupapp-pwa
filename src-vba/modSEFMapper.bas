Attribute VB_Name = "modSEFMapper"
Option Explicit

Public Function BuildSEFInvoiceDto(ByVal fakturaID As String) As clsSEFInvoiceSnapshot
    On Error GoTo EH

    Const SRC As String = "modSEFMapper.BuildSEFInvoiceDto"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "FakturaID is required."
    End If

    Dim dto As clsSEFInvoiceSnapshot
    Set dto = New clsSEFInvoiceSnapshot

    Dim fakture As Variant
    Dim stavke As Variant

    Dim colFakturaID As Long
    Dim colKupacID As Long
    Dim colBroj As Long
    Dim colDatum As Long
    Dim colIznos As Long

    fakture = GetTableData(TBL_FAKTURE)

    If IsEmpty(fakture) Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "TBL_FAKTURE is empty."
    End If

    colFakturaID = RequireColumnIndex(TBL_FAKTURE, "FakturaID", SRC)
    colKupacID = RequireColumnIndex(TBL_FAKTURE, "KupacID", SRC)
    colBroj = RequireColumnIndex(TBL_FAKTURE, "BrojFakture", SRC)
    colDatum = RequireColumnIndex(TBL_FAKTURE, "Datum", SRC)
    colIznos = RequireColumnIndex(TBL_FAKTURE, "Iznos", SRC)

    Dim kupacID As String
    Dim i As Long
    Dim found As Boolean

    Dim taxPercent As Double
    taxPercent = GetDefaultTaxPercent()

    If taxPercent < 0 Then
        Err.Raise ERR_SEF_CONFIG, SRC, "SEF tax percent cannot be negative."
    End If

    For i = 1 To UBound(fakture, 1)
        If CStr(fakture(i, colFakturaID)) = fakturaID Then
            found = True

            If Not IsDate(fakture(i, colDatum)) Then
                Err.Raise ERR_SEF_VALIDATION, SRC, _
                          "Invoice date is not valid for faktura " & fakturaID
            End If

            If Not IsNumeric(fakture(i, colIznos)) Then
                Err.Raise ERR_SEF_VALIDATION, SRC, _
                          "Invoice amount is not numeric for faktura " & fakturaID
            End If

            dto.fakturaID = fakturaID
            dto.InvoiceNumber = Trim$(CStr(fakture(i, colBroj)))
            dto.InvoiceDate = CDate(fakture(i, colDatum))
            dto.DeliveryDate = GetInvoiceDeliveryDate(fakturaID)
            
            dto.TotalNet = CDbl(fakture(i, colIznos))
            dto.TotalVat = Round(dto.TotalNet * taxPercent / 100, 2)
            dto.TotalGross = Round(dto.TotalNet + dto.TotalVat, 2)
            

            kupacID = Trim$(CStr(fakture(i, colKupacID)))
            Exit For
        End If
    Next i

    If Not found Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "Faktura not found: " & fakturaID
    End If

    If Len(Trim$(dto.InvoiceNumber)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "Invoice number is missing."
    End If

    If dto.TotalNet <= 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "Invoice net amount must be > 0."
    End If

    If Len(kupacID) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "KupacID is missing."
    End If

    dto.BuyerID = kupacID
    dto.BuyerName = CStr(LookupValue(TBL_KUPCI, "KupacID", kupacID, "Naziv"))
    dto.BuyerPIB = CStr(LookupValue(TBL_KUPCI, "KupacID", kupacID, "PIB"))

    If Len(Trim$(dto.BuyerName)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "Buyer name missing."
    End If

    If Len(Trim$(dto.BuyerPIB)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "Buyer PIB missing."
    End If

    dto.SellerName = GetConfigValue("SELLER_NAME")
    dto.SellerPIB = GetConfigValue("SELLER_PIB")

    If Len(Trim$(dto.SellerName)) = 0 Then
        Err.Raise ERR_SEF_CONFIG, SRC, "SELLER_NAME missing in tblConfig."
    End If

    If Len(Trim$(dto.SellerPIB)) = 0 Then
        Err.Raise ERR_SEF_CONFIG, SRC, "SELLER_PIB missing in tblConfig."
    End If

    dto.CurrencyCode = "RSD"
    dto.versionNo = GetNextSEFVersionNo(fakturaID)

    Set dto.Lines = New Collection

    stavke = GetTableData(TBL_FAKTURA_STAVKE)

    If IsEmpty(stavke) Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "tblFakturaStavke is empty."
    End If

    Dim colStFakturaID As Long
    Dim colPrijemnicaID As Long
    Dim colKolicina As Long
    Dim colCena As Long
    Dim colKlasa As Long
    Dim colBrojPrijemnice As Long

    colStFakturaID = RequireColumnIndex(TBL_FAKTURA_STAVKE, "FakturaID", SRC)
    colPrijemnicaID = RequireColumnIndex(TBL_FAKTURA_STAVKE, "PrijemnicaID", SRC)
    colKolicina = RequireColumnIndex(TBL_FAKTURA_STAVKE, "Kolicina", SRC)
    colCena = RequireColumnIndex(TBL_FAKTURA_STAVKE, "Cena", SRC)
    colKlasa = RequireColumnIndex(TBL_FAKTURA_STAVKE, "Klasa", SRC)
    colBrojPrijemnice = RequireColumnIndex(TBL_FAKTURA_STAVKE, "BrojPrijemnice", SRC)

    Dim line As clsSEFLine
    Dim prijemnicaID As String
    Dim brojPrijemnice As String
    Dim vrstaVoca As String
    Dim sortaVoca As String
    Dim opis As String

    Dim qty As Double
    Dim price As Double
    Dim lineNet As Double
    Dim lineVat As Double
    Dim lineGross As Double

    Dim calcTotalNet As Double
    Dim calcTotalVat As Double
    Dim calcTotalGross As Double

    For i = 1 To UBound(stavke, 1)

        If CStr(stavke(i, colStFakturaID)) = fakturaID Then

            prijemnicaID = Trim$(CStr(stavke(i, colPrijemnicaID)))
            brojPrijemnice = Trim$(CStr(stavke(i, colBrojPrijemnice)))

            If Len(prijemnicaID) = 0 Then
                Err.Raise ERR_SEF_VALIDATION, SRC, _
                          "PrijemnicaID missing in invoice line for " & fakturaID
            End If

            If Not TryParseDouble(CStr(stavke(i, colKolicina)), qty) Or qty <= 0 Then
                Err.Raise ERR_SEF_VALIDATION, SRC, _
                          "Invalid line quantity for faktura " & fakturaID
            End If

            If Not TryParseDouble(CStr(stavke(i, colCena)), price) Or price < 0 Then
                Err.Raise ERR_SEF_VALIDATION, SRC, _
                          "Invalid line price for faktura " & fakturaID
            End If

            vrstaVoca = CStr(LookupValue(TBL_PRIJEMNICA, "PrijemnicaID", prijemnicaID, "VrstaVoca"))
            sortaVoca = CStr(LookupValue(TBL_PRIJEMNICA, "PrijemnicaID", prijemnicaID, "SortaVoca"))

            opis = Trim$(vrstaVoca & " " & sortaVoca)
            opis = Trim$(opis & " po prijemnici " & brojPrijemnice)

            If Len(Trim$(opis)) = 0 Then
                opis = "Roba po prijemnici " & brojPrijemnice
            End If

            lineNet = Round(qty * price, 2)
            lineVat = Round(lineNet * taxPercent / 100, 2)
            lineGross = Round(lineNet + lineVat, 2)

            Set line = New clsSEFLine

            line.prijemnicaID = prijemnicaID
            line.brojPrijemnice = brojPrijemnice
            line.naziv = opis
            line.kolicina = qty
            line.cena = price
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
        Err.Raise ERR_SEF_VALIDATION, SRC, "Invoice has no lines."
    End If

    ' Za sada ostaje soft check kao i ranije.
    ' Ako kasnije želiš stroži SEF režim, ovo može postati Err.Raise.
    If Round(calcTotalNet, 2) <> Round(dto.TotalNet, 2) Then
        ' Soft mismatch: header total differs from line total.
    End If

    If Round(calcTotalVat, 2) <> Round(dto.TotalVat, 2) Then
        ' Soft mismatch: header VAT differs from line VAT.
    End If

    If Round(calcTotalGross, 2) <> Round(dto.TotalGross, 2) Then
        ' Soft mismatch: header gross differs from line gross.
    End If

    Set BuildSEFInvoiceDto = dto
    Exit Function

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.Description
End Function

' Legacy/debug JSON snapshot serializer.
' Not used for actual SEF UBL submission.
Public Function SerializeSEFRequest(ByVal dto As clsSEFInvoiceSnapshot) As String
    On Error GoTo EH
    
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
        sb = sb & """Naziv"":" & JsonString(ln.naziv) & ","
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
    Exit Function
    
EH:
    LogErr "SerializeSEFRequest"
    Err.Raise Err.Number, "SerializeSEFRequest", Err.Description

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
    On Error GoTo EH
        
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
    
    ValidateSEFDtoForUBL dto
    
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
    
    paymentDueDays = GetSEFPaymentDueDays()
    
    dueDate = DateAdd("d", paymentDueDays, dto.InvoiceDate)
    
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
    xml = xml & "    <cbc:ActualDeliveryDate>" & Format$(dto.DeliveryDate, "yyyy-mm-dd") & "</cbc:ActualDeliveryDate>" & vbCrLf
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
        xml = xml & "      <cbc:Name>" & XmlEscape(ln.naziv) & "</cbc:Name>" & vbCrLf
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
    
    ValidateGeneratedUBL xml, "modSEFMapper.SerializeUBLInvoice"
    SerializeUBLInvoice = xml
    Exit Function
    
EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    LogErr "SerializeUBLInvoice"
    On Error GoTo 0

    If errNum <> 0 Then
        Err.Raise errNum, "SerializeUBLInvoice", _
                  "Source=" & errSrc & " | " & errDesc
    Else
        Err.Raise ERR_SEF_VALIDATION, "SerializeUBLInvoice", _
                  "Unexpected error during UBL serialization."
    End If
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

Private Function GetInvoiceDeliveryDate(ByVal fakturaID As String) As Date
    On Error GoTo EH

    Const SRC As String = "modSEFMapper.GetInvoiceDeliveryDate"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "FakturaID is required."
    End If

    Dim InvoiceDate As Date
    InvoiceDate = GetFakturaDateForSEF(fakturaID)

    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)

    If IsEmpty(stavkeData) Then
        GetInvoiceDeliveryDate = InvoiceDate
        Exit Function
    End If

    Dim colStFakturaID As Long
    Dim colPrijemnicaID As Long

    colStFakturaID = RequireColumnIndex(TBL_FAKTURA_STAVKE, "FakturaID", SRC)
    colPrijemnicaID = RequireColumnIndex(TBL_FAKTURA_STAVKE, "PrijemnicaID", SRC)

    RequireColumnIndex TBL_PRIJEMNICA, "PrijemnicaID", SRC
    RequireColumnIndex TBL_PRIJEMNICA, "Datum", SRC

    Dim latestDeliveryDate As Date
    latestDeliveryDate = 0

    Dim i As Long
    For i = 1 To UBound(stavkeData, 1)

        If Trim$(CStr(stavkeData(i, colStFakturaID))) = fakturaID Then

            Dim prijemnicaID As String
            prijemnicaID = Trim$(CStr(stavkeData(i, colPrijemnicaID)))

            If Len(prijemnicaID) > 0 Then

                Dim v As Variant
                v = LookupValue(TBL_PRIJEMNICA, "PrijemnicaID", prijemnicaID, "Datum")

                If Not IsEmpty(v) And Not IsNull(v) And IsDate(v) Then
                    If CDate(v) > latestDeliveryDate Then
                        latestDeliveryDate = CDate(v)
                    End If
                Else
                    Err.Raise ERR_SEF_VALIDATION, SRC, _
                              "Prijemnica date missing or invalid. PrijemnicaID=" & prijemnicaID
                End If

            End If

        End If

    Next i

    If latestDeliveryDate = 0 Then
        GetInvoiceDeliveryDate = InvoiceDate
    Else
        GetInvoiceDeliveryDate = latestDeliveryDate
    End If

    Exit Function

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.Description
End Function

Private Function GetFakturaDateForSEF(ByVal fakturaID As String) As Date
    On Error GoTo EH

    Const SRC As String = "modSEFMapper.GetFakturaDateForSEF"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "FakturaID is required."
    End If

    RequireColumnIndex TBL_FAKTURE, "FakturaID", SRC
    RequireColumnIndex TBL_FAKTURE, "Datum", SRC

    Dim v As Variant
    v = LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "Datum")

    If IsEmpty(v) Or IsNull(v) Or Not IsDate(v) Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "Faktura date is missing or invalid for " & fakturaID
    End If

    GetFakturaDateForSEF = CDate(v)
    Exit Function

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.Description
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
        
        ' manuelles Modulo, damit h nie zu groß wird
        h = h - Int(h / 2147483647#) * 2147483647#
        
    Next i
    
    ComputePayloadHash = Hex$(CLng(h))

End Function

Private Function GetSEFPaymentDueDays() As Long
    On Error GoTo EH

    Const SRC As String = "modSEFMapper.GetSEFPaymentDueDays"

    Dim rawValue As String
    rawValue = Trim$(GetConfigValue("SEF_PAYMENT_DUE_DAYS"))

    If rawValue = "" Then
        GetSEFPaymentDueDays = 15
        Exit Function
    End If

    Dim daysValue As Long

    If Not TryParseLong(rawValue, daysValue) Then
        Err.Raise ERR_SEF_CONFIG, SRC, _
                  "SEF_PAYMENT_DUE_DAYS must be numeric."
    End If

    If daysValue < 0 Then
        Err.Raise ERR_SEF_CONFIG, SRC, _
                  "SEF_PAYMENT_DUE_DAYS cannot be negative."
    End If

    GetSEFPaymentDueDays = daysValue
    Exit Function

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.Description
End Function

Private Function GetFakturaIssueDate(ByVal fakturaID As String, _
                                     ByVal sourceName As String) As Date
    On Error GoTo EH

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, sourceName, "FakturaID is required."
    End If

    RequireColumnIndex TBL_FAKTURE, "FakturaID", sourceName
    RequireColumnIndex TBL_FAKTURE, COL_FAK_DATUM, sourceName

    Dim v As Variant
    v = LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, COL_FAK_DATUM)

    If IsEmpty(v) Or IsNull(v) Or Not IsDate(v) Then
        Err.Raise ERR_SEF_VALIDATION, sourceName, _
                  "Faktura date is missing or invalid for " & fakturaID
    End If

    GetFakturaIssueDate = CDate(v)
    Exit Function

EH:
    LogErr sourceName
    Err.Raise Err.Number, sourceName, Err.Description
End Function



Private Sub ValidateSEFDtoForUBL(ByVal dto As clsSEFInvoiceSnapshot)
    Const SRC As String = "modSEFMapper.ValidateSEFDtoForUBL"

    If dto Is Nothing Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "DTO is Nothing."
    End If

    If Len(Trim$(dto.fakturaID)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "FakturaID is missing."
    End If

    If Len(Trim$(dto.InvoiceNumber)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "InvoiceNumber is missing."
    End If

    If Len(Trim$(dto.CurrencyCode)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "CurrencyCode is missing."
    End If

    If Len(Trim$(dto.SellerName)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "SellerName is missing."
    End If

    If Len(Trim$(dto.SellerPIB)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "SellerPIB is missing."
    End If

    If Len(Trim$(dto.BuyerName)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "BuyerName is missing."
    End If

    If Len(Trim$(dto.BuyerPIB)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "BuyerPIB is missing."
    End If

    If dto.TotalNet <= 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "TotalNet must be > 0."
    End If

    If dto.TotalVat < 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "TotalVat cannot be negative."
    End If

    If dto.TotalGross <= 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "TotalGross must be > 0."
    End If

    If dto.Lines Is Nothing Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "DTO lines collection is Nothing."
    End If

    If dto.Lines.count = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "DTO has no lines."
    End If
    
    If dto.InvoiceDate = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "modSEFMapper.ValidateSEFDtoForUBL", _
              "InvoiceDate is missing."
    End If

    If dto.DeliveryDate = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "modSEFMapper.ValidateSEFDtoForUBL", _
              "DeliveryDate is missing."
    End If

    If dto.DeliveryDate > dto.InvoiceDate Then
        Err.Raise ERR_SEF_VALIDATION, "modSEFMapper.ValidateSEFDtoForUBL", _
              "DeliveryDate must not be later than InvoiceDate. DeliveryDate=" & _
              Format$(dto.DeliveryDate, "yyyy-mm-dd") & _
              " InvoiceDate=" & Format$(dto.InvoiceDate, "yyyy-mm-dd")
    End If
    
    Dim i As Long
    Dim ln As clsSEFLine

    For i = 1 To dto.Lines.count
        Set ln = dto.Lines(i)

        If Len(Trim$(ln.naziv)) = 0 Then
            Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "Line " & i & " has no item name."
        End If

        If ln.kolicina <= 0 Then
            Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "Line " & i & " quantity must be > 0."
        End If

        If ln.cena < 0 Then
            Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "Line " & i & " price cannot be negative."
        End If

        If ln.Neto < 0 Then
            Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "Line " & i & " net amount cannot be negative."
        End If

        If ln.PDV < 0 Then
            Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "Line " & i & " VAT amount cannot be negative."
        End If

        If ln.iznos < 0 Then
            Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "Line " & i & " gross amount cannot be negative."
        End If
    Next i
End Sub

Private Function GetRequiredSEFConfig(ByVal keyName As String, _
                                      ByVal sourceName As String) As String
    On Error GoTo EH

    Dim valueText As String
    valueText = Trim$(GetConfigValue(keyName))

    If valueText = "" Then
        Err.Raise ERR_SEF_CONFIG, sourceName, _
                  keyName & " missing in configuration."
    End If

    GetRequiredSEFConfig = valueText
    Exit Function

EH:
    LogErr "modSEFMapper.GetRequiredSEFConfig"
    Err.Raise Err.Number, "modSEFMapper.GetRequiredSEFConfig", Err.Description
End Function

Private Sub ValidateGeneratedUBL(ByVal xml As String, ByVal sourceName As String)
    If Len(Trim$(xml)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, sourceName, "Generated UBL XML is empty."
    End If

    If InStr(1, xml, "<Invoice ", vbTextCompare) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, sourceName, "Generated UBL XML has no Invoice root."
    End If

    If InStr(1, xml, "<cbc:ID>", vbTextCompare) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, sourceName, "Generated UBL XML has no invoice ID."
    End If

    If InStr(1, xml, "<cac:InvoiceLine>", vbTextCompare) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, sourceName, "Generated UBL XML has no invoice lines."
    End If
End Sub

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
    Debug.Print "Naziv: "; ln.naziv
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

Public Sub DebugSEFDates(ByVal fakturaID As String)
    On Error GoTo EH

    Dim dto As clsSEFInvoiceSnapshot
    Set dto = BuildSEFInvoiceDto(fakturaID)

    Debug.Print "=============================="
    Debug.Print "SEF DATE DEBUG"
    Debug.Print "FakturaID=" & fakturaID
    Debug.Print "DTO InvoiceDate=" & Format$(dto.InvoiceDate, "yyyy-mm-dd")
    Debug.Print "DTO DeliveryDate=" & Format$(dto.DeliveryDate, "yyyy-mm-dd")

    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)

    If IsEmpty(stavkeData) Then Exit Sub

    Dim colFakturaID As Long
    Dim colPrijemnicaID As Long

    colFakturaID = GetColumnIndex(TBL_FAKTURA_STAVKE, "FakturaID")
    colPrijemnicaID = GetColumnIndex(TBL_FAKTURA_STAVKE, "PrijemnicaID")

    Dim i As Long
    For i = 1 To UBound(stavkeData, 1)

        If Trim$(CStr(stavkeData(i, colFakturaID))) = fakturaID Then

            Dim prjID As String
            prjID = Trim$(CStr(stavkeData(i, colPrijemnicaID)))

            Debug.Print "Stavka row=" & i & _
                        " | PrijemnicaID=" & prjID & _
                        " | Prijemnica.Datum=" & _
                        CStr(LookupValue(TBL_PRIJEMNICA, "PrijemnicaID", prjID, "Datum"))

        End If

    Next i

    Debug.Print "=============================="
    Exit Sub

EH:
    Debug.Print "DebugSEFDates ERR " & Err.Number & " - " & Err.Description
End Sub
