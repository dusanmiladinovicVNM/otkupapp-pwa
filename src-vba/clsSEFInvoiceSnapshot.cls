VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSEFInvoiceSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public fakturaID As String
Public versionNo As Long

Public BuyerID As String
Public BuyerName As String
Public BuyerPIB As String

Public SellerName As String
Public SellerPIB As String

Public InvoiceNumber As String
Public InvoiceDate As Date
Public DeliveryDate As Date
Public CurrencyCode As String

Public TotalNet As Double
Public TotalVat As Double
Public TotalGross As Double

Public Lines As Collection
