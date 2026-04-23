VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSEFResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public HttpStatus As Long
Public apiStatus As String
Public Success As Boolean
Public Accepted As Boolean
Public Rejected As Boolean
Public CorrelationId As String
Public sefDocumentId As String
Public SEFInvoiceNumber As String
Public errorCode As String
Public errorMessage As String
Public RawBody As String
