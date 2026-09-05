Attribute VB_Name = "modLogo"
Option Explicit

'=====================================================================
' modLogo - brend logotip kao KOD, ne kao .frx
'
' GENERISAN FAJL. Ne menjaj rukom: pokreni  python tools/logo_to_vba.py
' (izvor: img/AgriX-Otkup-Logo-Final.png).
'
' ZASTO OVAKO: MSForms sliku uzima samo kroz LoadPicture, a LoadPicture cita
' FAJL. Ranije je logotip ziveo u .frx-u forme, sto znaci dve stvari koje ne
' zelimo: .frx se ne pravi iz koda (CLAUDE.md par.3), i .frx NE putuje kroz
' self-update -- svaka promena logotipa bi trazila REINSTALL na svakoj masini.
' Ovako logotip putuje kao kod, obicnim self-update-om.
'
' Format je GIF jer LoadPicture ne cita PNG (zna BMP, RLE, ICO, WMF, EMF, GIF,
' JPEG). Na ovom znaku je GIF bez gubitka -- ima svega nekoliko boja.
'
' POZADINA JE PECENA U SLIKU. MSForms ne zna per-pixel alfu, pa bi providan
' PNG svejedno bio spljosten -- samo na boju koju bi MSForms izabrao umesto
' nas. Zato je kompozit uradjen unapred, na tacno onu boju na kojoj slika
' stoji, a ista ta boja izlazi kao LOGO_BG_* -- crtanje i slika se ne mogu
' razici.
'
' Fajl je 100% ASCII (Base64 i jeste ASCII).
'=====================================================================

' SPLASH: 353x115, 5447 bajtova GIF-a
' SPLASH2: 705x230, 9595 bajtova GIF-a
' KARTICA: 107x35, 2013 bajtova GIF-a
' KARTICA2: 215x70, 3530 bajtova GIF-a
' MINI: 83x27, 1716 bajtova GIF-a
' MINI2: 166x54, 2817 bajtova GIF-a

' Osnovni kljucevi (1x). Varijantu za OVAJ ekran daje LogoKljuc.
Public Const LOGO_SPLASH As String = "SPLASH"
Public Const LOGO_KARTICA As String = "KARTICA"
Public Const LOGO_MINI As String = "MINI"

' Boja na koju je slika pecena. Ploca iza slike se crta BAS ovim, pa se
' pravougaonik oko znaka ne vidi.
Public Const LOGO_BG_SPLASH As Long = &H12281A   ' RGB(26, 40, 18)
Public Const LOGO_BG_KARTICA As Long = &HEEF4F7   ' RGB(247, 244, 238)
Public Const LOGO_BG_MINI As Long = &HEEF4F7   ' RGB(247, 244, 238)

' Ucitane slike po kljucu -- dekodiranje i upis na disk idu jednom po sesiji.
Private mKes As Object

'=====================================================================
' Slika za dati kljuc, ili Nothing ako je ucitavanje palo.
'
' NOTHING JE OCEKIVAN ISHOD, ne greska: MSXML ili ADODB mogu da nedostaju, a
' TEMP ume da bude nedostupan. Pozivalac tada crta tekstualni znak -- zato
' modUiFaze i dalje nosi natpise AX / OtkupApp.
'=====================================================================
Public Function LogoSlika(ByVal kljuc As String) As Object
    Dim p As String
    On Error GoTo EH
    If mKes Is Nothing Then Set mKes = CreateObject("Scripting.Dictionary")
    If mKes.Exists(kljuc) Then
        Set LogoSlika = mKes(kljuc)
        Exit Function
    End If
    p = UpisiPrivremeni(kljuc)
    If Len(p) = 0 Then Exit Function
    Set mKes(kljuc) = LoadPicture(p)
    Set LogoSlika = mKes(kljuc)
    Exit Function
EH:
    LogErr "modLogo.LogoSlika"
End Function

' Otpusti ucitane slike (self-update rusi runtime, pa i ovaj kes).
Public Sub LogoOtpusti()
    Set mKes = Nothing
End Sub

' Base64 -> bajtovi -> privremeni GIF. Vraca putanju ili "".
'
' Binarni upis ide istim obrascem kao modDrive.DriveDownloadToFile
' (ADODB.Stream, Type = 1, SaveToFile ... 2) -- nema drugog nacina u ovom
' projektu i ne uvodi se treci.
Private Function UpisiPrivremeni(ByVal kljuc As String) As String
    Dim dom As Object, cvor As Object, stm As Object, p As String, b64 As String
    On Error GoTo EH
    b64 = Base64Za(kljuc)
    If Len(b64) = 0 Then Exit Function
    p = Environ$("TEMP") & "\AgriX_logo_" & kljuc & ".gif"

    ' Fallback ide kroz Resume Next, ne kroz 'If dom Is Nothing': CreateObject
    ' nad nepostojecim ProgID-em DIZE gresku, ne vraca Nothing -- provera na
    ' Nothing se nikad ne bi izvrsila, a stara masina bi ostala bez logotipa.
    On Error Resume Next
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    If dom Is Nothing Then
        Err.Clear
        Set dom = CreateObject("MSXML2.DOMDocument")
    End If
    On Error GoTo EH
    If dom Is Nothing Then Exit Function
    Set cvor = dom.createElement("b")
    cvor.DataType = "bin.base64"
    cvor.text = b64

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1                 ' adTypeBinary
    stm.Open
    stm.Write cvor.nodeTypedValue
    stm.SaveToFile p, 2          ' adSaveCreateOverWrite
    stm.Close

    UpisiPrivremeni = p
    Exit Function
EH:
    LogErr "modLogo.UpisiPrivremeni"
End Function

'=====================================================================
' Odnos stranica izabrane slike (sirina / visina).
'
' Okvir se racuna po NJEMU, ne po zajednickom broju: varijante se razlikuju
' u cetvrtoj decimali zbog zaokruzivanja piksela, a okvir koji odnosu ne
' odgovara ostavlja pojas pozadine sa strane -- vidljiv pravougaonik oko
' znaka na gradijentu splash-a.
'=====================================================================
Public Function LogoOdnos(ByVal kljuc As String) As Single
    Select Case kljuc
        Case "SPLASH": LogoOdnos = 3.06957
        Case "SPLASH2": LogoOdnos = 3.06522
        Case "KARTICA": LogoOdnos = 3.05714
        Case "KARTICA2": LogoOdnos = 3.07143
        Case "MINI": LogoOdnos = 3.07407
        Case "MINI2": LogoOdnos = 3.07407
    End Select
End Function

' Visina slike u pikselima -- po njoj LogoKljuc bira varijantu.
Private Function LogoPxH(ByVal kljuc As String) As Long
    Select Case kljuc
        Case "SPLASH": LogoPxH = 115
        Case "SPLASH2": LogoPxH = 230
        Case "KARTICA": LogoPxH = 35
        Case "KARTICA2": LogoPxH = 70
        Case "MINI": LogoPxH = 27
        Case "MINI2": LogoPxH = 54
    End Select
End Function

'=====================================================================
' Varijanta slike za OVAJ ekran: 1x ili 2x.
'
' ZASTO UOPSTE POSTOJI: MSForms skalira sliku bez uglacavanja (StretchBlt,
' COLORONCOLOR), pa smanjivanje prosto ISPUSTA redove i kolone. Na kosim
' potezima znaka -- a "A" i "X" su same kosine -- to izlazi kao stepenice.
' Zato slika mora da ima priblizno onoliko piksela koliko okvir STVARNO
' pokriva na ekranu, a to zavisi i od rezolucije i od DPI-ja.
'
' Bira se varijanta cija je mera BLIZA U ODNOSU, ne u razlici: rastezanje
' 1.5x i skupljanje 1.5x nisu isti gubitak, a poredjenje kvadrata sa
' proizvodom (geometrijska sredina) ih izjednacava bez logaritma.
'=====================================================================
Public Function LogoKljuc(ByVal osnovni As String, ByVal visinaTacaka As Single) As String
    Dim pxTrazeno As Single, po1000 As Single, px1 As Long, px2 As Long
    On Error GoTo EH
    LogoKljuc = osnovni
    px1 = LogoPxH(osnovni)
    px2 = LogoPxH(osnovni & "2")
    If px1 <= 0 Or px2 <= 0 Then Exit Function
    po1000 = PixelsToPointsY(1000)          ' 1000 piksela u tackama, na ovom ekranu
    If po1000 <= 0 Then Exit Function
    pxTrazeno = visinaTacaka * 1000 / po1000
    If pxTrazeno * pxTrazeno > CSng(px1) * CSng(px2) Then LogoKljuc = osnovni & "2"
    Exit Function
EH:
    LogoKljuc = osnovni
End Function

Private Function Base64Za(ByVal kljuc As String) As String
    Select Case kljuc
        Case "SPLASH": Base64Za = B64_SPLASH()
        Case "SPLASH2": Base64Za = B64_SPLASH2()
        Case "KARTICA": Base64Za = B64_KARTICA()
        Case "KARTICA2": Base64Za = B64_KARTICA2()
        Case "MINI": Base64Za = B64_MINI()
        Case "MINI2": Base64Za = B64_MINI2()
    End Select
End Function

'=====================================================================
' Slike. Svaka je svoja procedura -- VBA ima granicu velicine procedure, a
' jedan zajednicki blok bi je s vremenom probio.
'=====================================================================

' SPLASH -- 353x115, 5447 B GIF-a, 7264 znakova Base64
Private Function B64_SPLASH() As String
    Dim s As String
    s = s & "R0lGODlhYQFzAPcAADxjI1BRJD1EHsCiSERxJ1CHLUh5KT1lJC5KHDBNHcWmSlqaM1yeNFSPMFiWMig/GTRVHzhdIStFG0p9KkyALI9+OCw1GFNSJVVUJZyIPB0qEx0sEx4sE0RyJ0Z1KKWOQKeQQDFPHjJRHkJuJkNwJ1KLLlONL06ELE+FLSY7GCc9GGZhK2hiK3ht"
    s = s & "MDdaIThcIT5DHj9EHkyAK0yBLEtMIkxNIh4uFB8vFCEuFENIIEVIIH9yM4B0M4F0NEBqJUBsJSUwFSYyFoR2NIR3NYZ4NSo0Fyo2GL2gR76hSHRqL5+KPsKkScSlSSM2FiQ3FxooEhspEhsqEjI6GjxBHT1CHVCJLlKLL2liK2tkLGtkLW1lLUl7Kkp9K4p7N4t8N8am"
    s = s & "SsioS1tYJ11ZKFubM1ydNF2fNF6hNbWaRbmdRrqeRpSCOjM8GjU8GzU8GjY9G1WQMFaSMViXMlqZM6mSQaqSQauSQayUQqyUQa2UQrCWQ7KYQ0hLIkpMInxwMn1yMzVWIDZYIGReKmRgKkFGH0JGHyxGGy1IGy02GC44GZ2JPZ+KPR8sFHhuMXluMSIuFYF0M4N2NCYw"
    s = s & "FiUxFigyF4d4NYh5NnRrMHZsMKCKPqGMPhwqExwrEzM6GjM8G7WaRLecRZWCOpaEOx8sEyEtFCMuFSIwFSgyFig0Fzc9Gzc+G1aSMFaTMUd3KUd4KTJSHjNTHypDGitEGo19OI5+OEdJIUhKID9oJT9pJbueR7ufR1JRJFJSJHJpL3NqL5uHPJyIPaKNP6KMPiAxFSEy"
    s = s & "FSMwFSMwFSI0FiI1Fi84GS85GTpAHDpAHaSOP6WPPz5nJD5oJYl6Nol6N25mLW5mLnBoLnFoLlZWJldVJlhWJlhWJl1aKF1aKLSZRLSaRLicRbicRpOBOZOBOpeEO5iFO5mGPJqGPKiRQKmRQa+WQq+WQylBGSpCGkhLIUlKIXtwMXtwMjpfIjpgImBcKWBdKWJdKWJd"
    s = s & "KmVhK2ZgKkBGH0FFH01PI05OI09RJFBQI5GAOZKBOiwAAAAAYQFzAAAI/wCfCBxI8EmUDQhtKFzIsKHDhxAjSpxIESLCTVEKatzIsaPHjyBDihxJsqTJkyhTqhS5aUPFlzBjypxpY0PGlThz6tzJs6fPnx030RxKtKjMDVCAKl3KtKnTpwVdGp1KtepCTVCzat3KlatU"
    s = s & "q2DDztzUtazZs2hJfhXLtq1ErGnjyp2rda3bu3gV3qTLt69flZryCs7L4a/hw4g52h3MGCzZxJAj0xXauLJYyZgzl11suTNRuJpDiwYaxbPpqRtGq16dk/Lp1zRZy56tFrZtmXtp6949kPPt3w5z8x4uG7jxt8STFz/OvCFo5dA1N5+u8Hn065CpT7eOvftf7c25e/8f"
    s = s & "Pxc8c/Hk0581fxy9+vdb2Rt3D7++09flzOk35+sGQy5WBBggK0R1IGCAFBRFn30MAnWaPmBEKKE2DJFgxoUYIkATLBhieIuCBVXT4IgOmjaOhBI2w1ACHV54Ak0TtEiGEyASJCKJOO5kmgAoosjOQsLI0aIZ6sj0wJAE1jjQjTk2qZJpRPQo4TgM+TCkBzKNMCQsRlnH"
    s = s & "pJNgluQZIktIKaEbC6VQRotlqACTE2S0iMJUXoZpJ0meXWKmhEQwxMqVMGnZogh0hnjnoSB1Row3e0Y4wCkLFTLkjBUJM0aLcPjXpaFhHlKLNGwgylFnLDQqYS8MnTCkMxWJMGQ8VNX/+d4s0PTQRzXVtNCHEBWQY0eZEeog6kadmWNqhGdoaoOrLcphDEUetDjGs4Xa"
    s = s & "CJ8ex6Io7LAFWXZBthFSuNAbq04UZIsjVCXrezhcgO2x23I7UGU4GCvlF1LqUcpCzHborERW+ktjrJz6tIgFnKCiMDKTaPCXKXTAKy9BlZUqpTm9mGlNqkNGEBGcLQJg1bo4aUIFFqEc0egnoKxABV8QmhrvxJXVYSY2ZEo5BEP9YvgiRIJiSEYTIxesEhU7pIFiJlnw"
    s = s & "oQwV+vQSMYqfCBGDXI5IPLFAjX0r5TfE2BBlj94oawO5HZYxcEMgd0gCWCSfpMMvPdrx8kb92CElCBgs/5LWJzJvzTVjvpiphUJuKCDlBQy9MKQPDzmzJdxGl7QM3XUX4ZEo5ZgJQj8+xQANMiOBELjgjPEoJROQKlSJlIkwZIyQ/lILJO0YyhBW3CJBcYniPYIABEim"
    s = s & "vCulF5rjxAYLzERIukimNzqzvIzJYmYoDCUuZQ4M2VIuQ/EMGcLulYu0g5l6HCISMmfsqUeoJclzTzU7qNE+is+HFP2e03M7WM5SWkFDvCClSjBETc0SBpAc0CITiIV3INGTlBTQP48QAnhSOgcOSqKyRuUPJPszUwVFNRhG7KkTDVEdihSAwoVEq0UuWIjjWgQr8lmL"
    s = s & "JNDY0w5MMo1GKWKDI+ngnv8++JEQSmmEiBIMMYSIIl885EQ9QtVCYjGkN2iqAQl8YPk80og9fQKIJNGE3vaUAb+JhIlSIqJHjNgjJB5KMBZb3EP2IKUjtE4hqmoRIJY1pB+wBYIcuUKjVoCSmO1JCEE0lRo7wkZtCe4JeTkFo6TkRIi8rkeWYIgE1tShVaRAFS1awDH+"
    s = s & "uMWN0ANfZsKDKFICClO14IyKLN3ptpaXd+xJABE5BRpZyJAXdggOgCLlDT+iA1OFQSVSOBY2QoLGHi2SI42UkBvvhBccNBMMlYyICaPIECoOaUgQaAsgCbII40lpCcNTSfMa5QkzeuSaEnrmRqIZrEfihR97YpxEEIH/wQjZkWPfFJrtbLgkkKzAVGrAyUFNRciPwNN5"
    s = s & "spSePe/SAzOZoyKXRNE8GBKCgGIoScIsqEcWAbhGyQMnhzhWNxb0UDDIUyP0BMM0f7II9aGkpirBqUZMho35seAC8PsIXswpIRZUxA2olFAezMZAj3JJnKUcyEIbJYWcxBQMyxwIG8LA1a4Ca08s6KpYw+DOgcS0f7sIQC10EANlsAEZcI1rKpShDELoIwgCiUEtqPBW"
    s = s & "uEpBGfkIgDwu4Q86KKAPGknFXvsKV2XEoAby6AU0QKAAImxEsXyN618DG9nJVrYgOpjFNfVAD6G6xZAoAttLEiGlADCEAB6dk1vGKZAx/+4pDTqxJUIJUg9wgSudBDmrRlpqJhgIZGzgQmyIfAsGy2qkGsx17hMWIYRs/UKedwGHmRgBE0HAjiGu8CihZhvVJ8TgWAnN"
    s = s & "idcapQDgWkAH8I3vAEyFgfjaVwfcES5ow/COkrKXHCtI5zK0QYSvNkq5BGGDNoRg4D1JN8ELbrCZnIuMdUq4R0vQx6jaokIULQERMEmGmfBxu29m6i60lWCjLqGTZB4rAO+MJfRmqRFRVKOfKCJCVTciBf/uCcEaQQZRpfTgIN+vUZYV8hFWgFdR8EER7IWxRtzCAzMZ"
    s = s & "MCb/kJIUFQLbIR0AL7TFw7GOqZMLo2iHHWnpSwui345oQ/9KdriaR8JwLCBrpBbHKvKd87wIOvzCpgMRRQUatYS7UYwtAOwRLmNCR0UzZAbf/Bl5h6mRImRLzlY9ViZi7MGI8g8k5iTCKj+iiZbaWSNDzjGoTUWEd4hj1AURRToadYeyQpItGZMSlWbCWgn944Ae1dCk"
    s = s & "RbqR9TYK0DgZtKmW4B41e1qEIPFHhBQQiJGQw1SnLghyJwySHDYqFOkAo0bo3ChqdEssa2gmEwZBk2XMFwwDaKFCfhBbFEe1BcdiQlJ00sVjtSHNMtYfjTeiATGDQBkk6fePP6IFVoOk4aaqgUeg8IFGDcANhw4LJMy0s6Hkess2EEYcPGqGPUKV0gX/ubap9MCTqTZK"
    s = s & "wxxx9owl6pGDggDWIoHugT8iD4d/JBCmAoFBTQWNjFsl0ShaNE3I9GGekdwMbxi2QL5UkFmbag48IbeptAHwTs/80x0hxBKWgPCS6Hzhc/Z52pEMklSYqr3zCkuue7RrokDCDw0pwdPNUIiTE3u4x1IETwKQrYZuROYCp/lG2vAJBdDgJGc3U7YJonVuf6TyRAaJJswc"
    s = s & "ISnf2iqLMpMfLkD60pv+9KgnvT1AzK+9m+EZfp96R6w7+GxRnSCIB+HAB0KIIyjgAiiJvJQmPxDM90jP41a7R9apw7hbxR7MjT4YutCQArjegSGV/UaAQPudoHZPtx9I/+6LuPsnEGK+ZIY8ti+vfI4YX9UgIaAPnV8Ve0kfXEq3gaSGtIVvCpug2qcR3Hcsv8ATNGB7"
    s = s & "XTdEz3ZEdzZfK5USwtcjxCcQ78cnIFGBEYJ8GqFwZrIE9DcVxnZ/pvJrDBEjDWQDwNQiuqNFKDcQA2gqBbgThHcsWJCAZrJmwTVwNPBVWQCB67d2DnaB7ccRcbQnyQMWUCSCx5J/HPIqNgAA3+QmAPgE4fcEmtB9OjGDpnJSMRdwuqd4T7CDKDIAyEYSEYgiE/gEGNhc"
    s = s & "Qsh2IPFmpnI3VtFhSrgnsVOCQyIHCjQ7wUQ5LTgQnBchzJB12YIBNphGC9hGAyGGPeIFwf/3gx2xhhpIeUNYbMdSC4NDFRmFInoQDp74iaAYiqL4iT+SJnGCLgthIWwihUXzhwLhY2ZiBzwBdMeCaYDndYn3aVAQCGaGiSZxhhKShpLYhkEIElq4J1JWFYPABGYiLlZh"
    s = s & "gh0SB6OkEE4wci0yAVNYhU+QAceCBDyRBMfyBaRwiM6UiCjiBV5gdXviCZHwi5DofpVYEMMIEsVkKslIFYVTR0EAFj2DITW0EFA4JP9HMK74BOdzLMCFE9FwLHTAaQr4dWYCi3tibmb3jhsxj0BoecT0YplYFHQoIXgHFigwKQNlA3zYIpJGkH9XED13LGWXE0pwLIhE"
    s = s & "jvhjjtIUBMr/xk4L8lwWmXxumJGZR4/HAnNUIX8YJm9UsX+o6BCq2CKG0IorSRDndSygoxMS2SPpd3heSH4SNQrqCFYVuXNAeXzEqJEeUY+NQggdORRI52thAY0dwooMoQKc1CFbAJUBuFPv1ijSoBNBkC3J04W4+IVgh5Z7cgRl+BHAGCHCGI+U+JMecYCmAj9TsU3b"
    s = s & "AxZN2CJVEBEGMCRtoi7lpQbHMgs6YZhmsmkOeYM2WU8CkZN7UnQjsZhg0JiQCY+1yREhKCVAZBTWZCbAEBb9NyQfAhEI8E1YopJ56ZONwnI54XJmYniC+ZC5CG0CIQU4xoA515Py6JjFx53d2Si4tZYz/4FPZqJPVZECp9giEiARI8kmKYCcVOgRQBCIYJCYJyFthJaQ"
    s = s & "WjmYXAl2AlFRpnIHOzl12vmYxTiW8PcRWGAqGfCBM+EH6GM2UxE0vzQRf/BNflQtUaltx8J1ONFS7wAS47dGNAYExMUC2SmWkeidasiiT9AHptIIDioTqWZUVtE2NEQRTVU7GpqciXUs/oATU7knSxCYNBlPqylTBdFDpoKYISGbtHmgK3qbGzELpgJzn0cTktkjR7CP"
    s = s & "VkFvJEkRrfBNrLIpBTkQrnlO+mkSc2cmISqiW0miYPgEopBqKAKbilmg3ymltsmn0JSfMwoT2iUlIVkVJ7mUE/EK3yQHw/9gphtaEFIQiB6aEnYKBl8UEkrDn3Lqn3s6pxshDXoqEHDop8pplhyhCXspJaFgdDLBB2aCBkYAFhSKIXKwNhIhDDvaIZCjJD66EdiQVFKS"
    s = s & "B6OQErkpIWhgix4hZprKSLsHBUa5J9xgAT8Xqk/ACyyKkR2RD6bSDqwaE/aHIjZ6ngElMi8BCN/0mZ9RXgRRhFJCkSZRTkSqliIRDHHKrJ76BCZ6LOTwERhArSpHqttJpQXRpj0Sg4FKEd/nT15aFfDwTZTyEufShzRBWwVhpYdppCOBgVklEqK5rPbZZgVBi/R1lqE6"
    s = s & "n9fKopkwp0MxqD1SqFUBaUPyNjEhOZ5pqzL/QbEEIQri0Ch4OhIktSd9SRIwuqwgkF45eK9W+JVm4gniVhAp1SiO2BHOKSVGO6UAOxDI0ChBem4y8ZFg0HQ3ygDf9FQwYQx12SFlOrHqGms7uydYKhKPcEgmsYaG1p01WHWmwgcdUaw9wmIdkalmIngcUQRpEFMNyX4C"
    s = s & "OxC6VUdqRBMQKiUkaBWzeiErKBO+5C8KpLZnuhGi4G1SkgY4OBDHiCL+gHMi8bR7Ap1PgAxH8An6uQjXiSJcyBGYw14vqRGeO0FrOl3koActuSdwh6AW2BHIYGZve7AR0ZYRkn9TYQzp2SHjJRMqEFCvMBQ4uxEHKSXogLEcIXbnhKIo/1G7UuIJ+ZMKngAGszsQ/XAs"
    s = s & "UcsRpmkmd2C6A/G+KPIIAkgOS6ADawgGk9qnploQ4osiQTtlXdsLBnzAvRCuVbEOztDADuwM5EoT8PDADpwA1ru2HCEPgMuJ3FsQqJCqEfILQXUSyxC7EbIEs9AHXjBfVSsQRaC0EySvG/GXCMgR/2omj/A8RcACgHNS+0u+wiuItiYQmuAOZqIA4Ess8vEb1zu47yBh"
    s = s & "dTDCBcEGyioh6QB8OCGyx6IEsJYKfNALV3lO1EAIRhoEtdCx+fYOeKURh1CpjpKpS8CF+wtvl1ADHby/mbAMBREDUPa5vrhhS2wbTdwRQLACzHfCXlALYP+0CHwADRikAGqgtzuhDfQJDOJGXOBHEFPLkRuBDFVsKpmAaXMsIeG3hmpAQHhACX0ADZ+MIiCAcR4RyLcx"
    s = s & "yB8hBSygBhusAHoAAnqAQUcwC9qwxj2BDNZzxEQgzALBAisgD2GAAQFAC21FV3QVAzrABxgQBvLAAjMTAywQCGGADfvAVtL8NDqgD9igDSwgxQRhCo/AeZjwxxT4tczgBZewAmEQAPoQX/qAAdogwwbKbYTgD2Z2BEQwU1kqy6ZhHajgE4eQDzVwDdpQD/WgDdjQD/mA"
    s = s & "DPu2FKZACBcQ0WKgC1MwrIkhCsqwD9gg0WLQDzFgChthCoeQ0XOrdqIAA9f/IA/1IAYXQAsvPRIIDRsD+kiYga040dOv8dNAHRlCvRJEfRpGfdSJkdRPstSe0dROfRhQnRJSPdVVLRtXjRJZ3RlUvdV+Mapm0r458dWWEdZizRfs2iOCqxNoXRlqvdZzkb1HvLteHdeM"
    s = s & "Mdd0HRdT0ygyqhO+oddiIRx9HRlk3Si9MMQnMdiEDRYwfdiHsQgxgA25eyxpQAQYEAN4HRKu8dhuIdmRoQwDgAt6YAfnAAyqvdqsDQx1UAd6cAa5MACeZxJQANp34TCiHSaOjdtFEdm7jSOl4dthkRrBHSYcQNxgYdjHTSK3rdxUYdzNDSaBAd1FId3TDSafbd0ykd2H"
    s = s & "LFLd3A0T2O3dYQLe4T0R403edrLd580QNqHe8gIFLZEQ1o0QGqAJwA3fIxIQADs="
    B64_SPLASH = s
End Function

' SPLASH2 -- 705x230, 9595 B GIF-a, 12796 znakova Base64
Private Function B64_SPLASH2() As String
    Dim s As String
    s = s & "R0lGODlhwQLmAPcAABooEhopEhopExsoEhspEhspExsqEhsqExsrExwpExwqExwqFBwrExwrFBwsFB0rEx0rFB0sEx0sFB0sFR0tFB0tFR4sEx4tFB4tFR8sFB8tFB8uFCAwFSAxFSI0FiI1FiMxFSUwFSUwFiY7GCY8GCc9GCg0Fyg/GSk1Fyk1GCo0FypDGi5KHC5K"
    s = s & "HS84GTA4GTA4GjBNHTBOHTJRHjJRHzJSHzM8GzQ9GzRVHzU8GzZYIDZZIDZZITpAHDpAHTpgIjpgIztAHTtAHjxkIz5EHj5EHz5mJD5nJUBEHkJuJkJuJ0JvJ0NIIERyJ0RyKEVIIEVIIUZ1KEZ1KUZ2KEZ2KUhMIklMIkp9Kkp9K0tMIk5QI06DLU6ELVBQI1BQJFCH"
    s = s & "LlFQI1FQJVKLL1KML1NTJVRUJVVUJVVUJlaSMVaTMViWMllYJ1qZM1qaM1tYJ1tYKFydNF1bKV2fNF5bKV6hNWVgK2ZgKmZgK2ZhK2dhKmhiLGhjLGljLGljLWtkLW1nLXFoLnFoL3NqMHNrL3NrMHZsMHdsMHhvMXtwMnxwMnxwM31wMn5yM35zM4F0M4F0NIZ4NYZ4"
    s = s & "NoZ4N4d4NYd4Noh6Noh6N4p7N4x8N45+OJGAOZGAOpGBOJGBOpKBOpOBOpOCOpeEOpeEO5yIPJyIPZ2IPZ2JPZ2JPqGMP6KMPqKMP6eQQKiQQKiQQaiRQKiRQamRQayUQrKYQ7KYRLOZRLecRbicRbicRr2gR72gSL2hR8KkScioSwAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACwAAAAAwQLmAAAI/wABCBxIsKBBgQQQNHhAocKFhxAjSpxIsaLFixgzatzIsaPHjw8rUGCAoMDBkyhTqlzJsqXLlzBjypxJs6bNmzhz6tzJs6fPnysVOABJtKjRo0iTKo1IoYECoFCjSp1KtarVq1izat1qUAAD"
    s = s & "h0vDih1LtizGCQdMcl3Ltq3bt3Djyp0r0KvZu3jz6gU5gYFauoADCx5MuLBhnQ/2Kl7MeDHaw5AjS55MuTLQAhYaa97MeemEv5ZDix5NuvTcxJ1Tq169kYHp17Bjy549kwHr27hzX3BNu7fv38ApD9VNvLjmBsGTK1/OPCsE49Cj66XQvLr169hfDpfOvXtY5NnDi/8f"
    s = s & "r9y29/PojYInz769e8uo08uf3/q9/fv45xagz7+/RdD5BSjggFBN4N+BCFpA4IIMNkjTAQhG6B9vDlZo4YUAFGCghBzOh+GHIApoXocknkdAiCimOJ6GJbbYHYUqxihjchC6aCN0Ecyo4460bXjjj7lJwOOQRIqmAJBI5pZAkUw2adh2SUbZ2XpOVmmlW/tJqWVnJ17p"
    s = s & "5ZdYGbDlmJodAOaZaEI1Ipls5kVlmnDGSVN8bdZZ1gNy5qnnSxTY6WdZewYq6EFg/WloUhMMquighzaa1KKQ6unopEVFaimclGbq0aWcgqnppxp1KqqVoJZq0aioMmnqqhKl6iqPrMb/esGrtMooK6u15oriravq6iuGvJr667AOBlsqscgSaCyoyTab37KfOivte9Bq"
    s = s & "Ou215FWbKbbcZqctpd2GW923k4prrnJjWmHJuuy2a8kcFh0hxbz01kvvCo2RYO++UnB27r+/jUkLLwQXbDAvuaRQERB0NOzwww5z0RgVEFdMhb8AZyzblnMc7DHBg1TkARsVl8zCYiSUDPHJm60USCAax3zYlqd87HHCFR2hMsQSK0bxznSM0ZnLMMtstGBaWmHzx/BS"
    s = s & "5AHQD7OcV8pQ8zC0Si8frTVdWlaytMeuWPQz1D3nNbbKaaRG9NZswyUlCl9/bEVFLEDtMAlT2/2D2lgX/93231xJSUbcHptiERd207FEXjpDjffVKWUN+ORZSXkI4R7PTVHddrfxAV5pQC2GamtTbnpVUn6C+cGVWBS63UfctYPdSZDe9+m4SyUlLqsfrDBFDHf+eVmv"
    s = s & "A62D7ZH7nfvyPEXZce8Gh+w0ybCXNbvdJyCPkuTMd59TlLBAbzDOFDUOtedkFb9z2tqfxL338NOUpNLiG7yHyInTEbtY11e9WunxC2BMkuS1+hUsbBU52/rGIga7pWF4fEte/EQAAwHCBEk3MODBNDcRztkNX0tpQeIu9r/b5Q4GYLhDJGTBC+VZcCVIKqAGCWa4iiDO"
    s = s & "bsdbyg0dxxoATusObgADFP+gAIMiFqSIRQwCFMDgBj9EghQs9JgLX5gSIGVwhgazAd0SF4WlUA1qJCyhBLE1CywubYpUPAmQnmdGXpDBIg2EGhqWokCVSa19B3mftMrYRimmkSVAGlgfBWGR4EHtjkX5ItCEdhsfSosUfOxjwdD4x4L8iI1trOH07FY2o9SxZFZrpAm7"
    s = s & "5YM7RNKMlKzkQH4kyD7u4nflSxwiP6LIBeLGkd0KRB9TqUoA3GgPkixY6yrCgS/YTQwdKIoHjGk3GeQGl916wikNyEtV2uiKq9PF6jjYwcTtDyTmA1oYRTnGwlAQCm5ww8vWGYh0DrGCDIICKnuJEhvJkHC0oB/hNEn/kR3uDH0gUd/OHnfLUc5FBFBQIStyIT5ZkGIR"
    s = s & "d4CCgEiBxWpW0kXYxBy8aoY5bkrEg0D7Zkf6B7ROknN7FsUKQhfBimAS7KEStc8dKkpPNbaoD72jxUMw+bVhViSOQANoRwSqshoQB5payUAYSMHQuLVCEu1M5x0eMc2l5YITbhBBe4RA05oaxEWtJJz0LhBWq8JyIoYMqUdWkDj26QapWIFCKQhHiqyipJRVXVokYjqe"
    s = s & "rnqVIC3iady0+BA+rG6sExlZ5zzyg8SF8q0G1QoUKBq3SATBJW7Iq81m0YPxaBZzKU1ji8r6NUtEBAXaJBz5Ymm3x2rkCg6E4DMje5XJ/xKOFXx9yR2aSjhcAAKe2KEsNf9qSRLpE3OEhQghVnc/kVEPaG7VCBrsNs7ZlrO2wv1aImrygs96bBZuOEwGFpFbnVBihqGl"
    s = s & "YolA0Tt+PgS1mENgAmW5ERHajaCQvS5VRNCJ1b0BJ4qAHngJAwVU8KK8OTmvBtP7QhJllHAefchyO2oRkO7MpIejLnTg+hM88PZruUAwTXQJvUgANy4iuINwRXwTBQ+XuKvs0D3jplOKwHefF/HnzmY5kVraccO0BQp/ewcGnpC4dwN+SxBU7DEW28TF9WOwBTv04Lg1"
    s = s & "jSIThnCFE4fhiXwSYow0Dod34gPvHuwOPjly7wChFRgEwf8NCv3wwZxcEyiLT8oC7JAgcnqRG8fNpxQBKtB4DBEfgzI6Y84JHqCHZ5cEWHysOPFPoCiLKEKPzjSxM6NhHGMJkXZpesBIlkF8Vomk9cIX+fLDolucRN8EEAKOSkvFNwtJ98TMX8P0TDTdu0bHj0OC/Vpy"
    s = s & "KwI3zCFWIoq9b0UMXbG9ITrIO4EE9HLxgqi8QM6Yw8VlgYLrpelaJrxena/hx6FP28y0GbEE5lY7kXDubHEUUfXD8NtqaONkyNBDs1TCYEBtc3uG345JuEHLaYFI6LiEG7ZFbMBciyQ7qLJ9CAmIajHuuHoms+4dLqoy8K/52yfdtlnAYdLxuI3bexL/4ijmPMERdePY"
    s = s & "IkrwZrsTR+9661cn0hbfI6oiAmxne9s8CfnHRv6Skn/t5N2LUJW/FmGLIHxpTX+IhVUm1IdQnGfduThMFmFAL1hlpgas9a0B7hOjn7HgvkTQjL9WY46o/GvulYiOVSZSku7MqBa3N01gbcAQXEXoB2OFVncC+DmXHb1oR9DSl3bljTzdZlG/wNRLJtSrr9o7Wm8JGDQY"
    s = s & "CazwW4PbJTzZe2J2myGdeQhqhMZB8valMSLHiQPCQ+yuMtc+++Y16bkGgW6VwhvMDqLXINFdUvqPnX55ByIC71Z3CZDYANe70EKF5dC5FaygDZxMZtb1HhNHaBARWkHC/y40uAt948T3BRt+S4rvx4IfyNw2eyNIHn+zUkdk7iXjAv4rVnMg434mMKBBs8AValY/wHd+"
    s = s & "o8cT7Hcwx5c7/lEG4mMCRdF6NtNcFGEE+TNC6JF5KoF+vNB5WyECWIQLGYCAwnd4C5Z4/UGBX8NyRUF/BxN3EBEDGWg3hCZm3OcSfGdA6scT2VU/WWCCPIiCL8Zp/bF4UHcUyxc3UTcCNbhI6cGBJxGAAsgWYAd6Qlg/PbgSC2gwDYg7/bF2bIcUqvNy/fSEtReFOcgS"
    s = s & "kTBDoccVIjB+GrQKWSg+W6gSXThJKjgfSGgzjUcUf7BNW4SGEMNq2/d/L0GFJ8gWHkgwtv8WE414hymRhwTzhafDH2K4NG1nFFWwOq9nQ4T4MM62gWuoEm2oQRvXFpqARSBYE5FIhFG2h/IBf0yTFCCQWnHDbhFxamjYf9IhhUeERZrgFoUwgoNHE69IeojnGzBgV1iB"
    s = s & "QtXWZmAQjSwhAmAQCJHACpVGMLngUBDFezVBH8EmbEoxal9zbBDxcGhYXZhXiihxihoUXm3hBvPkigm4E5TYQrGBQnewCEzliFPhZm6wCJzQVJIIEwJJkAa5Ej4ACD/4NbOwVzZBH7ToMeiWFH5GajkTinTgi7/ojgehiLvnFk9gRql4Ej5ACiq5kiypkli0Ci0ZkzL5"
    s = s & "iAWRjykVCZz/EAmRsAjslE7uNERAGZRAyUQ+CWeBsAiRcIAE4QOTEAhA5AZCOZQ++TI6SQpLeDA0WRBmwJNPGZVAOZXY+ES1IHIuwQlc+ZNRSZTqhI1Q5HMHlhIu8JC9AwnhKB9EID4KdxTmuDQWiGzPVYNddogoVRMFKD4nyRYu0EZOJk8u1UZZORA2yRKNuDqtOBCM"
    s = s & "2ZiPORA76FIj15jphxJQMJm8IHYyMR9xAD0XqRQZuTQy+BAxh4Z4Jx/AOBCTSQpvkQFu2TtvaBCX6ZkalJkAEJkrIZqEU5kC0ZvBBJwAsJnB1Jm+KWIw0F+SxGYDJB+MAD2RVxR7CXkVwVZP2Ab8MZsA/4CcsfgWcgk9h3lE7LSe65Sbxcme8Lmex4gSwqkSd5AIndAK"
    s = s & "npkLdWUQMLAJ+ulSyukFpBCgzekSklCgjYlgUOCeYaecgCUfrsBnY7GaNtOaF7B/O4MF4QmSBQGPGtQFcPFoWHSHyaiAy/gSUJCPBZMLnQAF83kQL5BZkgShAiECNNpH6oejxEkw5cWcBEOcswCOKiEffegxf7gU2yk3g5iBQ9ChiDicZkSkXEGP9SgTJ4qPKQoTL0CJ"
    s = s & "uRAIMboSJIpFNkoQhTmEM3GmWkgQm5kLb3CMM3qeP+cS8pGJH7OJYsFwmPOJZ5iBMwClgykT5GmYccFVZmSbM5Gl5rWlMf/hBg4qRWG6fo6ZpopZE2O6iMtZMKTgZEHggbgAoemBAhVpMEkaFi6Xi/Y3ezXoAYDqPr52qfVDh3HRR5Eqmfe4qClYEz7wqC9FjTKhe2RK"
    s = s & "ExmHqRhnRjFlpblgfijxAsMqYLVaXOcxjjaTl2GhpzRGrRegjlDYqnnkawaqQcMYF8S5hYqaYIwqE1doVco6YmZUpgUxqJdmE/DaOxIFAww1Cz7QEiLQrL2DqEWKHqNaMKlZFix4MAMbEfaVPyPArQahR/raRpaYE57QRuvqEuWKE/VZE+45C76ae+2qsSV6E7zqbQBQ"
    s = s & "Rhz7EtemQUpZT9EKPdgaFjBYMC+LBTWYQx7/4qEC8XlY9F9xAav1E66QeKvmmqs24bMEkwjPKhNyijnuWhBGuzoHCQBPiznoNJodyxJq2luZiR4F+zEHS7ClZRHMtjNdRB/AOLVQKxdWikUDiKVCi7HnKhNrezCi0BNoGzdNSxBzu6Y2sbd2WEb5GhPAWj/+alPdEbOZ"
    s = s & "sxcx+7LyVjJzZLY4CwBLuzpUuhY6i0VJixIX22Jx26gfkwsu0BN+Cz15OxAvELI2Yaicxwu7CRNZSzh05h0oMKGrA2h5sWcfEwcWQXuJgwOQG6UnkQFyiEUlGBc50EdP4LbEiqtFWBOJcDC+5Xc9cbwzVLoCIbxvCxPYO0OkIL0ykQGq/yCAoWu40mGnHoOnekGBGDq2"
    s = s & "5xNx7Qi8BzGvSEYXfVSxtrq8Q9u8AHgwJwsUwWqP+Ku8GnS1MCG/cVO40BodNxCwBFOqePF0kde4aiWbOJuuGoTAb0GcrdsSm/tknfsSINq//lu9NtHBL9GIGCy4I/sxtdod0nozL3sXb7e+lic8FAy/DdtGKdwWxGmcHJy9Hky0+1swIjzCv2kTk5trOPGtd4YTSXx2"
    s = s & "B9EdDMwLX6sXmOTAF8C7NShS7xuoLzGxZgS0cfHEByzAfIuiQhwTMBBJRWzEBmS9AkHGQ+fEQKxbM5SenQYdL+wxMYwXgoS+EVHDiVN1gumqMSHHekUXiP9sM7IKEyacaR/MEhnXxm5cP3AsuXXsEotsMFE7nqibwMbRtRbZGc+DxVr8hLJHijhcEExsQBvsFiwapGZs"
    s = s & "h7DYxDORc1ZbFf9LE5v8mTfRyz56b34VodCBuBvUGagFyBAhyPljiB+5ygSxwlAsF7E8mrMcr8qYxi6By9RmFbs8E8D8lr+cyRbLvVEcHeZ7MPLFGYSAxd7JkQ8Tm4XcrTGxS3RRzbxwzfRay5sGEyIgXLkQuLpMwjURzp1s0PlrQGGqwNDzCu7y0BAd0RI90ezCpxTR"
    s = s & "WPD8MIGJg9A8EPY8F/isz2mbzfq7EvvaogI90Edc0OTMEggNYJ+cx8SRzr7/aUboGBGwldFRcx4ceLoQe899JNJUy8+99hInzY0prdJvjMQtvRIvfROj2zsVCx03gIs13Zi5EMPTpdMRw9Me6tNX2rNBHbQBzLnajBJHTTA8ixXfrLRNrRJPLa+s+FXGscdXPUNV/BAJ"
    s = s & "CzVp8JeD1sWGjLIfLRev2ztCDbtELW76Oqw+rNSWzNRlDRNxzV1sS9fFMcV37bKplj884G6oNs8NK2VgPcxwEdJkjaZoXNIhOU24kLmVLD6XPNng/NYqIYJYZNm6YdeZXT95fQHse3naumOgXRAOqxKjHclcYdqOTNvEh9wAsMYf09hT0daHzNwnIds0Ic0eI2nGgdm7"
    s = s & "/41cm504zvaaGpZ3HQ0Ax33WbVHNbbvckV1nyC0Cn9XJ9UzQvGzdB4HdiRrTaacbxvzdqHkRvz1vEpeB7ntUHmrbYV3abdTeJ4zfXIjcr8sKWUHdkg3hBaHfMtHK0LPWBjfTAB5MMSzBDTNOGgoxXGxzXuwSgx0X1bzDHYjhkxjJMHCVH7OyVGHhL6HhOy7jGY5F60oc"
    s = s & "Rxri7SXgGYhfk0d5B25dK94SffTKbRHOmnrYcUPfNfnB0B03rX0VOq7JPh7HXz4QYJzGIE7kbZSdJE4HGHbiD5Pi+dXk9zvXcyHlBCPdMf7ekKzeWV6cXG7fs43nLcHjJIdFrasbQ27mhP+zzj3GzA1je1mcP4T85oH94GZk52vB4bbs3qitpWddZtic437u1oDu0mEu"
    s = s & "ECxqnLmdU3Ow6qze6q7+6rAe67Le6kRwEadciBXh1zvj5gV13ph8qHRBnB5ezqMucIwqAkB6W94c6tVd7HBd6sEp5zK9GlPc26zB6HTg6A/h2VR3e3Ae4WbUyLNaqad9xpz+Yiul3QWD41DR5YEe5oL+EoVdWaCsGjRdMK8ggdAxBfnDBdo3ERywBfnjBP737fbZRnjc"
    s = s & "FiHQRrtQvJpu7sxbP7awCjZuQLvgddPN7Bfu7CkR7y4x71+D6rfx3wWTnatx6xDjkb6dgTeIR6EdEwa8Oq7/HVc6vN8c39yIPsdS4e6kfvPXDe0gn8jEvBrstTqzIB3Yzo4TweZdreKT7hIKjkWVyxWXi4VUrsQkXT+zQOcGA+M7wfNODe9AL+0fvhqHXjCfEB00wPIZ"
    s = s & "keQm4/T0HBPq7jEiGhdB/zV1X+60nPW2vFR9xO49AfbP7vP5PfYzJPKrce8F05fEwfQN48x9yklw//IxMeYzZL9rcbeEQ8BxDvEJnekA4KgmOfM1IfgdL/aE7z6EXu+bcfYFU+vGsdf+oxEov9OSHvcwoflxA+VcQecO/vB7n9rlaZltxPs5Yfoo4fFhn/p5RNqsgVM5"
    s = s & "pe/FkeaQXxFwkD8cevuUDxNV/29AYuwWomnpd77pEQ/6AmHB5M8TyP/zzE8Qyo+1QM76moHZ1p4aA+4wo6gRjt8wKt8ykRv1AMFL4ECCBQvOApBQ4UKGDR0+hPhQhEGKFQd2iZhxlkWOAqFkBPmQUkeOgRySItmRVEiWDlNahNHyIcqXBj/KnFnTJs6IgXQafMPwwlCi"
    s = s & "RY0eRZpUadE5P20shRpValQqdKxexZqVxNShRrJ+xbqE61iyR0EGMskTQKufFEWohRsSSluDb+FupMvrZlyReXmlZTjXr0A7fDMOjsmXJt29iv02NjzysdCylaHS0mnJ8mayJMB+pkKWxWewbT5wRp30LGCcigZ7NByb4f+d1yvj4mUsu6FkuqwX8s6Ly67uhYgNL24L"
    s = s & "GS7yn8rjAm8bhHJqzlZ+WqGeHWlV0lrJeuj+9Yj27KvhCh58h7hs14Pd8MWdfL1C6Dp9K3yR63Wi+QqNO87LOZ6Y00lAtTwZrCHyLDtFJ1MWXNCz8K7iojIuJrzKNAg5Mw8u/QaLpD8A/XoBvsn6q6+m+xTy6TUXRPwvLgJrMlCmGV+qEacbU8JFwQ25uuGnOX7MTgkM"
    s = s & "reKhstGOpGM8IsvqUK32/EJIRLheC9HEAEVM8aUVExLhQ7/e6y/G5U48Ds3YdiTJtoWelKoSnWiBEzUS0jgyjc0uPFLDOqeKkif0/BrOypb/XHgtR5Dia45Lv75MqEW/QoHRr8RkVBPTLXVjVCf+pvvzqCB1GjLUynZgMknLZmCSDiBMjSpQnjr9KShDZXpisCq1zA3F"
    s = s & "RzPCZbBdCtXNTLXY7EjRkJDlSFmW0vMR1qIEmTMFacnCE8M0TtuMDSb1vFa1jNDiS9K8sryVJTfci41WHB3NC9KEItmvTEvT3BTfXmObyC/pQL0Ws5o0C3cqVI8MjbMjWq2hYLPGlRekMAfrMV2WAPGrYsPcTclZnLpMKeJ1E53P2AEzPTNfw4RIMNpriXDKYamyxXAr"
    s = s & "zjzw9sgKZSZKVp6mVNliiOjNSxPZOCbJY5lAJiliRGsr+d4R/+VbE+Xz/HLzzYL3yIxnqA7GMGHUFmaSBa8v8BmnF16LWGgAVvFLvXatjo3pjtru5DVeCiPOZB3p/jvouGjLS26AYc1EpyLOVormCc1OTUKdz077Y13dBtavHo4GnC+7SwJp7deE61tqTfedulH2rHYY"
    s = s & "6YIeZBwpGphsQzs+jxzB68plym8wpUVkOa9dN+78uV/P0vtT2fy2sXiWmLXod4gQpGt4ra8d9aVSYzfq9gmx0G7JI3XQHWLZzG1LecwZIpquf+cWXLbPLWobgIlJXt70lFE/nerY2KILug4XKjm9hE7cM0oLWvUD8ngvPE0oX0/qJzHXcYR062NIBSsiQP/DAJB/uplf"
    s = s & "RSaIPrpkjS9iasul4OJB/6WuQLLp1wcTEq411AQXTEBgUTwghlaVgDwngMOR1NABnu0OJznQYEUegcGF1CEvjggBcaBXEemFJIQUmWAGHqE3TGQAfnRRoVqCFb+TkREnZ8iLCa8HK4GlhGA5HAp3EAYhsmHISQ4zotp24TwRwSBe87niTnz1x5ZMkSPqW6H+eBI8Mzav"
    s = s & "kS2hXltK9JBrNaUmT4HjBSR3JJuRB2d94lbB8oiT9rVFjUIrZU1YMUgZ1u14LBmUv/SVQr6EwXeyMaQgiWfKiARsYJmMY6t2Rkcm3TFco5RJFgbzovXBIIkEmcUk10NCGsH/qzeBG8wpl6ZIUt6yao8MCSNVR0lYWfIli8vkJjGkKgipszt+Oqb51pPLgpAJcxj7SS58"
    s = s & "ICLCtTIyr2SJ6F5jTylxUybPpEgVG0JPgih0ISP7iTYVIq0GvQR2mayjtp7kwO4YU1rIlIkIEEoQo2HOmW0Bg5Ug2sKQ5MGKAGUJ0IJDLJn0k5ZX8+Ysx9m/avbSVNlLyfZyeKdWsXND4ptQ7UQpz/W8YKQCGRbmUpkSw/VHnDsFybpmkdKIBLIgE1zI/QaDSJnYEpxD"
    s = s & "e42tdPpCvvixLRx0CKwKmJIDZhJs2grljzhKGqPCCqRqeyovyGolt9rnVjFkqUYEAleGeJUg/2BdCDWjAxeBtkWtOCls3HDJx4istCbShMhPSQXMCzhuQmIjElLDM8yPMnU+TvXnekZ6WSsFlhe0jQgeBgJah0SyLZBViEj1JlGQoFAnVZVJ3nhhW8Fu9qwReWr9TDVX"
    s = s & "ktQVjnet2Z9yhiHItVaChoLtT2ZB0/ng8yW5wK01f4JciMDgQ8RNCEMHwtiQeDYvBHXkW9UyskjIVyDwfR5nH2LfNrHEVG0kyRvhaNrwoPZJGV1tPL8L3qdy4lax7Mgs9ik0AqcEwPEdiKKMq5OSHktvF8SJTXWisZa4VyAv8C8vWMxTtt5FvLwN7Z/MmRJMwnEFraJD"
    s = s & "d+H0ySMJOVR/hf+LCBwLCMI+kxQ4NhRiV0xehpiXvgvJbFtmLBMM08XCPKksVjNy0r8AIMZ6+WZsM2Lemmw4JKFCcEcUnMMftApcoYJwd1h7ZNcaSrIc4dt6hPuSwabLv4VWyCIIAmWGdLgmYczvYAIdaZ2skiVkrtKZEV1IATMky1SViY510mM4SqGosAJP2bwLEXIJ"
    s = s & "Lbw1wa9hRMCKl8zCoXzp8kuY7BAo0Hog9WOuQIC7kFz/BBfva0mxSbJphUDhQ9EEccaozOnnNmSkwP1TRVNy0Uyi4VvX2utnwOfXPqfrzxXZNb98vexpW0y51bsDFGAAgyDcgTmWzohuR9fuiISZSpAGSYz/0+0QEZRyL2durgtrghGeQGK9PCFtxCWOQCTL5gW+XXZs"
    s = s & "YLBujpDi1jAcsd4EAu2MOPonrOD3Q66qK4BHxAe6ckMYRQCGWBDksgjnxcDLWL2WS8ThNUGvWiY+dKIvdcIYdMMzZ9FzkEChgriINRMTomy/5ILRCvHBVGGOk1mLfCCziHpEVEyXXMiCFLIYMW1xzotVKPTMtoYlx0lCcogX3e53f1LF15P0mgCC6QRXNElykQWpQ2Ts"
    s = s & "r9GnQ8LghkBEQhZeN0guOBGIO4QdAGG4QyRCrretyuTculLO2kc+eTe4OeB6iwTToaB1jqQi5TnGe+xlnxq9zwcKIZyF30Mi/wJAbJ4gpLB84RNyeOGZfiHBpgtEkE+SKz/k8yWEsui/Sm26JAI3rFgEGIIw7xy4YRFy9zAz4TJ78pefLLXvzwvccIuXcMINyFYIDNzA"
    s = s & "rFwk4uPpejXWrr78nygf8hsUFP4TCESACOl7LOpLjhdwLI4Au9gwvweEQHE5OuFTiBCwN6AjBVKIhEgYBd8biFyQBCh4PQpkEdvKBfbKoFyYhRXMwAyUhBdMBLQIBDegwRq0QcaTwRdsQVJohRXUD+VTwVlohRbsBBiUwRu0wTtAi0SQhE5oQR9svp5APlLgqpxgwFUg"
    s = s & "hReUhBiUwSXUQR7MBbCKsZuAgnfDmvuboQhUw/8IRL/DgoJEIAUPBLpOiDcSjIv504kwHEE75MN+4zvxCgRFWYxZ2EI3kLc9VLh3wQ9KeCpSuANEBIk1lETza0O38YEwCARJAMO5y8I7CIO/68OQUD9KaIWQm4VOCMRQVMXY8IJOYCGDmAVSSISYawnKE8HCG0OH8IE7"
    s = s & "6AT2owgVpMNb7I9JJEbZW0VBEwEXUEZlVAER8KJjFJEMSEYXaMZnhMZr5AtpXEZqhERVzMWIEAEVWMZqFJpiNEe7w8Z0VMd1ZMd2xJxvPMZzlMeJc8d6tMd7xMd8VAh4XMV59Edg0seAFMiBJEgM4kdV/MeEpLiCZMiGdMiHREAx60OFpEj/yoHIi8TIjCTIgwzFivRI"
    s = s & "PNLIkBTJkcRGjpzIj0RJciPJlWTJllwfk+TDlJTJP3HJmrTJm9QNmLTDmeRJIsHJnwTKoIwInSTBnjTKBRHKpFTKnyRKCjzKp6SOpZTKqSTJphQ+qMTKzaDKreTKh3xFnUivmMzKsRyLrjTLs9RH5kJBsSTLtoQKtITLuFxHOeSIKLxKt8RLCZTLveRLEpQyY7vGvBRM"
    s = s & "o+jLwjRMDKK6pIHGwWRMtDnMx4RMfoKaeGxMwYzMy8TMXZK0Y6yAysTLCcjM0BTN+vK6SeNDCvBMvBzN1WTNZqPLmki9UHyA1GzLB2jN28xM3vu/gXA/ULQY/wagTbJkANwkzr6Uv8YDv93kBVbgBEMwRKkrgOAcywIozuqMS5NTToKwSyvpTOmESusEz7NMOiHMwi1E"
    s = s & "i8pDwvRUzy7Uwieche0UEeD0zqNsgPC0z/s8xgSYz6NMAPz0z/+kwO7cz5kE0AI1ULeRzwFNSds80AZ10PUIAA1Q0JSszwe10AuFiwSd0IokAAz10A8FiQjd0IrEABA10RNliNkcUYVUABR10RO1gBX9xxJ90Rr10AKQUBk9xw2gThv1UQtVUR0txv780SJ1UAUQUmJs"
    s = s & "USNl0gOFgCRdQwho0ik1UChVQyrFUgB9UislPynN0i/FTyTl0thbADA10/t8gGAcHdOh2wAiPdM3tc4CwIA1nTgL6FE4xdPqVIA5pVM4atM8BdQ43dM+jZ0NUAABCNRErc4CGFRCdZgKOFRFlVTrHIADYIAHiIAIuADUdFTyQM0HYAAFOIACGIBJNVW+DAgAOw=="
    B64_SPLASH2 = s
End Function

' KARTICA -- 107x35, 2013 B GIF-a, 2684 znakova Base64
Private Function B64_KARTICA() As String
    Dim s As String
    s = s & "R0lGODlhawAjAPcAAGt0YYiOfmZhK2hjLFJdSD9FHoySg0NHIDtIMTVYIHNqL52ilJOZip2hk56jlaernpWDOqaPQEdLIYt8N5KBOUNPOTBNHSw3GIR3NUhMIU5QI9K4bUV0KD9qJYV4NYp7N/Dt5/Xx6E6ELFCHLVlYJ3twMn9zMz9MNkNOOdS8d+XWr1ydNN3KlvLt"
    s = s & "3smrUcqsVCI0FiYzFtnEiLGYQ7SaREZ1KEh5KWljLGxmLTtCHT5EHu/n0x4tFMamStvHjdvIj+jdvenevy85GTI7GjQ9GzY+G/Dp1vDp2ObYs/Tv4/Tv5e3kze7lz0yAKypDGitGG86zYs+0Zc+0ZtC1aPXy69/Nmt/NnODQoCM3FjhcITlfItfAf9jBgrmdRryfR1WR"
    s = s & "MFeUMdK5cNO6c8yuWc6xX86yYda+e9a/feTVrOTWrlqaM93Kk+fbuFpYJ11aKN7Ll97MmdnEidrGjB8vFGNfKmVhK3BoLnJqLy5KHC9MHEJwJ0RyJ5uHPJ2IPeziyOzjykp8Kkp+Kyg/GSlBGcmpTsmqTzJSHjNUH72gR76gR1eVMViWMcGiSMKkScSlSVKLLlONL1OO"
    s = s & "L1SPMM2wXc2xXyMwFerfwergw+rgxOvhxdzIkdzJk/Dq2fHr2/Hr3PLs3UFtJkJvJuHRpOLSptS7dNS8duXXseXYsl2fNF2gNPPt4PPu4e/o1O/o1cioS8ioTObYtObZtUyBLE2DLPbz7Pbz7eDOnuDPnyQ4FyU6F9fAgNfBgdjCg9jChMyvWsywXFubM1udM93KlN3L"
    s = s & "lufbuejcuyEvFSAvFsSlSsanSiMxFSUyFsutVsutV7K0qbK1qeHQouHRojZaIDdaIDRVHzRWH0RIIEVJIFBRJFFSJIB0M4F0NCo1GCo2GKWOP6WPQD5nJD5oJa2VQq6VQtG3bNG3bUlUP0lVQCY9GCc+GDphIjthI9nDhtnEh7CXQ7GXQ8uuV8yuWNW9edW9eiAxFSEy"
    s = s & "FePUquPVq1mYMlqZM+batufatvXx6fXy6eLTp+LTqCwAAAAAawAjAAAI/wBr1XJm7hiPgwgTKlyY8Ni5BwIjSpxIsaLFixgzYgzAsKNHhgCoaBxJsqRJgc8+qlzZ4KTLly7PrZzZsQLMmzgvGqTJE+GxnECD9hx6MKjRm0SHHl16cqaxd16IPMk3S+EcSWDoLazVgqnX"
    s = s & "jDNJuHKFgQekVE8SUku1h2EtI1/jVpzpbSyyIdNSAUr4KJUgt3CPrhLlD9qmUXIjrpQwNpGrEvQW+T14KFWTjm8tJskUKxamVSU7NRvbK7HAlRRcRajjClEMcKlq8JjzJZUFzIEl7tDFDI0qI2sIkbpkUUnuWmdIm66lskgyV20q0XA1AEuwFbkSpIrkMbNEU4VcBf+R"
    s = s & "aG+ssIk+NhSCJpGLctMqtbmiYYyHG1fv5nBIFepslu7HmfKKK7pQFMZYVUgUxVjsReSeK6XB59Eyjt2BkDuuYONEKqikokY9AEbESXiuAEGRKWMRckREC7rSoEAPRpjYR3S44ogQCAngCgQ8jJBKKh185F0tmqSoBEWqDOjKefdAE89Y8kADTVcx1sKJFVdAc4UtnSAB"
    s = s & "zS1vcMFJLfpoOcwucC1xxRVwtBOEJz9MMcUbRzLn0TiufJCQMl0kw400P+IiZGC0wDOWFBaVMRYULDLY3nurpOBKKTvUQkscY7lSKS1yZMqEQMRkWsU/O9wyFjn82MlQNq4ko4NCOPD/wYMNP2oxqECtZDqPRZKOVWeLL9ZSpUBW/EBLRMVkWmktQXgqEC1K+hDRFGOJ"
    s = s & "oupCfbgizjXcdqtBBuh0mAo+IOImkCWZ8mLRLpmOWQuwj0IoUBC3TNTsWMtm4qxAJJoS0SZjTXFtQtY8l+nB3/CwRyoiqJGKOiHWAkSmMk70oCsrvuuog8oV40sS9iorkL5jfcrvWP4K9M9Yr1DRUbYz2CHzzHYcgMsKqeQBSiqKzGFuLSS7soVFZmT668YwjlVKNIS4"
    s = s & "wk7I+I68by39RoTiWKowVMBYN3QEmyQ8pPNjNT93kmkYFpEz1iSNuhivGPJkmolE92oqdckRVS0QEplm/72QB64w0g1DkaWyzkGApDLCz7VMMhY8Fr0wlgxtBxvjKpK7Us6x84oM9NR616JCii4rdEEjrmzTURapAAPDQYP8iAdgEVUxdURMZLpE5fGWJkqmaERUd76g"
    s = s & "oxzRFQEPzIMCrvRQREe16ZFQIKnIQrtASZBhHkVvjCWHRNS6zbG8IVAyVjOg1XKJ5xPjfbIrKdfSy1hWDMynKxN0tFYq6SSEx49O2MpxjGA+F2RMIC2QXApSFZG4uaJerZibLt4Di0x9rxas0J1AijSWP+TNeLXoxxhcgb6B5cAEJiDC6kJxOIWEIxSGEOBEWjBBKbir"
    s = s & "BeqpgkgkUkFXQOEKvegHK8J+MZYyjEkJzGBZLASiNqFZwgq/G4sP6kQiONSCCkVixtyUlxQhfaIirYCGDKIhCh/YQxUWYYMuSjGMJCQBDabojCnuUQs/pAEW+0CCCgSyiirIQx4sOMIO3mCPU5yCOFSjHzR0YYZbdOU0XaTJcoASuolEUpKTxAmJ9lgRn13yI3PI5E08"
    s = s & "kakt7HAiJ/jkRxAgypfs4AftkIEM2PGDLUrEAar0iAFaKRcA5HIhBAABL+NChQVUYCefPAYKGCDMYXolIAA7"
    B64_KARTICA = s
End Function

' KARTICA2 -- 215x70, 3530 B GIF-a, 4708 znakova Base64
Private Function B64_KARTICA2() As String
    Dim s As String
    s = s & "R0lGODlh1wBGAPcAAMyvWsyvXKmsn+HRo+LSpdLSyWNfKmRgK0yAK2RgKlVfS2FdKa6xpUdLIUdKITdbIW11YzpHMElMIjRVH0JHH6ernjI7GitGGy06IzZDLEJOOH6FdIWMfFmYMp+KPSc1HSY8GGNsWe7m0tfBgebYtNvHj+newN/Nm/Pu4qiQQCQ4F1lXJ11bKB4t"
    s = s & "FO/o1fDp2NjCg+jbu+ndvtnEiNrGi1FSJFRUJd3Klt7LmMqrUixGGy5KHC9MHDFPHUVzKOTj2+nn4NC1aXNqL3huMTtiI31yM4F0NLyfR7+hRyEzFVCILVaSMVeWMWRtWmlyXyI1Fry+s8DCty03GC85GTldITpgIktNIuvix+zjycPEucbHvfTv5fXw5/Xy6vbz7Y19"
    s = s & "OJGAOUZSPEpVQMywXM2xXs6yYc6zY2ljLG1mLbGXQ+LSpuLTqOrgwuvhxcKkScWmSiM3FyU0G8qsVMutVk9aRVFcSJugkp6ilDU+Gzg/HPHq2vHr3Cg0Fys2GODf1+Pi2tG3a9G3bW9oLnJqLz9oJUBqJXpvMXxxMricRrqeRkBrJUFtJkJvJkNxJ4R3NYd5NjhAHDtC"
    s = s & "HVKLLlONL1SPL1WRMLWaRbecRePTqePUq9S7ddS8d9W9eda+e1yeNF2gNCY9GCg/GeTVrOTWruXXr+XYsu3lze7lz9a+fNbAfubZtufbuNzJkdzJk9/OneDPn8enSsioTPLs3vPt4KqSQa2UQtnDhdnEh8mpTsmqUEZ1KEd4Kc+0ZdC1ZzxkIz5nJCYyFk6ELVCHLUxP"
    s = s & "I09QJLKYRLWaRB8vFCIvFSAwFSUyFiMyGSAyFSIxGMuuWMyuWeDPoOHQoezjy+3kzOro4ezp45CWh5KXidTUy9XVzD1KND9MNjVYIDZZIDJSHjNUH3uDcn2Ec8vLwczMw6GLPqKNP0l7Kkp8KqSOP6WPQJSCOpaDOik3Hys5IWZhK2hiLIl6Not8Nz1EHj9FHk2CLE2D"
    s = s & "LJmGO5uHPFqbM1ucMylBGSpCGtK5b9K5cNO7c9S7dCwAAAAA1wBGAAAI/wC9CBQI5I4CDC0SKlzIsKHDhxAdLtMALtvAixgzatzIsaPHjyBDihxJ8qKAdhFTqlypEpmTHyVjypxJs6bNgdUgsNzJs2cEcTeDCh1KNKTOnkiTQvxQoKjTp1BnVlBKtarCOlGzat16kRpC"
    s = s & "q2CTMuBKtuzQa2HT9hRjtq3bmHTUymX5563duxy/zt37cBzev4D5CnYYBbBhu4MTKyx8uHFZxYoZO56cFXJiyZQzF7U8GLPmzzfluitG2kZCQktSK1L5K/WSCSsZc9kCuvZMueli6U6XkMen356eRHyS7/cnFbEF6hFhuzlJtVZ0S3fQ4pgk41UiUjHugyXj5c7Dg/9U"
    s = s & "W066bngJvRmvFJGScRDelTMXT39j2hrmdSORktD9b24PTWDcOTt9Nx9lKGzRhRddbLHHNKtgUl9GadGSn26GJESEcUsc41AlxvVQoHyaoXKhbmRMiFFYEpwYCxJ9tKBCcb/BxtA363kYnxfgaabHJhemqOJAYX1hHjnmvZNQI8YF0xACxnXDk4EjyUIKDZsEwQsgm9CQ"
    s = s & "yQtPTRPkkERapQwS5kGSiHQeJASKcZ+EstCbv3WQzJQkghQDP7qVkYkILsRwi3SAqFJUF2OS6QVYg5gXTwsJSAcLBQn5MOBC5mDXE5UerRKEdCcseNE0vEhXRjS0icRFDKtwNEZ+QpL/aRUwaOr2Bh4JWajbPQmpAOc+CYViXAc6jsjjgRpxAYN50Gy0RanSBSFLRyjI"
    s = s & "QIIrqOQSCysckQGrootWtYB56yg0jKTyJJTpb40saRwjSHGqEQqamEdDR6ZoK90zWHCExYXcbuStebEOaRU65pmmUG66FZFQD8bh80Qy+BgH7KZ5ZrQFIOYBICpHr+TnzB4b/ZtfwBoNLF3BKlbVgHlpILMQC9IlAkx1/n2yiCLG0ZOUvBdx0U9+rnzUxavm9YKCRiab"
    s = s & "h3JGKqMIblVISscCQ8hYIh0aCeEInCfG6fBzxhcNcKELIOFw4QhMA9ztt4pSBYl5iMjMEBrSGSPzMR3A//nbJEoBLVAXAeRnRkhsnIg2Rk1L9zRGUcfC8oRUHWLeAQ4Bc4l0mLfAs99EBE62QGbnh0NIXcjhNoOntF7KhTC0LvtFkResxwt6yILC7ijIIssLXPC4B+97"
    s = s & "uNAGKTd8LMvwuxffRik4pCqQC7r3rscpbGQioUCnEGDLJpzQQELwKyZ1zBHSIVIsQ0KkLzOdcMIh+rEYdTHHhW2INMKFAAiEAi8ALMOF5gDAApLPC7W7yAxw4aJcnMIL+3PRAUvgoljM4iIAcJE/GLSsC43BBOVDyhnMMwSITMEN0lmBuuDUrvn1aCCJu9ABP7KGE0mD"
    s = s & "catLGdwG0gUSOMM8/FiF8v9ScaIZzqKD+bngRUQgwPz4Qw/8wAEb2qC2/IziIklRBiLSFJEiSGcBCakCnKhAFcGdIFEhucKJinaRxunmcbTb4UVWIR02YgQF+jLPDBmkuiRmJAYX8gcNwDQQUuQHFwdKiiDMA4aUSAF9llBGC5LABDh1gBkuRJYXOHEhfoxkFifqBA5P"
    s = s & "9jaCacQFsZhDDDgytPzscZMXUuJFtsC/VmHkU+YJgqiQciZJ5UElXkxAQh7gt08AaGz0u4gZLoQKkuRROodrYw6hJseB3CAWbOhIBPWYERrEUiPPjIUzNOKKC2UzXDyxgXnKsRIKIEKSLYCS3+qRSYz00TwwIEkG8yP/h1E6rZQry4g0cmEobcowIxT0Y0buqZsN/vFC"
    s = s & "rRAIUuxhHnUY4KIYzWhGFaaDYv7mYhhL5kBONAOSRE43/nQcQKV2kSvwAoQe2aZ0XplQ88jyIgyNhUMxgsr8eBKdLMkDLCpYQUvoiB7D6ttvnBQvsnGBpCRpYn5S+saVSu4ibcCFHQ3qSoR+c6FOTNaFciBRnoCBqBXkWgs6ahwqbMg4F2iqSBkE1ZGclKzSJKXAdsiG"
    s = s & "+5nhlRmRqW5o+lV7hlUjOY3FtHgyN7Se6A38acEuhlUdpX5CF3J9oUAYmJ98juR+1XTjtqwqpBjq5gYfEWwsCKtQw5pnp5C70AN5MgTz/zhiCrjNrW51GyMZVew3i0gII4yTD+Tgaa5eQJp5OEGSEzE3r//cq3nmEA0YhDN/XOUmRmoqnZsOJKewVeaFpgHUlBxji9KZ"
    s = s & "R08sZRw5tUAfLAypZr3QSvMAYiR6ONEJqDpa6UonByMIZyzGANiBqJa1NkXsYTMCLfOQdycrME8KeqICsP1GCQtRgnE8YdwdzXdQ+RnDSNpwIhnwF44DSaAIcnAhUWR3pl5tLU4XjJEGS4c5LDlGLcxTg57oAk48WMgO4mus+WLiREsLCQEuJIc9ihbFAkmgF0IGK+kF"
    s = s & "9qDbLeyMX7sRqUpnaSwhxpF6Ar9PAI4h13lPkTXZ0/zAFP8kqmXbiUkrkC0k9l4bOXCMEwxWLmsEtNLpZ3khclbpSKAnhYBTdhiyHeMQYs0Z8ceF8Gy0xL55IE+ms0BUcaJ+aUTPWZbxd2k8kKf6tKwpscAbpMMbniTDsp8QDkOSYOFPEMvDmvSCmPIzh491RLT3zUim"
    s = s & "/ctSgfCJ1AbG8kW4qxvvCgS8GnkBRFEdkUeYxwA98QWc4OUQQsDJF7jWCBKlY+KPVFE6c8i1F0wxzdiaciAyOFEpNGILZQ9kBloetZ8xYgJzUvsheFi1fvjAE+JsWNYNecJvbZ2E5CBXIHtgsXma6ZGj5eecGWmzSol91YtIOj85SPJFWGFvgQQi38//RrYXoBHiMkXE"
    s = s & "COYxQk+6AScERGRd/3H4fAcCSPPgYnEc4XR0NQJKvQrEFIREYDXZfaGtCoTK5rGyF2Rxomn1WTrhFYiJzOPifzfkhNJxwxR6Igw4fSMivmmSztUtkL7aV+QZ2cMPHVdggexTOpTeghlWKZC5SycAV+Y1yS5iWumYAiOpWGZ+PO1arMdd4rrpBxYjgjfpfKEnaa/T+hwy"
    s = s & "CTjtQCWCw8gplBuLfhQYSNK5okdITjDadAEGFPfCHpAs+hMFwtdd8LtuSjqQE9CA2bFA7dUbmhFOBjrpg2ZIJBrA/AaMnSeg2IH0d+DeiER/+qAA/egywoUT3PMfQFeO//FzUYLwcwQFig80IOYQhKWhAAvGz88m2qDELhTePJpAVg3zk4ponCAQqdAFwJcLooB8XpBT"
    s = s & "c4BdXoAC3MUJ5scZghF6GtEFmYBLsWALMYAFMVACLAYAN2B1IfECHJMfgWB1nSAHAFAGWwIIgBAEZvAMcrAJAiEDOfAMZLCCgMALYzAHtoARogB56LYGAlFTc6CCvEAGcqAHWyYdgXALZsAJgMBZA0YCGgGBfMEYKOBs1MIGouAKrcAKOBANqiACvjYSWEAAN9AK0YBx"
    s = s & "USEL18IKrTAAJiA9WHAFIMgR0PYCpIADYJgJjEchVjgXngEubQFtIxGIgkiId2GIIoGIcq8xiIpIFowYEo6oFpAYiVsBhLHwU41YiWFxiZgYFShwIbkAdx8RB54IFloQimZRThcyA2XYEWGQilbhB6zIFa/wcSfiDKhwOh/BAbRIFdpwi1xxBVcgArijB8qojCIwDWzA"
    s = s & "hhyBDZsXjCxhDcQoHhtAjT2RAdVwjeHRBRqgjTsxFt4YHtmwDeKYEstgB+VIH9UQAun4EBgABe04IVkQDmLADunYDBnQBBXQjfXYHAEBADs="
    B64_KARTICA2 = s
End Function

' MINI -- 83x27, 1716 B GIF-a, 2288 znakova Base64
Private Function B64_MINI() As String
    Dim s As String
    s = s & "R0lGODlhUwAbAPcAAJidj82wXV6hNa+yptzIkefat1FbR1FcR2tlLczMwt7d1TtIMkxOIzleIkpNIkBMNqqSQaCklqqtoFBRJFlYJ6GMPqmRQa2UQrWaRC9NHTRVH3VsMHxxMi47JDRBKio2GCg2HkdTPUtXQpSCOpmGO0BsJkp9K1eUMTVWH31yMyU6F2xlLXBoLlCI"
    s = s & "LlKML9rFi+fbuOjcvEFtJkNxJzQ9Gzc/HD5oJUBrJVONL1SQMEh6KUp9Kt3LljxCHSlCGtC2adG3bPDo1uDPoMqrU8utVuPUqUt/K0yCLPLs3CU7Fyc+GO7lz9rGjNvHjyAxFSIwGI19OI9+OF1bKF9cKURIIEZKITtjIzxkI7WaRbecRd7Nmt/Nm5uHPJ2JPSI1FiM3"
    s = s & "FnJpL3RrL7+hSMGjSMWmSsanSuHRpOLSpdzJkt3KlMioS8ioTezjy+3kzPDp2PHq2UZ2KEd4Kc6yYc+zY8+zZM+0ZfXy6u7m0u/n09jChNnDhtW+eta+fFiXMlmZMvHr2vHr3NO6c9S7dH9zM4F1NOnewOrfwdfAftfAgPPt4PPu4smqUMmqUevix+bYsubYsypDGitF"
    s = s & "G/bz7Pbz7eDQoeHRo+PUquTVrPLs3fLt3x8vFCEvFCIwFSEvF+7m0O7n0vXx6fTx6+vhx+ziyc2xX86yYF2fNF2gNObZtufatuDf1+Hg2L7Atb/AtTdbIThcIU5PI09QJC5LHC9MHb6hR7+iSFaSMFaTMSg0Fyk1Fzg/HDhAHERyJ0RzKOrfw+rgxNrFidrFiihAGSlB"
    s = s & "GdC1Z9C2aNnEidnFiejdvendvd/OnuDPntS8d9W9eMutV8uuWOLTp+LTqFubM1ucM1yeNFyfNDZZIDdaIB4tFB8tFdO6cdO6cundvunev9a/fNa/ffTv4/Tv5MmpTsmqT+3lzu7mz+vgxevhxSEzFSIzFcyuWcyvWsyvXMywXPTw5fTw5vXw5/Xx597MmN7MmT5EHj9F"
    s = s & "HuTWreTWruTWr+XXsOXXseXYsvXy7Pby6+/o1O/o1SwAAAAAUwAbAAAI/wAnJTDQCZvBgwgTJnwwYJLDhxAjSpxIsaJFhwlAKNzI0WA2CRdDihwp8UDHkwo98CPJsuXEgihjGlzlsqZLmTgV2Nw58iSnMA6E2VCC8JWGhApA8Vxa8eQUNRfS+dlxUNipowgVAGLKNeJJ"
    s = s & "CGrUxLpxCpJBOCc0Id26sxKiZny6OuzIQA2JMVxUUNOBTe8VhVojJmOC7xEwLW8mSnIoSZ2aOnIndRyhpgqhMrvgmEpSgtoXwGwngfqm5pxDcHPGOXroTou3Sg+JPY7MkQYZC9hqkBlklZefOBsDOzykRg9EQ2G7OQQUllLs2XI5plAjxSCUWh9MCBAQKThbVGF/Rf+U"
    s = s & "o4aUnXDlws5L9G6SbMjx3oECBW/+uyDhJNXPFEQ0qHduZAKKL9y4M5dCuYiRBScG9VAGGLII0AJHwmmjxhpKQcREWIaQ80hYxhSCh3uz4ZGHGtv8EkQgYTUCSiVheTOJG4KEtcwZTaixjieSKbSBGiwgNAIDrgjQhxPeibaGGgFIBKMaQkzCnBrOOfSeQ6KgsdgkH6rR"
    s = s & "iEMBqCHjJKmEtYVDLALRI0K40KIGIRvEuQEC2ORQjQANJPlGWJBFdElYaUjZ3HOQvQHbQwW06BB5Y8YQ1iMObRGWGwqtoAYWFmSaKSzWCHDFCbaotdYkf4RFh0SWqCcolYT6QwoiECX/6uWiYjrkqBqQTpKqGoYktAkGZeiSkCa3TOOFFQKgANok/YyjhjMSSarGPqtW"
    s = s & "SSI7zHCIqKKTMGrrow7dk1xCFKgRhULXCDADNl5I48Kyk2SrBiYRIaKGOJlUS6giRKhBhx0Oyfplt7VOcmuu0YS1REIVlEGPQjgIMIxBvQgwy6gGh4UPRJI4o8YLDiExqJWzCREWNA7BwG0pBR/sEA9qDGEHQvWU0YVCGQhwxEE+nGIExpPoocYP/Tz0YR2KOBRPWFF+"
    s = s & "Mkkds72zzrOJTOJLWOZMksgiauzxLa4OWUjAmthMwAEVCqEgg1kHWSEDOlmFJokyQxwTzyRsPKNH0g/tmaFGMfIYcg4jMbMxSTBhCQLOO+2ogcgjRZD3DDkZq7EFKPmoIUh7OMmkQHsQvYPMPQU8EtpDoDxSSRuTtHHHP3jw2AYe/iwBziSZWGLGKJN0w8YdI94KjT1m"
    s = s & "iHdg5yfpFBlLt1IbUTbIn6TK8iydERYTRUMUQvQcdRAK9SO5cU8++eCjT0SsPMG9QgCA31UrIsDE/QIRuM9VQAA7"
    B64_MINI = s
End Function

' MINI2 -- 166x54, 2817 B GIF-a, 3756 znakova Base64
Private Function B64_MINI2() As String
    Dim s As String
    s = s & "R0lGODlhpgA2APcAAMqrUmhjLGpkLFZgTGhiLEt/K2RfKlFbRzdELTZYIGJrWDtHMUtOIj9FH5ugkqmtoO/n04iPf0BrJTtCHTI/KO/n1Ojm3y04GEZ1KC45GSs5IfLt4FCILSc9GMDCtyYzFio2GCo4ICg/GSs2GCk3H4V3NVeUMero4e/t5kyBLCg0Fyk1Fyc2HURI"
    s = s & "IElMIqaPQKqSQTU+GzpBHY6UhZOYiVNTJergw39zM93Jk9O6cujdvZCAOZWDOq+WQyEzFSQyGiI1FiQ5F1FcSFVfSzNTH52JPaCLPjA6GuHg2M6zYsjJv7qeRk1PI+bYtFiXMnpvMX1yM0BMNkRQOtzIkdG3bejbu9S9eNW+e/Pu4vTv5SAxFdC1aNG2a9/NnODOn02D"
    s = s & "LFCHLeLg2OXj3Kuuoa2wpC9MHTBOHe7lz+7m0UdTPUpWQJeFO5qHPKyUQq6VQi1IGy5KHM+zZM+0ZsWmSsioS8ipTsmqUEd3KUh5KUl7Kkp9K/Xw5/Xx6fXy6/bz7byfR72gR8CiSMKkSeHRpOLSplyeNF2gNNa/fNa/flSPMFWSMCAuFVVVJVhXJuvhxezix+3ky+3k"
    s = s & "zdjDhdnEht3Kld7Ll9O7dNS8dunev+rfwfLs3PLs3iQxFSMxFyIwFSEwF1xZKF1aKHF5Z3J6aTFQHTJSHmtkLWxlLVKML1ONLzdaITdbITthIztiI0FtJkJvJipDGitFGsXGvMfIvm5nLnBoLnNqL3VsMIl6Not8NzxkIz1lJD5oJD9qJd7Nmt/Nm+DPoeHQolubM1uc"
    s = s & "M9fAgNfBgR4tFB8uFOziyezjytfCgtjCg9nEiNrFidrGi9vGjcqsVMutVcutV8uuWMyuWcyvWsywXM2wXuTWruTWr+XXsebYsoJ1NIR3Ne/o1fDo1vDp2PHq2fHr2vHr27KYRLSZROPTqePUquPVq+TVrTRVHzVXIDI7GjQ9G82xX86yYMvLwczMw7ebRbicRk9QJFFS"
    s = s & "JObatufat1mZMlqaM9vHjtvHj9K5b9O5cefauOfbuSwAAAAApgA2AAAI/wD9CPTzYAgJYwgTKlzIsKFDhZ2kRBAzsKLFixgzatzIsaPHjyCRCHlIsqRJhRQ8gFzJsqXLlxvDRDlJsybDTw5g6tzJc2efkTaDBlXSs6jRoxfHCF1aUwHSp1B1DmBKteQPFFGzau0YoqpX"
    s = s & "h/G2ih078KvZhSrJqs16tq2xtGvjHnV7Fq7cuzvpmrWLt29LquzkLZnHzpgEfPgkPOSFeBXJtOH8SgZJlRudy7aMwTLEOUhDIIUMFQLyWGCkyag5Ms0Q6DKdcp6MfeGMoaEuzq9Kpj2duvdFprRcXw5l7A1nQ54XmhCd/OFuvP4wOcLkz9wUYL4rLl3UQzgdN4uMcf/g"
    s = s & "rFjhKs54TD6/m+2Sd2bZywpt5Lr75XrGVHHGpyXhseWGmKGeaR1lUUUXzjDzTBf+YMFTFgAIB198fgh1DAyutdDGZTwY44MTnO2SUCucpXLSehj1oYMVdbyTDjgbQKIPHXX0k00fH/GBkRUSUlhhUEy4ZoQxNVw2RwzG7LKfD8ZokQ9n65xIIEbgbHFZEhtYhEkdl71z"
    s = s & "xkZYOIIDFRgt0yOFQuHiWg0I2XeDMUAMw1kCxiSw3zFS+sGbRd5Uc5kdFWDkhWvQQIIRJFxeFkeZZ8YXFGuXvZBQAJcFcoFhnHEgHmeu0ITiQGhI4xolGfFh5WUAYHJRMsItepGZrk3/6KhNT6yZ0AeAXGYKnKEZ4spxHXg6ZUWnXjaORjoIB8CxFbHqmqsWwXqZrNnZ"
    s = s & "BEJr3y10ymXgGYPBcZzpUdOnfvgi3CUcuSNcMxY5qyijsfpoEwGuCbAQJ0tcxqZx4MYy7rB+bJLoZdhwNKNrdWziRxZnnLGNcNc03HAWAklLx4R8VBDOBlh0vIE4DoLjMTiPpKODQHyAjMUG3thAyJcCjTPyI9s0IRA4vxxiyTM2WGQThnQEMgJDtkQaHhjHFWDTp4R4"
    s = s & "Bw5HTXiHXQVddNGMcNJUXTWzFk+4AQ78CFfHIcnwMYW6wlUikCbOTONdFQMNIod3y/ChzxZxCHeIgwLV/1SPa1A0dMEclzVgDBHHkbI0wO65Fk1HZ3jHT7OtwjutRc9cJgcaFVXgndoDQfB2RXz46doyzjwikDbCUYHjjye9cJkgGTgEBR1DGoPHcXksrmdF0LTeERbe"
    s = s & "1fG6H+7SAW1FXVuULDM6WpS3a6APNP1lcFckiXDuwCzQJMIR0vdJf1+Wy0MqBOKCMSKAawgs//4ukDjeHeKRHd6JM1Dyyw/UfEXPYBdGLJE2i+RAONkbyBTO1S7hvGN8JimCa+hBklMgpBfuS4+w5Ic870zCI24TDucEwj/LXawiVbBfRghIPYuwEHsWoYRwlmGRPgyM"
    s = s & "DhCAHUkaIJw/LOGHQAwiKP+aBCJD4OM4ItjgnjLxHo+YzjV7KuGrGgUMOVBshQWsyAvpkECByPB0F0mCcM6hw4eUwDtodM0SOGGM24hmM7hR4kCY2KiNPPEyI+zgs0w4oS7QKDJYbKEWERjDGV4kbK6ZQhkbsgJsLWEHkIykJHcwRCAc0RDpycN+SDMgDkLCO8XwSISE"
    s = s & "o4n9VW6KscLCF+kgCY1ssXoC2WIX/bBKOtDQIjwCo0lu4ZpanEQCxymDMcpwHF7kaU/EE04OOsIH7wCAcntE5WWC5x1kBPIysPSDLAsJRoscwpAlUUGu6CCII5gECMIoUUJSwRlhcLI0HPTDHa3REW94xwrQfJc0L3b/Bmq0ag8DzOJAtlmRWt6yIt+MV0kM4JoSnMSN"
    s = s & "hlBHQhDHGVZ0ck9+IIZ3srSRZAmnYKaMZrR6ZA/vBCOgghwoIQtqSIs0DpuLXIjsyNkOkxxDEZxRBJ4Sgoqc7tQ5AHuYcGy2ET+6RhpXJOEpRxovPyDyMtJQmAsFGsuVDsSgF6HCUGOaEAa4phsnOU+IFkJRQ9AJnhjlQwgvQwyOvLKBImXemTDhnbZONaVVdc0ssWqR"
    s = s & "O2aJJGswkgxOkghgLeQ/nEmEbgDmB3QIxw58w0gWBgYAqQ4kEkuVa1P9oFXhJMOlVHWqVb3Y0oFkQThbgCBDJkA4OrDhJKXgXUN4kTi0/14koZcZhEbM9adMXCQcmX3d//xwhhsusyLFoGof1spFbl7moALx6GXEx1Vj3MA1jDhJCo4DP4YEIWm2tUgWuOC4yFYkHImy"
    s = s & "g28vYkPXPDC61E2uaz54Ve9soyK/EI4zKrIN/LkmG861pUU0ehlrAJSrH8gXHcgRm5LE4jjicsgduAvUeJqWGa7hh3kXdkA6yOGzGXlpHTLBh0xsIUt9eIdwuFCRLIjKcXlEg3CqwdEmTKLDlwklS7sZXddMI496+Qq5KmKDDm/BexWwEjSAcTyMnMG/rgGAoZpwBTRO"
    s = s & "Yr3YGKXjeiaQfQhnGlaIwzL6sEU6OMOapPXxNjpGiETFIeaH8gkyVdISqI1UgBA3LoY+JnEFHPzjwBwJByX4EQcqNGOE2mjCP3SQiUzo4B9NOJkfEl0FTGQCE1VoAkbtoQwuxIEfPBNIFypxjiZUQQf+UF2an4uJSXiaH8y4R5OrK+c8+UgrfFVNrZnCl1s/Jdcb2TWv"
    s = s & "fa0VHJQ22MIWSq+JbRRnKNMjP0h2UGbBbKhsQMXhS2pG0iDtmiziBNVGyhSq8Y44mDsOSbhGNY6bERp0myYHCHdqUICAd5uEDPJODTw0YO+HiGLW+e6LLPjdb4UsYgYB940FZqAGFtj7EwsYRVgSPpmAAAA7"
    B64_MINI2 = s
End Function
