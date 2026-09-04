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

' SPLASH: 480x157, 5808 bajtova GIF-a
Public Const LOGO_SPLASH As String = "SPLASH"
' KARTICA: 300x98, 3326 bajtova GIF-a
Public Const LOGO_KARTICA As String = "KARTICA"
' MINI: 160x52, 1665 bajtova GIF-a
Public Const LOGO_MINI As String = "MINI"

' Boja na koju je slika pecena. Ploca iza slike se crta BAS ovim, pa se
' pravougaonik oko znaka ne vidi.
Public Const LOGO_BG_SPLASH As Long = &H12281A   ' RGB(26, 40, 18)
Public Const LOGO_BG_KARTICA As Long = &HEEF4F7   ' RGB(247, 244, 238)
Public Const LOGO_BG_MINI As Long = &HEEF4F7   ' RGB(247, 244, 238)

' Odnos stranica (sirina / visina). Okvir slike se racuna po njemu, pa Zoom
' nema sta da doda sa strane -- inace bi se oko znaka video pojas pozadine.
Public Const LOGO_ODNOS_SPLASH As Single = 3.0573
Public Const LOGO_ODNOS_KARTICA As Single = 3.0612
Public Const LOGO_ODNOS_MINI As Single = 3.0769

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

Private Function Base64Za(ByVal kljuc As String) As String
    Select Case kljuc
        Case LOGO_SPLASH: Base64Za = B64_SPLASH()
        Case LOGO_KARTICA: Base64Za = B64_KARTICA()
        Case LOGO_MINI: Base64Za = B64_MINI()
    End Select
End Function

'=====================================================================
' Slike. Svaka je svoja procedura -- VBA ima granicu velicine procedure, a
' jedan zajednicki blok bi je s vremenom probio.
'=====================================================================

' SPLASH -- 480x157, 5808 B GIF-a, 7744 znakova Base64
Private Function B64_SPLASH() As String
    Dim s As String
    s = s & "R0lGODlh4AGdAPUAAFqaM4J2NIR2NV2gNCYxFj5EHhooEryfRzJSHkJGHzRVH1GKLqWPPzA5GUVIIIZ4No9+OClBGTZYIDpgIkV1KEt/K06ELC1IGx4tFD5nJEFtJXpvMX9zM7+iSMenSpuHPJ6KPSU6F7KYQ7abRGljLGxlLFtYJ15bKKiRQK2UQjg+HDpAHCIwFSI0"
    s = s & "FiY8GC02GFJSJFZUJUtMIk9QI0Z2KEl7KS9MHTFPHXFoLnRrL2FcKWVfKlWRMFeVMZOBOZeEOiwAAAAA4AGdAAAG/0CDcEgsCjHIpHLJbDqf0Kh0Sq0mjdisdsvter/gsHhMLpvP6LR6Pba63/C4/M2u2+/4vH7P7+/ngIGCg1R+hoeIiYqLjGWEj5CRgI2UlZaXmJlF"
    s = s & "kpydnlGaoaKjpKVcn6ipnqasra6vfqqys4+wtre4uVq0vL1xusDBwqO+xcaFw8nKy33Hzs9LzNLT1I7Q18/V2tvcQ9jfxt3i48ng5r7k6eqw5+2z6/Dxoe70qvL3+Iv1+6v5/v96+AmMBLCgQTUDEw46yLAhGIUQAzmcSHFTxItwKmqciLGjlY0gD3ocKSWkyX8kUzo5"
    s = s & "yfKeypdKWspcB7PmzJvjasLEyXObzv+XPYNO+6lSqFFlRFMeXRosKUmmUHE5HRm16qupHq1qNSVwBoqvYFF8YLHkwoKzaNFeeCQhLdoKkrbKFSXQh4e7eO/mWMICwIC/gP9SINQicODBBI04wDG3cSJ+DQ7kzduhwRIbhgEDaDFoQua/N+IqZuy4dLN6LFBMnrx3SYXP"
    s = s & "AyYIatHjcw1OWBab3h2wHozVk0UwkQB7gaAMsF3gHs27uZ19D4BPdrCkhd/PoeeEuG4Y7vIiup2LT7NPsnS8G5h4/gyArJz1mSV0yk16vH0y9QSczyvCfRIWPMAmnxwLsMfZd0SEd9+CD7nTwH6TkcCEBrBZIAdxnw2I4BAKMuj/4Rb0kABhXiMQUBZsA0QAB22f8eCf"
    s = s & "aODV59wLDuyww4enuKPaiHidwIQFFcKBYWYafEIfVDvE4MAKDQxBQAMNrCCDCSU8sONdH+AIYjsP8ojXD0y4wJ1ha7nx2mdlzsccUyJ4uVqWWmbhjohuetDBC0ycmRliVbhQHCpHMvWBeXV6AGecRuhYqAc/vIjBkIYBoBwVLBT4mYpGrgnVCzC06eahiBLRjgyL3kWd"
    s = s & "En3BRkMVCMDmXaYxauVAnaCGeoQ5qe3XwXmtKQFfZpNKQQNssgGqaSIEqLCCA8w6UACTBFCDw6e2WgTOb+cdQKp0wjER4GcZUDFmYAfCmqCMfaiw/8MPu+53wAc4NLmMpyPWaus5EOzHAQYMnCfhEsi16GgTw35mYSqBBrSYDyNM1gEKP/jgwwcNryaCDyVQJ0y+PNob"
    s = s & "6jmEAicDBnQCd4CjF6AYLhQRoFissbHqoUPIeIlgAgZGrMCxdDcCM23H1YoKTnTnoUBWA+0C56NrKIYABYXsOY3wsWw0gAO9JJbAwhYrcJD0ZAdsoPEtOnjpMaLgzJqtiUjMUDQTLNT2WZFOiCmgLAmvEcMH55UABgH6Fb3DC4iwMMMDDpThNtBB34oNB/s9sAQI5428"
    s = s & "BAWw8fAEkJ8ZhzfVaNglXQeJi8H3fijIywcGO3ha+hiL19u449C8UP+xdDAsccJ5IrCdRMqwrbxEy8G/A3oZL+wMXM9iNHB6tn7bQcCyOuTwQ8ivixE7hGfH+U0O+3WLKs155b6E3J8FmwTUlxofMxoFQIiuGP1CGD0aH4iAtXTZh7H9ft3TEjYis58ZNOF/eZHc5VxW"
    s = s & "HfQZ5jbuO1caZrC/yQTQCwn4GnA6kIOtnaGC/FOc2WZnAGyUDDgoeMKVHIYnJQCvc0soWPpokTcymGBEBUgDtiC0ATSAEDj9AwMCpXPBD2GDcv56wgknE4AfoSg7GIjAuADzqs+9jwwz0OBqcrAG5Z1nAx4kww9XE8QvDBE4RfTQNVQAIcs4oUvScePvUNQeJPz/yjAa"
    s = s & "sqIEy7At1LEhfjyaHxjGOB0RMq5x0MjVeRT4BCQCp1dJ4FyGkGCpzGyGFzUEgwoIeRfmqSEAPOrADMrASbyU0QtnfBMJoZFKDxxACruTDgOGgyK4EG+SmDxeFwhAPuAQrg69BE4Ox1BKUxlSdrODBuQWKQUCjNGASvATiiLAPsO4qBeZ9MIyeXgHUPLIB2L00im70Mq8"
    s = s & "pJFB0AjmXcwXhT6u5mQAQ5EFvkUkdOhyC3B0Fx7UxiNPfqGYHhgnF8qJl3Mu6Bkb2E8KBuYER64Gmv+hJ4oykyYa3jMLBBzRD/IA0A6oQAwAFegWCIqlVR7jBRCy3BSWmEAmYG6i/5nxHDYvioWE9jMPO3BTD8MQ0mNyz6TGYGleUlgFpEnnAHJEwgthChhMzfSKu7wd"
    s = s & "hIZ5BwyoczIl4qk4fQpAoBZjhav5VxXAJ519LUGSTK1iLqHKBaKN6AB7EB2PHqBVHolUCyQ1lFd7UQAt5iWpU2DBDz16Iqb+5WVP3aMXVuAmLupBqKNTXRd6isURimJ6MKjSL9GAgQbIQAcb8OcZpjeDEghgBVl4wQxy4AMUHEAyHRgBAyCwA6p6wRgZlY4P4BBL4DAy"
    s = s & "CS+F6TWLkc0seHE/m8UDY3VqhA3o77nPddMIoEtd/QnSAHntnglSgAIQfOADEvPBA8ZLXvJCILwfmP8tEQogFh9AoLwPOO8PPoACEWhQskVIAQgkBt8HSIy+KaAZXbWwAwaA973llS99R6DBIDrgeRBCwQ5w1gVj9BZ3cGBBMMWXBMwwFQHHKC4Wrsqf1ZE4LwwwgltL"
    s = s & "xaOdGiG7BGYxXlCQIBnjlwgyHnAWflaq7BFArtSqcDGA/E7fuSGVKkUCgGAqU+LSdAg59RI4+XDc/eBXBc3KsgOkup8daPnLNxYCjFM7pQ2AFUIdgJcMiECAGOjgAX61Ml5PIIATT0bHWFABlXwQ5/O8zgH1u0uf3znKLfWCjftZWhwc2tLCqswZIi4CQD0QgD7c0E3X"
    s = s & "NQJlYWfZLlxao9Hagjv/eRRmI0B4RHjeAj+3agB+/uCjBmiATUfUAdTuohcsYPRkiCoHyPaHCZVsX4ifbABAekm0d0Cpm0TwT1Zz+pBdWIGu8/KDuwohBnUqdREawOXIgUEGgy7kCuy7gRu/gMijG6cvjF1WQBDgqIDFQMBgCGlil81NMfCDnfGibSJsWnud9gK7JzNl"
    s = s & "MJxazl5YNYRS3VY3OYCXB7A1FjBwZuAwW0692KZ0ChAIshK8CWilaL3ZigVvesm2eqh45bzwb/8F3AsHH0EJKAwGyAKn35L2EsO38OkROWAD2uICAVLwTZoLjReE/O0cKsjOJEgTRXxyMsmNEGhS+wHdXWa5swEO/20v9PwuKP/CqBHuBaz7NgzLtWvEvcBjCJlg4rxQ"
    s = s & "uMPiHQeP32UEjqoUTC8ZDmLv+y6h5sOKR+TiLbRciC/vwr2xhHMt5LONYJj1wsPwbi+hYOdYMCoyrTULja9mt4N4vFmV0Cqm0k3qit1FoQwh+RFBQOt25ep5DCoEufbODJW3+hdKoHMx/L3QXhi8nxNFixcQsukdz0uSkaAnmKpvranPAqItbwjeB5kLhzdj4rn2pcCX"
    s = s & "wU2Nh3Lv6zoi4WjSSxc/uiwuXGRIeAoEYZpi8VDPoUxziFaGiDL1J7t1l3edCzZ1e2gAfmDwdYskBiq3GvaHBS1HC7mlNJEAOR3Acf8xZFh/MVyJVX9dgG1u8np+YID7MQKw53OyR0RgUAK78gPeZwYE+AUgeHZhcHAhFAZVBhyFV0KzwH67JgnxA3pKcEsWiFgWNXVE"
    s = s & "sHgtZgh5dRdwxX+xV1n/lwWLI4Bp0IJeN34G139dYHddpX6p8AP7AUmEwAI/gHwYUE2GxXfQp4FcYH1esoBskIR2MoIQYm1YMGZcYAINAwIreAZU6GlW+AUyCERi8ILvZHSzMH3S8QP9tYiM2IjjNXpKIH8wlUcRpIZb0HYj4oZrAIceIIf7QYcvtn1FEGUHkFxT6CXh"
    s = s & "JwSEeGemg4VcMHbD5w2qwAJVJ2OlQoZmaIEHk4ZC0CH/WdB68oOEq8eEJOiEm5cFGeQBKGCKp6h7VThXrdiE5+cljkU7nzBwtlgovOZ0E+VAIseLreaGwLgfmqgGcgchnhiLz3aMRkAAFQN8bNCHd/iHMOeKW5B75ceFnuB52VgnYpUEd2QYERBcmcED5aJHlqgF43ge"
    s = s & "5ThBw4h99khOomgAMkAoHlgH8shz9NgFgUhGYxBueUFzqqB5/aiNjsIiqhJFEyWEU0OEQ7CQ0tGQaACHSwiR0sh17HgEOKBB1bgGGakFq5hA0ViMYUB0XgJrOIgKwleSpMYEMvQZIIYBEmUYu4iQveiGmBiMfgCH6Wd4ETlQ25eVYIeRqFiAG8kF/x1ZSEOJQ7J4jUy5"
    s = s & "KPA3PBM1XPOGHUMYfVjAhjxSaR9YJymQjjO4jj8lKjspHRepBj+ZBUGJF5inBWmZF6A4BDUoiNbICfz4lhAyYfGEIhoChJbEUGrikkKgf9DoB0aYkwz4lSP1clrokT5Zli54llvwmKY0BkspHcCXCgnoASVQAL75m8AZnMI5nMFpZJSEIk2GAc2XGQpQiVe5gXUCAtVX"
    s = s & "J4epBdmHSpbVACdAmyPAjN8Hm8+Iamv5iWMAkw/Vlp2gg5NhnJ7gmYYhPEhQGLXknOHYBefIO4YglrwCmJQpmCH4AH/nATK5BYmJBYt5F42ZBbRpTGKgn8CRm57AS//78Y+okIuB4VRJUAN09Hz9IJoGgIg8Yggmd1PEOIclyJQjkIpcUKBGcKAekKBYsKABNQYOep6V"
    s = s & "CQmryGGpIIkDsCq0JE9WWZ+nUCjeeQeTKTL86Zr+mS0yak6ICZ5+WJpXeJNfUKOTAaGdYHZ3AYkVOlEcigEtMJVkYg/E1m37IXF7MG3nUaQ5R6WIBzQEYAK7uRr5NoBQOo9SCoiqiQVWmhdYKgkguhoUmApP9xlqpQSQsidk6qEGUIsp5QeT1pVe6abad0iPByEpZqfO"
    s = s & "GKXiGYN7agR9Wps3SghHCgKgKQkBGRhR6QRiqhmnWgvENqIjcgLNAJJ4cT82SZQ4OZj/QlAAfzegRsCiReCiMGpqn1oEoTqWoyoI2DgZdNcJwZYZB8kEdZkZUNSheIkFpwkhPZkH9ykde5gF1ymRXXebG8eHd6qReVqPlNpw4NkJ5gkmspCoeBQF7nkYLZmtRtCs51Gd"
    s = s & "eJCsHiCdzdau5MqOK2CrCIqum4qnnTqluvoFRzoZYdQJhKRoqCCfmTMFIWcY0woj+moEf0djfNCkNzipD/umOQmwM/qdC6uuDaunBOuYXlKTSQkJ6okXlTGvE3V6UGA3rgIzH1sE23oeHWB0d8ACdRJ21nmsofh/GGCUPCICbLqi6QqUsimzMZsFc4oXhykJynYePpgK"
    s = s & "yymQVKCh/yhSUR6bkFzwd2+nB99qMeRnosbIq0RAmpnIsiOiosQ6nuoIBm7iSZIQlDogC/cKGFUZBQqws+aitltgrqshsDjlJbUWt+Q5t1tYBIIlXSqKY1WrmFeroEwrBPgIIcklCbtpNKqwHSjSAx0LBWOrGV9KCJFWBIF6pnrghUdIuX3rf6gZa38ntWMgrESwt56a"
    s = s & "tUZQu8CRqeg5CBx4HoOKCqkKGFE3BYWaGYcLCbNbBI47Gci2BvuGpkmqlktqgr9YJ8AqvENAvA4rt99Gj5CQuecxS7Kwsar6BvQLGGgLq4zqJFALIQVnB0P7SMQUukRgh0ZAkiOCd2KAvqr4uTFKwP8u2rbLGwg7JB0U6gkeBhs9ICQTNb36G7R5KbnhmgbFJLIISMBD"
    s = s & "YMCgWicS/AX7pqIBvBrFWgQyGpkG0JpxxHmCcKRINYvRahjXWgUY+xk98KpzkL1F4Exewpd1wK+rAQNkUMMnqkq74Khvo7s5vHsObKzGWwQRW1I6HAhXpXSeUHp/EgfV+p4b8pzNg7Cb6wXmiRfAKgTbK6rji0YMm7tfsLV3oaKyeoDFe7Jc0L/ncWOPUMGrUSKy8LqA"
    s = s & "wZJVsGQoEruTQGxGYHOTMcdaoMT+awZxbMe7erlZoGF1kgBgoKXOur7eFsjs6wUvAJIe8wimLACyUL2ZgaFvYKGAAZ//skvJXAwhkooGfVqKZmC3u4uydFu+bvK/ANi5I7bFNBy6B9rCEywHL3BVEPUJLMDIf+GjcrBUBZm2bNwGuAsh/loGK0BiKYCUZACL/YkFXaPO"
    s = s & "KTyRQ6ACCHsX8OiyEAK+MeYl5TybTIsBpWQ0cCcIjquj7TlRaCgH9/sXlCgISJwFmvyoadDJNtyOX9kmO6fCO0akXnCp0iHNEF0nG6XKlbtYLbusbuDEdvK8n4DLf3GocEDLgWGQ2MvLWEAA43we8IcGDjBoKKDP4TQiOmC1BxCuzQshkAsifIwX3aoFdtbUWAA+S23C"
    s = s & "MCvIyLwfJTvNcGDJ24gKPDoAqwoICz0A/w09yfubyaWkzM0DIcKsBqZcUEuLeXqpULHpJh0wwkWQ09nSBUZFzKPDt4HZBSQmhcQXCLt5wZ3wlJ/RuhwMpB/MuG2Aw5OBq2GAAXo9GQygtGVgyTgLz7F2OkBtAHWMF3hNBPDrJSkwtUIw18VMBBVJOi7qAUONyq2dKI+Z"
    s = s & "2oY2BzdL2qog04EB03HAAjCVv0ds0/hEYiBdYbuZ1Wfg0XicIBWTab/6BW9rg3ztJcp7wLvCGLH9SrQd2Ptc0hgHCCeQA+Z93uhtsZ8QARnQ3u793hlA3BcC3/DdnLt81nz91lC9C5P5AWt2By56ABCwAQEQaGotBC9w2bpCyl2AvP9EiwNGawSsTUQzEHgY8HMN4wNb"
    s = s & "E9sesN8PzCP7MlBmOmOarY9YkRAPXYBXRdhbsJAHkNxsMNomE64vsNR58eKnYONCyQUKfh4jQDF5UXAcrozhR7JbINk1U9pafeL8kOJg8ALrwjsmgNcYYAJWDAIl8MZj8AKE3DGmGAA/gAL1nMjg9WI/MOI8cgAMjlFN+nlhNOR2kgI/UKegyyMM8DAlIAMNQABPIgM5"
    s = s & "sJsHICFfwOQX4eRiQAAzsAFWHLABsAMmMAMxsAMboGsj8AAnoOVmwDpdLh0pcDOa1gEHMF0i8BXeFTHhFV7zRV/P9VqdOIqgLuqk/gGmfurfxQApoD//ByBKXSADH1DPIlBuw7pBBzDq3fVdHwACKHDrB2B/NbwCAEorg1PZhA4Rhl4GBKADvV4nHSA2lGAChNQBEL4N"
    s = s & "DlBKNuO5HiACAWACK6DkVS23BPDsEBIA7D7Q0z4QmVfRZ8ACKiADJ1ACObAB5bUBOVACJuAAKhDh+rACMVACAP8AARBaMxBG3KACMUACDf8AHIADOqDn+FQA8w7Y7cwCBWACORAAj5gDJBADB5/p9Y7iJOQQUhwLLW/vL98QMX8aM9/kNc8QN786Oa/zO28QPf8HP78P"
    s = s & "QX8QQ98bRU8PRy/0KGwNS+8OTV8QSZ8HUc/0U/8PVY8HVy/1We8PSw3j1VXV9efw9f6A5qxhCGRf9mbvEm5Cs0S/9t/Q9vfg4O2s9HJ/DXQvD359HiPt83mv93u/Dg2g44wS2mwQ+Ngw+OKAAVLyZjKWAg+gAzNQAJieI4o/bIzPDWn3ljMsBpk/cpuvDec87ChgYLN+"
    s = s & "6qqPXsee7CIwAgfw+dIe+vQ3+rtB+/Zk+86B+xmo+6bB++Do+44B/EEq/MNP/Plq/LeP/Niq/LzB/Gvs/MsP/fct/eJB/Qth/feB/Wat/fbB/b/g/WoE/lUg/jhC/iVh/h+D/kig/uIRBAA7"
    B64_SPLASH = s
End Function

' KARTICA -- 300x98, 3326 B GIF-a, 4436 znakova Base64
Private Function B64_KARTICA() As String
    Dim s As String
    s = s & "R0lGODlhLAFiAPUAAM+0ZcuuWOfbusyvXM2wXc2xX6irnuPUq12gNE6ELGtkLVKMLl1mU3luMTxkIx4tFFOOL3ZtMIyRgqGmmFVfS0hTPnNqL1dhTeHRpOHQosXGvODQoamsoL6/tJ+KPa6VQio7GoJ1NI5+OL6gR8ioS1JSJF5bKNrGjd7Mmu/o1W11ZHiAby5KHNO7"
    s = s & "dNbAfzVYIDc/HDpEJuzkzOrgxExOIj5LND5SIUlMIvHs3/bz7URyJ0p9KjFQHTtOH+vhxt3d1CwAAAAALAFiAAAG/8CccCjEcVQ10GPJbDqf0Kh0SoXGGBMcccvter/gsHhMLpvP6LR6bIhV3/C4fAqSaNf4vH7P7/vNHRVzg4SFUzEGf4qLjI2OeAaGkpOSEo+XmJma"
    s = s & "ehyUnp9xK5ujpKWjHaCpqlGJpq6vsHmCq7SrIBqxubq7XJG1v6k1vMPEpjhuwMmeHcXNzo6dytKSos/W13oq09uEFNjf4GY13ORxMeHo6V1K5e1U6vDq7vNS8fbg9PlO9/zW+v8P+gksBlDfwIO7CuZDyBCWQnoNI5Z6OE+iRU3baHzYuJHdAx0QQoZkMSeBSAg7CG25"
    s = s & "c7Fln20iSMiUGYIJCAQ4caaMwyMnTv8eKokIcEl0zzQYI2bKHOExgU8EHqvs8Lmg0JahRbOqmRZC6UwLTF481QHnpk8HVoVqXXtGGgyvMz80gfA0qpSpPu3Gucq27xhpFuDOLMHEwdgqLA6nHYLVr+Mu0j4IlimXCd2cEKroqGuIb0MZGA4IGH3gAIYNJ1wAIJHiMRdl"
    s = s & "JSbPpMFk81kqT9F2VtvwgGylrV0TUSb5twibT6FKsZ1zkueIMlr8Di5cSLLYv0kwZYI3p+4nZnOSlfR8T4oZ6FOwfJVievUhyWJ6NZFUKdgliX0miCLWJ1DyvK2RQgYuECBYAC2cMMMrq01GXXXAgFBfXA8o4NV2S3SHE0lPLOD/U2bOBXhGCig0SEIBKMyQQgoyHOCC"
    s = s & "UgG40NgmKMj2oHDANACXAktMKJMJyFH1RE8+vUBJeWTgcIJXJ3jhA1wDzJgHDut5IYCN71lXi4RwsROYUh404ZRPHFr2oSdIioFBADAu+AUOL3rVgg944CCACwHc2MWVDmaZwy8WelXTEiDAdUNYT+20BJHeoSliGDhI55WUXsTp1QFlpHDaCS2wOZOeXPApGKiO/eLB"
    s = s & "hVHpCOZcnGXYaoiMkZGCiTM1OQYOBQiGAhmiwkWqUFhmWcsNcB3nhI8kIPNAf+IRmugnaXaRgqdKAWBGe4K1MEavXv3KWLDv1XKqV7Q50ZVSgy5x/1lOSmionKOxhiEDtUphasaSgtkKBrfAicHvp37SYgJclTnxllIjNGGYTzowqhMo0W5B60wBoCGDbPZ+8a9M3gqx"
    s = s & "MWsBr1KcUjxCMe5MQOL3FARjNgrto1z4lm8auR5Y5VXgapwzjqoQ61XBTxxMIRMeJpcXxDATcXGfaMis67477xn1Y6vIp1QDU5ws030P5Gc0AuO9HK8X+MJlrRr0wgg102uPGjIoXHoFwxTYzYThA0UbrReAY3OB7cxqWArXrlZOjTPbi5yX8a0yCCDDiDNQOkMGqbmA"
    s = s & "wgHeqnKuUsZOMfKPiBqtqNgefyH4pGtg8BsRKBTgegFpK0XA67Qf7v82EQAA0EILLrhwwu+/9+5CC25y6nvww9csExe8H59aCwBQu7gQKPAOfO8tKK/tEAfELhMApKYSt1JzU+HzTGE6a/R/SPdNBA7uqTHDb40JgML9KJyu1An492+7r1tAzcS8EqPguMh7SuECp7Iz"
    s = s & "vRwc8DfbE4CkZBOADUAGFIHiHBy0lqzCJAdEqYjYx0hQp98QLlSGAxbiVDi4LuBgAKuTGsa8gAMDTUZbVyKeEGYwQJmcED6g4CAJUlaFgZEsSLdRRcTKZrY8IJAE20PhCmV4uy7Mb3+/uuJkTDdDL/wNLi2YlZuGgIMekmB6oKABXBIWhwmxcQkLe8oqImZGErj/IA8T"
    s = s & "hEvFqAjAMHzMWyaiFBfq+IUnTaaBuLshAB7nt5yB4nMyKRkcjEgCrKkvOUYKYdJyYMI8MBEuN8vBH/2VwhyspgCMDMMnE/iFLlZKNj/cgv5koq8/eUKNPxtEccr3gDg+ZT+adJ8QvggXRJohA78ZIwv72DZmcmFpqQxDjbb4BQQaUwirlMkeveA0uAzHE6pC1yAiQIL0"
    s = s & "LaFlydnbboQpSvrloZtwEeQo/ZhC6ShTmrIBg/IuBYZsQhEMSxNMYz6BLBIcag5I6QETvCa6YJYuZsnMwwgROc9mdotsJBDkF6YpGH0esp83DINsouiJDNqtBChNqUpXulIi4u1r/+xq30O3wFHB3BMNI4ylx0q50ynm4EUa3Wg+v7DPeoE0W2Go4x0oMb7sOPVuDP1a"
    s = s & "2I6UNH/K5KZn0OLTpFhFnU1xSRhAQ0294lHBXDMH/owiRifDSEpQ0qlOTddHEsWc5pBOlF2wKgmwaoac8vGi9FyhpMJ6hrGy0gtFnclZ04rPyWCFEkKE62QUesmcvCA8LqMqOw2rFL6WYaJ/7VdgqzirmQwglGDg7PKI+tEvMBYM8FQKYSchNMn+xpwPcBcC4JhOeM2U"
    s = s & "CMiUTVDJEFulyJOn7WTbxO5oBtWSkLVmPSoYR9tCWxpCiCFQgHa3y93udpeXmMXJd5LzHb79lnvuxP+D6mQTzW/59H/dslMPPdsF55a1mNKVE3W9QjhJ9GCNnqgrTpogYBCaF6/P/E0G8uDc53LVmYWbjA1DWgb7Qhe/rqXwF7TKXy0V4kv28cS6cDJVh21Is+ctwm9q"
    s = s & "mYZZyoQAEe5qjPWYHTqRwcKIbe1apwsGDiulv4aA5N0MwayclEldK0MxgrkAQw2roY7MfTBgLQqcFOTRK6e98VBzHN0MI7XHsPTwIOo2E7kaYsQIMPASivwTWKVYCFdWygDy8BsLhhZg++0XDp6oUy/guAuJlclinWxF2dg5yIaihIARwD5W+WR0QWFnDhqM2jEQ0ys2"
    s = s & "lrJoqYznjwWg0qzbMqD/dZxXQmtasWKWg0llgttChDfNUfAlTtQJh4gZcjKZxqls5jzjKXu1inGeiVqFSk0uY3jH+oWtbNxEiKbK5KCSkDUCytuEV0N6DhHLQZMFQ9g0OPeaFf01M0c4XOqJmguBPmN+lTJsLmwgWIR4K6tF3NsoLPrI2N6kA2UTZTS4uACc5hgpERfs"
    s = s & "74nhz+gmtQJNvQV/bpMQQnRpIaQ9VfB8cJ1vHoIZt4kGBHb7zgLPM56HyefGdvTC/PQyj79QcH0NorZLoXUc9EYF3TZaDtnet8KJOxlebxi54dY5rsHw7mJ3YcIpRza7w4DAVA7CamWmhIkRAMwpsBlsi8l4ERB4/7YzFLzct4YwyEG2hXSToOtdWO/JveDKUn+ZhpPh"
    s = s & "OELXyEtD6DaTVEDzu/It6SHMa+hmCCiTxCB4X/d60+htO3wPm+CdE+G1h6flN+NATq90zhCv3u0bKB5prQ9B7V7p9xiyifYvXHrkZFw86oeAdK/A2As+ngmxu6x0YacWcXOAJAkIMwndUnsKepd5FHI+BNWWmwindwGoibAzXH086DtcMdwnE8o9G3rdtv9CD2MpB3mT"
    s = s & "AGiFiOqA47Do31eB+EOIPdnFkFgWh+Hrsly/0OHS3iG4WCa5JkK6sZqadLv/8QzXU8k2eW+ge7wnCYtWcVTwamr2Bug3BD4wQC2wfP8/hWVnBXITWHwkEEsFp25dEHYr13CT4X6tgyuTAXAqN4ArETsBUH9xAAMWEIMyKEnR5gA2eIPC9wQ3uIM5uA/65mdwUQDecjoZ"
    s = s & "iAYFRwC/wyaiN0KlR0bZ0W45cHq0dAc+YE85YHZ9hk0n+EquF0oU4Q4PuAUycAJpgwI3IwC00gLH9wUysG2TEWUykD9OVQAnYGdCgAEdOBPgEypP5CkF4CZmd3YnQCkOBwCYowUpcAATEwAngFpf2A5h2AWKGHqicQBkSDEn0DFkcH8/tCYDUAC60zzCszsAADsOlgMB"
    s = s & "EACgqDvCMzzQAzuCJAN1FAA/tE8BkDu7Az0E0G5pJYv/v6F8YPCI5SCGUzIDB4ACz7M7vrMBjvMHLZKMJ4ABmrgJM5A/uWg5AlAlGOA4FAiAX+YDGMApyogC2SgGwkgOfqIOkLcG58gN6ZgO67gV7TgN74gO8ZgG80iP9QgO94gG+SgN+8iPAWgG/6gMAfkN/tSE/liQ"
    s = s & "wHCQ2OBiHCePDFkLDnkNrfdjeaAsEyk+FekMSpIdJ1B/ZXABG7kKFdCRxCBBtoUg/xcGElCSqqACKDkMaxKKrXiTpAg7UBgGPwCTqTABM8kW2uCTTPUDQbkWP9CDRDkFlnCUa/GSS1kIF+CUbIEDDBCVg1AB3UiVDUEBWGkORsmVVbkCX1kFDBCWI2LZFxowDmX5BDHA"
    s = s & "AWkpHB2wAhSgkUQJAhWgAhywlXE5EEEAADs="
    B64_KARTICA = s
End Function

' MINI -- 160x52, 1665 B GIF-a, 2220 znakova Base64
Private Function B64_MINI() As String
    Dim s As String
    s = s & "R0lGODlhoAA0APUAANK4b8yvXMenS52hk5OYicusVeTWrl2gNFNTJb/BttfBguHRpFmYMl1bKFVfS3twMuLTqM7PxXxxMlmXMjtiIkV0J2pzYX+GdvDp2fbz7aCLPrOZRB4tFDpSH0NJJkBMNk2CK1SQLyc6GDI8G4N2NI59OGtlLXRrL0pVQGFdKSxHGzZLHejdvevh"
    s = s & "xtvIkN7NmwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACwAAAAAoAA0AAAG/8CMMENAiTjIpHLJbDqTn0tkSK1ar9isdsvtereJz3NMLiNFhK96zW67rQmzfM68vO/4vFtM788TeoGCg0IEfodmDoSLjGwoiJBjGI2UlVhHkZlL"
    s = s & "U5adlpqgSYCepIyhoaOlqoGnoKmrsG+QCCkpSCsUFJhNHRQdY6+xwmqIIwLHCBwUBwcVTirMzk+AGJPD11yIJMcCGkgTzLtKIMwqwBnV2OpZh8bcAiPKzCBMIswhZNTW1+nrVIcn3gkogYRBuCUVmK3Ih25fKQMAIgYoIECBv39+NgiExyFhMyX2DuBj2I9LNYd3MFDk"
    s = s & "ZvGiED8IuJ3QKOAEB2gHGOzySKGMPv8tGF4ACODihYsAARQYwOPiXUuXfjQc28AhxdQj5D5yCKnTZ0MsGBQceyrkxbECLkwaYGEFglOXL+ms4PYACU0THDowO3DEozSSKIW0WAngStNjAAJngDD02AIrBt7C7VOCWzwOD6YiCcFM115zXksOGXxWcQYA3AIEPuwYsmSo"
    s = s & "c9wNTCK7gbwDEzyODI1SJbe0WFi8K1yFtYDHVSKzhJuBTuZjC5NIOOaNA7i9B6LzriKWWwstAd4tpWIcORXlY5nPEcGNIEiaHm7f+/N1iPDUW8xyKyDkZPdjL5xkDXoVDSFaP9ZUwwJb/fWHQQsG7JPgg+NlwIJRL3yHkRkBIcP/xHMkbIVdT3L8NARqy2nRgkBLGaCA"
    s = s & "AuFxA8CLCjBIYEthbWQRCzHKeOJGGmKA4jsZtNBjekPIIQJNVDHhTpMe8UVfSRgIZF4WAgEnRHmupdgfigo4NKQAxAkmkIZCEChACy5MwsJKZCZphgnc4NXEdCFywBkzv5RYXwZqMgjecOS9c2War5V1XJeIVXEkmkII9MJ57wAnB00CiAOSAPHptRc9fpakn3dcjMnf"
    s = s & "EFwml+hpFVKxoo9UjAlpBnAKgFKtk5jhQXsj9Orrr3Zm9Vmo+xhna6kCFcrNoYAmqgAEWLzaaKzvzFqrFWMuZUZlG3VrG045zUPsEP8dYxoV5Qqg/2xrqnqpgJZWSBsntaRScS13lTZHhmzdvtNkViEsU44ZJmZg7Lk/EomqoYwWaLAAAWQhb5lCyFrFvesWWMY2x5xg"
    s = s & "wscghxzfCnv9ch2ogC38DsIVvxOxyuxSylILQ8LrKqH0HmOtwlSMGue+3FRHRlYMICGwlCkL4dY7gmZxJMWpynwMUgIpNnEVFtvL8xA+F0ZGh5yWEdIBJHJgkFbnlCSvANBuUeukMC/abnpjUjwazgnrfPHWii5HBk1Cj+FR0UkMvt0QtZJ1RZXVZszsjUWyGC/eLdeL"
    s = s & "ON8P/6avEw1wYwsZIpz91017le1EwZlPvcXS08b9uGTlvnwzrHmvuf83N1YYB+0YTJoRpaaDa7oE6r5ZfkXWrjeMo6RVXJ2z7VrjjvXKmzPR+TESiP0pE+CazgTqGSyQL1jj98zw3A6nLkABDjlf+87SU9Hjy09IdcxlYxwN2hIApx3Yf+zDAmvsJgTxLcs+iPIS42iH"
    s = s & "DsotUG/RO0bzmBYXJuzqGO4ZQ+jE1QSSdWYaf1qXaqxwHwHY7G6aAxRwfEYWA3KjbULo0amEoIBs3U6COdOSEzxwgh7i7wkqyAUF9scEIXpPCeAbggEIU4X7FKBVVjiSC2rYnyMFsD8CKQCaWNgCFgCAZsNxSK3Q5EKytEITSaQCYwoQgBcYwEUFAAAEWCakMAKmKiI0"
    s = s & "qiEAJiWkGdEoIvt4Qa2+eJo4+rFGl+OGUWqIGCieMRPB2MKDFsQyK2CAkoToYgsqKYR7aVIxj4xEJJlDCIxpIZSQGCUpBWFKdqDyEKpcpR6qlo1X+iGWsrzD2hR3BT7Ycg6czKUaFFCrqYUpCxf45Rw+IExCCOiZCMPAD5X5hAE0cx0DoCYZFHFNbApPmxywQDC72YkW"
    s = s & "XMCX1ByBBXBJTlIEAQA7"
    B64_MINI = s
End Function
