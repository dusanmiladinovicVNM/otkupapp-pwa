Attribute VB_Name = "modModernUI"
Option Explicit

'--- Farben (RGB!)
Private Const UI_BG As Long = &HF8F6F5          'RGB(245,246,248)
Private Const UI_CARD As Long = &HFFFFFF        'weiß
Private Const UI_TEXT As Long = &H37291F        'RGB(31,41,55)
Private Const UI_MUTED As Long = &H8A847E       'grauer Text

Public Sub ModernUI_ApplyTheme(ByVal frm As Object)
    On Error Resume Next

    frm.BackColor = UI_BG
    frm.Font.Name = "Segoe UI"
    frm.Font.Size = 10

    'Optional: etwas „Luft“ – wenn du kein pixelgenaues Layout hast
    'frm.Width = frm.Width + 10

    ModernUI_StyleControls frm

    On Error GoTo 0
End Sub

Private Sub ModernUI_StyleControls(ByVal parent As Object)
    Dim c As Object
    For Each c In parent.Controls
        On Error Resume Next

        'Grundfont für alles
        c.Font.Name = "Segoe UI"
        c.Font.Size = 10
        c.foreColor = UI_TEXT

        Select Case TypeName(c)

            Case "TextBox"
                c.BackColor = UI_CARD
                c.BorderStyle = fmBorderStyleSingle
                c.SpecialEffect = fmSpecialEffectFlat
                'Einheitliche Höhe (optional)
                'c.Height = 20

            Case "ComboBox"
                c.BackColor = UI_CARD
                c.SpecialEffect = fmSpecialEffectFlat
                c.style = fmStyleDropDownList

            Case "ListBox"
                c.BackColor = UI_CARD
                c.SpecialEffect = fmSpecialEffectFlat

            Case "CommandButton"
                c.SpecialEffect = fmSpecialEffectFlat
                c.Font.Size = 10
                'Optional: Standardhöhe
                'c.Height = 26

                'Schnelle “Primary Button”-Regel über Namen/Caption:
                'Passe das an deine Buttons an (z.B. cmdUnos, cmdStampaj)
                If LCase$(c.Name) Like "*unos*" Or LCase$(c.Caption) Like "*unos*" Or LCase$(c.Caption) Like "*save*" Then
                    ModernUI_PrimaryButton c
                ElseIf LCase$(c.Name) Like "*povrat*" Or LCase$(c.Caption) Like "*povrat*" Or LCase$(c.Caption) Like "*cancel*" Then
                    ModernUI_SecondaryButton c
                Else
                    ModernUI_SecondaryButton c
                End If

            Case "Label"
                c.BackStyle = fmBackStyleTransparent
                'Wenn es eher “Hinweistext” ist:
                If Len(c.Caption) > 0 And (InStr(1, LCase$(c.Name), "valid") > 0 Or InStr(1, LCase$(c.Name), "manjak") > 0) Then
                    c.foreColor = UI_MUTED
                End If

            Case "Frame"
                'Frames wirken alt – mach sie „unsichtbar“
                c.SpecialEffect = fmSpecialEffectFlat
                c.BorderStyle = fmBorderStyleNone
                c.BackColor = UI_BG
                'Controls innerhalb des Frames ebenfalls stylen
                ModernUI_StyleControls c

            Case "MultiPage"
                ModernUI_StyleControls c

            Case "Page"
                c.BackColor = UI_BG
                ModernUI_StyleControls c

            Case Else
                'Falls ein Control kein Font/Color hat, einfach ignorieren
        End Select

        On Error GoTo 0
    Next c
End Sub

Private Sub ModernUI_PrimaryButton(ByVal btn As Object)
    'MSForms-Buttons können nicht sauber „rund“ + farbig wie Web,
    'aber flach + dunkler Text wirkt schon deutlich moderner.
    btn.BackColor = RGB(37, 99, 235)   'Akzent Blau
    btn.foreColor = vbWhite
    btn.Font.Bold = True
End Sub

Private Sub ModernUI_SecondaryButton(ByVal btn As Object)
    btn.BackColor = RGB(243, 244, 246) 'hellgrau
    btn.foreColor = UI_TEXT
    btn.Font.Bold = False
End Sub

