# 2-Excel-munkalap-1-1-oszlopa-osszevetese
Az Excel remek program adatkezelésre, de két külön munkalap oszlopai közötti összehasonlításra nem igazán alkalmas, a kézi keresgélés gyorsan időrablóvá válik. Szerencsére  makró program (VBA) lehetővé teszi, hogy néhány sor kóddal automatizáljuk ezeket a feladatokat.

Kód:

Sub EmailKereses_Dinamikus()

    Dim ws2 As Worksheet, ws3 As Worksheet
    Dim col2 As String, col3 As String, colOut As String
    Dim lastRow2 As Long, lastRow3 As Long
    Dim i As Long
    Dim dict As Object
    Dim emailCim As String

    ' --- Oszlopok bekérése ---
    col2 = InputBox("Add meg, melyik oszlopban vannak az email címek a 2. lapon (pl. A):")
    If col2 = "" Then Exit Sub

    col3 = InputBox("Add meg, melyik oszlopban vannak az email címek a 3. lapon (pl. A):")
    If col3 = "" Then Exit Sub

    colOut = InputBox("Add meg, melyik oszlopba írjam a találatokat a 2. lapon (pl. D):")
    If colOut = "" Then Exit Sub

    ' --- Lapok beállítása ---
    Set ws2 = ThisWorkbook.Sheets(2)
    Set ws3 = ThisWorkbook.Sheets(3)

    lastRow2 = ws2.Cells(ws2.Rows.Count, col2).End(xlUp).Row
    lastRow3 = ws3.Cells(ws3.Rows.Count, col3).End(xlUp).Row

    ' --- Dictionary létrehozása ---
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1   ' TextCompare (nem érzékeny kis/nagybetűre)

    ' --- 2. lap emailjeinek betöltése ---
    For i = 1 To lastRow2
        emailCim = Trim(ws2.Cells(i, col2).Value)
        If emailCim <> "" Then
            dict(emailCim) = i
        End If
    Next i

    ' --- 3. lap emailjeinek keresése ---
    For i = 1 To lastRow3
        emailCim = Trim(ws3.Cells(i, col3).Value)

        If emailCim <> "" Then
            If dict.Exists(emailCim) Then
                Dim sor As Long
                sor = dict(emailCim)

                ' Csak akkor ír, ha még üres
                If ws2.Cells(sor, colOut).Value = "" Then
                    ws2.Cells(sor, colOut).Value = emailCim
                End If
            End If
        End If
    Next i

    MsgBox "Kész!", vbInformation

End Sub
