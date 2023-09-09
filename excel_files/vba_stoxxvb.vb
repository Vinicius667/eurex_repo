Sub Ueberhaenge()

    Dim UeberhaengeSheet, PutFront, CallFront, Hedge, Summery, Info As Worksheet
    
'    Dim Prng, Crng As Range
    
    
    'Kurstabelle erzeugen
        Dim import_sheet As Worksheet
    Set UeberhaengeSheet = ActiveWorkbook.Sheets("Ueberhaenge")
    Set Hedge = ActiveWorkbook.Sheets("STOXX_HedgeBedarf")
    Set Summery = ActiveWorkbook.Sheets("STOXX_Summery")
    Set Info = ActiveWorkbook.Sheets("STOXX_Infos")
    Set PutFront = ActiveWorkbook.Sheets("STOXX_Put_Front")
    Set CallFront = ActiveWorkbook.Sheets("STOXX_Call_Front")
    Set Put01 = ActiveWorkbook.Sheets("STOXX_Put+01")
    Set Call01 = ActiveWorkbook.Sheets("STOXX_Call+01")
    Set PutLast_1 = ActiveWorkbook.Sheets("STOXX_PutFront-1")
    Set CallLast_1 = ActiveWorkbook.Sheets("STOXX_CallFront-1")
   
    
    If MsgBox("Wollen Sie die Überhänge wirklich neu berechnen? Alle Ergebnisse werden gelöscht!", vbOKCancel) = vbCancel Then
        Exit Sub
    End If
    
Dim LetzteZeile As Long
Dim x As Long
Dim KontraktWert As Integer

'Zentralkurs
Do Until Abs(UeberhaengeSheet.Range("C3").Value - UeberhaengeSheet.Range("X4").Value) <= 350
    If UeberhaengeSheet.Range("C3").Value - UeberhaengeSheet.Range("X4").Value > 350 Then
        UeberhaengeSheet.Range("C3").Value = UeberhaengeSheet.Range("C3").Value - 350
    ElseIf UeberhaengeSheet.Range("C3").Value - UeberhaengeSheet.Range("X4").Value < -350 Then
         UeberhaengeSheet.Range("C3").Value = UeberhaengeSheet.Range("C3").Value + 350
    End If
Loop

' Zahlenformat ändern
With PutFront
    LetzteZeile = .Range("B65536").End(xlUp).Row - 1
    For x = 2 To LetzteZeile
        .Cells(x, 1) = CLng(.Cells(x, 2))
        .Cells(x, 1).NumberFormat = "General"
        .Cells(x, 10) = CLng(.Cells(x, 5))
        .Cells(x, 10).NumberFormat = "General"
    Next
End With

With CallFront
    LetzteZeile = .Range("B65536").End(xlUp).Row - 1
    For x = 2 To LetzteZeile
        .Cells(x, 1) = CLng(.Cells(x, 2))
        .Cells(x, 1).NumberFormat = "General"
        .Cells(x, 10) = CLng(.Cells(x, 5))
        .Cells(x, 10).NumberFormat = "General"
    Next
End With

With Put01
    LetzteZeile = .Range("B65536").End(xlUp).Row - 1
    For x = 2 To LetzteZeile
        .Cells(x, 1) = CLng(.Cells(x, 2))
        .Cells(x, 1).NumberFormat = "General"
        .Cells(x, 10) = CLng(.Cells(x, 5))
        .Cells(x, 10).NumberFormat = "General"
    Next
End With

With Call01
    LetzteZeile = .Range("B65536").End(xlUp).Row - 1
    For x = 2 To LetzteZeile
        .Cells(x, 1) = CLng(.Cells(x, 2))
        .Cells(x, 1).NumberFormat = "General"
        .Cells(x, 10) = CLng(.Cells(x, 5))
        .Cells(x, 10).NumberFormat = "General"
    Next
End With


With PutLast_1
    LetzteZeile = .Range("B65536").End(xlUp).Row - 1
    For x = 2 To LetzteZeile
        .Cells(x, 1) = CLng(.Cells(x, 2))
        .Cells(x, 1).NumberFormat = "General"
        .Cells(x, 10) = CLng(.Cells(x, 5))
        .Cells(x, 10).NumberFormat = "General"
    Next
End With

With CallLast_1
    LetzteZeile = .Range("B65536").End(xlUp).Row - 1
    For x = 2 To LetzteZeile
        .Cells(x, 1) = CLng(.Cells(x, 2))
        .Cells(x, 1).NumberFormat = "General"
        .Cells(x, 10) = CLng(.Cells(x, 5))
        .Cells(x, 10).NumberFormat = "General"
    Next
End With

With UeberhaengeSheet
        Minkurs = .Range("C3").Value - (.Range("C4").Value / 2)
        Maxkurs = Minkurs + (.Range("C4").Value)
        Schritte = (.Range("C4").Value) / (.Range("C5").Value)
        KontraktWert = .Range("C6").Value
        
        If Info.Range("O30").Value = "Adjusted" Then OIadjusted = True Else OIadjusted = False
        
        
        
        If Summery.Range("C3").Value >= 1 Then
            Delta = 0.5
        Else
            Delta = 1
        End If
        Summery.Range("F3").Value = Delta
        
        
        'alte Summenüberhänge retten um sie zur Berechnung der Änderung des Überhangs zu verwenden
        For I = 0 To Schritte
            Summery.Range("T" & 10 + I).Value = Summery.Range("P" & 10 + I).Value
        Next I
        
        'alte Tabelle löschen
        .Range("A10:S" & .Range("A1000000").End(xlUp).Row + 1).Value = ""
        .Range("A10:S100").Interior.ColorIndex = 2  'weiss
       
        Summery.Range("A10:S40").Value = ""         '& .Range("A1000000").End(xlUp).Row + 1
        Summery.Range("A10:S40").Interior.ColorIndex = 2  'weiss
       
        
        'Basiswerte schreiben + Puts
        Z = 2
        
        For I = 0 To Schritte
            BasisI = Minkurs + I * .Range("C5").Value
            .Range("A" & 10 + Schritte - I).Value = BasisI
            Summery.Range("A" & 10 + Schritte - I).Value = BasisI
            Summery.Range("N" & 10 + Schritte - I).Value = BasisI

            'Set Prng = PutFront.Range("A1:A200").Find(BasisI)
            For K = Z To 500
                '.Range("E2").Value = PutFront.Range("A" & K).Value

                If (PutFront.Range("A" & K).Value = BasisI) Then
                    If OIadjusted Then
                        .Range("D" & 10 + Schritte - I).Value = PutFront.Range("J" & K).Value
                    Else
                        .Range("D" & 10 + Schritte - I).Value = PutFront.Range("I" & K).Value
                    End If
                    Z = K + 1
                    Exit For
                Else
                If (PutFront.Range("A" & K).Value > BasisI) Or K = 499 Then
                    .Range("D" & 10 + Schritte - I).Interior.ColorIndex = 3                    ' rot
                    
                    .Range("D" & 10 + Schritte - I).Value = ""
                    Z = K
                    Exit For
                End If
                End If
            Next K
        Next I
        
 'Calls
         Z = 2
        
        For I = 0 To Schritte
            BasisI = Minkurs + I * .Range("C5").Value

            For K = Z To 500
                '.Range("E2").Value = CallFront.Range("A" & K).Value

                If (CallFront.Range("A" & K).Value = BasisI) Then
                    If OIadjusted Then
                        .Range("D" & 10 + Schritte - I).Value = .Range("D" & 10 + Schritte - I).Value - CallFront.Range("J" & K).Value
                    Else
                        .Range("D" & 10 + Schritte - I).Value = .Range("D" & 10 + Schritte - I).Value - CallFront.Range("I" & K).Value
                    End If
                    Summery.Range("C" & 10 + Schritte - I).Value = (.Range("D" & 10 + Schritte - I).Value) * (1 / KontraktWert) * Delta
                    Z = K + 1
                    Exit For
                Else
                If (CallFront.Range("A" & K).Value > BasisI) Or K = 499 Then
                    .Range("D" & 10 + Schritte - I).Interior.ColorIndex = 3                    ' rot
                    .Range("D" & 10 + Schritte - I).Value = ""
                    Z = K
                    Exit For
                End If
                End If
            Next K
        Next I
' + 01
        'Basiswerte schreiben + Puts
        Z = 2
        
        For I = 0 To Schritte
            BasisI = Minkurs + I * .Range("C5").Value

            For K = Z To 500
                '.Range("E2").Value = Put01.Range("A" & K).Value

                If (Put01.Range("A" & K).Value = BasisI) Then
                    If OIadjusted Then
                        .Range("E" & 10 + Schritte - I).Value = Put01.Range("J" & K).Value
                    Else
                        .Range("E" & 10 + Schritte - I).Value = Put01.Range("I" & K).Value
                    End If
                    Z = K + 1
                    Exit For
                Else
                If (Put01.Range("A" & K).Value > BasisI) Or K = 499 Then
                    .Range("E" & 10 + Schritte - I).Interior.ColorIndex = 3                    ' rot
                    
                    .Range("E" & 10 + Schritte - I).Value = ""
                    Z = K
                    Exit For
                End If
                End If
            Next K
        Next I
        
 'Calls
         Z = 2
        
        For I = 0 To Schritte
            BasisI = Minkurs + I * .Range("C5").Value

            For K = Z To 500
                '.Range("E2").Value = Call01.Range("A" & K).Value

                If (Call01.Range("A" & K).Value = BasisI) Then
                    If OIadjusted Then
                        .Range("E" & 10 + Schritte - I).Value = .Range("E" & 10 + Schritte - I).Value - Call01.Range("J" & K).Value
                    Else
                        .Range("E" & 10 + Schritte - I).Value = .Range("E" & 10 + Schritte - I).Value - Call01.Range("I" & K).Value
                    End If
                    Z = K + 1
                    Exit For
                Else
                If (Call01.Range("A" & K).Value > BasisI) Or K = 499 Then
                    .Range("E" & 10 + Schritte - I).Interior.ColorIndex = 3                    ' rot
                    .Range("E" & 10 + Schritte - I).Value = ""
                    Z = K
                    Exit For
                End If
                End If
            Next K
        Next I

' Front-1
        'Basiswerte schreiben + Puts
        Z = 2
        
        For I = 0 To Schritte
            BasisI = Minkurs + I * .Range("C5").Value

            For K = Z To 500
                '.Range("E2").Value = Put01.Range("A" & K).Value

                If (PutLast_1.Range("A" & K).Value = BasisI) Then
                    If OIadjusted Then
                        Summery.Range("D" & 10 + Schritte - I).Value = PutLast_1.Range("J" & K).Value
                    Else
                        Summery.Range("D" & 10 + Schritte - I).Value = PutLast_1.Range("I" & K).Value
                    End If
                    Z = K + 1
                    Exit For
                Else
                If (PutLast_1.Range("A" & K).Value > BasisI) Or K = 499 Then
                    Summery.Range("D" & 10 + Schritte - I).Interior.ColorIndex = 3                    ' rot
                    
                    Summery.Range("D" & 10 + Schritte - I).Value = ""
                    Z = K
                    Exit For
                End If
                End If
            Next K
        Next I
        
 'Calls
         Z = 2
        
        For I = 0 To Schritte
            BasisI = Minkurs + I * .Range("C5").Value

            For K = Z To 500
                '.Range("E2").Value = Call01.Range("A" & K).Value

                If (CallLast_1.Range("A" & K).Value = BasisI) Then
                     If OIadjusted Then
                        Summery.Range("D" & 10 + Schritte - I).Value = (Summery.Range("D" & 10 + Schritte - I).Value - CallLast_1.Range("J" & K).Value) * (1 / KontraktWert) * Delta
                    Else
                        Summery.Range("D" & 10 + Schritte - I).Value = (Summery.Range("D" & 10 + Schritte - I).Value - CallLast_1.Range("I" & K).Value) * (1 / KontraktWert) * Delta
                    End If
                    Z = K + 1
                    Exit For
                Else
                If (CallLast_1.Range("A" & K).Value > BasisI) Or K = 499 Then
                    Summery.Range("D" & 10 + Schritte - I).Interior.ColorIndex = 3                    ' rot
                    Summery.Range("D" & 10 + Schritte - I).Value = ""
                    Z = K
                    Exit For
                End If
                End If
            Next K
        Next I

        
        For I = 0 To Schritte
            .Range("C" & 10 + I).Value = WorksheetFunction.Sum(.Range("D" & (10 + I) & ":I" & (10 + I)).Value)
            Summery.Range("P" & 10 + I).Value = (.Range("C" & 10 + I).Value) * (1 / KontraktWert) * Delta                                'Berechnung Summe über alle Verfallsdaten (der hier abgefragten)
            Summery.Range("B" & 10 + I).Value = Summery.Range("C" & 10 + I).Value - Summery.Range("D" & 10 + I).Value       'Berechnung der Änderung von gestern auf heute Front
            Summery.Range("Q" & 10 + I).Value = Summery.Range("T" & 10 + I).Value                                           ' alte Werte wieder zurück kopieren
            Summery.Range("O" & 10 + I).Value = Summery.Range("P" & 10 + I).Value - Summery.Range("Q" & 10 + I).Value       'Berechnung der Änderung von gestern auf heute Summe
            If Delta = 1 Then
                Summery.Range("D" & 10 + I).Value = Summery.Range("D" & 10 + I).Value / 2
            End If
        Next I
        
    End With
        
    
        
End Sub

Sub HedgeBerechnung()

    Dim UeberhaengeSheet, Hedge As Worksheet
    Dim d1, h1, h2, h3, h4, h5, h2_1, d1_1 As Double
    Dim KontraktWert As Integer
    
    Set UeberhaengeSheet = ActiveWorkbook.Sheets("Ueberhaenge")
    Set Hedge = ActiveWorkbook.Sheets("STOXX_HedgeBedarf")
    Set Hedge_1 = ActiveWorkbook.Sheets("STOXX_HedgeBedarf+01")
    
    UeberhaengeSheet.Range("R1").Value = "Start Hedge"
    UeberhaengeSheet.Range("R1:S1").Interior.ColorIndex = 6
    
    KontraktWert = UeberhaengeSheet.Range("C6").Value

With Hedge_1
        .Range("A4:BB" & .Range("A1000000").End(xlUp).Row + 1).Value = ""
        .Range("A4:BB200").Interior.ColorIndex = 2  'weiss
End With

With Hedge
        
        'alte Tabellen löschen
        .Range("A4:BB" & .Range("A1000000").End(xlUp).Row + 1).Value = ""
        .Range("A4:BB200").Interior.ColorIndex = 2  'weiss
       
        'Hedge_1.Range("A4:BB" & .Range("A1000000").End(xlUp).Row + 1).Value = ""
        'Hedge_1.Range("A4:BB200").Interior.ColorIndex = 2  'weiss
       
        Minkurs = UeberhaengeSheet.Range("C3").Value - (UeberhaengeSheet.Range("C4").Value / 2)
        Maxkurs = Minkurs + (UeberhaengeSheet.Range("C4").Value)
        Schritte = (UeberhaengeSheet.Range("C4").Value) / (UeberhaengeSheet.Range("C5").Value)
        
        SchrittWeite = 10
        InterestRate = UeberhaengeSheet.Range("N3").Value
        VDAX = UeberhaengeSheet.Range("N4").Value
        VDAX_Laufzeit = UeberhaengeSheet.Range("N6").Value
        
        Tage = UeberhaengeSheet.Range("G2").Value
        If Tage = 0 Then
            Tage = 0.5
        End If
            
        Tage_1 = UeberhaengeSheet.Range("R5").Value
        
        For I = 0 To (UeberhaengeSheet.Range("C4").Value / SchrittWeite)
            Kurs = Maxkurs - I * SchrittWeite
            .Range("A" & I + 5).Value = Kurs
            Hedge_1.Range("A" & I + 5).Value = Kurs
        Next I
        
        For K = 0 To Schritte
            Basis = UeberhaengeSheet.Range("A" & 10 + K).Value
            .Cells(4, K + 4) = Basis
            Hedge_1.Cells(4, K + 4) = Basis
            Kontrakte = UeberhaengeSheet.Range("D" & 10 + K).Value
            Kontrakte_1 = UeberhaengeSheet.Range("E" & 10 + K).Value
       
            For I = 0 To (UeberhaengeSheet.Range("C4").Value / SchrittWeite)
                Kurs = Maxkurs - I * SchrittWeite
                h1 = Application.WorksheetFunction.Ln(Kurs / Basis)
                sigma = VDAX    '* ((Tage / VDAX_Laufzeit) ^ 0.5)                  '* ((Tage / 365) ^ 0.5)
                sigma_1 = VDAX * ((Tage_1 / VDAX_Laufzeit) ^ 0.5)              '* ((Tage_1 / 365) ^ 0.5)
                h2 = InterestRate + sigma * sigma / 2
                h2_1 = InterestRate + sigma_1 * sigma_1 / 2
                d1 = (h1 + (h2 * (Tage / 365))) / (sigma * ((Tage / 365) ^ 0.5))
                d1_1 = (h1 + (h2_1 * (Tage_1 / 365))) / (sigma_1 * ((Tage_1 / 365) ^ 0.5))
                Phi = Application.WorksheetFunction.Norm_Dist(d1, 0, 1, False)               'Dichtefunktion
                Phi_1 = Application.WorksheetFunction.Norm_Dist(d1_1, 0, 1, False)               'Dichtefunktion
               'Phi = Application.WorksheetFunction.NormDist(d1, 0, 1, False)               'Dichtefunktion
                Gamma = Phi / (Kurs * (sigma * (Tage / 365) ^ 0.5))
                Gamma_1 = Phi_1 / (Kurs * (sigma_1 * (Tage_1 / 365) ^ 0.5))
                .Cells(I + 5, K + 4).Value = Gamma * Kontrakte / KontraktWert
                Hedge_1.Cells(I + 5, K + 4).Value = Gamma_1 * Kontrakte_1 / KontraktWert
            Next I
        Next K
        
        For I = 0 To (UeberhaengeSheet.Range("C4").Value / SchrittWeite)
            HedgeSum = 0
            HedgeSum_1 = 0
            For K = 0 To Schritte
                HedgeSum = HedgeSum + .Cells(I + 5, K + 4).Value
                HedgeSum_1 = HedgeSum_1 + Hedge_1.Cells(I + 5, K + 4).Value
            Next K
            .Range("C" & I + 5).Value = HedgeSum / 2                                ' Faktor 0,5 weil nicht alles in einer Hand ist
            Hedge_1.Range("C" & I + 5).Value = HedgeSum_1 / 2                         ' Faktor 0,5 weil nicht alles in einer Hand ist
            If I = UeberhaengeSheet.Range("X5").Value Then
                UeberhaengeSheet.Range("Y4").Value = HedgeSum / 2
                UeberhaengeSheet.Range("Y5").Value = (HedgeSum / 2) + (HedgeSum_1 / 2)
                Hedge_1.Range("B" & I + 5).Value = (HedgeSum_1 + HedgeSum) / 2
            End If
        Next I
        
        
    End With
    
    UeberhaengeSheet.Range("R1").Value = "Hedge fertig"
    UeberhaengeSheet.Range("R1:S1").Interior.ColorIndex = 4

End Sub
