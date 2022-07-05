Sub EnterValues_Eingabefeld_Button()
' Enter Values Script
' Author: Hannes Duve

'------------------------------------------------------------------------------------------------
'-------- WORKSHEETS, VARIABLES AND RANGES --------------------------------------------------------------
'------------------------------------------------------------------------------------------------

'Set worksheets
Dim Eingabefeld As Worksheet, Haushaltsbuch As Worksheet, Budget As Worksheet
Set Budget = Sheets("Budget pro Land")
Set Eingabefeld = Sheets("Eingabefeld")
Set Haushaltsbuch = Sheets("Haushaltsbuch")

'Entries of the Input fields
Dim DatCell As Range, RegCell As Range, CatCell As Range
Dim TxtCell As Range, PriCell As Range

Set DatCell = Eingabefeld.Range("C6")
Set RegCell = Eingabefeld.Range("D6")
Set CatCell = Eingabefeld.Range("E6")
Set TxtCell = Eingabefeld.Range("F6")
Set PriCell = Eingabefeld.Range("G6")

'Dynamically set categories by reference
Dim Cat1Cell As Range, Cat2Cell As Range, Cat3Cell As Range
Dim Cat4Cell As Range, Cat5Cell As Range, Cat6Cell As Range
Dim Cat7Cell As Range, Cat8Cell As Range, Cat9Cell As Range

Set Cat1Cell = Budget.Range("C99")
Set Cat2Cell = Budget.Range("C100")
Set Cat3Cell = Budget.Range("C101")
Set Cat4Cell = Budget.Range("C102")
Set Cat5Cell = Budget.Range("C103")
Set Cat6Cell = Budget.Range("C104")
Set Cat7Cell = Budget.Range("C105")
Set Cat8Cell = Budget.Range("C106")
Set Cat9Cell = Budget.Range("C107")

'Dynamically set 'other' categories by reference
Dim Sonst1Cell As Range, Sonst2Cell As Range, Sonst3Cell As Range, Sonst4Cell As Range, Sonst5Cell As Range

Set Sonst1Cell = Budget.Range("C108")
Set Sonst2Cell = Budget.Range("C109")
Set Sonst3Cell = Budget.Range("C110")
Set Sonst4Cell = Budget.Range("C111")
Set Sonst5Cell = Budget.Range("C112")

'Set possibility of multiple Regions/Countries (via offset? added a index to every region)
Dim offsetInt As Integer, offsetDistance As Integer
'Have to change the offsetDistance whenever the 'Haushaltsbuch' tables get changed
offsetDistance = 9
'The offsetInt is a multiplier to indicate in which 'Region' the entry will be made
offsetInt = 0
'Set maximum days accordingly
Dim maxDays As Integer
maxDays = 90

'Check the first two/three characters for their rank in the region category
If InStr(1, RegCell.Value, "1.") = 1 Then
    offsetInt = 0
ElseIf InStr(1, RegCell.Value, "2.") = 1 Then
    offsetInt = 1
ElseIf InStr(1, RegCell.Value, "3.") = 1 Then
    offsetInt = 2
ElseIf InStr(1, RegCell.Value, "4.") = 1 Then
    offsetInt = 3
ElseIf InStr(1, RegCell.Value, "5.") = 1 Then
    offsetInt = 4
ElseIf InStr(1, RegCell.Value, "6.") = 1 Then
    offsetInt = 5
ElseIf InStr(1, RegCell.Value, "7.") = 1 Then
    offsetInt = 6
ElseIf InStr(1, RegCell.Value, "8.") = 1 Then
    offsetInt = 7
ElseIf InStr(1, RegCell.Value, "9.") = 1 Then
    offsetInt = 8
ElseIf InStr(1, RegCell.Value, "10.") = 1 Then
    offsetInt = 9
ElseIf InStr(1, RegCell.Value, "11.") = 1 Then
    offsetInt = 10
ElseIf InStr(1, RegCell.Value, "12.") = 1 Then
    offsetInt = 11
End If

'Define and set the ranges where we will try to make an entry
'These sadly have to be hardcoded for now and have to be adjusted as the table changes
Dim AnreiRange As Range, InlanRange As Range, FortbRange As Range
Dim ToureRange As Range, UnterRange As Range, EssenRange As Range
Dim VisumRange As Range, SimkaRange As Range, VergnRange As Range
Dim Sons1Range As Range, Sons2Range As Range, Sons3Range As Range
Dim Sons4Range As Range, Sons5Range As Range

Set AnreiRange = Haushaltsbuch.Range("C13:C18")
Set InlanRange = Haushaltsbuch.Range("C24:C29")
Set FortbRange = Haushaltsbuch.Range("C35:C124")

Set ToureRange = Haushaltsbuch.Range("C130:C158")
Set UnterRange = Haushaltsbuch.Range("C164:C253")
Set EssenRange = Haushaltsbuch.Range("C259:C348")

Set VisumRange = Haushaltsbuch.Range("C353")
Set SimkaRange = Haushaltsbuch.Range("C358")
Set VergnRange = Haushaltsbuch.Range("C364:C453")

Set Sons1Range = Haushaltsbuch.Range("C459:C468")
Set Sons2Range = Haushaltsbuch.Range("C474:C483")
Set Sons3Range = Haushaltsbuch.Range("C489:C498")

Set Sons4Range = Haushaltsbuch.Range("C504:C593")
Set Sons5Range = Haushaltsbuch.Range("C599:C688")

'------------------------------------------------------------------------------------------------
'------- THE ENTRY WRITER ALGO --------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

Dim myRange As Range, tempRange As Range
Dim entryCell As Range

'Find out which category we are in via CatString comparison
Dim CatString As String
CatString = CatCell.Value

'Some logic bools to see if we are a 'daily' or 'added' or 'finite' category
Dim dailyBool As Boolean
dailyBool = False
Dim addedBool As Boolean
addedBool = False
Dim finiteBool As Boolean
finiteBool = False
'Some logic bools to see if we dont wanna write the entry or log (history)
Dim writtenBool As Boolean
writtenBool = False
Dim logBool As Boolean
logBool = False

'String comparison -> Categories -> Ranges
Select Case CatString
    Case Cat1Cell.Value
    'Anreise
         Set myRange = AnreiRange
         finiteBool = True
    Case Cat2Cell.Value
    'Inlandsflug
         Set myRange = InlanRange
         finiteBool = True
    Case Cat3Cell.Value
    'Fortbewegung
         Set myRange = FortbRange
         dailyBool = True
    Case Cat4Cell.Value
    'Touren&Aktivitäten
         Set myRange = ToureRange
         finiteBool = True
    Case Cat5Cell.Value
    'Unterkunft
         Set myRange = UnterRange
         dailyBool = True
    Case Cat6Cell.Value
    'Essen&Trinken
         Set myRange = EssenRange
         dailyBool = True
    Case Cat7Cell.Value
    'Visum
         Set myRange = VisumRange
         addedBool = True
    Case Cat8Cell.Value
    'Sim Karte
         Set myRange = SimkaRange
         addedBool = True
    Case Cat9Cell.Value
    'Vergnügen
         Set myRange = VergnRange
         dailyBool = True
    Case Sonst1Cell.Value
         Set myRange = Sons1Range
         finiteBool = True
    Case Sonst2Cell.Value
         Set myRange = Sons2Range
         finiteBool = True
    Case Sonst3Cell.Value
         Set myRange = Sons3Range
         finiteBool = True
    Case Sonst4Cell.Value
         Set myRange = Sons4Range
         dailyBool = True
    Case Sonst5Cell.Value
         Set myRange = Sons5Range
         dailyBool = True
    Case Else
         MsgBox _
         "Bitte wähle eine passende Kategorie aus dem Drop-Down Menü aus! Eintrag abgebrochen."
         writtenBool = True
    End Select

'Offset for different Regions
Set tempRange = myRange.Offset(0, offsetDistance * offsetInt)
Set myRange = tempRange

'------------------------------------------------------------------------------------------------
'------- DAILY CATEGORY --------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

If dailyBool = True And writtenBool = False Then
    If myRange.Item(1) = "-" Or myRange.Item(1) = " - " Or IsEmpty(myRange.Item(1)) Then
        ' if empty, make the first entry
        Dim answer As Integer
        answer = MsgBox("Ist dieses Datum: " & DatCell.Value & " dein erster Tag in " & RegCell.Value & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Erster täglicher Eintrag in " & RegCell.Value & "!")
        If answer = vbYes Then
            myRange.Item(1).Value = DatCell.Value
            myRange.Item(1).Offset(0, 2).Value = PriCell.Value
            'fill all cells according to maximum number of Days
            'for each range in dailyranges
            Dim laterDate As Date
            For i = 1 To maxDays
                laterDate = DateAdd("d", i - 1, DatCell.Value)
                FortbRange.Offset(0, offsetDistance * offsetInt).Item(i).Value = DateSerial(Year(laterDate), Month(laterDate), Day(laterDate))
            Next i
            For i = 1 To maxDays
                laterDate = DateAdd("d", i - 1, DatCell.Value)
                UnterRange.Offset(0, offsetDistance * offsetInt).Item(i).Value = DateSerial(Year(laterDate), Month(laterDate), Day(laterDate))
            Next i
            For i = 1 To maxDays
                laterDate = DateAdd("d", i - 1, DatCell.Value)
                EssenRange.Offset(0, offsetDistance * offsetInt).Item(i).Value = DateSerial(Year(laterDate), Month(laterDate), Day(laterDate))
            Next i
            For i = 1 To maxDays
                laterDate = DateAdd("d", i - 1, DatCell.Value)
                Sons4Range.Offset(0, offsetDistance * offsetInt).Item(i).Value = DateSerial(Year(laterDate), Month(laterDate), Day(laterDate))
            Next i
            For i = 1 To maxDays
                laterDate = DateAdd("d", i - 1, DatCell.Value)
                Sons5Range.Offset(0, offsetDistance * offsetInt).Item(i).Value = DateSerial(Year(laterDate), Month(laterDate), Day(laterDate))
            Next i
            MsgBox _
            "Erfolgreich ersten Eintrag in " & RegCell.Value & " mit Datum: " & DatCell.Value & " und Preis: " & PriCell.Value & "€ eingetragen und alle Daten ab dem ersten Reisetag eingefügt!"
            writtenBool = True
            logBool = True
        Else
            MsgBox "Bitte gib erst einen Eintrag mit deinem ersten Reisetag in " & RegCell.Value & " an. Eintrag abgebrochen."
            writtenBool = True
        End If
    End If
    ' if not empty, we find the fitting date entry and add price to it
    For Each entryCell In myRange
        If writtenBool = True Then
            Exit For
        ElseIf entryCell = DatCell.Value Then
            If entryCell.Offset(0, 2).Value > 0 Then
                entryCell.Value = DatCell.Value
                MsgBox _
                "Erfolgreich " & DatCell.Value & " " & PriCell.Value & "€ zu " & PriCell.Value + entryCell.Offset(0, 2).Value & "€ aufsummiert!"
                entryCell.Offset(0, 2).Value = entryCell.Offset(0, 2).Value + PriCell.Value
                writtenBool = True
                logBool = True
            Else
                entryCell.Value = DatCell.Value
                MsgBox _
                "Erfolgreich " & DatCell.Value & " " & PriCell.Value & "€ eingetragen!"
                entryCell.Offset(0, 2).Value = entryCell.Offset(0, 2).Value + PriCell.Value
                writtenBool = True
                logBool = True
            End If
            Exit For
    End If
    Next entryCell
    'If we still have not found any fitting date, some mistake has to be occured
    If writtenBool = False Then
        MsgBox _
        "Dieses Datum ist nicht mehr verfügbar in dieser Region. Überprüfe gerne noch einmal das Datum und den Anreisetag in " & RegCell & "."
        writtenBool = True
    End If
    
End If

'------------------------------------------------------------------------------------------------
'------- SUMMED CATEGORY --------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

If addedBool = True And writtenBool = False Then
    For Each entryCell In myRange
    If entryCell = "-" Or entryCell = " - " Or IsEmpty(entryCell) Then
        If writtenBool = False Then
            entryCell.Value = TxtCell.Value
            entryCell.Offset(0, 2).Value = PriCell.Value
            writtenBool = True
            logBool = True
            MsgBox _
            "Erfolgreich " & TxtCell.Value & " für " & PriCell.Value & "€ eingetragen!"
            Exit For
        End If
    End If
    ' For any other string we just add the price ontop
    If writtenBool = False Then
        'Function
        'entryCell.Value = TxtCell.Value
        MsgBox _
        "Erfolgreich " & TxtCell.Value & " " & PriCell.Value & " zu " & entryCell.Value & " " & PriCell.Value + entryCell.Offset(0, 2).Value & " aufsummiert!"
        entryCell.Offset(0, 2).Value = PriCell.Value + entryCell.Offset(0, 2).Value
        writtenBool = True
        logBool = True
        Exit For
    End If
    Next entryCell
End If

'------------------------------------------------------------------------------------------------
'------- FINITE CATEGORY --------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

If finiteBool = True Then
    For Each entryCell In myRange
        If writtenBool = True Then
            Exit For
        End If
        If entryCell = "-" Or entryCell = " - " Or IsEmpty(entryCell) Then
            If writtenBool = False Then
                entryCell.Value = TxtCell.Value
                entryCell.Offset(0, 2).Value = PriCell.Value
                writtenBool = True
                logBool = True
                MsgBox _
                "Erfolgreich " & TxtCell.Value & " " & PriCell.Value & " eingetragen!"
                Exit For
            End If
        End If
    Next entryCell
    
    'Only option left for writtenBool to be false is if no free cells in finite category
    If writtenBool = False Then
        Dim answer1 As Integer
        answer1 = MsgBox("Keine freien Felder mehr in der Kategorie: " & CatString & " verfügbar, möchtest du den Preis in dem letzten Feld aufsummieren?", vbQuestion + vbYesNo + vbDefaultButton2, CatString & " ist vollständig ausgefüllt!")
        If answer1 = vbYes Then
            myRange.End(xlDown).Value = myRange.End(xlDown).Value & " + " & TxtCell.Value
            myRange.End(xlDown).Offset(0, 2).Value = myRange.End(xlDown).Offset(0, 2).Value + PriCell.Value
            MsgBox myRange.End(xlDown).Value & " wurde erfolgreich zu " & myRange.End(xlDown).Offset(0, 2).Value & "€ aufsummiert!"
            writtenBool = True
            logBool = True
        Else
            MsgBox "Der Eintrag wurde abgebrochen."
            writtenBool = True
        End If
    End If
End If


'------------------------------------------------------------------------------------------------
'------- THE HISTORY LOG WRITER ALGO --------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

' assume the history is inside these corners
' "C14" --- "G14"
'   |         |
' "C54" --- "G54"
' "C55"clear"G55"

If logBool = True Then
    ' excel FIFO stack
    Eingabefeld.Range("C14:G54").Copy Range("C15")
    ' new entry in the top
    Eingabefeld.Range("C14").Value = DatCell.Value
    Eingabefeld.Range("D14").Value = RegCell.Value
    Eingabefeld.Range("E14").Value = CatCell.Value
    Eingabefeld.Range("F14").Value = TxtCell.Value
    Eingabefeld.Range("G14").Value = PriCell.Value
    ' clear the last entry when history is full
    Range("C55:G55").ClearContents
End If
End Sub


