Sub BedingteFormatierungHinzu()

    ' This program is free software: you can redistribute it and/or modify
    ' it under the terms of the GNU General Public License as published by
    ' the Free Software Foundation, either version 3 of the License, or
    ' (at your option) any later version.
    '
    ' This program is distributed in the hope that it will be useful,
    ' but WITHOUT ANY WARRANTY; without even the implied warranty of
    ' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    ' GNU General Public License for more details.
    '
    ' You should have received a copy of the GNU General Public License
    ' along with this program. If not, see <http://www.gnu.org/licenses/>.
    '
    ' Date: 26.04.2019
    ' Autor: Peter Siemer
    
    ' Update: 29.04.2019
    ' - "Elementänderungsstatus" überprüfen auf nicht freigegebene Artikel.
    ' - PosSpalte prüfen auf "Kleiner als Vorgänger"
    
    ' Update: 24.05.2019
    ' - "Strukturtyp" überprüfen auf einzelne Ebenen
    ' - TYP, HBG, MBG
    
    Dim Spaltenbeschriftung As Range: Set Spaltenbeschriftung = Application.Range("A1:Z1")
    
    Dim ArtikelnummerSpalte As String
    ArtikelnummerSpalte = "-1"
    
    Dim ElementaenderungsstatusSpalte As String
    ElementaenderungsstatusSpalte = "-1"
    
    Dim PosSpalte As String
    PosSpalte = "-1"
    
    Dim StrukturtypSpalte As String
    StrukturtypSpalte = "-1"
    
    Dim LetzteZeile As Integer
    LetzteZeile = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    '
    'Spalte mit Artikelnummer und Pos finden
    '
    
    Dim zelle As Range
    For Each zelle In Spaltenbeschriftung.Cells
    
        If zelle.Text = "Artikelnummer" Then
            ArtikelnummerSpalte = Split(zelle.Address, "$")(1)
        End If
        
        If zelle.Text = "Elementänderungsstatus" Then
            ElementaenderungsstatusSpalte = Split(zelle.Address, "$")(1)
        End If
        
        If zelle.Text = "Pos." Then
            PosSpalte = Split(zelle.Address, "$")(1)
        End If
        
        If zelle.Text = "Strukturtyp" Then
            StrukturtypSpalte = Split(zelle.Address, "$")(1)
        End If
        
    Next zelle
    
    If ArtikelnummerSpalte = "-1" Or ElementaenderungsstatusSpalte = "-1" Or PosSpalte = "-1" Or StrukturtypSpalte = "-1" Then
        MsgBox "Die Spalten 'Artikelnummer', Elementänderungsstatus und 'Pos.' müssen mit exportiert werden"
        Exit Sub
    End If
    
    
    '
    'Bedinge Formatierungen löschen
    '
    
    Range("$1:$" & LetzteZeile).FormatConditions.Delete
    
    '
    'Bedinge Formatierungen einfügen
    '

    'Ganze Zeilen markieren
    With Range("$2:$" & LetzteZeile)
    
        'Klammerbaugruppen markieren
        '=$D2="000.90000"
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$" & ArtikelnummerSpalte & "2=" & Chr(34) & "000.90000" & Chr(34)
        .FormatConditions(Range("$2:$" & LetzteZeile).FormatConditions.Count).Interior.ColorIndex = 4
        .FormatConditions(Range("$2:$" & LetzteZeile).FormatConditions.Count).StopIfTrue = False
                
        'SPI markieren
        '=LINKS($D2;3)="SPI"
        .FormatConditions.Add Type:=xlExpression, Formula1:="=links($" & ArtikelnummerSpalte & "2;3)=" & Chr(34) & "SPI" & Chr(34)
        .FormatConditions(Range("$2:$" & LetzteZeile).FormatConditions.Count).Interior.ColorIndex = 7
        .FormatConditions(Range("$2:$" & LetzteZeile).FormatConditions.Count).StopIfTrue = False
        
        'SPL markieren
        '=LINKS($D2;3)="SPL"
        .FormatConditions.Add Type:=xlExpression, Formula1:="=links($" & ArtikelnummerSpalte & "2;3)=" & Chr(34) & "SPL" & Chr(34)
        .FormatConditions(Range("$2:$" & LetzteZeile).FormatConditions.Count).Interior.ColorIndex = 13
        .FormatConditions(Range("$2:$" & LetzteZeile).FormatConditions.Count).StopIfTrue = False
        
        'HBG markieren
        '=$G2="TYP"
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$" & StrukturtypSpalte & "2=" & Chr(34) & "TYP" & Chr(34)
        .FormatConditions(Range("$2:$" & LetzteZeile).FormatConditions.Count).Interior.Color = RGB(0, 200, 0)
        .FormatConditions(Range("$2:$" & LetzteZeile).FormatConditions.Count).StopIfTrue = False
        
        'HBG markieren
        '=$G2="HBG"
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$" & StrukturtypSpalte & "2=" & Chr(34) & "HBG" & Chr(34)
        .FormatConditions(Range("$2:$" & LetzteZeile).FormatConditions.Count).Interior.Color = RGB(0, 150, 0)
        .FormatConditions(Range("$2:$" & LetzteZeile).FormatConditions.Count).StopIfTrue = False
        
        'MBG markieren
        '=$G2="MBG"
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$" & StrukturtypSpalte & "2=" & Chr(34) & "MBG" & Chr(34)
        .FormatConditions(Range("$2:$" & LetzteZeile).FormatConditions.Count).Interior.Color = RGB(0, 100, 0)
        .FormatConditions(Range("$2:$" & LetzteZeile).FormatConditions.Count).StopIfTrue = False
        
    End With
    
    'Innerhalb der PosSpalte
    With Range("$" & PosSpalte & "$2:$" & PosSpalte & LetzteZeile)
        
        'Leere Pos markieren
        '=$H2=""
        '.FormatConditions.Add Type:=xlExpression, Formula1:="=UND($" & ArtikelnummerSpalte & "2<>" & Chr(34) & Chr(34) & ";$" & PosSpalte & "2=" & Chr(34) & Chr(34) & ")"
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$" & PosSpalte & "2=" & Chr(34) & Chr(34)
        .FormatConditions(Range("$" & PosSpalte & "$2:$" & PosSpalte & LetzteZeile).FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Interior.ColorIndex = 6
        .FormatConditions(1).StopIfTrue = False
        
        'Kleiner als Vorgänger
        '=ZAHLENWERT($H2)>ZAHLENWERT($H3)
        .FormatConditions.Add Type:=xlExpression, Formula1:="=ZAHLENWERT($" & PosSpalte & "1)>ZAHLENWERT($" & PosSpalte & "2)"
        .FormatConditions(Range("$" & PosSpalte & "$2:$" & PosSpalte & LetzteZeile).FormatConditions.Count).Interior.ColorIndex = 8
        .FormatConditions(Range("$" & PosSpalte & "$2:$" & PosSpalte & LetzteZeile).FormatConditions.Count).StopIfTrue = False
    
        'Doppelte Pos markieren
        .FormatConditions.AddUniqueValues
        .FormatConditions(Range("$" & PosSpalte & "$2:$" & PosSpalte & LetzteZeile).FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1)
            .DupeUnique = xlDuplicate
            .Font.Color = -16383844
            .Font.TintAndShade = 0
            .Interior.PatternColorIndex = xlAutomatic
            .Interior.Color = 13551615
            .Interior.TintAndShade = 0
            .StopIfTrue = False
        End With
        
    End With
    
    'Innerhalb der ElementaenderungsstatusSpalte
    With Range("$" & ElementaenderungsstatusSpalte & "$2:$" & ElementaenderungsstatusSpalte & LetzteZeile)
    
        'Nicht freigegeben
        '=UND(RECHTS(LINKS($F2;2);1)<>"F";RECHTS(LINKS($F2;3);1)<>"F";$F2<>"Veraltet")
        .FormatConditions.Add Type:=xlExpression, Formula1:="=UND(RECHTS(LINKS($" & ElementaenderungsstatusSpalte & "2;2);1)<>" & Chr(34) & "F" & Chr(34) & ";RECHTS(LINKS($" & ElementaenderungsstatusSpalte & "2;3);1)<>" & Chr(34) & "F" & Chr(34) & ";$" & ElementaenderungsstatusSpalte & "2<>" & Chr(34) & "Veraltet" & Chr(34) & ")"
        .FormatConditions(Range("$" & ElementaenderungsstatusSpalte & "$2:$" & ElementaenderungsstatusSpalte & LetzteZeile).FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Interior.ColorIndex = 6
        .FormatConditions(1).StopIfTrue = False
        
    End With

End Sub
