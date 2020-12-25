Attribute VB_Name = "Module1"
Sub Main()
    
    Dim NumWS As Integer

    'Get number of Worksheets in Workbook
    NumWS = ThisWorkbook.Sheets.Count

    'Iterate through all existing worksheets to create and fill out summary tables
    For i = 1 To NumWS
    
        'Activate Worksheet for wich of subsequent operations will happen
        Worksheets(i).Activate
        
        'Call Macro To Create Table Headers
        CreateTableHeaders
        
        'Call Macro To Create Summary Table One
        CreateSummaryTableOne
    
        'Call Macro To Create Summary Table Two
        CreateSummaryTableTwo
        
        'Add Formats to Both Summary Tables
        FormatTables
    
    Next i

End Sub

Sub CreateTableHeaders()
'MACRO: Create Headers For Summary Tables
    
    'Declaration of Variables
    Dim HeaderTitles As Variant
    
    'Initialize Array of Strings for Headers of Summary Tables
    HeaderTitles = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", "Value", _
                                     "Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")

    'Create Headers for Summary Table 1
    Range("I1").Select
    
    For j = 0 To 3
    
        ActiveCell.Value = HeaderTitles(j)
        ActiveCell.Offset(0, 1).Select
    
    Next j
    
    'Create Headers for Summary Table 2
    Range("O2").Select
        
    For j = 5 To 7
    
        ActiveCell.Value = HeaderTitles(j)
        ActiveCell.Offset(1, 0).Select
        
    Next j
    
    Range("P1").Value = HeaderTitles(0)
    Range("Q1").Value = HeaderTitles(4)

End Sub

Sub CreateSummaryTableOne()
'MACRO: Create Values for Summary Table 1
    
    'Declare Variables
    Dim LastRow As Long
    Dim TotalVolume As LongLong
    Dim Flag As Integer
    Dim TOp As Double
    Dim TCls As Double
    Dim TCount As Integer
    
    'Get Number of Rows in the Active Worksheet
    Range("A1").Select
    LastRow = ActiveCell.End(xlDown).Row

  
    'Initialize Ticker Counter and First Iteration Flag
    TCount = 2  'Counter to Identify the Row  for New Tickers
    Flag = 0       'Flag to Identify the First Row of a New Ticker
    
    'Fill Out Values for Summary Table 1
    For r = 2 To LastRow
    
        'Identify Ticker Name Changes and Calculate Values
        If Cells(r, 1).Value <> Cells(r + 1, 1) Then
        
            TCls = Cells(r, 6).Value 'Get Value of Ticker Close
            Cells(TCount, 9).Value = Cells(r, 1).Value 'Get Name of Current Ticker
            Cells(TCount, 10).Value = TCls - TOp 'Calculate Yearly Change
            Cells(TCount, 12).Value = TotalVolume 'Get Total Volume for Current Ticker
            
            'Calculate % Change and Handle Division by Zero
            If TOp = 0 Then
                
                Cells(TCount, 11).Value = 0
                    
            Else
                
                Cells(TCount, 11).Value = (TCls - TOp) / TOp
            
            End If
            
            'Reinitialize Variables For Next Ticker
            TCount = TCount + 1
            Flag = 0
            TotalVolume = 0
            TOp = 0
            TCls = 0
            
        'If There is No Ticker Name Change, Accumulate Total Volume
        Else
        
            TotalVolume = TotalVolume + Cells(r, 7).Value
            Flag = Flag + 1
            
            'Identify Row for a New Ticker to Record Ticker Open Value
            If Flag = 1 Then
            
                TOp = Cells(r, 3).Value
            
            End If
        
        End If
    
    Next r

End Sub

Sub CreateSummaryTableTwo()
'MACRO: Create Values for Summary Table 2
   
    'Declare Variables
    Dim MaxRow As Long
    Dim MinRow As Long
    Dim MaxVolRow As Long
    
    'Calculate Greatest Percentage Increase
    Range("Q2").Value = Application.WorksheetFunction.Max(Columns("K"))
    MaxRow = Application.WorksheetFunction.Match(Range("Q2").Value, Columns("K"), 0)
    Range("P2").Value = Cells(MaxRow, 9).Value
    
    'Calculate Greatest Percentage Decrease
    Range("Q3").Value = Application.WorksheetFunction.Min(Columns("K"))
    MinRow = Application.WorksheetFunction.Match(Range("Q3").Value, Columns("K"), 0)
    Range("P3").Value = Cells(MinRow, 9).Value
    
    'Calculate Greatest Total Volume
    Range("Q4").Value = Application.WorksheetFunction.Max(Columns("L"))
    MaxVolRow = Application.WorksheetFunction.Match(Range("Q4").Value, Columns("L"), 0)
    Range("P4").Value = Cells(MaxVolRow, 9).Value

End Sub

Sub FormatTables()
'MACRO: Add Formating To Summary Tables

    'Declare Variables
    Dim TLastRow As Integer
    
    'Get Last Row of Summary Table Onw
    Range("J1").Select
    TLastRow = ActiveCell.End(xlDown).Row
    
    'Implement Conditional Formating for Yearly Change
    For k = 2 To TLastRow
    
        If k Mod 2 = 0 Then
        
            Range(Cells(k, 9), Cells(k, 12)).Interior.Color = RGB(232, 232, 232)
        
        End If
        
        If Cells(k, 10).Value >= 0 Then
            
            Cells(k, 10).Interior.Color = RGB(80, 220, 100)
            
        Else
            
            Cells(k, 10).Interior.Color = RGB(250, 128, 114)
            
        End If
    
    Next k
    
    'Format Percent Ranges As "0.00%"
    Columns("K").NumberFormat = "0.00%"
    Range("Q2:Q3").NumberFormat = "0.00%"
    
    'Format Cells Backgrounds
    Range("A1:G1").Interior.Color = RGB(0, 128, 255)
    Range("A1:G1").Font.Color = vbWhite
    Range("A1:G1").Font.Name = "ArialBlack"
    
    Range("I1:L1").Interior.Color = RGB(0, 128, 255)
    Range("I1:L1").Font.Color = vbWhite
    Range("I1:L1").Font.Name = "ArialBlack"
    
    Range("P1:Q1").Interior.Color = RGB(0, 128, 255)
    Range("P1:Q1").Font.Color = vbWhite
    Range("P1:Q1").Font.Name = "ArialBlack"
    
    Range("O2:O4").Interior.Color = RGB(0, 128, 255)
    Range("O2:O4").Font.Color = vbWhite
    Range("O2:O4").Font.Name = "ArialBlack"
    
    'AutoFit Headers
    Columns("A:Q").AutoFit

End Sub


