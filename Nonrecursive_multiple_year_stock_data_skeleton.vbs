{\rtf1\ansi\ansicpg1252\cocoartf2761
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub Stock_analysis()\
\
\
'-- Assigning names for column headers\
    Range("I1").Value = "Ticker"\
    Range("J1").Value = "Quarterly Change"\
    Range("K1").Value = "Percent Change"\
    Range("L1").Value = "Total Stock Volume"\
    Range("P1").Value = "Ticker"\
    Range("Q1").Value = "Value"\
    Range("O2").Value = "Greatest % Increase"\
    Range("O3").Value = "Greatest % Decrease"\
    Range("O4").Value = "Greatest Total Volume"\
\
\
'-- Setting Values from headers\
        Dim total As Double\
        Dim i As Long\
        Dim change As Double\
        Dim j As Integer\
        Dim start As Long\
        Dim rowCount As Long\
        Dim percentChange As Double\
        Dim dailyChange As Double\
        Dim averageChange As Double\
        \
    '--Set initial values\
        j = 0\
        total = 0\
        change = 0\
        start = 2\
        \
    '--Now we have to get the last row containing data in the worksheet\
        rowCount = Cells(Rows.Count, "A").End(xlUp).Row\
        \
        For i = 2 To rowCount\
        \
    '--Determining when ticker changes to print the results\
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then\
        \
    '--Collect results in variables\
        total = total + Cells(i, 7).Value\
        \
    '--We also have to check for non-divisibility since this is a percentage calculation\
        If total = 0 Then\
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value\
                ws.Range("J" & 2 + j).Value = 0\
                ws.Range("K" & 2 + j).Value = "%" & 0\
                ws.Range("L" & 2 + j).Value = 0\
        \
        Else\
            '--What is the first value after 0?\
            If Cells(start, 3) = 0 Then\
                For find_value = start To i\
                    If ws.Cells(find_value, 3).Value <> 0 Then\
                         start = find_value\
                            Exit For\
                         End If\
                     Next find_value\
            End If\
\
    '--Next step is to calculate the change\
        change = (Cells(i, 6) - Cells(start, 3))\
        percentChange = change / Cells(start, 3)\
\
    '--Start of the next stock ticker\
    start = i + 1\
    \
    '--Printing the results\
    \
    Range("I" & 2 + j).Value = Cells(i, 1).Value\
    Range("J" & 2 + j).Value = change\
    Range("J" & 2 + j).NumberFormat = "0.00"\
    Range("K" & 2 + j).Value = percentChange\
    Range("K" & 2 + j).NumberFormat = "0.00%"\
    Range("L" & 2 + j).Value = total\
\
    '--Conditional formatting shows green when result is positive and red when negative\
        Select Case change\
                Case Is > 0\
                    Range("J" & 2 + j).Interior.ColorIndex = 4\
                Case Is < 0\
                    Range("J" & 2 + j).Interior.ColorIndex = 3\
                Case Else\
                    Range("J" & 2 + j).Interior.ColorIndex = 0\
        End Select\
\
End If\
\
'--Reset variables when starting new ticker\
    total = 0\
    change = 0\
    j = j + 1\
    'days = 0\
    \
'--If ticker name is still the same, keep adding the results\
    Else\
        total = total + Cells(i, 7).Value\
    End If\
    \
    Next i\
'--Setting the maximum and minimum values\
        Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100\
        Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100\
        Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowCount))\
\
'--Deducting header row\
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)\
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)\
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0)\
\
 '--Final ticker symbol for total, greatest % of increase and decrease, and average\
    Range("P2") = Cells(increase_number + 1, 9)\
    Range("P3") = Cells(decrease_number + 1, 9)\
    Range("P4") = Cells(volume_number + 1, 9)\
\
\
End Sub\
}