{\rtf1\ansi\ansicpg1252\cocoartf2761
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww20920\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub Challenge()\
\
    'Define Variables and Values\
    Dim total As Double\
    total = 0\
    Dim i As Long\
    Dim change As Double\
    change = 0\
    Dim j As Integer\
    j = 0\
    Dim start As Long\
    start = 2\
    Dim lastrow As Long\
    Dim percentchange As Double\
    \
    Dim dailychange As Double\
    Dim averagechange As Double\
\
\
    'Create Column Headings\
    Range("I1").Value = "Ticker"\
    Range("I1").Font.Bold = True\
    Range("J1").Value = "Quarterly_Change"\
    Range("J1").Font.Bold = True\
    Range("K1").Value = "Percent_Change"\
    Range("K1").Font.Bold = True\
    Range("L1").Value = "Total_Stock_Volume"\
    Range("L1").Font.Bold = True\
    Range("P1").Value = "Ticker"\
    Range("P1").Font.Bold = True\
    Range("Q1").Value = "Value"\
    Range("Q1").Font.Bold = True\
    Range("O2").Value = "Greatest_%_Increase"\
    Range("O2").Font.Bold = True\
    Range("O3").Value = "Greatest_%_Decrease"\
    Range("O3").Font.Bold = True\
    Range("O4").Value = "Greatest_Total_Volume"\
    Range("O4").Font.Bold = True\
    Range("O:O").EntireColumn.AutoFit\
    \
\
    'Define Last Row\
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row\
\
    'Start Loop\
    For i = 2 To lastrow\
\
       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then\
    \
            total = total + Cells(i, 7).Value\
        \
            If total = 0 Then\
                          \
                'Print the results in the Summary Table\
                Range("I" & 2 + j).Value = Cells(i, 1).Value\
                Range("J" & 2 + j).Value = 0\
                Range("K" & 2 + j).Value = "%" & 0\
                Range("L" & 2 + j).Value = 0\
            \
            Else\
                \
                If Cells(start, 3) = 0 Then\
                    For find_value = start To i\
                        If Cells(find_value, 3).Value <> 0 Then\
                            start = find_value\
                            Exit For\
                        End If\
                    Next find_value\
                End If\
                \
                change = (Cells(i, 6) - Cells(start, 3))\
                percentchange = change / Cells(start, 3)\
                \
                start = i + 1\
                \
                'Print the results in the Summary Table\
                Range("I" & 2 + j).Value = Cells(i, 1).Value\
                Range("J" & 2 + j).Value = change\
                Range("J" & 2 + j).NumberFormat = "0.00"\
                Range("K" & 2 + j).Value = percentchange\
                Range("K" & 2 + j).NumberFormat = "0.00%"\
                Range("L" & 2 + j).Value = total\
                Range("L:L").EntireColumn.AutoFit\
                \
                \
                \
                'Color Changes\
                If change > 0 Then\
                    Range("J" & 2 + j).Interior.ColorIndex = 4\
                End If\
                \
                If change < 0 Then\
                    Range("J" & 2 + j).Interior.ColorIndex = 3\
                End If\
                \
                If change = 0 Then\
                    Range("J" & 2 + j).Interior.ColorIndex = 0\
                End If\
            \
            End If\
            \
            total = 0\
            change = 0\
            j = j + 1\
            \
        Else\
            total = total + Cells(i, 7).Value\
            \
        End If\
        \
    Next i\
    \
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastrow)) * 100\
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & lastrow)) * 100\
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & lastrow))\
    \
    increaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)\
    decreaseNumber = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)\
    volumeNumber = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), 0)\
    \
    Range("P2") = Cells(increaseNumber + 1, 9)\
    Range("P3") = Cells(decreaseNumber + 1, 9)\
    Range("P4") = Cells(volumeNumber + 1, 9)\
    \
    \
    \
    \
End Sub}