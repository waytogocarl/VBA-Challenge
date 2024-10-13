{\rtf1\ansi\ansicpg1252\cocoartf2761
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub Challenge()\
\
    'Run Code Through all Worksheets\
    Dim ws As Worksheet\
    For Each ws In ThisWorkbook.Worksheets\
        ws.Activate\
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
        ws.Range("I1").Value = "Ticker"\
        ws.Range("I1").Font.Bold = True\
        ws.Range("J1").Value = "Quarterly_Change"\
        ws.Range("J1").Font.Bold = True\
        ws.Range("K1").Value = "Percent_Change"\
        ws.Range("K1").Font.Bold = True\
        ws.Range("L1").Value = "Total_Stock_Volume"\
        ws.Range("L1").Font.Bold = True\
        ws.Range("P1").Value = "Ticker"\
        ws.Range("P1").Font.Bold = True\
        ws.Range("Q1").Value = "Value"\
        ws.Range("Q1").Font.Bold = True\
        ws.Range("O2").Value = "Greatest_%_Increase"\
        ws.Range("O2").Font.Bold = True\
        ws.Range("O3").Value = "Greatest_%_Decrease"\
        ws.Range("O3").Font.Bold = True\
        ws.Range("O4").Value = "Greatest_Total_Volume"\
        ws.Range("O4").Font.Bold = True\
        ws.Range("O:O").EntireColumn.AutoFit\
    \
\
        'Define Last Row\
        lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row\
\
        'Start Loop\
        For i = 2 To lastrow\
\
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then\
    \
                total = total + ws.Cells(i, 7).Value\
        \
                If total = 0 Then\
                          \
                    'Print the results in the Summary Table\
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value\
                    ws.Range("J" & 2 + j).Value = 0\
                    ws.Range("K" & 2 + j).Value = "%" & 0\
                    ws.Range("L" & 2 + j).Value = 0\
            \
                Else\
                \
                    If ws.Cells(start, 3) = 0 Then\
                        For find_value = start To i\
                            If ws.Cells(find_value, 3).Value <> 0 Then\
                                start = find_value\
                                Exit For\
                            End If\
                        Next find_value\
                    End If\
                \
                    change = (ws.Cells(i, 6) - ws.Cells(start, 3))\
                    percentchange = change / ws.Cells(start, 3)\
                \
                    start = i + 1\
                \
                    'Print the results in the Summary Table\
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value\
                    ws.Range("J" & 2 + j).Value = change\
                    ws.Range("J" & 2 + j).NumberFormat = "0.00"\
                    ws.Range("K" & 2 + j).Value = percentchange\
                    ws.Range("K" & 2 + j).NumberFormat = "0.00%"\
                    ws.Range("L" & 2 + j).Value = total\
                    ws.Range("J:J").EntireColumn.AutoFit\
                    ws.Range("K:K").EntireColumn.AutoFit\
                    ws.Range("L:L").EntireColumn.AutoFit\
                \
                \
                \
                    'Color Changes\
                    If change > 0 Then\
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 4\
                    End If\
                \
                    If change < 0 Then\
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 3\
                    End If\
                \
                    If change = 0 Then\
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 0\
                    End If\
            \
                End If\
            \
                total = 0\
                change = 0\
                j = j + 1\
            \
            Else\
                total = total + ws.Cells(i, 7).Value\
            \
            End If\
        \
        Next i\
    \
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) * 100\
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastrow)) * 100\
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))\
    \
        increaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)\
        decreaseNumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)\
        volumeNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)\
    \
        ws.Range("P2") = ws.Cells(increaseNumber + 1, 9)\
        ws.Range("P3") = ws.Cells(decreaseNumber + 1, 9)\
        ws.Range("P4") = ws.Cells(volumeNumber + 1, 9)\
    \
    \
    Next ws\
    \
End Sub}