VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub VBAChallenge()

For Each ws In Worksheets


   'definition of variables
   Dim Ticker As String
   Dim Yearly_Change As Long
   Dim Percent_Change As Long
   Dim Total_Stock_Volume As LongLong
   Dim Open_Value As Double
   Dim Close_Value As Double
   Dim Volume As Long
   Dim Summary_Table_Row As Integer
   Dim Last_Row As Long
   
   'Initial Values
   Summary_Table_Row = 2
   Volume = 0
   Total_Stock_Volume = 0
   Open_Value = ws.Cells(2, 3).Value
   Close_Value = 0
   Yearly_Change = 0
   Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
   'Headers
   ws.Cells(1, 9).Value = "Ticker"
   ws.Cells(1, 10).Value = "Yearly Change"
   ws.Cells(1, 11).Value = "Percent Change"
   ws.Cells(1, 12).Value = "Total Stock Volume"
   
   For i = 2 To Last_Row
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then Ticker = ws.Cells(i, 1).Value
           Volume = ws.Cells(i, 7).Value
           Total_Stock_Volume = Total_Stock_Volume + Volume
           
           ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
           ws.Range("I" & Summary_Table_Row).Value = Ticker
           
           Close_Value = ws.Cells(i, 6).Value
           Yearly_Change = Close_Value - Open_Value
           
           ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        Next i
           
           'Calculate Percent Change
            If Open_Value <> 0 Then
            Percent_Change = (Yearly_Change / Open_Value) - 1
            
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            
            Else
                ws.Range("K" & Summary_Table_Row).Value = 0
            End If
            
            
        'Conditional Formatting
        If Yearly_Change > 0 Then
          ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
          
        Else
          ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
          
          End If
          
          'Iterate and update Ticker
          
        Open_Value = ws.Cells(i + 1, 3).Value
        Summary_Table_Row = Summary_Table_Row + 1
        
      

    
 Next Each
 
   

   
End Sub

Private Sub Workbook_Open()

End Sub
