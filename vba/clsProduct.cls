VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1

Private Sub btn_Click()
    

    Dim Data As Variant
    
    Data = Split(btn.Tag, "|")
    
   
    Dim name As String
    Dim price As Double
    Dim found As Boolean
    Dim i As Long
    Dim total As Double
    Dim qty As Long
    name = Data(0)
price = CDbl(Replace(Data(1), ",", Application.DecimalSeparator))
    qty = 0
    total = 0

    With frmPOS.lstOrder

    For i = 0 To .ListCount - 1
    
     
        If .List(i, 0) = name Then
         .List(i, 1) = .List(i, 1) + 1
         .List(i, 2) = .List(i, 2) + price
          
          
            found = True
    Exit For
    End If
  
   Next i
    
   If found = False Then
    .AddItem name
    .List(.ListCount - 1, 0) = name
    .List(.ListCount - 1, 1) = 1
    .List(.ListCount - 1, 2) = CDbl(Replace(price, ".", Application.DecimalSeparator))
    
     

    End If
   For i = 0 To .ListCount - 1
    total = total + CDbl(Replace(.List(i, 2), ".", Application.DecimalSeparator))
    qty = qty + CLng(.List(i, 1))
Next i
     frmPOS.lblAmount.Caption = Format(total, "0.00")
    frmPOS.txtPrice.Text = "$ " & Format(total, "0.00")
     frmPOS.lblQI.Caption = qty
 
End With
End Sub
'========================================================
' Project: Coffee Shop POS
' Author: Mariano Ferrer
' Role: Excel VBA Developer
' Date: 2026
'
' Description:
' Point of Sale (POS) system developed in Excel using VBA.
' Includes product management, order processing, automatic
' calculations, receipt generation and printable tickets.
'
' GitHub: https://github.com/cracka07
'========================================================
