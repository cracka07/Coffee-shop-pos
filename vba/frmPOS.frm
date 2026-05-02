VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPOS 
   Caption         =   "Coffee Shop"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16665
   OleObjectBlob   =   "frmPOS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim buttom_btn As Collection

Private Sub btnBakery_Click()
        Dim lastRow As Long
        Dim ws As Worksheet
        Dim i As Long
        Dim btn As MSForms.CommandButton
        Dim topPos As Long
        Dim leftPos As Long
        Dim evt As clsProduct
        Dim displayName As String
        Set ws = Sheets("Product")
         

        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        
        topPos = 50
        leftPos = 10
        Dim name As String
        Call LimpiarFrame
        For i = 3 To lastRow
        
            If ws.Cells(i, 2).Value = "Bakery" Then
            
           
                Set btn = Me.fraProduct.Controls.Add("Forms.CommandButton.1")
                
                 With btn
                    name = ws.Cells(i, 3).Value
                   displayName = name
                    .Width = 88
                    .Height = 105
                    
                    .Left = leftPos
                    .Top = topPos
                    If Len(displayName) > 12 Then
                        displayName = Replace(displayName, " ", vbCrLf, 1, 1)
                        End If
                    .Caption = displayName & vbCrLf & "$" & ws.Cells(i, 4).Value
                    .Tag = name & "|" & ws.Cells(i, 4).Value
        '
                    
                    .Picture = LoadPicture("D:\Image_coffee_shop\bakery\" & ws.Cells(i, 5).Value & ".jpg")
                    
                    .PicturePosition = fmPicturePositionAboveCenter
                    
                    Set evt = New clsProduct
                    Set evt.btn = btn
                    buttom_btn.Add evt
                    
                    
                    .BackColor = vbBlack
                    .ForeColor = vbWhite
                    .Font.Size = 10
                    .Font.name = "Segoe UI Black"
                End With
                
                
                leftPos = leftPos + 92
                
                If leftPos > 280 Then
                
                    leftPos = 10
                    topPos = topPos + 110
                    
                End If
                
            End If
        
        Next i
        
        With Me.fraProduct
        
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = topPos + 120
        
        End With
        
        Me.lblCate.Caption = "BAKERY"
End Sub

Private Sub btnBakery_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
            Me.btnBakery.BackColor = RGB(171, 142, 79)
            Me.imgBakery.Picture = LoadPicture("D:\Image_coffee_shop\br_bakery.jpg")
End Sub

Private Sub btnBakery_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Me.btnBakery.BackColor = vbBlack
         Me.imgBakery.Picture = LoadPicture("D:\Image_coffee_shop\bakery.jpg")
End Sub

Private Sub btnCancel_Click()
        
       Me.txtReceived.Value = "0"
       Me.txtChange.Value = "0"

End Sub
Private Sub btnCash_Click()

        Me.txtReceived.Enabled = True
        Me.txtChange.Enabled = True
        Me.btnCancel.Enabled = True
        Me.btnPay.Enabled = True
End Sub

Private Sub btnCash_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
                Me.btnCash.BackColor = RGB(171, 142, 79)
                Me.imgCash.Picture = LoadPicture("D:\Image_coffee_shop\br_cash.jpg")
End Sub

Private Sub btnCash_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
             Me.btnCash.BackColor = vbBlack
               Me.imgCash.Picture = LoadPicture("D:\Image_coffee_shop\cash.jpg")
End Sub

Private Sub btnClr_Click()
    Dim i As Long

    Range("G6").Value = ""
    Range("G7").Value = ""
    Range("F9:H1000").Value = ""
    
    Me.lblAmount = 0
    Me.lblQI = 0
    Me.txtPrice.Value = 0
    Me.lstOrder.Clear
End Sub

Private Sub btnCoffee_Click()
    
        Dim lastRow As Long
        Dim ws As Worksheet
        Dim i As Long
        Dim btn As MSForms.CommandButton
        Dim evt As clsProduct
        Dim topPos As Long
        Dim leftPos As Long
        Dim name As String
        Dim quantity As Long
        
        Set ws = Sheets("Product")
             
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        
        topPos = 50
        leftPos = 10
      
        
        For i = 3 To lastRow
        
            If ws.Cells(i, 2).Value = "Coffee" Then
            
           
                Set btn = Me.fraProduct.Controls.Add("Forms.CommandButton.1")
                
                 With btn
                    name = ws.Cells(i, 3).Value
                    
                    .Width = 88
                    .Height = 99
                    
                    .Left = leftPos
                    .Top = topPos
                    If Len(name) > 12 Then
                        name = Replace(name, " ", vbCrLf, 1, 1)
                        End If
                    .Caption = name & vbCrLf & "$" & ws.Cells(i, 4).Value
                    .Tag = name & "|" & ws.Cells(i, 4).Value
                   
                    
                    .Picture = LoadPicture("D:\Image_coffee_shop\images\" & ws.Cells(i, 5).Value & ".jpg")
                    
                    .PicturePosition = fmPicturePositionAboveCenter
                      
                    Set evt = New clsProduct
                    Set evt.btn = btn
                    buttom_btn.Add evt
                    
                    
                    .BackColor = vbBlack
                    .ForeColor = vbWhite
                    .Font.Size = 10
                    .Font.name = "Segoe UI Black"
                End With
          
                leftPos = leftPos + 92
                
                If leftPos > 280 Then
                
                    leftPos = 10
                    topPos = topPos + 100
                    
                End If
                
            End If
        
        Next i
        
        With Me.fraProduct
           .ScrollBars = fmScrollBarsVertical
           .ScrollHeight = topPos + 120
        End With
        
        Me.lblCate.Caption = "COFFEE"
End Sub

Private Sub btnCoffee_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
                Me.btnCoffee.BackColor = RGB(171, 142, 79)
                Me.imgCoffee.Picture = LoadPicture("D:\Image_coffee_shop\br_coffee.jpg")
End Sub

Private Sub btnCoffee_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
            Me.btnCoffee.BackColor = vbBlack
            Me.imgCoffee.Picture = LoadPicture("D:\Image_coffee_shop\coffee.jpg")
End Sub

Private Sub btnPay_Click()
        Dim row As Long
        Dim i As Long
        Dim ws As Worksheet
        Dim ID As Range
        Dim product As Range
        Dim category As Range
        
        Set ws = Worksheets("Sales")
        Set product = Sheets("Product").Range("tblProductq[PRODUCT]")
        Set ID = Sheets("Product").Range("tblProductq[ID]")
        Set category = Sheets("Product").Range("tblProductq[CATEGORY]")
        row = 3
        With Me.lstOrder
        

        For i = 0 To .ListCount - 1
              
            Do While ws.Cells(row, 1) <> ""
                    row = row + 1
            Loop
             ws.Cells(row, 1) = Me.txtOrderNumber.Value
             ws.Cells(row, 2) = Me.lblTime.Caption
              ws.Cells(row, 5) = .List(i, 0)
              ws.Cells(row, 6) = .List(i, 1)
              ws.Cells(row, 7) = .List(i, 2)
              ws.Cells(row, 8) = CDbl(Me.lblAmount.Caption)
              ws.Cells(row, 3).Value = Application.WorksheetFunction.XLookup( _
                                                  ws.Cells(row, 5).Value, product, ID)
              ws.Cells(row, 4) = Application.WorksheetFunction.XLookup( _
                                                  ws.Cells(row, 5).Value, product, category)
              
        Next i
    
        End With
       
        MsgBox "Successful Payment", vbInformation, "The Passenger Coffee"
        
        Me.lstOrder.Clear
        Me.lblAmount.Caption = 0
        Me.lblQI.Caption = 0
        
       
             If Application.WorksheetFunction.Count(ws.Range("A:A")) = 0 Then
                    Me.txtOrderNumber.Value = 1
        Else
                   Me.txtOrderNumber.Value = Application.WorksheetFunction.Max(ws.Range("A:A")) + 1
      
      End If
       
    
End Sub

Private Sub btnPayment_Click()

         Me.txtReceived.Enabled = False
         Me.txtChange.Enabled = False
         Me.btnCancel.Enabled = False
         
End Sub

Private Sub btnPayment_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
            Me.btnPayment.BackColor = RGB(171, 142, 79)
            Me.imgPayment.Picture = LoadPicture("D:\Image_coffee_shop\br_card.jpg")
End Sub

Private Sub btnPayment_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
             Me.btnPayment.BackColor = vbBlack
            Me.imgPayment.Picture = LoadPicture("D:\Image_coffee_shop\card.jpg")
End Sub

Private Sub btnPrint_Click()
            
            Dim i As Long
            Dim row As Long
            Dim ws As Worksheet
            Set ws = Worksheets("Ticket")
            
            row = 9
            
              ws.Range("G6").Value = Me.lblTime.Caption
              ws.Range("G7").Value = Me.txtOrderNumber.Value
            
            With Me.lstOrder
                
              
                For i = 0 To .ListCount - 1
                
                    ws.Cells(row, 6).Value = .List(i, 0)
                    ws.Cells(row, 7).Value = .List(i, 1)
                    ws.Cells(row, 8).Value = .List(i, 2)
                    row = row + 1
                Next i
            End With
            
            With ws.Cells(row, 6)
                    .Value = "----------------------------------------"
                    .Font.Bold = True
            End With
            
              With ws.Cells(row + 2, 7)
                    .Value = "Thank you for your visit"
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter
            End With
            
            ws.PageSetup.PrintArea = "F3:H14"
            ws.PrintOut
            
            MsgBox "Printed Ticket", vbInformation, "The Passenger Coffee"
            
End Sub
Private Sub btnPrint_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
                Me.btnPrint.BackColor = RGB(171, 142, 79)
                Me.imgPrint.Picture = LoadPicture("D:\Image_coffee_shop\br_ticket.jpg")
End Sub

Private Sub btnPrint_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
                Me.btnPrint.BackColor = vbBlack
                Me.imgPrint.Picture = LoadPicture("D:\Image_coffee_shop\print.jpg")
End Sub

Private Sub btnTea_Click()

        Dim lastRow As Long
        Dim ws As Worksheet
        Dim i As Long
        Dim btn As MSForms.CommandButton
        Dim topPos As Long
        Dim leftPos As Long
        Dim name As String
         Dim evt As clsProduct
        Dim displayName As String
       
        Set ws = Sheets("Product")


        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        
        topPos = 50
        leftPos = 10
        
        Call LimpiarFrame
        For i = 3 To lastRow
        
            If ws.Cells(i, 2).Value = "Tea" Then
            
           
                Set btn = Me.fraProduct.Controls.Add("Forms.CommandButton.1")
                
                 With btn
                    name = ws.Cells(i, 3).Value
                    displayName = name
                    .Width = 88
                    .Height = 99
                    
                    .Left = leftPos
                    .Top = topPos
                    If Len(displayName) > 10 Then
                        displayName = Replace(displayName, " ", vbCrLf, 1, 1)
                        End If
                    .Caption = displayName & vbCrLf & "$" & ws.Cells(i, 4).Value
                    .Tag = name & "|" & ws.Cells(i, 4).Value
        
                    
                    .Picture = LoadPicture("D:\Image_coffee_shop\tea\" & ws.Cells(i, 5).Value & ".jpg")
                    
                    .PicturePosition = fmPicturePositionAboveCenter
                    
                    
                     Set evt = New clsProduct
                    Set evt.btn = btn
                    buttom_btn.Add evt
                    
                    .BackColor = vbBlack
                    .ForeColor = vbWhite
                    .Font.Size = 10
                    .Font.name = "Segoe UI Black"
                End With
          
                
                leftPos = leftPos + 92
                
                If leftPos > 280 Then
                
                    leftPos = 10
                    topPos = topPos + 100
                    
                End If
                
            End If
        
        Next i
            
            With Me.fraProduct
                .ScrollBars = fmScrollBarsVertical
                .ScrollHeight = topPos + 120
            End With
        Me.lblCate.Caption = "TEA"
End Sub

Private Sub btnTea_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
            Me.btnTea.BackColor = RGB(171, 142, 79)
            Me.imgTea.Picture = LoadPicture("D:\Image_coffee_shop\br_tea.jpg")
End Sub

Private Sub btnTea_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
            Me.btnTea.BackColor = vbBlack
            Me.imgTea.Picture = LoadPicture("D:\Image_coffee_shop\tea.jpg")
End Sub
Private Sub fraCategory_Layout()
    Me.fraCategory.BackColor = RGB(17, 17, 17)
End Sub


Private Sub txtReceived_Change()

    Dim received As Double
    Dim total As Double
    
    If Me.txtReceived.Value = "" Then
        received = 0
    Else
        received = CDbl(Me.txtReceived.Value)
    End If
    
    If Me.txtPrice.Value = "" Then
        total = 0
    Else
        total = CDbl(Replace(Me.txtPrice.Value, "$", ""))
    End If
    
    If received <> 0 Or total <> 0 Then
        Me.txtChange.Value = Format(received - total, "0.00")
    End If

End Sub

Private Sub UserForm_Initialize()
  
        Dim lastRow As Long
        Dim ws As Worksheet
         Dim wsSales As Worksheet
        Dim i As Long
        Dim btn As MSForms.CommandButton
        Dim topPos As Long
        Dim leftPos As Long
        Dim name As String
        Dim evt As clsProduct
        Dim displayName As String
              
        With Me.txtPrice
            .ForeColor = RGB(212, 175, 99)
            .Value = 0
        End With
        
        With Me
         .BackColor = RGB(18, 18, 18)
         .BorderColor = RGB(18, 18, 18)
         .btnClr.BackColor = RGB(153, 40, 40)
         .lblPayment.ForeColor = RGB(212, 175, 99)
         .btnPay.BackColor = RGB(56, 142, 60)
         .lblCate.ForeColor = RGB(212, 175, 99)
         .lblTime.BackColor = RGB(40, 40, 40)
         .lblOrder.BackColor = RGB(40, 40, 40)
         .txtOrderNumber.BackColor = RGB(40, 40, 40)
         .lblName.BackColor = RGB(40, 40, 40)
         .lblQty.BackColor = RGB(40, 40, 40)
         .lblMoney.BackColor = RGB(40, 40, 40)
         .txtReceived.BackColor = RGB(240, 240, 240)
         .txtChange.BackColor = RGB(240, 240, 240)
         .lblTotal.BackColor = RGB(40, 40, 40)
         .lblItem.BackColor = RGB(40, 40, 40)
         .lblQI.BackColor = RGB(40, 40, 40)
         .lblAmount.BackColor = RGB(40, 40, 40)
         .btnCancel.BackColor = RGB(153, 40, 40)
         .lblTime.Caption = Date & "   -  " & " " & Time
        End With
        
       
        
        With Me.lstOrder
        .Clear
              .ColumnCount = 3
              .ColumnWidths = "120 pt;30 pt;30 pt"
              .BackColor = RGB(40, 40, 40)
        End With
        
        Set buttom_btn = New Collection
        Set ws = Sheets("Product")
        Set wsSales = Sheets("Sales")
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        
        topPos = 50
        leftPos = 10
      
        
        For i = 3 To lastRow
        
            If ws.Cells(i, 2).Value = "Coffee" Then
            
           
                Set btn = Me.fraProduct.Controls.Add("Forms.CommandButton.1")
                
                 With btn
                    name = ws.Cells(i, 3).Value
                    displayName = name
                    .Width = 88
                    .Height = 99
                    
                    .Left = leftPos
                    .Top = topPos
                    If Len(displayName) > 12 Then
                        displayName = Replace(displayName, " ", vbCrLf, 1, 1)
                        End If
                    .Caption = displayName & vbCrLf & "$" & ws.Cells(i, 4).Value
                       .Tag = name & "|" & ws.Cells(i, 4).Value
                  
                       
                    .Picture = LoadPicture("D:\Image_coffee_shop\images\" & ws.Cells(i, 5).Value & ".jpg")
                    
                    .PicturePosition = fmPicturePositionAboveCenter
                    
                       Set evt = New clsProduct
                       Set evt.btn = btn
                       buttom_btn.Add evt
                    
                       .BackColor = vbBlack
                       .ForeColor = vbWhite
                       .Font.Size = 10
                       .Font.name = "Segoe UI Black"
                End With
          
                
                leftPos = leftPos + 92
                
                If leftPos > 280 Then
                
                    leftPos = 10
                    topPos = topPos + 100
                    
                End If
                
            End If
        
        Next i
        
        With Me.fraProduct
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = topPos + 120
        End With
        
            Me.btnCancel.Enabled = False
            Me.btnPay.Enabled = False
        
          If Application.WorksheetFunction.Count(ws.Range("A:A")) = 0 Then
                Me.txtOrderNumber.Value = 1
          Else
                Me.txtOrderNumber.Value = Application.WorksheetFunction.Max(wsSales.Range("A:A")) + 1
      
      End If

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
