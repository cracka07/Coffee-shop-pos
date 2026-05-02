Attribute VB_Name = "Module1"
Option Explicit

Sub LimpiarFrame()

    Dim ctrl As Control

For Each ctrl In frmPOS.fraProduct.Controls
    If TypeName(ctrl) = "CommandButton" Then
        frmPOS.fraProduct.Controls.Remove ctrl.name
    End If
Next ctrl
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
