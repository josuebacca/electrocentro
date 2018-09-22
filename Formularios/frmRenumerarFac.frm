VERSION 5.00
Begin VB.Form frmRenumerarFac 
   Caption         =   "Renumerar Factura"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "frmRenumerarFac.frx":0000
      Height          =   735
      Left            =   3120
      Picture         =   "frmRenumerarFac.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1600
      Width           =   900
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Cancelar"
      DisabledPicture =   "frmRenumerarFac.frx":0614
      Height          =   735
      Left            =   4110
      Picture         =   "frmRenumerarFac.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1600
      Width           =   900
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   2415
      Begin VB.TextBox txtNueNumero 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtNueSuc 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2415
      Begin VB.TextBox txtActNumero 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtActSuc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmRenumerarFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBorrar_Click()
    Unload Me
End Sub

Private Sub CmdGrabar_Click()
    sql = "SELECT * FROM FACTURA_CLIENTE "
    sql = sql & " WHERE FCL_SUCURSAL = " & Int(txtNueSuc.Text)
    sql = sql & " AND FCL_NUMERO = " & Int(txtNueNumero.Text)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        MsgBox "La Factura " & txtNueSuc.Text & " - " & txtNueNumero.Text & " ya existe!!", vbExclamation, TIT_MSGBOX
    Else
        ' FACTURA CLIENTE
        sql = "UPDATE FACTURA_CLIENTE "
        sql = sql & " SET FCL_SUCURSAL = " & txtNueSuc.Text
        sql = sql & " ,FCL_NUMERO = " & txtNueNumero.Text
        sql = sql & " WHERE FCL_SUCURSAL = " & txtActSuc.Text
        sql = sql & " AND FCL_NUMERO =" & txtActNumero.Text
        DBConn.Execute sql
        ' DETALLE FACTURA CLIENTE
        sql = "UPDATE DETALLE_FACTURA_CLIENTE "
        sql = sql & " SET FCL_SUCURSAL = " & txtNueSuc.Text
        sql = sql & " ,FCL_NUMERO = " & txtNueNumero.Text
        sql = sql & " WHERE FCL_SUCURSAL = " & txtActSuc.Text
        sql = sql & " AND FCL_NUMERO =" & txtActNumero.Text
        DBConn.Execute sql
        ' CTA_CTE CLIENTE
        sql = "UPDATE CTA_CTE_CLIENTE "
        sql = sql & " SET COM_SUCURSAL = " & txtNueSuc.Text
        sql = sql & " ,COM_NUMERO = " & txtNueNumero.Text
        sql = sql & " WHERE COM_SUCURSAL = " & txtActSuc.Text
        sql = sql & " AND COM_NUMERO =" & txtActNumero.Text
        DBConn.Execute sql
        ' FACTURAS_RECIBO_CLIENTE
        sql = "UPDATE FACTURAS_RECIBO_CLIENTE "
        sql = sql & " SET FCL_SUCURSAL = " & txtNueSuc.Text
        sql = sql & " ,FCL_NUMERO = " & txtNueNumero.Text
        sql = sql & " WHERE FCL_SUCURSAL = " & txtActSuc.Text
        sql = sql & " AND FCL_NUMERO =" & txtActNumero.Text
        DBConn.Execute sql
        ' FACTURA_NOTA_CREDITO_CLIENTES
        sql = "UPDATE FACTURAS_NOTA_CREDITO_CLIENTE "
        sql = sql & " SET FCL_SUCURSAL = " & txtNueSuc.Text
        sql = sql & " ,FCL_NUMERO = " & txtNueNumero.Text
        sql = sql & " WHERE FCL_SUCURSAL = " & txtActSuc.Text
        sql = sql & " AND FCL_NUMERO =" & txtActNumero.Text
        DBConn.Execute sql
        
        MsgBox "La Factura se renumeró correctamente", vbInformation, TIT_MSGBOX
        frmAnulaDocumentos.GrdModulos.TextMatrix(frmAnulaDocumentos.GrdModulos.RowSel, 1) = Format(txtNueSuc.Text, "0000") & "-" & Format(txtNueNumero.Text, "00000000")
        
        Unload Me
        
    End If
    rec.Close
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Call Centrar_pantalla(Me)
    txtActSuc.Text = Left(frmAnulaDocumentos.GrdModulos.TextMatrix(frmAnulaDocumentos.GrdModulos.RowSel, 1), 4)
    txtActNumero.Text = Right(frmAnulaDocumentos.GrdModulos.TextMatrix(frmAnulaDocumentos.GrdModulos.RowSel, 1), 8)
    
End Sub

Private Sub txtNueNumero_LostFocus()
    txtNueNumero.Text = Format(txtNueNumero.Text, "00000000")
End Sub

Private Sub txtNueSuc_LostFocus()
    txtNueSuc.Text = Format(txtNueSuc, "0000")
End Sub
