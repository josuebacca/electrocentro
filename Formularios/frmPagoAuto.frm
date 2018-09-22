VERSION 5.00
Begin VB.Form frmPagoAuto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago Automatico"
   ClientHeight    =   2460
   ClientLeft      =   7830
   ClientTop       =   4920
   ClientWidth     =   3585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3585
   Begin VB.TextBox txtResta 
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "frmPagoAuto.frx":0000
      Height          =   615
      Left            =   2520
      Picture         =   "frmPagoAuto.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   900
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "frmPagoAuto.frx":0614
      Height          =   600
      Left            =   1560
      Picture         =   "frmPagoAuto.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3375
      Begin VB.Frame Frame2 
         Caption         =   "Forma de Pago"
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   3135
         Begin VB.OptionButton optContado 
            Caption         =   "Contado"
            Height          =   255
            Left            =   360
            TabIndex        =   1
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optCheque 
            Caption         =   "Cheque"
            Height          =   255
            Left            =   1680
            TabIndex        =   2
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox txtEntrega 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Entrega:"
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmPagoAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdGrabar_Click()
Dim I As Integer
Dim Saldo As Double
Dim Entrega As Double
Dim ReciboSaldo As Double
Dim Resta As Double
Dim SaldoAplicar As Double
    If txtEntrega.Text <> "0,00" Then
        Entrega = txtEntrega.Text
        ReciboSaldo = frmReciboCliente.lblSaldo
        If Entrega <= ReciboSaldo Then
             Resta = Valido_Importe(txtEntrega.Text)
             I = 1
             'Comprobantes
             Do While Resta > 0.1
                 Screen.MousePointer = vbHourglass
                 SaldoAplicar = frmReciboCliente.GrillaAplicar.TextMatrix(I, 4)
                 If Resta >= SaldoAplicar Then
                     Resta = Resta - Valido_Importe(frmReciboCliente.GrillaAplicar.TextMatrix(I, 4))
                     Resta = IIf(Resta = 0, 0, Format(Resta, "##.##"))
                     frmReciboCliente.txtImporteApagar.Text = frmReciboCliente.GrillaAplicar.TextMatrix(I, 4)
                     Saldo = 0
                     
                 Else
                     With frmReciboCliente
                         .txtSaldo.Text = .GrillaAplicar.TextMatrix(I, 4)
                         .txtImporteApagar.Text = Resta
                         txtResta.Text = 0
                         Resta = 0
                         Saldo = .txtSaldo.Text - .txtImporteApagar.Text
                     End With
                 End If
                 With frmReciboCliente
                     '.txtImporteApagar = .GrillaAplicar.TextMatrix(i, 4)
                     
                     .GrillaAplicar1.AddItem .GrillaAplicar.TextMatrix(I, 0) & Chr(9) & _
                                Valido_Importe(.txtImporteApagar.Text) & Chr(9) & _
                                .GrillaAplicar.TextMatrix(I, 1) & Chr(9) & _
                                .GrillaAplicar.TextMatrix(I, 2) & Chr(9) & _
                                Valido_Importe(CStr(Saldo)) & Chr(9) & _
                                .GrillaAplicar.TextMatrix(I, 5)
                     
                 End With
                 I = I + 1
             Loop
            
             With frmReciboCliente
             '.GrillaAplicar.Rows = 1
             .GrillaAplicar1.HighLight = flexHighlightAlways
             .txtImporteApagar.Text = ""
             .txtSaldo.Text = ""
             .tabComprobantes.Tab = 0
             
             End With
            'valores
            If optContado.Value = True Then
                 With frmReciboCliente
                 
                 .grillaValores.AddItem "EFT" & Chr(9) & Valido_Importe(txtEntrega.Text) & Chr(9) & "" _
                                    & Chr(9) & "PESOS" & Chr(9) & "" & Chr(9) & _
                                    1
            
                 .txtTotalValores.Text = Valido_Importe(txtEntrega.Text)
                 .grillaValores.HighLight = flexHighlightAlways
                 .GrillaEfectivo.Rows = 1
                 .txtTotalEfectivo.Text = ""
                 .tabValores.Tab = 0
                 End With
            Else
                frmReciboCliente.tabValores.Tab = 1
            End If
            Screen.MousePointer = vbNormal
            Unload Me
         Else
            MsgBox "El monto de entrega ingresado supera al saldo", vbInformation, TIT_MSGBOX
            txtEntrega.SetFocus
            Exit Sub
         End If
    Else
        MsgBox "Debe ingresar un monto de entrega", vbInformation, TIT_MSGBOX
        txtEntrega.SetFocus
        Exit Sub
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    txtEntrega.Text = "0,00"
End Sub

Private Sub txtEntrega_GotFocus()
    SelecTexto txtEntrega
End Sub

Private Sub txtEntrega_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtEntrega, KeyAscii)
End Sub

Private Sub txtEntrega_LostFocus()
    txtEntrega.Text = Valido_Importe(txtEntrega.Text)
End Sub
