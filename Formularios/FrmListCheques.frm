VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmListCheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Cheques"
   ClientHeight    =   6645
   ClientLeft      =   1365
   ClientTop       =   975
   ClientWidth     =   8760
   Icon            =   "FrmListCheques.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraImp 
      Caption         =   "Impresión de Reporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6510
      Left            =   105
      TabIndex        =   15
      Top             =   75
      Width           =   8595
      Begin VB.Frame Frame1 
         Height          =   4035
         Left            =   3945
         TabIndex        =   19
         Top             =   390
         Width           =   4485
         Begin VB.ComboBox CboEstado 
            Height          =   315
            Left            =   1575
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2145
            Width           =   2790
         End
         Begin VB.TextBox TxtNroCheque 
            Enabled         =   0   'False
            Height          =   330
            Left            =   1575
            TabIndex        =   6
            Top             =   990
            Width           =   1080
         End
         Begin VB.ComboBox CboBanco 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1575
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   615
            Width           =   2775
         End
         Begin MSComCtl2.DTPicker TxtFecVtoD 
            Height          =   315
            Left            =   1560
            TabIndex        =   34
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55443457
            CurrentDate     =   43367
         End
         Begin MSComCtl2.DTPicker TxtFecVtoH 
            Height          =   315
            Left            =   3120
            TabIndex        =   35
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55443457
            CurrentDate     =   43367
         End
         Begin MSComCtl2.DTPicker TxtFecIngresoD 
            Height          =   315
            Left            =   1560
            TabIndex        =   36
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55443457
            CurrentDate     =   43367
         End
         Begin MSComCtl2.DTPicker TxtFecIngresoH 
            Height          =   315
            Left            =   3120
            TabIndex        =   37
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55443457
            CurrentDate     =   43367
         End
         Begin MSComCtl2.DTPicker Fecha1 
            Height          =   315
            Left            =   1560
            TabIndex        =   38
            Top             =   2520
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55443457
            CurrentDate     =   43367
         End
         Begin MSComCtl2.DTPicker Fecha2 
            Height          =   315
            Left            =   3120
            TabIndex        =   39
            Top             =   2520
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55443457
            CurrentDate     =   43367
         End
         Begin MSComCtl2.DTPicker Fecha3 
            Height          =   315
            Left            =   1560
            TabIndex        =   40
            Top             =   2880
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55443457
            CurrentDate     =   43367
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "(cheques en cartera)"
            Height          =   195
            Left            =   45
            TabIndex        =   30
            Top             =   405
            Width           =   1470
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   930
            TabIndex        =   29
            Top             =   3000
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Left            =   915
            TabIndex        =   28
            Top             =   2625
            Width           =   510
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Index           =   2
            Left            =   2940
            TabIndex        =   27
            Top             =   2595
            Width           =   120
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Index           =   1
            Left            =   2955
            TabIndex        =   26
            Top             =   300
            Width           =   120
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Index           =   0
            Left            =   2955
            TabIndex        =   25
            Top             =   1500
            Width           =   120
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   885
            TabIndex        =   24
            Top             =   2220
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   915
            TabIndex        =   23
            Top             =   675
            Width           =   510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nro de Cheque:"
            Height          =   195
            Left            =   300
            TabIndex        =   22
            Top             =   1020
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Ingreso:"
            Height          =   195
            Left            =   135
            TabIndex        =   21
            Top             =   1500
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Vto.:"
            Height          =   195
            Left            =   375
            TabIndex        =   20
            Top             =   240
            Width           =   1050
         End
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         DisabledPicture =   "FrmListCheques.frx":27A2
         Height          =   720
         Left            =   7500
         Picture         =   "FrmListCheques.frx":2BEC
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5685
         Width           =   915
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Aceptar"
         DisabledPicture =   "FrmListCheques.frx":2EF6
         Height          =   720
         Left            =   5640
         Picture         =   "FrmListCheques.frx":37C0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5685
         Width           =   915
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         DisabledPicture =   "FrmListCheques.frx":3ACA
         Height          =   720
         Left            =   6570
         Picture         =   "FrmListCheques.frx":4394
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5685
         Width           =   915
      End
      Begin VB.Frame fraSentido 
         Caption         =   "Sentido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   165
         TabIndex        =   18
         Top             =   3465
         Width           =   3660
         Begin VB.OptionButton oDescendente 
            Caption         =   "Descendente"
            Height          =   255
            Left            =   1965
            TabIndex        =   9
            Top             =   435
            Width           =   1335
         End
         Begin VB.OptionButton oAscendente 
            Caption         =   "Ascendente"
            Height          =   255
            Left            =   210
            TabIndex        =   8
            Top             =   435
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame fraOrden 
         Height          =   3000
         Left            =   165
         TabIndex        =   17
         Top             =   390
         Width           =   3660
         Begin VB.OptionButton Option5 
            Caption         =   " en Cartera al ..."
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2430
            Width           =   2910
         End
         Begin VB.OptionButton Option0 
            Caption         =   "... por Fecha de Vencimiento"
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   225
            Value           =   -1  'True
            Width           =   2910
         End
         Begin VB.OptionButton Option4 
            Caption         =   "... por Estado"
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1875
            Width           =   2910
         End
         Begin VB.OptionButton Option1 
            Caption         =   "... por Banco y Nro de Cheque"
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   765
            Width           =   2910
         End
         Begin VB.OptionButton Option3 
            Caption         =   "... por Fecha de Ingreso"
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1320
            Width           =   2910
         End
      End
      Begin VB.Frame fraImpresion 
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   165
         TabIndex        =   16
         Top             =   4440
         Width           =   8250
         Begin VB.CommandButton CmdCambiarImp 
            Caption         =   "&Configurar Impresora"
            Height          =   495
            Left            =   195
            TabIndex        =   32
            Top             =   600
            Width           =   1890
         End
         Begin VB.OptionButton oImpresora 
            Caption         =   "Impresora"
            Height          =   255
            Left            =   2220
            TabIndex        =   11
            Top             =   270
            Width           =   990
         End
         Begin VB.OptionButton oPantalla 
            Caption         =   "Pantalla"
            Height          =   255
            Left            =   1185
            TabIndex        =   10
            Top             =   270
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label LBImpActual 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impresora Actual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2235
            TabIndex        =   33
            Top             =   840
            Width           =   1440
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Destino:"
            Height          =   195
            Left            =   480
            TabIndex        =   31
            Top             =   270
            Width           =   585
         End
      End
      Begin Crystal.CrystalReport Rep 
         Left            =   4515
         Top             =   5760
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin MSComDlg.CommonDialog CDImpresora 
         Left            =   4950
         Top             =   5700
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Flags           =   64
      End
   End
End
Attribute VB_Name = "FrmListCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Limpio_Campos()
   Me.TxtFecVtoD.Value = Date
   Me.TxtFecVtoH.Value = Date
   Me.CboBanco.ListIndex = -1
   Me.TxtNroCheque.Text = ""
   Me.TxtFecIngresoD.Value = Date
   Me.TxtFecIngresoH.Value = Date
   Me.CboEstado.ListIndex = -1
   Me.Fecha1.Value = Date
   Me.Fecha2.Value = Date
   Me.Fecha3.Value = Date
End Sub

Private Sub CboEstado_LostFocus()
    If Me.Option4.Value = True Then Fecha1.SetFocus
     If Me.Option0.Value = True Then Me.CmdAgregar.SetFocus
End Sub

Private Sub CmdAgregar_Click()
    sql = ""
    'VALIDO LAS FECHAS
    If Option0.Value = True Then
        If TxtFecVtoD.Value = Date Then
            MsgBox "Falta ingresar la fecha desde la cual quiere consultar", vbExclamation, TIT_MSGBOX
            TxtFecVtoD.SetFocus
            Exit Sub
        End If
    ElseIf Option3.Value = True Then
        If TxtFecIngresoD.Value = Date Then
            MsgBox "Falta ingresar la fecha desde la cual quiere consultar", vbExclamation, TIT_MSGBOX
            TxtFecIngresoD.SetFocus
            Exit Sub
        End If
    ElseIf Option4.Value = True Then
        If Fecha1.Value = Date Then
            MsgBox "Falta ingresar la fecha desde la cual quiere consultar", vbExclamation, TIT_MSGBOX
            Fecha1.SetFocus
            Exit Sub
        End If
    ElseIf Option5.Value = True Then
        If Fecha3.Value = Date Then
            MsgBox "Falta ingresar la fecha desde la cual quiere consultar", vbExclamation, TIT_MSGBOX
            Fecha3.SetFocus
            Exit Sub
        End If
    End If
    
   On Error GoTo ErrorTrans
   
   Screen.MousePointer = 11
   
   'Sentido del Orden
   If oAscendente.Value = True Then
      wSentido = "+"
      Rep.Formulas(1) = "sentido ='Sentido: ASCENDENTE'"
   Else
      wSentido = "-"
      Rep.Formulas(1) = "sentido ='Sentido: DESCENDENTE '"
   End If
   
   If Me.Option0.Value = True Then 'Por Fecha de Vencimiento
       
       If Me.TxtFecVtoD.Value = Date Or Me.TxtFecVtoH.Value = Date Then
          If Me.TxtFecVtoD.Value = Date Then
            Me.TxtFecVtoD.Value = Format(Date, "dd/mm/yyyy")
          ElseIf Me.TxtFecVtoH.Value = Date Then
            Me.TxtFecVtoH.Value = Format(Date, "dd/mm/yyyy")
          End If
       End If
       
       '{ChequeEstadoVigente.ECH_CODIGO} = 1 Unicamente Cheques en Cartera
       sql = sql & "({ChequeEstadoVigente.ECH_CODIGO} = 1 and {ChequeEstadoVigente.CHE_FECVTO} >= DATE(" & Mid(TxtFecVtoD.Value, 7, 4) & "," & _
                                                            Mid(TxtFecVtoD.Value, 4, 2) & "," & _
                                                            Mid(TxtFecVtoD.Value, 1, 2) & ") and " & _
                      "{ChequeEstadoVigente.CHE_FECVTO} <= DATE(" & Mid(TxtFecVtoH.Value, 7, 4) & "," & _
                                                                    Mid(TxtFecVtoH.Value, 4, 2) & "," & _
                                                                    Mid(TxtFecVtoH.Value, 1, 2) & "))"
       wCondicion = wSentido & " {ChequeEstadoVigente.CHE_FECVTO}"
       wCondicion1 = wSentido & " {ChequeEstadoVigente.CHE_NUMERO}"
       Rep.Formulas(0) = "orden ='Ordenado por: FECHA DE VTO. Y NRO DE CHEQUE'"
       
   ElseIf Me.Option1.Value = True Then 'por Banco y Nº de Cheque
      
       If TxtNroCheque.Text = "" Then
          TxtNroCheque.SetFocus
          Exit Sub
       End If
       
       sql = sql & "{ChequeEstadoVigente.BAN_CODINT} =  " & XN(Me.CboBanco.ItemData(Me.CboBanco.ListIndex)) & " AND {ChequeEstadoVigente.CHE_NUMERO} =  " & XS(TxtNroCheque.Text)
       wCondicion = wSentido & " {ChequeEstadoVigente.CHE_NUMERO}"
       wCondicion1 = ""
       Rep.Formulas(0) = "orden ='Ordenado por: NÚMERO DE CHEQUE'"
          
   ElseIf Me.Option3.Value = True Then 'por Fecha de Ingreso
   
       If Me.TxtFecIngresoD.Value = Date Or Me.TxtFecIngresoH.Value = Date Then
          If Me.TxtFecIngresoD.Value = Date Then
            Me.TxtFecIngresoD.Value = Format(Date, "dd/mm/yyyy")
          ElseIf Me.TxtFecIngresoH.Value = Date Then
            Me.TxtFecIngresoH.Value = Format(Date, "dd/mm/yyyy")
          End If
       End If
       
       sql = sql & "{ChequeEstadoVigente.CHE_FECENT} >= DATE(" & Mid(TxtFecIngresoD.Value, 7, 4) & _
                                                      "," & Mid(TxtFecIngresoD.Value, 4, 2) & _
                                                      "," & Mid(TxtFecIngresoD.Value, 1, 2) & ")and " & _
                   "{ChequeEstadoVigente.CHE_FECENT} <= DATE(" & Mid(TxtFecIngresoH.Value, 7, 4) & "," & _
                                                            Mid(TxtFecIngresoH.Value, 4, 2) & "," & _
                                                            Mid(TxtFecIngresoH.Value, 1, 2) & ")"
       
       wCondicion = wSentido & " {ChequeEstadoVigente.CHE_FECENT}"
       wCondicion1 = wSentido & " {ChequeEstadoVigente.CHE_FECVTO}"
       
       Rep.Formulas(0) = "orden ='Ordenado por: FECHA DE INGRESO y FECHA DE VTO.'"
   
   ElseIf Me.Option4.Value = True Then 'por Estado y Fecha de Ingreso
   
       If Fecha1.Value = Date Or Fecha2.Value = Date Then
          If Fecha1.Value = Date Then
            Fecha1.Value = Format(Date, "dd/mm/yyyy")
          ElseIf Fecha2.Value = Date Then
            Fecha2.Value = Format(Date, "dd/mm/yyyy")
          End If
       End If
    
       sql = sql & " {ChequeEstadoVigente.CHE_FECENT} >= DATE(" & Mid(Fecha1.Value, 7, 4) & "," & _
                                                                    Mid(Fecha1.Value, 4, 2) & "," & _
                                                                    Mid(Fecha1.Value, 1, 2) & ") and " & _
                   "{ChequeEstadoVigente.CHE_FECENT} <= DATE(" & Mid(Fecha2.Value, 7, 4) & "," & _
                                                                    Mid(Fecha2.Value, 4, 2) & "," & _
                                                                    Mid(Fecha2.Value, 1, 2) & ")"
       'por Estado
       If Me.CboEstado.List(Me.CboEstado.ListIndex) <> "(Todos)" Then
           If Me.CboEstado.List(Me.CboEstado.ListIndex) = "RECHAZADOS TODOS" Then
              sql = sql & " AND {ChequeEstadoVigente.ECH_CODIGO} >= 8 " & _
                            " AND {ChequeEstadoVigente.ECH_CODIGO} <= 24 "
           Else
              sql = sql & " AND {ChequeEstadoVigente.ECH_CODIGO} =  " & XN(Me.CboEstado.ItemData(Me.CboEstado.ListIndex))
           End If
       End If
       wCondicion = wSentido & " {ChequeEstadoVigente.CHE_FECVTO}"
       wCondicion1 = wSentido & " {ChequeEstadoVigente.CHE_NUMERO}"
       Rep.Formulas(0) = "orden ='Ordenado por: FECHA DE VTO. Y NRO. DE CHEQUE'"
   
   ElseIf Me.Option5.Value = True Then 'en Cartera a Fecha
   
       If Fecha3.Value = Date Then Fecha3.Value = Format(Date, "dd/mm/yyyy")
       sql = sql & " {ChequeEstadoVigente.ECH_CODIGO} = 1"
       sql = sql & " AND {ChequeEstadoVigente.CHE_FECENT} <= DATE(" & Mid(Fecha3.Value, 7, 4) & "," & _
                                                            Mid(Fecha3.Value, 4, 2) & "," & _
                                                            Mid(Fecha3.Value, 1, 2) & ")" 'And " &" _
'                      "{ChequeEstadoVigente.CES_FECHA} <= DATE(" & Mid(Fecha3.Value, 7, 4) & "," & _
'                                                                   Mid(Fecha3.Value, 4, 2) & "," & _
'                                                                   Mid(Fecha3.Value, 1, 2) & ")"
       
       wCondicion = wSentido & " {ChequeEstadoVigente.CHE_FECVTO}"
       wCondicion1 = wSentido & " {ChequeEstadoVigente.CHE_NUMERO}"
       Rep.Formulas(0) = "orden ='Ordenado por: FECHA DE VTO. Y NRO. DE CHEQUE'"
   
   End If
   
   If oImpresora = True Then
       Rep.Destination = 1
   Else
       Rep.Destination = 0
       Rep.WindowMinButton = 0
       Rep.WindowTitle = "Consulta de Cheques"
       Rep.WindowBorderStyle = 2
   End If
   
   Rep.SortFields(0) = wCondicion
   Rep.SortFields(1) = wCondicion1
   
   Rep.SelectionFormula = sql
   Rep.WindowState = crptMaximized
   Rep.WindowBorderStyle = crptNoBorder
   Rep.Connect = "Provider=MSDASQL.1;Persst Security Info=False;Data Source=SIELECTROCENTRO"
   
   Rep.ReportFileName = DRIVE & DirReport & "CHEQUE.rpt"
   Rep.Action = 1
   
   Rep.Formulas(0) = ""
   Rep.Formulas(1) = ""
   Rep.Formulas(2) = ""
   Rep.Formulas(3) = ""
   
   Screen.MousePointer = 1
   Exit Sub

ErrorTrans:
  Screen.MousePointer = 1
  MsgBox "Error intentando armar el reporte. " & Chr(13) & Err.Description, 16, TIT_MSGBOX
End Sub

Private Sub CmdCambiarImp_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    LBImpActual.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub CmdCancelar_Click()
    Limpio_Campos
    Option0.Value = True
    Option1.Value = False
    Option3.Value = False
    Option4.Value = False
    oAscendente.Value = True
    oPantalla.Value = True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    Set FrmListCheques = Nothing
End Sub

Private Sub fecha1_LostFocus()
    'If Me.Option4.Value = True And Fecha1.Value= date Then Fecha1.Value = Format(Date, "dd/mm/yyyy")
    'If Me.Option4.Value = True Then Fecha2.SetFocus
End Sub

Private Sub fecha2_LostFocus()
    'If Me.Option4.Value = True And Fecha2.Value= date Then Fecha2.Value = Format(Date, "dd/mm/yyyy")
    'If Me.Option4.Value = True Then CmdAgregar.SetFocus
End Sub

Private Sub fecha3_LostFocus()
   'If Me.Option5.Value = True And Fecha3.Value= date Then Fecha3.Value = Format(Date, "dd/mm/yyyy")
   'If Me.Option5.Value = True Then CmdAgregar.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then 'avanza de campo
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    KeyPreview = True
    
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.Enabled = False
    Me.CboEstado.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    
    Call Centrar_pantalla(Me)

    Set rec = New ADODB.Recordset
    
    CboEstado.Clear
    CboEstado.AddItem "(Todos)"
    sql = "SELECT ECH_CODIGO, ECH_DESCRI FROM ESTADO_CHEQUE"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While Not rec.EOF
            CboEstado.AddItem Trim(rec!ECH_DESCRI)
            CboEstado.ItemData(CboEstado.NewIndex) = Trim(rec!ECH_CODIGO)
            rec.MoveNext
        Loop
        Me.CboEstado.ListIndex = -1
    End If
    rec.Close
    CboEstado.AddItem "RECHAZADOS TODOS"
    Me.MousePointer = 1
    
    LBImpActual.Caption = "Impresora Actual: " & Printer.DeviceName
    
    Option0_Click
End Sub

Private Sub oImpresora_Click()
  Me.LBImpActual.Caption = "Impresora Actual: " & Printer.DeviceName
  Me.LBImpActual.Visible = True
End Sub

Private Sub oPantalla_Click()
 ' Me.CDImpresora.Visible = False
  Me.LBImpActual.Visible = False
End Sub

Private Sub Option0_Click()
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = True
    Me.TxtFecVtoH.Enabled = True
    Me.CboBanco.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    Me.CboEstado.Enabled = False
    Me.Fecha1.Enabled = False
    Me.Fecha2.Enabled = False
    Me.Fecha3.Enabled = False
    If Me.TxtFecVtoD.Visible = True Then Me.TxtFecVtoD.SetFocus
End Sub

Private Sub Option1_Click()
    Me.CboBanco.Clear
    Set rec = New ADODB.Recordset
    sql = "SELECT DISTINCT B.BAN_CODINT,B.BAN_DESCRI"
    sql = sql & " FROM BANCO B, CHEQUE C"
    sql = sql & " WHERE C.BAN_CODINT = B.BAN_CODINT"
    sql = sql & " ORDER BY B.BAN_DESCRI"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        Do While Not rec.EOF
            Me.CboBanco.AddItem Trim(rec!BAN_DESCRI)
            Me.CboBanco.ItemData(Me.CboBanco.NewIndex) = rec!BAN_CODINT
            rec.MoveNext
        Loop
        Me.CboBanco.ListIndex = -1
    End If
    rec.Close
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.ListIndex = 0
    Me.CboBanco.Enabled = True
    Me.TxtNroCheque.Enabled = True
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    Me.CboEstado.Enabled = False
    Me.Fecha1.Enabled = False
    Me.Fecha2.Enabled = False
    Me.Fecha3.Enabled = False
    Me.CboBanco.SetFocus
End Sub

Private Sub Option3_Click()
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = True
    Me.TxtFecIngresoH.Enabled = True
    Me.CboEstado.Enabled = False
    Me.Fecha1.Enabled = False
    Me.Fecha2.Enabled = False
    Me.Fecha3.Enabled = False
    Me.TxtFecIngresoD.SetFocus
End Sub

Private Sub Option4_Click()
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    Me.CboEstado.ListIndex = 0
    Me.CboEstado.Enabled = True
    Me.Fecha1.Enabled = True
    Me.Fecha2.Enabled = True
    Me.Fecha3.Enabled = False
    Me.CboEstado.SetFocus
End Sub

Private Sub Option5_Click()
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    Me.CboEstado.Enabled = False
    Me.Fecha1.Enabled = False
    Me.Fecha2.Enabled = False
    Me.Fecha3.Enabled = True
    Me.Refresh
    Me.Fecha3.SetFocus
End Sub

Private Sub TxtFecIngresoD_LostFocus()
   'If Me.Option3.Value = True And TxtFecIngresoD.Value = date Then TxtFecIngresoD.Value = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub TxtFecIngresoH_LostFocus()
'If Me.Option3.Value = True And TxtFecIngresoH.Value = date Then TxtFecIngresoH.Value = Format(Date, "dd/mm/yyyy")

  If IsDate(TxtFecIngresoD.Value) And IsDate(TxtFecIngresoH.Value) Then
    
    If CVDate(TxtFecIngresoD.Value) > CVDate(TxtFecIngresoH.Value) Then
      MsgBox "La Fecha Hasta no puede ser inferior a la Fecha Desde. Verifique!", 16, TIT_MSGBOX
      TxtFecIngresoH.Value = Date
      TxtFecIngresoH.SetFocus
      Exit Sub
    Else
      If Not IsDate(TxtFecIngresoD.Value) Then TxtFecIngresoD.Value = Date
      If Not IsDate(TxtFecIngresoH.Value) Then TxtFecIngresoH.Value = Date
    End If
    
 End If
End Sub

Private Sub TxtFecVtoD_LostFocus()
  'If Me.Option0.Value = True And TxtFecVtoD.Value = date Then TxtFecVtoD.Value = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub TxtFecVtoH_LostFocus()

  'If Me.Option0.Value = True And TxtFecVtoH.Value = date Then TxtFecVtoH.Value = Format(Date, "dd/mm/yyyy")
  
  If IsDate(TxtFecVtoD.Value) And IsDate(TxtFecVtoH.Value) Then
  
    If CVDate(TxtFecVtoD.Value) > CVDate(TxtFecVtoH.Value) Then
      MsgBox "La Fecha Hasta no puede ser inferior a la Fecha Desde. Verifique!", 16, TIT_MSGBOX
      TxtFecVtoH.Value = Date
      TxtFecVtoD.SetFocus
      Exit Sub
    Else
      CmdAgregar.SetFocus
    End If
 End If
End Sub

Private Sub TxtNroCheque_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub TxtNroCheque_LostFocus()
    If Me.Option1.Value = True And Me.TxtNroCheque.Text = "" Then
        MsgBox "El Número de Cheque no puede ser Nulo. Verifique!", 16, TIT_MSGBOX
        TxtNroCheque.SetFocus
    Else
        If Len(TxtNroCheque.Text) < 10 Then TxtNroCheque.Text = CompletarConCeros(TxtNroCheque.Text, 10)
        CmdAgregar.SetFocus
    End If
End Sub

