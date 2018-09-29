VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListadoCantidadesVendidas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Cantidades Vendidas"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport Rep 
      Left            =   1320
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame FrameImpresora 
      Caption         =   "impresora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   60
      TabIndex        =   13
      Top             =   1785
      Width           =   6435
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   3
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   2
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   435
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   210
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   3885
      Picture         =   "frmListadoCantidadesVendidas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2565
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   5640
      Picture         =   "frmListadoCantidadesVendidas.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2565
      Width           =   840
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoCantidadesVendidas.frx":0BD4
      Height          =   750
      Left            =   4755
      Picture         =   "frmListadoCantidadesVendidas.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2565
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listar por..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   60
      TabIndex        =   8
      Top             =   -15
      Width           =   6435
      Begin VB.ComboBox cboRubro 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   780
         Width           =   3420
      End
      Begin VB.ComboBox cboLinea 
         Height          =   315
         Left            =   1695
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   3420
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   55443457
         CurrentDate     =   43367
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   4200
         TabIndex        =   17
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   55443457
         CurrentDate     =   43367
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Rubro:"
         Height          =   195
         Left            =   1110
         TabIndex        =   12
         Top             =   825
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Línea:"
         Height          =   195
         Left            =   1125
         TabIndex        =   11
         Top             =   345
         Width           =   465
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   3165
         TabIndex        =   10
         Top             =   1410
         Width           =   960
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   585
         TabIndex        =   9
         Top             =   1410
         Width           =   1005
      End
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2850
      Top             =   2580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   750
   End
End
Attribute VB_Name = "frmListadoCantidadesVendidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub cboLinea_Click()
    cboRubro.Clear
End Sub

Private Sub cboLinea_LostFocus()
    If cboLinea.List(cboLinea.ListIndex) <> "<Todas>" Then
        cargocboRubro
    Else
        cboRubro.Clear
        cboRubro.AddItem "<Todos>"
        cboRubro.ListIndex = 0
    End If
End Sub

Private Sub cmdListar_Click()
    lblestado.Caption = "Buscando Listado..."
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SIELECTROCENTRO"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    
   
    If cboLinea.List(cboLinea.ListIndex) <> "<Todas>" Then
        Rep.SelectionFormula = " {LINEAS.LNA_CODIGO}=" & XN(cboLinea.ItemData(cboLinea.ListIndex))
    End If
    If cboRubro.List(cboRubro.ListIndex) <> "<Todos>" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {RUBROS.RUB_CODIGO}=" & XN(cboRubro.ItemData(cboRubro.ListIndex))
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {RUBROS.RUB_CODIGO}=" & XN(cboRubro.ItemData(cboRubro.ListIndex))
        End If
    End If
    
    If FechaDesde.Value <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {DETALLE_REMITO_CLIENTE.RCL_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {DETALLE_REMITO_CLIENTE.RCL_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        End If
    End If
    If FechaHasta.Value <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {DETALLE_REMITO_CLIENTE.RCL_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                           
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {DETALLE_REMITO_CLIENTE.RCL_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
        End If
    End If
    
    If Rep.SelectionFormula = "" Then 'ESTADO DEFINITIVO
        Rep.SelectionFormula = " {REMITO_CLIENTE.EST_CODIGO}<>2"
    Else
        Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_CLIENTE.EST_CODIGO}<>2"
    End If
    
    If FechaDesde.Value <> "" And FechaHasta.Value <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value <> "" And FechaHasta.Value = Date Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
    ElseIf FechaDesde.Value = Date And FechaHasta.Value <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value = Date And FechaHasta.Value = Date Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
    Rep.WindowTitle = "Listado de Cantidades Vendidas"
    Rep.ReportFileName = DRIVE & DirReport & "rptlistadocantidadesvendidas.rpt"

    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
     Rep.WindowState = crptMaximized
     Rep.Action = 1
     
     lblestado.Caption = ""
     Rep.SelectionFormula = ""
     Rep.Formulas(0) = ""
End Sub

Private Sub CmdNuevo_Click()
    cboLinea.ListIndex = 0
    cboRubro.Clear
    FechaDesde.Value = Date
    FechaHasta.Value = Date
End Sub

Private Sub cmdSalir_Click()
    Set frmListadoCantidadesVendidas = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    cargocboLinea
    cboRubro.AddItem "<Todos>"
    cboRubro.ListIndex = 0
    'FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub cargocboLinea()
    lblestado.Caption = ""
    sql = "SELECT * FROM LINEAS  ORDER BY LNA_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboLinea.AddItem "<Todas>"
        Do While rec.EOF = False
            cboLinea.AddItem rec!LNA_DESCRI
            cboLinea.ItemData(cboLinea.NewIndex) = rec!LNA_CODIGO
            rec.MoveNext
        Loop
        cboLinea.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub cargocboRubro()
    sql = "SELECT * FROM RUBROS "
    sql = sql & " WHERE LNA_CODIGO= " & cboLinea.ItemData(cboLinea.ListIndex)
    sql = sql & " ORDER BY RUB_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboRubro.AddItem "<Todos>"
        Do While rec.EOF = False
            cboRubro.AddItem rec!RUB_DESCRI
            cboRubro.ItemData(cboRubro.NewIndex) = rec!RUB_CODIGO
            rec.MoveNext
        Loop
        cboRubro.ListIndex = 0
    End If
    rec.Close
End Sub


