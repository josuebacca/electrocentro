VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListadoVentasPorVendedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Ventas por Vendedor"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport Rep 
      Left            =   1680
      Top             =   2880
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
      TabIndex        =   12
      Top             =   1725
      Width           =   6690
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   4
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   3
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   435
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   210
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   4140
      Picture         =   "frmListadoVentasPorVendedor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2505
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   5895
      Picture         =   "frmListadoVentasPorVendedor.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2505
      Width           =   840
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoVentasPorVendedor.frx":0BD4
      Height          =   750
      Left            =   5010
      Picture         =   "frmListadoVentasPorVendedor.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2505
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
      Height          =   1740
      Left            =   60
      TabIndex        =   9
      Top             =   -15
      Width           =   6690
      Begin VB.OptionButton optMonto 
         Caption         =   "Por Monto"
         Height          =   195
         Left            =   1710
         TabIndex        =   1
         Top             =   1365
         Width           =   1305
      End
      Begin VB.OptionButton optCantidad 
         Caption         =   "Por Cantidad"
         Height          =   195
         Left            =   3990
         TabIndex        =   2
         Top             =   1365
         Width           =   1380
      End
      Begin VB.ComboBox CboVend 
         Height          =   315
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   375
         Width           =   3435
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   53673985
         CurrentDate     =   43367
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   375
         Left            =   4320
         TabIndex        =   18
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   53673985
         CurrentDate     =   43367
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ver:"
         Height          =   195
         Left            =   1350
         TabIndex        =   16
         Top             =   1365
         Width           =   285
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   900
         TabIndex        =   14
         Top             =   420
         Width           =   735
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   3210
         TabIndex        =   11
         Top             =   870
         Width           =   960
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   630
         TabIndex        =   10
         Top             =   870
         Width           =   1005
      End
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2835
      Top             =   2640
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
      Left            =   150
      TabIndex        =   15
      Top             =   2820
      Width           =   750
   End
End
Attribute VB_Name = "frmListadoVentasPorVendedor"
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

Private Sub cmdListar_Click()
    lblestado.Caption = "Buscando Listado..."
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SIELECTROCENTRO"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.SelectionFormula = ""
    
    If CboVend.List(CboVend.ListIndex) <> "<Todos>" Then
        Rep.SelectionFormula = "{FACTURA_CLIENTE.VEN_CODIGO}=" & CboVend.ItemData(CboVend.ListIndex)
    End If
    
    If FechaDesde.Value <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {FACTURA_CLIENTE.FCL_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.FCL_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        End If
    End If
    If FechaHasta.Value <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {FACTURA_CLIENTE.FCL_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                           
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.FCL_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
        End If
    End If
    If Rep.SelectionFormula = "" Then
        Rep.SelectionFormula = " {FACTURA_CLIENTE.EST_CODIGO}=3"
    Else
        Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.EST_CODIGO}=3"
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
    
    Rep.WindowTitle = "Listado de Ventas por Vendedor"
    If optMonto.Value = True Then
        Rep.ReportFileName = DRIVE & DirReport & "rptlistadoventasxvendedor.rpt"
    ElseIf optCantidad.Value = True Then
        Rep.ReportFileName = DRIVE & DirReport & "rptcantidadesvendidasporvendedor.rpt"
    End If
    
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
     Rep.Action = 1
     
     lblestado.Caption = ""
     Rep.SelectionFormula = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
End Sub

Private Sub CmdNuevo_Click()
    CboVend.ListIndex = 0
    FechaDesde.Value = Date
    FechaHasta.Value = Date
    optMonto.Value = True
    CboVend.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Set frmListadoVentasPorVendedor = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    LLenarComboVendedor
    optMonto.Value = True
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    lblestado.Caption = ""
End Sub

Private Sub LLenarComboVendedor()
    sql = "SELECT * FROM VENDEDOR ORDER BY VEN_NOMBRE"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        CboVend.AddItem "<Todos>"
        Do While rec.EOF = False
            CboVend.AddItem rec!VEN_NOMBRE
            CboVend.ItemData(CboVend.NewIndex) = rec!VEN_CODIGO
            rec.MoveNext
        Loop
        CboVend.ListIndex = 0
    End If
    rec.Close
End Sub

