VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Begin VB.Form frmListadoComprasProveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Compras a Proveedores...."
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
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
      Height          =   975
      Left            =   45
      TabIndex        =   21
      Top             =   4905
      Width           =   8325
      Begin VB.CommandButton CBImpresora 
         Caption         =   "&Configurar Impresora"
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   570
         Width           =   1665
      End
      Begin VB.OptionButton optExcel 
         Caption         =   "Excel"
         Height          =   225
         Left            =   4020
         TabIndex        =   13
         Top             =   270
         Width           =   780
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   11
         Top             =   270
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   12
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label lblImpresora 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1920
         TabIndex        =   23
         Top             =   690
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   30
      TabIndex        =   28
      Top             =   5820
      Width           =   11685
   End
   Begin VB.TextBox txtTotalGastos 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10170
      TabIndex        =   30
      Top             =   4950
      Width           =   1350
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoComprasProveedores.frx":0000
      Height          =   705
      Left            =   10065
      Picture         =   "frmListadoComprasProveedores.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5985
      Width           =   810
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   4155
      Top             =   6075
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   705
      Left            =   9240
      Picture         =   "frmListadoComprasProveedores.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5985
      Width           =   810
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   705
      Left            =   10890
      Picture         =   "frmListadoComprasProveedores.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5985
      Width           =   810
   End
   Begin VB.Frame Frame2 
      Caption         =   "   Proveedor ......"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   60
      TabIndex        =   17
      Top             =   -15
      Width           =   11580
      Begin VB.CheckBox chkFecha 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha"
         Height          =   195
         Left            =   960
         TabIndex        =   3
         Top             =   1245
         Width           =   810
      End
      Begin VB.CheckBox chkTipoGasto 
         Alignment       =   1  'Right Justify
         Caption         =   "Por Gasto"
         Height          =   195
         Left            =   765
         TabIndex        =   2
         Top             =   960
         Width           =   1005
      End
      Begin VB.ComboBox CboGastos 
         Height          =   315
         Left            =   3405
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   950
         Width           =   5535
      End
      Begin VB.CheckBox chkRazSoc 
         Alignment       =   1  'Right Justify
         Caption         =   "Por Raz. Soc."
         Height          =   195
         Left            =   405
         TabIndex        =   1
         Top             =   690
         Width           =   1365
      End
      Begin VB.CheckBox chkPorTipo 
         Alignment       =   1  'Right Justify
         Caption         =   "Por Tipo"
         Height          =   195
         Left            =   810
         TabIndex        =   0
         Top             =   405
         Width           =   960
      End
      Begin VB.ComboBox cboBuscaTipoProv 
         Height          =   315
         Left            =   3405
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   5535
      End
      Begin VB.TextBox txtBuscaProv 
         Height          =   315
         Left            =   3405
         MaxLength       =   30
         TabIndex        =   5
         Top             =   595
         Width           =   5535
      End
      Begin VB.CommandButton cmdBusProv 
         Height          =   1350
         Left            =   10560
         Picture         =   "frmListadoComprasProveedores.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   255
         Width           =   510
      End
      Begin FechaCtl.Fecha FechaHasta 
         Height          =   285
         Left            =   5910
         TabIndex        =   8
         Top             =   1305
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin FechaCtl.Fecha FechaDesde 
         Height          =   330
         Left            =   3405
         TabIndex        =   7
         Top             =   1305
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   4845
         TabIndex        =   27
         Top             =   1350
         Width           =   960
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   2325
         TabIndex        =   26
         Top             =   1335
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Gasto:"
         Height          =   195
         Left            =   2865
         TabIndex        =   25
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   225
         Index           =   28
         Left            =   2970
         TabIndex        =   19
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor:"
         Height          =   195
         Index           =   26
         Left            =   2550
         TabIndex        =   18
         Top             =   630
         Width           =   780
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgBuscaProv 
      Height          =   3165
      Left            =   45
      TabIndex        =   10
      Top             =   1710
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   5583
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColorSel    =   8388736
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   3720
      Top             =   6045
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Gastos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8670
      TabIndex        =   29
      Top             =   4980
      Width           =   1410
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
      Height          =   240
      Left            =   165
      TabIndex        =   20
      Top             =   6165
      Width           =   750
   End
End
Attribute VB_Name = "frmListadoComprasProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CBImpresora_Click()
  CDImpresora.PrinterDefault = True
  CDImpresora.ShowPrinter
  lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub chkFecha_Click()
    If chkFecha.Value = Checked Then
        FechaDesde.Enabled = True
        FechaHasta.Enabled = True
        FechaDesde.Text = ""
        FechaHasta.Text = ""
    Else
        FechaDesde.Enabled = False
        FechaHasta.Enabled = False
        FechaDesde.Text = ""
        FechaHasta.Text = ""
    End If
End Sub

Private Sub chkPorTipo_Click()
    If chkPorTipo.Value = Checked Then
        cboBuscaTipoProv.Enabled = True
        cboBuscaTipoProv.ListIndex = 0
    Else
        cboBuscaTipoProv.Enabled = False
        cboBuscaTipoProv.ListIndex = -1
    End If
End Sub

Private Sub chkRazSoc_Click()
    If chkRazSoc.Value = Checked Then
        txtBuscaProv.Enabled = True
    Else
        txtBuscaProv.Enabled = False
    End If
End Sub

Private Sub chkTipoGasto_Click()
    If chkTipoGasto.Value = Checked Then
        CboGastos.Enabled = True
        CboGastos.ListIndex = 0
    Else
        CboGastos.Enabled = False
        CboGastos.ListIndex = -1
    End If
End Sub

Private Sub cmdBusProv_Click()
    Screen.MousePointer = vbHourglass
    Me.lblestado.Caption = "Buscando Proveedores...."
    Me.Refresh
    
    fgBuscaProv.Rows = 1
    txtTotalGastos.Text = ""
    sql = "SELECT C.*"
    sql = sql & " FROM COMPRAS_V C, TIPO_PROVEEDOR TP"
    sql = sql & " WHERE TP.TPR_CODIGO=C.TPR_CODIGO"
   If chkPorTipo.Value = Checked Then
    sql = sql & " AND C.TPR_CODIGO=" & XN(cboBuscaTipoProv.ItemData(cboBuscaTipoProv.ListIndex))
   End If
   If chkRazSoc.Value = Checked Then
    sql = sql & " AND C.PROV_RAZSOC LIKE '%" & Trim(txtBuscaProv.Text) & "%'"
   End If
   If chkTipoGasto.Value = Checked Then
    sql = sql & " AND C.TGT_CODIGO=" & XN(CboGastos.ItemData(CboGastos.ListIndex))
   End If
   If FechaDesde <> "" Then sql = sql & " AND C.FPR_FECHA>=" & XDQ(FechaDesde)
   If FechaHasta <> "" Then sql = sql & " AND C.FPR_FECHA<=" & XDQ(FechaHasta)
   sql = sql & " ORDER BY C.TPR_CODIGO,C.PROV_CODIGO,C.FPR_FECHA"
    
   rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
   If rec.EOF = False Then
        fgBuscaProv.HighLight = flexHighlightAlways
        rec.MoveFirst
        txtTotalGastos.Text = "0,00"
        Do While rec.EOF = False
            fgBuscaProv.AddItem rec!TPR_DESCRI & Chr(9) & rec!PROV_RAZSOC _
                        & Chr(9) & rec!TGT_DESCRI & Chr(9) & rec!FPR_NROSUCTXT & "-" & rec!FPR_NUMEROTXT _
                        & Chr(9) & rec!FPR_FECHA & Chr(9) & Valido_Importe(rec!FPR_TOTAL) & Chr(9) & rec!PROV_CODIGO _
                        & Chr(9) & rec!TPR_CODIGO & Chr(9) & rec!TGT_CODIGO
            
            txtTotalGastos.Text = Valido_Importe(CStr(CDbl(txtTotalGastos.Text) + CDbl(rec!FPR_TOTAL)))
            rec.MoveNext
        Loop
         
        fgBuscaProv.SetFocus
        Screen.MousePointer = vbNormal
        Me.lblestado.Caption = ""
    
   Else ' SI NO ENCONTRO NINGUNO
        Screen.MousePointer = vbNormal
        Me.lblestado.Caption = ""
        fgBuscaProv.HighLight = flexHighlightNever
        fgBuscaProv.Rows = 1
        MsgBox "No se han encontrado Compras...", vbExclamation, TIT_MSGBOX
        'cboBuscaTipoProv.SetFocus
   End If
   rec.Close
End Sub

Private Sub cmdListar_Click()
    lblestado.Caption = "Buscando Listado..."
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SIELECTROCENTRO"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    
    If chkPorTipo.Value = Checked And (chkRazSoc.Value = Unchecked) Then
        Rep.SelectionFormula = "{COMPRAS_V.TPR_CODIGO}=" & cboBuscaTipoProv.ItemData(cboBuscaTipoProv.ListIndex)
    ElseIf chkRazSoc.Value = Checked Then
        Rep.SelectionFormula = "{COMPRAS_V.TPR_CODIGO}=" & fgBuscaProv.TextMatrix(fgBuscaProv.RowSel, 3) _
                               & " AND {COMPRAS_V.PROV_CODIGO}=" & fgBuscaProv.TextMatrix(fgBuscaProv.RowSel, 1)
    End If
    If chkTipoGasto.Value = Checked Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = "{COMPRAS_V.TGT_CODIGO}=" & CboGastos.ItemData(CboGastos.ListIndex)
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {COMPRAS_V.TGT_CODIGO}=" & CboGastos.ItemData(CboGastos.ListIndex)
        End If
    End If
    
    If FechaDesde.Text <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {COMPRAS_V.FPR_FECHA}>= DATE (" & Mid(FechaDesde.Text, 7, 4) & "," & Mid(FechaDesde.Text, 4, 2) & "," & Mid(FechaDesde.Text, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {COMPRAS_V.FPR_FECHA}>= DATE (" & Mid(FechaDesde.Text, 7, 4) & "," & Mid(FechaDesde.Text, 4, 2) & "," & Mid(FechaDesde.Text, 1, 2) & ")"
        End If
    End If
    
    If FechaHasta.Text <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {COMPRAS_V.FPR_FECHA}<= DATE (" & Mid(FechaHasta.Text, 7, 4) & "," & Mid(FechaHasta.Text, 4, 2) & "," & Mid(FechaHasta.Text, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {COMPRAS_V.FPR_FECHA}<= DATE (" & Mid(FechaHasta.Text, 7, 4) & "," & Mid(FechaHasta.Text, 4, 2) & "," & Mid(FechaHasta.Text, 1, 2) & ")"
        End If
    End If
    
    
    If FechaDesde.Text <> "" And FechaHasta.Text <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Text & "   Hasta: " & FechaHasta.Text & "'"
    ElseIf FechaDesde.Text <> "" And FechaHasta.Text = "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Text & "   Hasta: " & Date & "'"
    ElseIf FechaDesde.Text = "" And FechaHasta.Text <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Text & "'"
    ElseIf FechaDesde.Text = "" And FechaHasta.Text = "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
    
    Rep.WindowTitle = "Compras a Proveedores..."
    Rep.ReportFileName = DRIVE & DirReport & "rptcomprasaproveedores.rpt"
    
    If optPantalla.Value = True Then
         Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
        Rep.PrintFileType = crptExcel50
    ElseIf optExcel.Value = True Then
        Rep.Destination = crptToFile
        Rep.PrintFileType = crptExcel50
    End If
    Rep.Action = 1
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    lblestado.Caption = ""
End Sub

Private Sub CmdNuevo_Click()
    fgBuscaProv.Rows = 1
    fgBuscaProv.HighLight = flexHighlightNever
    txtBuscaProv.Text = ""
    txtTotalGastos.Text = ""
    chkPorTipo.Value = Unchecked
    chkRazSoc.Value = Unchecked
    chkTipoGasto.Value = Unchecked
    chkFecha.Value = Unchecked
    chkPorTipo.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Set frmListadoComprasProveedores = Nothing
    Unload Me
End Sub

Private Sub fgBuscaProv_DblClick()
    cmdListar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    lblestado.Caption = ""
    txtBuscaProv.Text = ""
    
    fgBuscaProv.FormatString = "Tipo Prov.|Proveedor|Gasto|^Comp.|^Fecha|>Total|nro proveedor|codigo tipo_proveedor|COD GASTO"
    
    fgBuscaProv.ColWidth(0) = 2250 'TIPO PROVEEDOR
    fgBuscaProv.ColWidth(1) = 3350 'PROVEEDOR
    fgBuscaProv.ColWidth(2) = 2250 'GASTO
    fgBuscaProv.ColWidth(3) = 1300 'NRO COMPROBANTE
    fgBuscaProv.ColWidth(4) = 1100 'FECHA COMP
    fgBuscaProv.ColWidth(5) = 1100 'TOTAL
    fgBuscaProv.ColWidth(6) = 0    'PROV_CODIGO
    fgBuscaProv.ColWidth(7) = 0    'TPR_CODIGO
    fgBuscaProv.ColWidth(8) = 0    'TGT_CODIGO GASTO
    fgBuscaProv.Rows = 2
    fgBuscaProv.Cols = 9
    fgBuscaProv.MergeCells = 1
    fgBuscaProv.MergeCol(0) = True
    fgBuscaProv.MergeCol(1) = True
    fgBuscaProv.MergeCol(2) = True
    'CARGO COMBO CON TIPO DE PROVEEDORES
    CargoComboTipoProveedor
    'CARGO COMBO CON LOS TIPO DE GASTOS
    LlenarComboGastos
    txtTotalGastos.Text = ""
    'impresora actual
    lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    cboBuscaTipoProv.Enabled = False
    txtBuscaProv.Enabled = False
    CboGastos.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
End Sub

Private Sub LlenarComboGastos()
    sql = "SELECT * FROM TIPO_GASTO ORDER BY TGT_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            CboGastos.AddItem rec!TGT_DESCRI
            CboGastos.ItemData(CboGastos.NewIndex) = rec!TGT_CODIGO
            rec.MoveNext
        Loop
        CboGastos.ListIndex = -1
    End If
    rec.Close
End Sub

Public Sub CargoComboTipoProveedor()
    'Cargo el combo Tipo de Proveedor
    cboBuscaTipoProv.Clear
    
    sql = "SELECT * FROM TIPO_PROVEEDOR ORDER BY TPR_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        Do While Not rec.EOF
            cboBuscaTipoProv.AddItem rec.Fields!TPR_DESCRI
            cboBuscaTipoProv.ItemData(cboBuscaTipoProv.NewIndex) = rec.Fields!TPR_CODIGO
            rec.MoveNext
        Loop
        cboBuscaTipoProv.ListIndex = -1
    End If
    rec.Close
End Sub

Private Sub txtBuscaProv_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub
