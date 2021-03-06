VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListadoFacturasPendientePorProveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas Pendientes de Pago a Proveedores"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport Rep 
      Left            =   3720
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   720
      Left            =   6765
      Picture         =   "frmListadoFacturasPendientePorProveedor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5130
      Width           =   840
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   720
      Left            =   8490
      Picture         =   "frmListadoFacturasPendientePorProveedor.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5130
      Width           =   825
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoFacturasPendientePorProveedor.frx":0BD4
      Height          =   720
      Left            =   7620
      Picture         =   "frmListadoFacturasPendientePorProveedor.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5130
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proveedor..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   60
      TabIndex        =   11
      Top             =   -15
      Width           =   9285
      Begin VB.CommandButton cmdBuscarProveedor1 
         Height          =   300
         Left            =   1845
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoFacturasPendientePorProveedor.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Buscar Proveedor"
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.CommandButton CmdBuscAprox 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
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
         Left            =   5640
         MaskColor       =   &H00000000&
         TabIndex        =   3
         ToolTipText     =   "Buscar "
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   1125
      End
      Begin VB.ComboBox cboTipoProveedor 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtCodProveedor 
         Height          =   300
         Left            =   960
         MaxLength       =   40
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtProvRazSoc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2265
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "Descripci�n"
         Top             =   720
         Width           =   4500
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Prov.:"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   390
         Width           =   780
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "C�digo:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   15
         Top             =   735
         Width           =   540
      End
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2835
      Top             =   5145
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   3240
      Left            =   45
      TabIndex        =   17
      Top             =   1095
      Width           =   9300
      Begin VB.TextBox txtTotalSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   5535
         MaxLength       =   40
         TabIndex        =   19
         Top             =   2820
         Width           =   1230
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   2625
         Left            =   75
         TabIndex        =   4
         Top             =   150
         Width           =   9030
         _ExtentX        =   15928
         _ExtentY        =   4630
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Deuda"
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
         Left            =   4320
         TabIndex        =   18
         Top             =   2880
         Width           =   1065
      End
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
      Top             =   4350
      Width           =   9285
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   6
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   5
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   435
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   10
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
      TabIndex        =   14
      Top             =   5310
      Width           =   750
   End
End
Attribute VB_Name = "frmListadoFacturasPendientePorProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SumaSaldo As Double

Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub CmdBuscAprox_Click()
    SumaSaldo = 0
    GrdModulos.HighLight = flexHighlightNever
    GrdModulos.Rows = 1
    
    sql = "SELECT TCO_ABREVIA,FPR_NROSUC,FPR_NUMERO,"
    sql = sql & "FPR_FECHA,FPR_TOTAL,FPR_SALDO,PROV_CODIGO,PROV_RAZSOC"
    sql = sql & " FROM SALDO_FACTURAS_PROVEEDOR_V"
    If cboTipoProveedor.List(cboTipoProveedor.ListIndex) <> "TODOS" Then
        sql = sql & " WHERE"
        sql = sql & " TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
        If txtCodProveedor.Text <> "" Then
            sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
        End If
    End If
    sql = sql & " ORDER BY FPR_FECHA"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!FPR_NROSUC, "0000") & Chr(9) & _
                               Format(rec!FPR_NUMERO, "00000000") & Chr(9) & rec!FPR_FECHA & Chr(9) & _
                               Valido_Importe(rec!FPR_TOTAL) & Chr(9) & Valido_Importe(rec!FPR_SALDO) & Chr(9) & _
                               rec!PROV_CODIGO & Chr(9) & rec!PROV_RAZSOC
            
            SumaSaldo = SumaSaldo + CDbl(rec!FPR_SALDO)
            rec.MoveNext
        Loop
        txtTotalSaldo.Text = Valido_Importe(CStr(SumaSaldo))
        GrdModulos.HighLight = flexHighlightAlways
        GrdModulos.SetFocus
    Else
        MsgBox "No se registran Facturas impagas al Proveedor", vbExclamation, TIT_MSGBOX
        GrdModulos.Rows = 1
        txtTotalSaldo.Text = ""
        GrdModulos.HighLight = flexHighlightNever
        cboTipoProveedor.SetFocus
    End If
    rec.Close
End Sub

Private Sub cmdBuscarProveedor1_Click()
        frmBuscar.TipoBusqueda = 5
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 1
        txtCodProveedor.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 2
        txtProvRazSoc.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 3
        Call BuscaCodigoProxItemData(CInt(frmBuscar.grdBuscar.Text), cboTipoProveedor)
        txtCodProveedor_LostFocus
        txtProvRazSoc.SetFocus
    Else
        txtCodProveedor.SetFocus
    End If
End Sub

Private Sub cmdListar_Click()
    lblEstado.Caption = "Buscando Listado..."
    Rep.WindowState = crptMaximized 'crptMinimized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SIELECTROCENTRO"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.SelectionFormula = ""
    
    If txtCodProveedor.Text <> "" Then
        Rep.SelectionFormula = "{SALDO_FACTURAS_PROVEEDOR_V.TPR_CODIGO}=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
        Rep.SelectionFormula = Rep.SelectionFormula & " AND {SALDO_FACTURAS_PROVEEDOR_V.PROV_CODIGO}=" & txtCodProveedor.Text
    End If
    
    Rep.WindowTitle = "Facturas Pendientes de pago por Proveedor"
    Rep.ReportFileName = DRIVE & DirReport & "facturaspendientespagoxproveedor.rpt"
    
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
     Rep.Action = 1
     
     lblEstado.Caption = ""
     Rep.SelectionFormula = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
End Sub

Private Sub CmdNuevo_Click()
    txtTotalSaldo.Text = ""
    txtCodProveedor.Text = ""
    cboTipoProveedor.ListIndex = 0
    GrdModulos.HighLight = flexHighlightNever
    GrdModulos.Rows = 1
    cboTipoProveedor.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmListadoFacturasPendientePorProveedor = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    'CARGO COMBO TIPO PROVEEDOR
    LlenarComboTipoProv
    'CONFIGURA GRILLA-------------------
    GrdModulos.FormatString = "Comprobante|^Nro Suc|^Nro Comp|^Fecha|>Total|>Saldo|Cod. Prov|Proveedor"
    GrdModulos.ColWidth(0) = 1500
    GrdModulos.ColWidth(1) = 800
    GrdModulos.ColWidth(2) = 1000
    GrdModulos.ColWidth(3) = 1100
    GrdModulos.ColWidth(4) = 1100
    GrdModulos.ColWidth(5) = 1100
    GrdModulos.ColWidth(6) = 800
    GrdModulos.ColWidth(7) = 2500
    GrdModulos.Rows = 2
    GrdModulos.HighLight = flexHighlightAlways
    '-----------------------------------
    txtTotalSaldo.Text = ""
    lblEstado.Caption = ""
End Sub

Private Sub txtCodProveedor_Change()
    If txtCodProveedor.Text = "" Then
        txtProvRazSoc.Text = ""
    End If
End Sub

Private Sub txtCodProveedor_GotFocus()
    SelecTexto txtCodProveedor
End Sub

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodProveedor_LostFocus()
    If txtCodProveedor.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        Rec1.Open BuscoProveedor(txtCodProveedor), DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtProvRazSoc.Text = Rec1!PROV_RAZSOC
            Call BuscaCodigoProxItemData(CInt(Rec1!TPR_CODIGO), cboTipoProveedor)
        Else
            MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
            txtCodProveedor.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtProvRazSoc_Change()
    If txtProvRazSoc.Text = "" Then
        txtCodProveedor.Text = ""
    End If
End Sub

Private Sub txtProvRazSoc_GotFocus()
    SelecTexto txtProvRazSoc
End Sub

Private Sub txtProvRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtProvRazSoc_LostFocus()
    If ActiveControl.Name = "txtCodProveedor" Or _
       cboTipoProveedor.List(cboTipoProveedor.ListIndex) = "TODOS" Then Exit Sub
       
    If txtCodProveedor.Text = "" And txtProvRazSoc.Text <> "" Then
        rec.Open BuscoProveedor(txtProvRazSoc), DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                frmBuscar.TipoBusqueda = 5
                frmBuscar.TxtDescriB.Text = txtProvRazSoc.Text
                frmBuscar.Show vbModal
                If frmBuscar.grdBuscar.Text <> "" Then
                    frmBuscar.grdBuscar.Col = 1
                    txtCodProveedor.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 2
                    txtProvRazSoc.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 3
                    Call BuscaCodigoProxItemData(CInt(frmBuscar.grdBuscar.Text), cboTipoProveedor)
                    txtCodProveedor_LostFocus
                Else
                    txtCodProveedor.SetFocus
                End If
            Else
                txtCodProveedor.Text = rec!PROV_CODIGO
                txtProvRazSoc.Text = rec!PROV_RAZSOC
                txtCodProveedor_LostFocus
            End If
        Else
            MsgBox "No se encontro el Proveedor", vbExclamation, TIT_MSGBOX
            txtCodProveedor.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Function BuscoProveedor(Pro As String) As String
    sql = "SELECT TPR_CODIGO,PROV_CODIGO, PROV_RAZSOC,"
    sql = sql & " PROV_DOMICI"
    sql = sql & " FROM PROVEEDOR "
    sql = sql & " WHERE"
    If txtCodProveedor.Text <> "" Then
        sql = sql & " PROV_CODIGO=" & XN(Pro)
    Else
        sql = sql & " PROV_RAZSOC LIKE '" & Pro & "%'"
    End If
    If cboTipoProveedor.List(cboTipoProveedor.ListIndex) <> "TODOS" Then
        sql = sql & " AND TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
    End If

    BuscoProveedor = sql
End Function

Private Sub LlenarComboTipoProv()
    sql = "SELECT * FROM TIPO_PROVEEDOR ORDER BY TPR_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboTipoProveedor.AddItem "TODOS"
        Do While rec.EOF = False
            cboTipoProveedor.AddItem rec!TPR_DESCRI
            cboTipoProveedor.ItemData(cboTipoProveedor.NewIndex) = rec!TPR_CODIGO
            rec.MoveNext
        Loop
        cboTipoProveedor.ListIndex = 0
    End If
    rec.Close
End Sub

