VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListadoRemitoCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Remito de Cliente"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3720
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   32
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txttotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   31
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoRemitoCliente.frx":0000
      Height          =   750
      Left            =   8865
      Picture         =   "frmListadoRemitoCliente.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5895
      Width           =   870
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   6615
      Top             =   5250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   -15
      TabIndex        =   24
      Top             =   6660
      Width           =   10650
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   9750
      Picture         =   "frmListadoRemitoCliente.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5895
      Width           =   840
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   7995
      Picture         =   "frmListadoRemitoCliente.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5895
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ver..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   90
      TabIndex        =   22
      Top             =   4800
      Width           =   10425
      Begin VB.OptionButton optDetallado 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado Detallado"
         Height          =   255
         Left            =   2370
         TabIndex        =   5
         Top             =   225
         Value           =   -1  'True
         Width           =   1770
      End
      Begin VB.OptionButton optGeneralTodos 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado General"
         Height          =   210
         Left            =   5265
         TabIndex        =   6
         Top             =   240
         Width           =   1620
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Impresora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   90
      TabIndex        =   18
      Top             =   5415
      Width           =   7845
      Begin Crystal.CrystalReport Rep 
         Left            =   5520
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   8
         Top             =   360
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "&Configurar Impresora"
         Height          =   435
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   660
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   360
         Width           =   585
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
         Left            =   1965
         TabIndex        =   20
         Top             =   840
         Width           =   840
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Remito de Cliente por..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   105
      TabIndex        =   12
      Top             =   75
      Width           =   10395
      Begin VB.Frame Frame5 
         Caption         =   "Estado Remito"
         Height          =   495
         Left            =   360
         TabIndex        =   25
         Top             =   1080
         Width           =   8535
         Begin VB.OptionButton optPen 
            Caption         =   "Pendientes"
            Height          =   195
            Left            =   1200
            TabIndex        =   29
            Top             =   200
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optDef 
            Caption         =   "Definitivos"
            Height          =   195
            Left            =   3075
            TabIndex        =   28
            Top             =   200
            Width           =   1455
         End
         Begin VB.OptionButton optAnu 
            Caption         =   "Anulados"
            Height          =   195
            Left            =   4845
            TabIndex        =   27
            Top             =   200
            Width           =   1455
         End
         Begin VB.OptionButton optTod 
            Caption         =   "Todos"
            Height          =   195
            Left            =   6600
            TabIndex        =   26
            Top             =   200
            Width           =   1455
         End
      End
      Begin VB.CheckBox chkCliente 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   300
         TabIndex        =   0
         Top             =   375
         Width           =   855
      End
      Begin VB.CheckBox chkFecha 
         Caption         =   "Fecha"
         Height          =   195
         Left            =   300
         TabIndex        =   1
         Top             =   720
         Width           =   810
      End
      Begin VB.TextBox txtCliente 
         Height          =   300
         Left            =   3360
         MaxLength       =   40
         TabIndex        =   2
         Top             =   255
         Width           =   975
      End
      Begin VB.TextBox txtDesCli 
         BackColor       =   &H8000000F&
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
         Height          =   300
         Left            =   4845
         MaxLength       =   50
         TabIndex        =   14
         Tag             =   "Descripción"
         Top             =   255
         Width           =   4620
      End
      Begin VB.CommandButton CmdBuscAprox 
         Height          =   1155
         Left            =   9660
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoRemitoCliente.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Buscar Nota de Pedido"
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.CommandButton cmdBuscarCli 
         Height          =   300
         Left            =   4410
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoRemitoCliente.frx":398A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Buscar"
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   375
         Left            =   3360
         TabIndex        =   33
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   53936129
         CurrentDate     =   43367
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   375
         Left            =   5880
         TabIndex        =   34
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   53936129
         CurrentDate     =   43367
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   2745
         TabIndex        =   17
         Top             =   300
         Width           =   525
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   2265
         TabIndex        =   16
         Top             =   765
         Width           =   1005
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   4815
         TabIndex        =   15
         Top             =   780
         Width           =   960
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   2715
      Left            =   90
      TabIndex        =   4
      Top             =   1845
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   4789
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorSel    =   8388736
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
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
      Left            =   2040
      TabIndex        =   30
      Top             =   4500
      Width           =   510
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
      TabIndex        =   23
      Top             =   6795
      Width           =   750
   End
End
Attribute VB_Name = "frmListadoRemitoCliente"
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

Private Sub CmdBuscAprox_Click()
Dim VTotal  As Double
Dim VSaldo  As Double
    GrdModulos.Rows = 1
    lblestado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    GrdModulos.HighLight = flexHighlightNever
    sql = "SELECT RC.RCL_NUMERO, RC.RCL_SUCURSAL,RC.RCL_SALDO, "
    sql = sql & "RC.RCL_FECHA, RC.RCL_TOTAL,C.CLI_RAZSOC, C.CLI_DOMICI,L.LOC_DESCRI, P.PRO_DESCRI"
    sql = sql & " FROM REMITO_CLIENTE RC, CLIENTE C, LOCALIDAD L, PROVINCIA P"
    sql = sql & " WHERE"
    sql = sql & " RC.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=P.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(txtCliente)
    If FechaDesde <> "" Then sql = sql & " AND RC.RCL_FECHA>=" & XDQ(FechaDesde)
    If FechaHasta <> "" Then sql = sql & " AND RC.RCL_FECHA<=" & XDQ(FechaHasta)
    If optPen.Value = True Then
        sql = sql & " AND (RC.EST_CODIGO = 1 OR RC.RCL_SALDO > 0)"
        sql = sql & " AND RC.EST_CODIGO <> 2 " 'Q NO ESTE ANULADO
    End If
    If optDef.Value = True Then
        sql = sql & " AND RC.EST_CODIGO = 3 "
    End If
    If optAnu.Value = True Then
        sql = sql & " AND RC.EST_CODIGO = 2 "
    End If
    sql = sql & " ORDER BY RC.RCL_NUMERO DESC"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    VTotal = 0
    If rec.EOF = False Then
        GrdModulos.HighLight = flexHighlightAlways
        Do While rec.EOF = False
            VTotal = VTotal + IIf(IsNull(rec!RCL_TOTAL), 0, rec!RCL_TOTAL)
            VSaldo = VSaldo + IIf(IsNull(rec!RCL_SALDO), 0, rec!RCL_SALDO)
            GrdModulos.AddItem Format(rec!RCL_SUCURSAL, "0000") & "-" & Format(rec!RCL_NUMERO, "00000000") _
                            & Chr(9) & rec!RCL_FECHA & Chr(9) & Format(rec!RCL_TOTAL, "0.00") & Chr(9) & Format(rec!RCL_SALDO, "0.00") _
                            & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!CLI_DOMICI _
                            & Chr(9) & rec!LOC_DESCRI & Chr(9) & rec!PRO_DESCRI
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
        lblestado.Caption = "Se encontraron " & GrdModulos.Rows - 1 & " remitos"
        txtTotal.Text = Format(VTotal, "0.00")
        txtSaldo.Text = Format(VSaldo, "0.00")
    Else
        lblestado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
    End If
    'lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    rec.Close
End Sub
Private Sub cmdListar_Click()
'    Rep.WindowState = crptNormal
'    Rep.WindowBorderStyle = crptNoBorder
'    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SIHDG"
'    Rep.SelectionFormula = ""
'    Rep.Formulas(0) = ""
'    Rep.Formulas(1) = ""
'    Rep.Formulas(2) = ""
'
'    If optGeneralTodos.Value = True Then
'        Rep.SelectionFormula = ""
'        If txtCliente.Text <> "" Then
'            Rep.SelectionFormula = "{REMITO_CLIENTE.CLI_CODIGO}=" & txtCliente.Text
'            Rep.Formulas(0) = "CLIENTE='" & "Cliente: " & txtDesCli & "'"
'        Else
'            Rep.Formulas(0) = "CLIENTE='" & "Cliente: Todos'"
'        End If
'        If FechaDesde.value <> "" Then
'            If Rep.SelectionFormula = "" Then
'                Rep.SelectionFormula = " {REMITO_CLIENTE.RCL_FECHA}>= DATE (" & Mid(FechaDesde.value, 7, 4) & "," & Mid(FechaDesde.value, 4, 2) & "," & Mid(FechaDesde.value, 1, 2) & ")"
'            Else
'                Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_CLIENTE.RCL_FECHA}>= DATE (" & Mid(FechaDesde.value, 7, 4) & "," & Mid(FechaDesde.value, 4, 2) & "," & Mid(FechaDesde.value, 1, 2) & ")"
'            End If
'        End If
'
'        If FechaHasta.value <> "" Then
'            If Rep.SelectionFormula = "" Then
'                Rep.SelectionFormula = " {REMITO_CLIENTE.RCL_FECHA}<= DATE (" & Mid(FechaHasta.value, 7, 4) & "," & Mid(FechaHasta.value, 4, 2) & "," & Mid(FechaHasta.value, 1, 2) & ")"
'            Else
'                Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_CLIENTE.RCL_FECHA}<= DATE (" & Mid(FechaHasta.value, 7, 4) & "," & Mid(FechaHasta.value, 4, 2) & "," & Mid(FechaHasta.value, 1, 2) & ")"
'            End If
'        End If
'        If optPen.Value = True Then
'            If Rep.SelectionFormula = "" Then
'                Rep.SelectionFormula = "{ESTADO_DOCUMENTO.EST_CODIGO}= 1 "
'            Else
'                Rep.SelectionFormula = Rep.SelectionFormula & "AND {ESTADO_DOCUMENTO.EST_CODIGO}= 1 "
'            End If
'        Else
'            If optDef.Value = True Then
'                If Rep.SelectionFormula = "" Then
'                    Rep.SelectionFormula = "{ESTADO_DOCUMENTO.EST_CODIGO}= 3 "
'                Else
'                    Rep.SelectionFormula = Rep.SelectionFormula & "AND {ESTADO_DOCUMENTO.EST_CODIGO}= 3 "
'                End If
'            Else
'                If optAnu.Value = True Then
'                    If Rep.SelectionFormula = "" Then
'                        Rep.SelectionFormula = "{ESTADO_DOCUMENTO.EST_CODIGO}= 2 "
'                    Else
'                        Rep.SelectionFormula = Rep.SelectionFormula & "{ESTADO_DOCUMENTO.EST_CODIGO}= 2 "
'                    End If
'                End If
'
'            End If
'        End If
'
'
'        If FechaDesde.value <> "" And FechaHasta.value <> "" Then
'            Rep.Formulas(2) = "FECHA='" & "Desde: " & FechaDesde.value & "   Hasta: " & FechaHasta.value & "'"
'        ElseIf FechaDesde.value <> "" And FechaHasta.Value = date Then
'            Rep.Formulas(2) = "FECHA='" & "Desde: " & FechaDesde.value & "   Hasta: " & Date & "'"
'        ElseIf FechaDesde.Value = date And FechaHasta.value <> "" Then
'            Rep.Formulas(2) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.value & "'"
'        End If
'
''            Rep.WindowTitle = "Remito de Cliente - General - Por cuenta y orden de Terceros"
''            Rep.ReportFileName = DRIVE & DirReport & "rptremitoclientegeneralTerceros.rpt"
'
'        Rep.WindowTitle = "Remito de Cliente - General..."
'        Rep.ReportFileName = DRIVE & DirReport & "rptremitoclientegeneral.rpt"
'    End If
'
'    If optDetallado.Value = True Then
'        If GrdModulos.TextMatrix(GrdModulos.RowSel, 0) = "" Then
'            MsgBox "Debe seleccionar un Remito", vbExclamation, TIT_MSGBOX
'            chkCliente.SetFocus
'            Exit Sub
'        End If
'        Rep.SelectionFormula = ""
'        Rep.SelectionFormula = "{REMITO_CLIENTE.RCL_NUMERO}=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8)) _
'                               & " AND {REMITO_CLIENTE.RCL_SUCURSAL}=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 4))
'
'        Rep.WindowTitle = "Remito de Cliente - Detallado..."
'        Rep.ReportFileName = DRIVE & DirReport & "rptremitoclientedetalle.rpt"
'    End If
'
'    If optPantalla.Value = True Then
'        Rep.Destination = crptToWindow
'    ElseIf optImpresora.Value = True Then
'        Rep.Destination = crptToPrinter
'    End If
'     Rep.Action = 1
'
'     Rep.SelectionFormula = ""
'     Rep.Formulas(0) = ""
'     Rep.Formulas(1) = ""
'     Rep.Formulas(2) = ""
 'Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SIELECTROCENTRO"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
    Rep.WindowState = crptMaximized
    If optGeneralTodos.Value = True Then
        'Rep.SelectionFormula = ""
        'Rep.SelectionFormula = "{REMITO_CLIENTE.EST_CODIGO} <> 2"
        If txtCliente.Text <> "" Then
            Rep.SelectionFormula = "{REMITO_CLIENTE.CLI_CODIGO}=" & txtCliente.Text
            Rep.Formulas(0) = "CLIENTE='" & "Cliente: " & txtDesCli & "'"
        Else
            Rep.Formulas(0) = "CLIENTE='" & "Cliente: Todos'"
        End If
        If FechaDesde.Value <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {REMITO_CLIENTE.RCL_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_CLIENTE.RCL_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
            End If
        End If
        If optPen.Value = True Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " ({REMITO_CLIENTE.EST_CODIGO}= 1 OR {REMITO_CLIENTE.RCL_SALDO} > 0)"
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_CLIENTE.EST_CODIGO} <> 2"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND ({REMITO_CLIENTE.EST_CODIGO}= 1 OR {REMITO_CLIENTE.RCL_SALDO} > 0)"
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_CLIENTE.EST_CODIGO} <> 2"
            End If
        End If
        If optDef.Value = True Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {REMITO_CLIENTE.EST_CODIGO}= 3 "
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_CLIENTE.EST_CODIGO}= 3 "
            End If
        End If
        If optDef.Value = True Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {REMITO_CLIENTE.EST_CODIGO}= 2 "
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_CLIENTE.EST_CODIGO}= 3 "
            End If
        End If
        
        If FechaHasta.Value <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {REMITO_CLIENTE.RCL_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_CLIENTE.RCL_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
            End If
        End If
        
        If FechaDesde.Value <> "" And FechaHasta.Value <> "" Then
            Rep.Formulas(2) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
        ElseIf FechaDesde.Value <> "" And FechaHasta.Value = Date Then
            Rep.Formulas(2) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
        ElseIf FechaDesde.Value = Date And FechaHasta.Value <> "" Then
            Rep.Formulas(2) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
        End If
        
            Rep.WindowTitle = "Remito de Cliente - General - Por cuenta y orden de Terceros"
            Rep.ReportFileName = DRIVE & DirReport & "rptremitoclientegeneralTerceros.rpt"
    
        Rep.WindowTitle = "Remito de Cliente - General..."
        Rep.ReportFileName = DRIVE & DirReport & "rptremitoclientegeneral.rpt"
    End If
    
    If optDetallado.Value = True Then
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 0) = "" Then
            MsgBox "Debe seleccionar un Remito", vbExclamation, TIT_MSGBOX
            chkCliente.SetFocus
            Exit Sub
        End If
        Rep.SelectionFormula = ""
        'Rep.SelectionFormula = "{REMITO_CLIENTE.RCL_NUMERO}=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8)) _
                               & " AND {REMITO_CLIENTE.RCL_SUCURSAL}=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 4))
 
        'Rep.SelectionFormula = "{REMITO_CLIENTE.EST_CODIGO} <> 2"
        If txtCliente.Text <> "" Then
            Rep.SelectionFormula = "{REMITO_CLIENTE.CLI_CODIGO}=" & txtCliente.Text
           ' Rep.Formulas(0) = "CLIENTE='" & "Cliente: " & txtDesCli & "'"
        'Else
           ' Rep.Formulas(0) = "CLIENTE='" & "Cliente: Todos'"
        End If
        If optPen.Value = True Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " ({REMITO_CLIENTE.EST_CODIGO}= 1 OR {REMITO_CLIENTE.RCL_SALDO} > 0)"
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_CLIENTE.EST_CODIGO} <> 2"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND ({REMITO_CLIENTE.EST_CODIGO}= 1 OR {REMITO_CLIENTE.RCL_SALDO} > 0)"
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_CLIENTE.EST_CODIGO} <> 2"
            End If
        End If
        If optDef.Value = True Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {REMITO_CLIENTE.EST_CODIGO}= 3 "
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_CLIENTE.EST_CODIGO}= 3 "
            End If
        End If
        If optDef.Value = True Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {REMITO_CLIENTE.EST_CODIGO}= 2 "
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_CLIENTE.EST_CODIGO}= 3 "
            End If
        End If
        If FechaDesde.Value <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {REMITO_CLIENTE.RCL_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_CLIENTE.RCL_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
            End If
        End If
        
        If FechaHasta.Value <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {REMITO_CLIENTE.RCL_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {REMITO_CLIENTE.RCL_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
            End If
        End If
        
        Rep.WindowTitle = "Remito de Cliente - Detallado..."
        Rep.ReportFileName = DRIVE & DirReport & "rptremitoclientedetalle.rpt"
    End If
    
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
     
     Rep.Action = 1
     
     Rep.SelectionFormula = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
     Rep.Formulas(2) = ""
End Sub

Private Sub cmdBuscarCli_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtCliente.Text = frmBuscar.grdBuscar.Text
        txtCliente.SetFocus
        txtCliente_LostFocus
    Else
        txtCliente.SetFocus
    End If
End Sub
Private Sub CmdNuevo_Click()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    FechaDesde.Value = Date
    FechaHasta.Value = Date
    GrdModulos.Rows = 1
    GrdModulos.Rows = 2
    txtCliente.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    chkCliente.Value = Unchecked
    chkFecha.Value = Unchecked
    optDetallado.Value = True
    optPantalla.Value = True
    chkCliente.SetFocus
    lblestado.Caption = ""
    txtTotal.Text = ""
    txtSaldo.Text = ""
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
       
    txtCliente.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    cmdBuscarCli.Enabled = False
    GrdModulos.Rows = 1
'    lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    lblestado.Caption = ""

    Call Centrar_pantalla(Me)
    GrdModulos.FormatString = "^Número|^Fecha|Importe|Saldo|Cliente|Domicilio|Localidad|Provincia"
    GrdModulos.ColWidth(0) = 1300
    GrdModulos.ColWidth(1) = 1200
    GrdModulos.ColWidth(2) = 1100
    GrdModulos.ColWidth(3) = 1100
    GrdModulos.ColWidth(4) = 3200
    GrdModulos.ColWidth(5) = 3200
    GrdModulos.ColWidth(6) = 3200
    GrdModulos.ColWidth(7) = 3200
    GrdModulos.Rows = 2
    '------------------------------------
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub chkCliente_Click()
    If chkCliente.Value = Checked Then
        txtCliente.Enabled = True
        cmdBuscarCli.Enabled = True
    Else
        txtCliente.Enabled = False
        cmdBuscarCli.Enabled = False
    End If
End Sub

Private Sub chkFecha_Click()
    If chkFecha.Value = Checked Then
        FechaDesde.Enabled = True
        FechaHasta.Enabled = True
    Else
        FechaDesde.Enabled = False
        FechaHasta.Enabled = False
    End If
End Sub
Private Sub txtCliente_Change()
    If txtCliente.Text = "" Then
        txtDesCli.Text = ""
    End If
End Sub

Private Sub txtCliente_GotFocus()
    SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCliente_LostFocus()
    If txtCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_RAZSOC FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(txtCliente)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesCli.Text = rec!CLI_RAZSOC
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtDesCli.Text = ""
            txtCliente.SetFocus
        End If
        rec.Close
    End If
    If chkFecha.Value = Unchecked _
        And ActiveControl.Name <> "cmdBuscarCli" _
        And ActiveControl.Name <> "cmdNuevo" And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub
Private Sub cmdSalir_Click()
    Set frmListadoRemitoCliente = Nothing
    Unload Me
End Sub


