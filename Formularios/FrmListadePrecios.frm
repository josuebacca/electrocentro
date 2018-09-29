VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmListadePrecios 
   Caption         =   "Lista de Precios"
   ClientHeight    =   7395
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmListadePrecios.frx":0000
   ScaleHeight     =   7395
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImportar 
      Height          =   495
      Left            =   11400
      Picture         =   "FrmListadePrecios.frx":22A2
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Importar Lista de Precios"
      Top             =   5160
      Width           =   450
   End
   Begin VB.CommandButton cmdPreCompra 
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   50
      Top             =   4560
      Width           =   450
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LISTA"
      Height          =   495
      Left            =   3840
      TabIndex        =   45
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      DisabledPicture =   "FrmListadePrecios.frx":25AC
      Height          =   750
      Left            =   9255
      Picture         =   "FrmListadePrecios.frx":28B6
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   6600
      Width           =   870
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      DisabledPicture =   "FrmListadePrecios.frx":2BC0
      Height          =   750
      Left            =   10125
      Picture         =   "FrmListadePrecios.frx":2ECA
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6600
      Width           =   870
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "FrmListadePrecios.frx":566C
      Height          =   750
      Left            =   8370
      Picture         =   "FrmListadePrecios.frx":5976
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6600
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "FrmListadePrecios.frx":5C80
      Height          =   750
      Left            =   10995
      Picture         =   "FrmListadePrecios.frx":5F8A
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6600
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "FrmListadePrecios.frx":6294
      Height          =   750
      Left            =   6630
      Picture         =   "FrmListadePrecios.frx":70D6
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6600
      Width           =   870
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      DisabledPicture =   "FrmListadePrecios.frx":7F18
      Height          =   750
      Left            =   7500
      Picture         =   "FrmListadePrecios.frx":8D5A
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6600
      Width           =   870
   End
   Begin VB.TextBox txtObservaciones2 
      Height          =   315
      Left            =   1500
      TabIndex        =   15
      Top             =   6120
      Width           =   9870
   End
   Begin VB.TextBox txtObservaciones1 
      Height          =   315
      Left            =   1500
      TabIndex        =   14
      Top             =   5790
      Width           =   9870
   End
   Begin VB.Frame frebotopc 
      Height          =   2445
      Left            =   11350
      TabIndex        =   33
      Top             =   2025
      Width           =   495
      Begin VB.CommandButton cmdAgregar 
         Height          =   615
         Left            =   0
         Picture         =   "FrmListadePrecios.frx":90E4
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Agegar producto"
         Top             =   1260
         Width           =   465
      End
      Begin VB.CommandButton cmdQuitar 
         Height          =   570
         Left            =   0
         Picture         =   "FrmListadePrecios.frx":93EE
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Quitar Producto"
         Top             =   1860
         Width           =   465
      End
      Begin VB.CommandButton cmdPrecios 
         Enabled         =   0   'False
         Height          =   915
         Left            =   15
         Picture         =   "FrmListadePrecios.frx":96F8
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Modificar Precios"
         Top             =   105
         Width           =   465
      End
   End
   Begin TabDlg.SSTab TabPrecios 
      Height          =   2265
      Left            =   3060
      TabIndex        =   29
      Top             =   3060
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   3995
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FrmListadePrecios.frx":9FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   2160
         Left            =   120
         TabIndex        =   30
         Top             =   15
         Width           =   3825
         Begin VB.OptionButton optPVenta 
            Caption         =   "Precio Venta"
            Height          =   225
            Left            =   480
            TabIndex        =   44
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optPCompra 
            Caption         =   "Precio Compra"
            Height          =   240
            Left            =   1875
            TabIndex        =   43
            Top             =   240
            Width           =   1440
         End
         Begin VB.CommandButton cmdSalirP 
            Caption         =   "&Cancelar"
            Height          =   465
            Left            =   2100
            TabIndex        =   22
            Top             =   1545
            Width           =   900
         End
         Begin VB.TextBox txtActual 
            Height          =   300
            Left            =   1590
            TabIndex        =   20
            Top             =   1050
            Width           =   1140
         End
         Begin VB.CommandButton cmdAceptarP 
            Caption         =   "&Aceptar"
            Height          =   465
            Left            =   1155
            TabIndex        =   21
            Top             =   1545
            Width           =   900
         End
         Begin VB.TextBox txtAnterior 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1605
            TabIndex        =   19
            Top             =   585
            Width           =   1125
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Precio Actual:"
            Height          =   195
            Left            =   225
            TabIndex        =   32
            Top             =   1080
            Width           =   990
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Precio Anterior:"
            Height          =   195
            Left            =   135
            TabIndex        =   31
            Top             =   600
            Width           =   1080
         End
      End
   End
   Begin VB.Frame freLista 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   30
      TabIndex        =   23
      Top             =   0
      Width           =   11820
      Begin TabDlg.SSTab tabLista 
         Height          =   855
         Left            =   6720
         TabIndex        =   47
         Top             =   120
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   1508
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Maquinarias"
         TabPicture(0)   =   "FrmListadePrecios.frx":9FDE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "cboListaPrecio"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Repuestos"
         TabPicture(1)   =   "FrmListadePrecios.frx":9FFA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cboLPrecioRep"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Adicionales"
         TabPicture(2)   =   "FrmListadePrecios.frx":A016
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cboAdicionales"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.ComboBox cboAdicionales 
            Height          =   315
            Left            =   -74760
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   420
            Width           =   4305
         End
         Begin VB.ComboBox cboListaPrecio 
            Height          =   315
            Left            =   260
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   400
            Width           =   4185
         End
         Begin VB.ComboBox cboLPrecioRep 
            Height          =   315
            Left            =   -74740
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   400
            Width           =   4185
         End
      End
      Begin VB.ComboBox cbodescri 
         Height          =   315
         Left            =   6735
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   4185
      End
      Begin VB.TextBox txtcodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2505
         TabIndex        =   0
         Top             =   255
         Width           =   1110
      End
      Begin VB.TextBox TxtDescriB 
         Height          =   285
         Left            =   6735
         MaxLength       =   40
         TabIndex        =   2
         Top             =   255
         Width           =   3660
      End
      Begin VB.CommandButton CmdBuscAprox 
         Height          =   360
         Left            =   10350
         MaskColor       =   &H000000FF&
         Picture         =   "FrmListadePrecios.frx":A032
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Buscar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker Fecha1 
         Height          =   315
         Left            =   2520
         TabIndex        =   53
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   53936129
         CurrentDate     =   43367
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   1755
         TabIndex        =   28
         Top             =   285
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Vigencia desde:"
         Height          =   195
         Left            =   1215
         TabIndex        =   27
         Top             =   660
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lista de Precios"
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
         Index           =   0
         Left            =   5160
         TabIndex        =   26
         Top             =   300
         Width           =   1380
      End
   End
   Begin VB.Frame freOpciones 
      Caption         =   "Opciones de Consulta"
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
      Left            =   30
      TabIndex        =   24
      Top             =   1050
      Width           =   11820
      Begin VB.CheckBox chkRubro 
         Caption         =   "Rubro"
         Height          =   285
         Left            =   6015
         TabIndex        =   6
         Top             =   240
         Width           =   825
      End
      Begin VB.CheckBox chkLinea 
         Caption         =   "Línea"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   615
         Width           =   780
      End
      Begin VB.CheckBox chkProducto 
         Caption         =   "Producto"
         Height          =   225
         Left            =   360
         TabIndex        =   4
         Top             =   270
         Width           =   990
      End
      Begin VB.CommandButton cmdfiltrar 
         Caption         =   "&Filtrar"
         Height          =   690
         Left            =   11085
         Picture         =   "FrmListadePrecios.frx":A33C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   225
         Width           =   570
      End
      Begin VB.ComboBox cboRep 
         Height          =   315
         Left            =   7005
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   585
         Width           =   3870
      End
      Begin VB.ComboBox cbolinea 
         Height          =   315
         Left            =   1485
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   585
         Width           =   4020
      End
      Begin VB.ComboBox cborubro 
         Height          =   315
         Left            =   6990
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   225
         Width           =   3870
      End
      Begin VB.CheckBox chkRepres 
         Caption         =   "Marca"
         Height          =   255
         Left            =   6015
         TabIndex        =   7
         Top             =   615
         Width           =   1380
      End
      Begin VB.TextBox txtproducto 
         Height          =   315
         Left            =   1485
         TabIndex        =   9
         Top             =   225
         Width           =   4005
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   3735
      Left            =   0
      TabIndex        =   13
      Top             =   2025
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorSel    =   8388736
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   5160
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "FrmListadePrecios.frx":CADE
      Height          =   750
      Left            =   5760
      Picture         =   "FrmListadePrecios.frx":CDE8
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6600
      Width           =   870
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "2 - Observaciones:"
      Height          =   195
      Left            =   60
      TabIndex        =   35
      Top             =   6165
      Width           =   1335
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "1 - Observaciones:"
      Height          =   195
      Left            =   60
      TabIndex        =   34
      Top             =   5835
      Width           =   1335
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
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
      Left            =   105
      TabIndex        =   25
      Top             =   6930
      Width           =   750
   End
End
Attribute VB_Name = "FrmListadePrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim resp As Integer
Dim CodigoProducto As String
Dim inicio As Integer
Dim j As Integer
Dim Rec1 As ADODB.Recordset

Private Sub cboAdicionales_Click()
    ''' adicionales
Dim BAND As Integer
    Screen.MousePointer = vbHourglass
    BAND = 0


        sql = "SELECT DISTINCT  P.PTO_DESCRI,L.LNA_DESCRI,R.RUB_DESCRI,"
        sql = sql & " RE.TPRE_DESCRI,P.PTO_PRECIO, P.PTO_PRECIOC, LP.LIS_CODIGO, LP.LIS_FECHA,P.PTO_CODIGO,P.PTO_PRECIVA"
        sql = sql & " FROM LISTA_PRECIO LP,PRODUCTO P,LINEAS L,RUBROS R, TIPO_PRESENTACION RE"
        ',DETALLE_LISTA_PRECIO D "
        sql = sql & " WHERE P.LIS_CODIGO = LP.LIS_CODIGO "
        'AND D.LIS_CODIGO = LP.LIS_CODIGO "
        sql = sql & " AND P.LIS_CODIGO = " & cboAdicionales.ItemData(cboAdicionales.ListIndex) & " "
        sql = sql & " AND L.LNA_CODIGO = P.LNA_CODIGO "
        sql = sql & " AND R.RUB_CODIGO = P.RUB_CODIGO  AND P.TPRE_CODIGO = RE.TPRE_CODIGO "
        'sql = sql & " AND LP.LIS_CODIGO LIKE '" & Me.cbodescri.Text & "%' "
        'sql = sql & " AND LP.LIS_DESCRI LIKE '" & Me.cbodescri.Text & "%' "
        sql = sql & "ORDER BY P.PTO_DESCRI"
 '   End If
    
    lblestado.Caption = "Buscando..."
    
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        GrdModulos.Rows = 1
        Do While Not Rec1.EOF
           GrdModulos.AddItem Rec1!PTO_CODIGO & Chr(9) & Rec1.Fields(0) & Chr(9) & Rec1.Fields(1) & Chr(9) & _
                              Rec1.Fields(2) & Chr(9) & Rec1.Fields(3) & Chr(9) & _
                              IIf(IsNull(Rec1.Fields(4)), "0,00", Valido_Importe(Rec1.Fields(4))) & Chr(9) & _
                              IIf(IsNull(Rec1.Fields(5)), "0,00", Valido_Importe(Rec1.Fields(5))) & Chr(9) & _
                              IIf(IsNull(Rec1!PTO_PRECIVA), "0,00", Valido_Importe(IIf(IsNull(Rec1!PTO_PRECIVA), 0, Rec1!PTO_PRECIVA)))
    
            Rec1.MoveNext
        Loop
        Rec1.MoveFirst
        If BAND = 0 Then
            TxtCodigo.Text = Rec1.Fields(6)
            Fecha1.Value = Rec1.Fields(7)
        Else
            TxtCodigo.Text = ""
            Fecha1.Value = Date
        End If
    Else
        lblestado.Caption = ""
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbInformation, TIT_MSGBOX
        GrdModulos.Rows = 1
        'me.txt
'        Me.cbodescri.SetFocus
    End If
    lblestado.Caption = ""
    Text1.Text = GrdModulos.Rows - 1
    lblestado.Caption = "Se encontraron " + Text1.Text + " registros"
    Rec1.Close
    Screen.MousePointer = vbNormal
End Sub

Private Sub cbodescri_Click()
'    Dim BAND As Integer
'    Screen.MousePointer = vbHourglass
'    BAND = 0
''    If cbodescri.ListIndex = 0 Then ' Busca en los productos
''        BAND = 1
''        sql = "SELECT P.PTO_DESCRI, L.LNA_DESCRI,R.RUB_DESCRI,RE.TPRE_DESCRI,"
''        sql = sql & "P.PTO_PRECIO,P.PTO_PRECIOC,P.PTO_CODIGO "
''        sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L, TIPO_PRESENTACION RE"
''        sql = sql & " WHERE"
''        sql = sql & " P.LNA_CODIGO=L.LNA_CODIGO AND P.RUB_CODIGO=R.RUB_CODIGO"
''        sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
''        sql = sql & " AND P.TPRE_CODIGO=RE.TPRE_CODIGO"
''        sql = sql & " AND P.PTO_ESTADO=1"
''        sql = sql & " ORDER BY P.PTO_DESCRI"
''    Else 'Busca en la Lista de Precios
'
'        sql = "SELECT DISTINCT  P.PTO_DESCRI,L.LNA_DESCRI,R.RUB_DESCRI,"
'        sql = sql & " RE.TPRE_DESCRI,P.PTO_PRECIO, P.PTO_PRECIOC, LP.LIS_CODIGO, LP.LIS_FECHA,P.PTO_CODIGO"
'        sql = sql & " FROM LISTA_PRECIO LP,PRODUCTO P,LINEAS L,RUBROS R, TIPO_PRESENTACION RE"
'        ',DETALLE_LISTA_PRECIO D "
'        sql = sql & " WHERE P.LIS_CODIGO = LP.LIS_CODIGO "
'        'AND D.LIS_CODIGO = LP.LIS_CODIGO "
'        If tabLista.Tab = 0 Then
'            sql = sql & " AND P.LIS_CODIGO = " & cboListaPrecio.ItemData(cboListaPrecio.ListIndex) & " "
'        Else
'            sql = sql & " AND P.LIS_CODIGO = " & cboLPrecioRep.ItemData(cboLPrecioRep.ListIndex) & " "
'        End If
'        sql = sql & " AND L.LNA_CODIGO = P.LNA_CODIGO "
'        sql = sql & " AND R.RUB_CODIGO = P.RUB_CODIGO  AND P.TPRE_CODIGO = RE.TPRE_CODIGO "
'        'sql = sql & " AND LP.LIS_CODIGO LIKE '" & Me.cbodescri.Text & "%' "
'        'sql = sql & " AND LP.LIS_DESCRI LIKE '" & Me.cbodescri.Text & "%' "
'        sql = sql & "ORDER BY P.PTO_DESCRI"
' '   End If
'
'    lblEstado.Caption = "Buscando..."
'
'
'    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If Rec1.EOF = False Then
'        GrdModulos.Rows = 1
'        Do While Not Rec1.EOF
'           GrdModulos.AddItem Rec1!PTO_CODIGO & Chr(9) & Rec1.Fields(0) & Chr(9) & Rec1.Fields(1) & Chr(9) & _
'                              Rec1.Fields(2) & Chr(9) & Rec1.Fields(3) & Chr(9) & _
'                              IIf(IsNull(Rec1.Fields(4)), "0,00", Valido_Importe(Rec1.Fields(4))) & Chr(9) & _
'                              IIf(IsNull(Rec1.Fields(5)), "0,00", Valido_Importe(Rec1.Fields(5)))
'
'            Rec1.MoveNext
'        Loop
'        Rec1.MoveFirst
'        If BAND = 0 Then
'            TxtCodigo.Text = Rec1.Fields(6)
'            Fecha1.Value = Rec1.Fields(7)
'        Else
'            TxtCodigo.Text = ""
'            Fecha1.Value= date
'        End If
'    Else
'        lblEstado.Caption = ""
'        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbInformation, TIT_MSGBOX
'        GrdModulos.Rows = 1
'        'me.txt
''        Me.cbodescri.SetFocus
'    End If
'    lblEstado.Caption = ""
'    Text1.Text = GrdModulos.Rows - 1
'    lblEstado.Caption = "Se encontraron " + Text1.Text + " registros"
'    Rec1.Close
'    Screen.MousePointer = vbNormal
'
End Sub

Private Sub cbodescri_GotFocus()
    SelecTexto TxtDescriB
End Sub

Private Sub cboLinea_LostFocus()
    If cboLinea.ListIndex <> -1 Then
        chkLinea.Value = 1
        cboRubro.Clear
        cargocboRubro (cboLinea.ItemData(cboLinea.ListIndex))
    Else
        cboRubro.Clear
        cargocboRubro (-1)
        chkRubro.Value = 0
    End If
End Sub

Private Sub cboListaPrecio_click()
''' maquinarias
Dim BAND As Integer
    Screen.MousePointer = vbHourglass
    BAND = 0


        sql = "SELECT DISTINCT  P.PTO_DESCRI,L.LNA_DESCRI,R.RUB_DESCRI,"
        sql = sql & " RE.TPRE_DESCRI,P.PTO_PRECIO, P.PTO_PRECIOC, LP.LIS_CODIGO, LP.LIS_FECHA,P.PTO_CODIGO,P.PTO_PRECIVA"
        sql = sql & " FROM LISTA_PRECIO LP,PRODUCTO P,LINEAS L,RUBROS R, TIPO_PRESENTACION RE"
        ',DETALLE_LISTA_PRECIO D "
        sql = sql & " WHERE P.LIS_CODIGO = LP.LIS_CODIGO "
        'AND D.LIS_CODIGO = LP.LIS_CODIGO "
        sql = sql & " AND P.LIS_CODIGO = " & cboListaPrecio.ItemData(cboListaPrecio.ListIndex) & " "
        sql = sql & " AND L.LNA_CODIGO = P.LNA_CODIGO "
        sql = sql & " AND R.RUB_CODIGO = P.RUB_CODIGO  AND P.TPRE_CODIGO = RE.TPRE_CODIGO "
        'sql = sql & " AND LP.LIS_CODIGO LIKE '" & Me.cbodescri.Text & "%' "
        'sql = sql & " AND LP.LIS_DESCRI LIKE '" & Me.cbodescri.Text & "%' "
        sql = sql & "ORDER BY P.PTO_DESCRI"
 '   End If
    
    lblestado.Caption = "Buscando..."
    
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        GrdModulos.Rows = 1
        Do While Not Rec1.EOF
           GrdModulos.AddItem Rec1!PTO_CODIGO & Chr(9) & Rec1.Fields(0) & Chr(9) & Rec1.Fields(1) & Chr(9) & _
                              Rec1.Fields(2) & Chr(9) & Rec1.Fields(3) & Chr(9) & _
                              IIf(IsNull(Rec1.Fields(4)), "0,00", Valido_Importe(Rec1.Fields(4))) & Chr(9) & _
                              IIf(IsNull(Rec1.Fields(5)), "0,00", Valido_Importe(Rec1.Fields(5))) & Chr(9) & _
                              IIf(IsNull(Rec1!PTO_PRECIVA), "0,00", Valido_Importe(IIf(IsNull(Rec1!PTO_PRECIVA), 0, Rec1!PTO_PRECIVA)))
    
            Rec1.MoveNext
        Loop
        Rec1.MoveFirst
        If BAND = 0 Then
            TxtCodigo.Text = Rec1.Fields(6)
            Fecha1.Value = Rec1.Fields(7)
        Else
            TxtCodigo.Text = ""
            Fecha1.Value = Date
        End If
    Else
        lblestado.Caption = ""
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbInformation, TIT_MSGBOX
        GrdModulos.Rows = 1
        'me.txt
'        Me.cbodescri.SetFocus
    End If
    lblestado.Caption = ""
    Text1.Text = GrdModulos.Rows - 1
    lblestado.Caption = "Se encontraron " + Text1.Text + " registros"
    Rec1.Close
    Screen.MousePointer = vbNormal

End Sub

Private Sub cboLPrecioRep_Click()
' aca pongo el clic de repuestos
Dim BAND As Integer
    Screen.MousePointer = vbHourglass
    BAND = 0
    
    sql = "SELECT DISTINCT TOP 800 P.PTO_DESCRI,L.LNA_DESCRI,R.RUB_DESCRI,"
    sql = sql & " RE.TPRE_DESCRI,P.PTO_PRECIO, P.PTO_PRECIOC, LP.LIS_CODIGO, LP.LIS_FECHA,P.PTO_CODIGO,P.PTO_PRECIVA"
    sql = sql & " FROM LISTA_PRECIO LP,PRODUCTO P,LINEAS L,RUBROS R, TIPO_PRESENTACION RE"
    ',DETALLE_LISTA_PRECIO D "
    sql = sql & " WHERE P.LIS_CODIGO = LP.LIS_CODIGO "
    'AND D.LIS_CODIGO = LP.LIS_CODIGO "
    sql = sql & " AND P.LIS_CODIGO = " & cboLPrecioRep.ItemData(cboLPrecioRep.ListIndex) & " "
    sql = sql & " AND L.LNA_CODIGO = P.LNA_CODIGO "
    sql = sql & " AND R.RUB_CODIGO = P.RUB_CODIGO  AND P.TPRE_CODIGO = RE.TPRE_CODIGO "
    'sql = sql & " AND LP.LIS_CODIGO LIKE '" & Me.cbodescri.Text & "%' "
    'sql = sql & " AND LP.LIS_DESCRI LIKE '" & Me.cbodescri.Text & "%' "
    sql = sql & "ORDER BY P.PTO_DESCRI"
 '   End If
    
    lblestado.Caption = "Buscando..."
    
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        GrdModulos.Rows = 1
        Do While Not Rec1.EOF
           GrdModulos.AddItem Rec1!PTO_CODIGO & Chr(9) & Rec1.Fields(0) & Chr(9) & Rec1.Fields(1) & Chr(9) & _
                              Rec1.Fields(2) & Chr(9) & Rec1.Fields(3) & Chr(9) & _
                              IIf(IsNull(Rec1.Fields(4)), "0,00", Valido_Importe(Rec1.Fields(4))) & Chr(9) & _
                              IIf(IsNull(Rec1.Fields(5)), "0,00", Valido_Importe(Rec1.Fields(5))) & Chr(9) & _
                              IIf(IsNull(Rec1!PTO_PRECIVA), "0,00", Valido_Importe(IIf(IsNull(Rec1!PTO_PRECIVA), 0, Rec1!PTO_PRECIVA)))
    
            Rec1.MoveNext
        Loop
        Rec1.MoveFirst
        If BAND = 0 Then
            TxtCodigo.Text = Rec1.Fields(6)
            Fecha1.Value = Rec1.Fields(7)
        Else
            TxtCodigo.Text = ""
            Fecha1.Value = Date
        End If
    Else
        lblestado.Caption = ""
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbInformation, TIT_MSGBOX
        GrdModulos.Rows = 1
        'me.txt
'        Me.cbodescri.SetFocus
    End If
    lblestado.Caption = ""
    Text1.Text = GrdModulos.Rows - 1
    lblestado.Caption = "Se encontraron " + Text1.Text + " registros"
    Rec1.Close
    Screen.MousePointer = vbNormal

End Sub

Private Sub cboRep_LostFocus()
    If cboRep.ListIndex <> -1 Then
        chkRepres.Value = 1
    End If
End Sub

Private Sub cboRubro_Change()
    If cboRubro.ListIndex <> -1 Then
        chkRubro.Value = 1
    End If
End Sub

Private Sub chklinea_Click()
    If chkLinea.Value = 1 Then
        cboLinea.Enabled = True
        cboLinea.ListIndex = 0
        cboLinea.SetFocus
    Else
         cboLinea.Enabled = False
        cboLinea.ListIndex = -1
    End If
End Sub

Private Sub chkProducto_Click()
    If chkProducto.Value = 1 Then
        txtproducto.Enabled = True
        txtproducto.SetFocus
    Else
        txtproducto.Enabled = False
        txtproducto.Text = ""
    End If
End Sub

Private Sub chkRepres_Click()
    If (chkLinea.Value = 1) And (chkRubro.Value = 1) Then
        cargocboRepres cboLinea.ItemData(cboLinea.ListIndex), cboRubro.ItemData(cboRubro.ListIndex)
    Else
        If chkLinea.Value = 1 Then
            cargocboRepres cboLinea.ItemData(cboLinea.ListIndex), -1
        Else
            If chkRubro.Value = 1 Then
                cargocboRepres -1, cboRubro.ItemData(cboRubro.ListIndex)
            Else
                    cargocboRepres -1, -1
            End If
        End If
        
    End If
    If chkRepres.Value = 1 Then
        cboRep.Enabled = True
'        cboRep.ListIndex = 0
        cboRep.SetFocus
    Else
        cboRep.Enabled = False
        cboRep.ListIndex = -1
    End If
End Sub

Private Sub chkRubro_Click()
    If chkLinea.Value = 0 Then
        cargocboRubro (-1)
    Else
        cargocboRubro (cboLinea.ItemData(cboLinea.ListIndex))
    End If
    If chkRubro.Value = 1 Then
        cboRubro.Enabled = True
        cboRubro.ListIndex = 0
        cboRubro.SetFocus
    Else
        cboRubro.Enabled = False
        cboRubro.ListIndex = -1
    End If
    
End Sub

Private Sub cmdAceptarP_Click()
Dim PRECIVA As String
Dim TOTAL As String
Dim resp As Integer
Dim precio1 As String
Dim precio2 As String

    freLista.Enabled = True
    freOpciones.Enabled = True
    'frebotones.Enabled = True
    frebotopc.Enabled = True
    
    If optPVenta.Value = True Then
        GrdModulos.TextMatrix(GrdModulos.row, 5) = Valido_Importe(txtActual.Text)
        TOTAL = Valido_Importe(txtActual.Text)
        If FrmListadePrecios.GrdModulos.TextMatrix(FrmListadePrecios.GrdModulos.RowSel, 2) = "MAQUINARIA" Then
            PRECIVA = TOTAL * 1.105
        Else
            PRECIVA = TOTAL * 1.21
        End If
        FrmListadePrecios.GrdModulos.TextMatrix(FrmListadePrecios.GrdModulos.RowSel, 7) = Valido_Importe(PRECIVA)
    Else
        GrdModulos.TextMatrix(GrdModulos.row, 6) = Valido_Importe(txtActual.Text)
    End If
    
    On Error GoTo HayError
        
        If TxtCodigo = "" Then
            'ENTRA ACA CUANDO ES UNA LISTA NUEVA
            If optPVenta.Value = True Then
                GrdModulos.TextMatrix(GrdModulos.row, 5) = Valido_Importe(txtActual.Text)
                TOTAL = Valido_Importe(txtActual.Text)
                If FrmListadePrecios.GrdModulos.TextMatrix(FrmListadePrecios.GrdModulos.RowSel, 2) = "MAQUINARIA" Then
                    PRECIVA = TOTAL * 1.105
                Else
                    PRECIVA = TOTAL * 1.21
                End If
                FrmListadePrecios.GrdModulos.TextMatrix(FrmListadePrecios.GrdModulos.RowSel, 7) = Valido_Importe(PRECIVA)
            Else
                GrdModulos.TextMatrix(GrdModulos.row, 6) = Valido_Importe(txtActual.Text)
            End If
            
        Else
            'ENTRA ACA CUANDO ACTUALIZO UN PRECIO DE LISTA
            Screen.MousePointer = vbHourglass
            lblestado.Caption = "Actualizando ..."
            DBConn.BeginTrans
                sql = "UPDATE PRODUCTO "
                If optPVenta.Value = True Then
                    sql = sql & " SET PTO_PRECIO=" & XN(txtActual.Text)
                    sql = sql & " ,PTO_PRECIVA = " & XN(PRECIVA)
                Else
                    sql = sql & " SET PTO_PRECIOC=" & XN(txtActual.Text)
                End If
                sql = sql & " WHERE LIS_CODIGO =" & XN(TxtCodigo)
                sql = sql & " AND PTO_CODIGO LIKE '" & XN(GrdModulos.TextMatrix(GrdModulos.row, 0)) & "'"
                DBConn.Execute sql
               
                Screen.MousePointer = vbNormal
                lblestado.Caption = ""
                DBConn.CommitTrans
                
        End If
    If optPCompra.Value = True Then
        precio1 = optPCompra.Caption
        precio2 = optPVenta.Caption
        If MsgBox("El " & precio1 & " fue actualizado. Desea modificar el " & precio2 & " ? ", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then TabPrecios.Visible = False
            optPVenta.Value = True
    Else
        precio1 = optPVenta.Caption
        precio2 = optPCompra.Caption
        If MsgBox("El " & precio1 & " fue actualizado. Desea modificar el " & precio2 & " ? ", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then TabPrecios.Visible = False
            optPCompra.Value = True
    End If
    
        
    Exit Sub
            
HayError:
    lblestado.Caption = ""
    DBConn.RollbackTrans
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdAgregar_Click()
    Dim I As Integer
    Dim BAND As Integer
    
    frmBuscar.CodListaPrecio = 0
    frmBuscar.TipoBusqueda = 2
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        If frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0) = "" Then Exit Sub
    
        CodigoProducto = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0)
    
        I = 0
        BAND = 0
        'OJO CONTROLAR SOBRE EL RECORDSET Y NO SOBRE LA GRILLA
        Do While GrdModulos.Rows - 1 <> I
            If GrdModulos.TextMatrix(I + 1, 0) = CodigoProducto Then
                BAND = 1
            End If
            I = I + 1
        Loop
        If BAND = 0 Then
            Agregoproducto
            Else
                MsgBox " El Producto ya existe en la Lista de Precios"
        End If
    End If
End Sub
Function Agregoproducto()
    
    On Error GoTo HayError
    Screen.MousePointer = vbHourglass
    lblestado.Caption = "Guardando ..."
    DBConn.BeginTrans
    
    sql = "SELECT DISTINCT  P.PTO_DESCRI,L.LNA_DESCRI,R.RUB_DESCRI,"
    sql = sql & " RE.TPRE_DESCRI,P.PTO_PRECIO,P.PTO_PRECIOC,P.PTO_CODIGO"
    sql = sql & " FROM PRODUCTO P,LINEAS L,RUBROS R, TIPO_PRESENTACION RE "
    sql = sql & " WHERE L.LNA_CODIGO = P.LNA_CODIGO  "
    sql = sql & " AND R.RUB_CODIGO = P.RUB_CODIGO  AND P.TPRE_CODIGO = RE.TPRE_CODIGO "
    sql = sql & " AND P.PTO_CODIGO LIKE '" & CodigoProducto & "' ORDER BY P.PTO_CODIGO"
        
    lblestado.Caption = "Buscando..."
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
            If TxtCodigo.Text = "" Then
                'ACA ENTRA CUANDO ESTOY CREANDO UNA NUEVA LISTA DE PRECIO
                GrdModulos.AddItem rec.Fields(6) & Chr(9) & rec.Fields(0) & Chr(9) & rec.Fields(1) & Chr(9) & _
                              rec.Fields(2) & Chr(9) & rec.Fields(3) & Chr(9) & _
                              Valido_Importe(rec.Fields(4)) & Chr(9) & _
                              Valido_Importe(rec.Fields(6))

            Else
                 'INSERTO EN LA LISTA DE PRECIO Y EN DETALLE DE LISTA DE PRECIO
'                 sql = "INSERT INTO DETALLE_LISTA_PRECIO(LIS_CODIGO,PTO_CODIGO,LIS_PRECIO,LIS_PRECIOC)"
'                 sql = sql & " VALUES ("
'                 sql = sql & XN(TxtCodigo) & ","
'                 sql = sql & XS(CodigoProducto, True) & ","
'                 sql = sql & XN(rec!PTO_PRECIO) & " ,"
'                 sql = sql & XN(rec!PTO_PRECIOC) & " )"
'                 DBConn.Execute sql
                
                sql = "UPDATE PRODUCTO "
                sql = sql & " SET LIS_CODIGO=" & XN(TxtCodigo.Text)
                'sql = sql & " ,PTO_PRECIO= " & XN(GrdModulos.TextMatrix(j, 5))
                'sql = sql & " ,PTO_PRECIOC= " & XN(GrdModulos.TextMatrix(j, 6))
                sql = sql & " WHERE PTO_CODIGO LIKE '" & CodigoProducto & "'"
                DBConn.Execute sql
                
                
                'INSERTO EN LA GRILLA
                GrdModulos.AddItem rec.Fields(6) & Chr(9) & rec.Fields(0) & Chr(9) & rec.Fields(1) & Chr(9) & _
                                   rec.Fields(2) & Chr(9) & rec.Fields(3) & Chr(9) & _
                                   Valido_Importe(rec.Fields(4)) & Chr(9) & _
                                   Valido_Importe(rec.Fields(5))

            End If
    End If
    rec.Close
        
    Screen.MousePointer = vbNormal
    lblestado.Caption = ""
    DBConn.CommitTrans
    Exit Function
    
HayError:
    lblestado.Caption = ""
    DBConn.RollbackTrans
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
End Function
Private Sub CmdBorrar_Click()
    On Error GoTo CLAVOSE
    If Trim(TxtCodigo) <> "" Then
        resp = MsgBox("Seguro desea eliminar La Lista de Precios: " & Trim(cbodescri.Text) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        lblestado.Caption = "Eliminando ..."
        
        'DBConn.Execute "DELETE FROM DETALLE_LISTA_PRECIO WHERE LIS_CODIGO = " & XN(TxtCodigo)
        
        For j = 1 To GrdModulos.Rows - 1
'                sql = "INSERT INTO DETALLE_LISTA_PRECIO(LIS_CODIGO,"
'                sql = sql & "PTO_CODIGO,LIS_PRECIO,LIS_PRECIOC)"
'                sql = sql & " VALUES ("
'                sql = sql & XN(txtcodigo) & ","
'                sql = sql & XS(GrdModulos.TextMatrix(j, 0), True) & ","
'                sql = sql & XN(GrdModulos.TextMatrix(j, 5)) & " ,"
'                sql = sql & XN(GrdModulos.TextMatrix(j, 6)) & " )"
                sql = "UPDATE PRODUCTO "
                sql = sql & " SET LIS_CODIGO=" & XN("0")
                'sql = sql & " ,PTO_PRECIO= " & XN(GrdModulos.TextMatrix(j, 5))
                'sql = sql & " ,PTO_PRECIOC= " & XN(GrdModulos.TextMatrix(j, 6))
                sql = sql & " WHERE PTO_CODIGO LIKE '" & GrdModulos.TextMatrix(j, 0) & "'"
                DBConn.Execute sql
            Next
        
        DBConn.Execute "DELETE FROM LISTA_PRECIO WHERE LIS_CODIGO = " & XN(TxtCodigo)
        lblestado.Caption = ""
        Screen.MousePointer = vbNormal
        'Refrescar
        Cancelando
    End If
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblestado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub
Function Refrescar()
        'cbodescri.Visible = True
        tabLista.Visible = True
        cbodescri.Clear
        cargocboLista
        CargoCboListaPrecio
        CargoCboLPrecioRep
        CmdBuscAprox_Click
End Function
Private Sub CmdBuscAprox_Click()
    
    LimpiarOpciones
    If inicio > 0 Then
        cbodescri.Visible = False
        TxtDescriB.Visible = True
        'TxtDescriB.Text
        'TxtDescriB.Text = cbodescri.Text
        If TxtDescriB.Text <> "" Then
            cbodescri.Text = TxtDescriB.Text
        End If
        TxtDescriB.Enabled = True
        TxtDescriB.SetFocus
        CmdBuscAprox.Enabled = False
        cmdGrabar.Enabled = True
        Fecha1.Enabled = True
    End If
    Screen.MousePointer = vbHourglass
    sql = "SELECT DISTINCT  P.PTO_DESCRI,L.LNA_DESCRI,R.RUB_DESCRI,"
    sql = sql & " RE.TPRE_DESCRI,D.LIS_PRECIO,D.LIS_PRECIOC, LP.LIS_CODIGO, LP.LIS_FECHA,P.PTO_CODIGO"
    sql = sql & " FROM LISTA_PRECIO LP,PRODUCTO P,LINEAS L,RUBROS R, TIPO_PRESENTACION RE,DETALLE_LISTA_PRECIO D "
    sql = sql & " WHERE P.PTO_CODIGO = D.PTO_CODIGO"
    sql = sql & " AND D.LIS_CODIGO = LP.LIS_CODIGO "
    sql = sql & " AND L.LNA_CODIGO = P.LNA_CODIGO "
    sql = sql & " AND R.RUB_CODIGO = P.RUB_CODIGO  AND P.TPRE_CODIGO = RE.TPRE_CODIGO "
    sql = sql & " AND LP.LIS_DESCRI LIKE '" & Me.cbodescri.Text & "%' ORDER BY P.PTO_DESCRI"
        
    lblestado.Caption = "Buscando..."
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        GrdModulos.Rows = 1
        Do While Not rec.EOF
           GrdModulos.AddItem rec!PTO_CODIGO & Chr(9) & rec.Fields(0) & Chr(9) & rec.Fields(1) & Chr(9) & _
                              rec.Fields(2) & Chr(9) & rec.Fields(3) & Chr(9) & _
                              Valido_Importe(rec!LIS_PRECIO) & Chr(9) & _
                              Valido_Importe(rec!LIS_PRECIOC)
    
            rec.MoveNext
        Loop
        rec.MoveFirst
        TxtCodigo.Text = rec.Fields(6)
        Fecha1.Value = rec.Fields(7)
         
    Else
        lblestado.Caption = ""
        'MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbInformation, TIT_MSGBOX
        GrdModulos.Rows = 1
        'me.txt
'        Me.cbodescri.SetFocus
    End If
    If rec.State = 1 Then rec.Close
    'rec.Close
    Screen.MousePointer = vbNormal
    lblestado.Caption = ""
    inicio = inicio + 1
    CmdCancelar_Click
    
End Sub

Private Sub CmdCancelar_Click()
    TxtDescriB.Text = ""
    cmdGrabar.Enabled = False
    CmdBuscAprox.Enabled = True
    SeteoInicial
    Cancelando
    freOpciones.Caption = ""
    freOpciones.Caption = "Opciones de Consulta"
    txtObservaciones1.Text = ""
    txtObservaciones2.Text = ""
    CmdBorrar.Enabled = True
    Fecha1.Enabled = False
    If Consulta = 2 Then
        cmdPrecios.Enabled = True
    End If
    LimpiarOpciones
    'cbodescri.SetFocus
    cmdModificar.Enabled = True
    CmdBorrar.Enabled = True
    cmdImprimir.Enabled = True
End Sub

Function Cancelando()
'    Dim BAND As Integer
'    'LIMPIA LA GRILLA Y VUELVE A CARGAR UNA LISTA
'    BAND = 0
'    Screen.MousePointer = vbHourglass
''    If cbodescri.ListIndex = 0 Then 'Busco en Productos
''        BAND = 1
''        sql = "SELECT P.PTO_DESCRI, L.LNA_DESCRI,R.RUB_DESCRI,RE.TPRE_DESCRI,"
''        sql = sql & "P.PTO_PRECIO,P.PTO_PRECIOC,P.PTO_CODIGO "
''        sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L, TIPO_PRESENTACION RE"
''        sql = sql & " WHERE"
''        sql = sql & " P.LNA_CODIGO=L.LNA_CODIGO AND P.RUB_CODIGO=R.RUB_CODIGO"
''        sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
''        sql = sql & " AND P.TPRE_CODIGO=RE.TPRE_CODIGO"
''        sql = sql & " AND P.PTO_ESTADO=1"
''        sql = sql & " ORDER BY P.PTO_DESCRI"
''    Else
'        sql = "SELECT DISTINCT  P.PTO_DESCRI,L.LNA_DESCRI,R.RUB_DESCRI,"
'        sql = sql & " RE.TPRE_DESCRI,P.PTO_PRECIO,P.PTO_PRECIOC, LP.LIS_CODIGO, LP.LIS_FECHA,P.PTO_CODIGO"
'        sql = sql & " FROM LISTA_PRECIO LP,PRODUCTO P,LINEAS L,RUBROS R, TIPO_PRESENTACION RE"
'        ', DETALLE_LISTA_PRECIO D "
'        sql = sql & " WHERE P.LIS_CODIGO = LP.LIS_CODIGO "
'        'AND D.LIS_CODIGO = LP.LIS_CODIGO  "
'        sql = sql & " AND L.LNA_CODIGO = P.LNA_CODIGO "
'        sql = sql & " AND R.RUB_CODIGO = P.RUB_CODIGO  AND P.TPRE_CODIGO = RE.TPRE_CODIGO "
'        sql = sql & " AND LP.LIS_DESCRI LIKE '" & Me.cbodescri.Text & "%' ORDER BY P.PTO_DESCRI"
' '   End If
'        lblEstado.Caption = "Buscando..."
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If rec.EOF = False Then
'            GrdModulos.Rows = 1
'            Do While Not rec.EOF
'               GrdModulos.AddItem rec!PTO_CODIGO & Chr(9) & rec.Fields(0) & Chr(9) & rec.Fields(1) & Chr(9) & _
'                                  rec.Fields(2) & Chr(9) & rec.Fields(3) & Chr(9) & _
'                                  Valido_Importe(rec.Fields(4)) & Chr(9) & _
'                                  Valido_Importe(rec.Fields(5))
'
'                rec.MoveNext
'            Loop
'            rec.MoveFirst
'            If BAND = 0 Then
'                TxtCodigo.Text = rec.Fields(6)
'                Fecha1.Value = rec.Fields(7)
'            Else
'                TxtCodigo.Text = ""
'                Fecha1.Value= date
'            End If
'
'        Else
'            lblEstado.Caption = ""
'            MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
'            Me.cbodescri.SetFocus
'        End If
'        rec.Close
'        Screen.MousePointer = vbNormal
'        lblEstado.Caption = ""
    
'    Else
'        MsgBox "No hay Listas de Precios Cargadas en el Sistema", vbInformation
'    End If
    SeteoInicial

End Function

Private Sub cmdfiltrar_Click()
    
    
    If TxtCodigo.Text <> "" Then
        'ENTRA ACA CUANDO CONSULTA UNA LISTA
        
        Screen.MousePointer = vbHourglass
        
        sql = " SELECT  TOP 800 P.PTO_DESCRI,L.LNA_DESCRI, "
        sql = sql & " R.RUB_DESCRI,RE.TPRE_DESCRI,P.PTO_PRECIO,P.PTO_PRECIOC,P.PTO_CODIGO,P.PTO_PRECIVA "
        sql = sql & " FROM PRODUCTO P,LINEAS L,RUBROS R,TIPO_PRESENTACION RE,LISTA_PRECIO LP"
        ',DETALLE_LISTA_PRECIO D "
        sql = sql & " WHERE P.LNA_CODIGO = L.LNA_CODIGO "
        sql = sql & " AND P.PTO_ESTADO <> 1 "
        sql = sql & " AND LP.LIS_CODIGO = P.LIS_CODIGO"
        sql = sql & " AND P.LIS_CODIGO <> " & 0 & " "
        sql = sql & " AND P.RUB_CODIGO = R.RUB_CODIGO AND P.TPRE_CODIGO = RE.TPRE_CODIGO "
        
        If TxtCodigo.Text <> "" Then
            sql = sql & " AND LP.LIS_CODIGO = " & TxtCodigo.Text & " "
        End If
        If chkProducto.Value = 1 Then
            'sql = sql & " AND (P.PTO_DESCRI LIKE '" & txtproducto.Text & "%' "
            'sql = sql & " OR P.PTO_CODIGO LIKE '" & txtproducto.Text & "%' )"
            txtproducto.Text = Replace(txtproducto, "'", "´")
            
            sql = sql & " AND (P.PTO_DESCRI LIKE '" & txtproducto.Text & "%' "
            sql = sql & " OR P.PTO_CODIGO LIKE '" & txtproducto.Text & "%' )"
        End If
        If chkLinea.Value = 1 Then
            sql = sql & " AND L.LNA_CODIGO = " & cboLinea.ItemData(cboLinea.ListIndex) & " "
        End If
        If chkRubro.Value = 1 Then
            sql = sql & " AND R.RUB_CODIGO = " & cboRubro.ItemData(cboRubro.ListIndex) & " "
        End If
        If chkRepres.Value = 1 Then
            sql = sql & " AND RE.TPRE_CODIGO = " & cboRep.ItemData(cboRep.ListIndex) & " "
        End If
        
        sql = sql & " ORDER BY P.PTO_DESCRI"
        
        lblestado.Caption = "Buscando..."
       
    Else
        'ENTRA ACA CUANDO CARGO UNA NUEVA LISTA
        Screen.MousePointer = vbHourglass
        sql = " SELECT P.PTO_DESCRI,L.LNA_DESCRI, "
        sql = sql & " R.RUB_DESCRI,RE.TPRE_DESCRI,P.PTO_PRECIO,P.PTO_PRECIOC,P.PTO_CODIGO,P.PTO_PRECIVA "
        sql = sql & " FROM PRODUCTO P,LINEAS L,RUBROS R,TIPO_PRESENTACION RE"
        sql = sql & " WHERE P.LNA_CODIGO = L.LNA_CODIGO "
        sql = sql & " AND P.LIS_CODIGO = " & 0 & " "
        sql = sql & " AND P.RUB_CODIGO = R.RUB_CODIGO AND P.TPRE_CODIGO = RE.TPRE_CODIGO "
        sql = sql & " AND P.PTO_ESTADO <> 1 "
        If chkProducto.Value = 1 Then
            'sql = sql & " AND (P.PTO_DESCRI LIKE '" & txtproducto.Text & "%' "
            'sql = sql & " OR P.PTO_CODIGO LIKE '" & txtproducto.Text & "%' )"
            sql = sql & " AND (P.PTO_DESCRI LIKE '" & txtproducto.Text & "%' "
            sql = sql & " OR P.PTO_CODIGO LIKE '" & txtproducto.Text & "%' )"
        End If
        If chkLinea.Value = 1 Then
            sql = sql & " AND L.LNA_CODIGO = " & cboLinea.ItemData(cboLinea.ListIndex) & " "
        End If
        If chkRubro.Value = 1 Then
            sql = sql & " AND R.RUB_CODIGO = " & cboRubro.ItemData(cboRubro.ListIndex) & " "
        End If
        If chkRepres.Value = 1 Then
            sql = sql & " AND RE.TPRE_CODIGO = " & cboRep.ItemData(cboRep.ListIndex) & " "
        End If
        
        sql = sql & " ORDER BY P.PTO_DESCRI "
        
        lblestado.Caption = "Buscando..."
    End If
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        GrdModulos.Rows = 1
        Do While Not rec.EOF
           GrdModulos.AddItem rec!PTO_CODIGO & Chr(9) & rec!PTO_DESCRI & Chr(9) & rec!LNA_DESCRI & Chr(9) & _
                              rec!RUB_DESCRI & Chr(9) & rec!TPRE_DESCRI & Chr(9) & _
                              Valido_Importe(rec.Fields(4)) & Chr(9) & _
                              Valido_Importe(rec.Fields(5)) & Chr(9) & _
                              IIf(IsNull(rec!PTO_PRECIVA), "0,00", Valido_Importe(IIf(IsNull(rec!PTO_PRECIVA), 0, rec!PTO_PRECIVA)))
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblestado.Caption = ""
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
        Me.cmdfiltrar.SetFocus
    End If
    rec.Close
    Screen.MousePointer = vbNormal
    lblestado.Caption = ""
    lblestado.Caption = ""
    Text1.Text = GrdModulos.Rows - 1
    lblestado.Caption = "Se encontraron " + Text1.Text + " registros"
    
End Sub

Function ValidarLista()
    If TxtDescriB.Text = "" Then
        MsgBox "No ha ingresado la Descricpción de la Lista de Precios", vbExclamation, TIT_MSGBOX
        TxtDescriB.SetFocus
        ValidarLista = False
        Exit Function
    End If
    If Fecha1.Value = Date Then
        MsgBox "No ha ingresado la Fecha de Vigencia", vbExclamation, TIT_MSGBOX
        Fecha1.SetFocus
        ValidarLista = False
        Exit Function
    End If
'    If GrdModulos.Rows = 1 Then
'        MsgBox "Debe haber al menos un producto en la Lista de Precios", vbExclamation, TIT_MSGBOX
'        CmdAgregar.SetFocus
'        ValidarLista = False
'        Exit Function
'    End If
    ValidarLista = True
    
End Function

Private Sub CmdGrabar_Click()
    
    On Error GoTo HayError
    
    CmdBorrar.Enabled = True
         
    If ValidarLista = False Then Exit Sub
       
    If TxtCodigo.Text <> "" Then
         ' Entra aca cuando hago una modificación
         
         Screen.MousePointer = vbHourglass
         lblestado.Caption = "Actualizando ..."
         DBConn.BeginTrans
         sql = "UPDATE LISTA_PRECIO "
         sql = sql & " SET LIS_DESCRI=" & XS(TxtDescriB.Text)
         sql = sql & " ,LIS_FECHA= " & XDQ(Fecha1)
         sql = sql & " WHERE LIS_CODIGO =" & XN(TxtCodigo)
         DBConn.Execute sql
        
         Screen.MousePointer = vbNormal
         lblestado.Caption = ""
         DBConn.CommitTrans
         TxtDescriB.Text = ""
         
   Else
        
        ' Entra aca cuando hago una nueva lista
        
         Screen.MousePointer = vbHourglass
         lblestado.Caption = "Guardando ..."
         DBConn.BeginTrans
         
         TxtCodigo = "1"
         sql = "SELECT MAX(LIS_CODIGO) as maximo FROM LISTA_PRECIO"
         rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
         If Not IsNull(rec.Fields!Maximo) Then TxtCodigo = XN(rec.Fields!Maximo) + 1
         rec.Close
            sql = "INSERT INTO LISTA_PRECIO(LIS_CODIGO,LIS_FECHA,LIS_DESCRI)    "
            sql = sql & " VALUES ("
            sql = sql & XN(TxtCodigo) & ","
            sql = sql & XDQ(Fecha1) & ","
            sql = sql & XS(TxtDescriB) & ")"
            DBConn.Execute sql
            For j = 1 To GrdModulos.Rows - 1
'                sql = "INSERT INTO DETALLE_LISTA_PRECIO(LIS_CODIGO,"
'                sql = sql & "PTO_CODIGO,LIS_PRECIO,LIS_PRECIOC)"
'                sql = sql & " VALUES ("
'                sql = sql & XN(txtcodigo) & ","
'                sql = sql & XS(GrdModulos.TextMatrix(j, 0), True) & ","
'                sql = sql & XN(GrdModulos.TextMatrix(j, 5)) & " ,"
'                sql = sql & XN(GrdModulos.TextMatrix(j, 6)) & " )"
                sql = "UPDATE PRODUCTO "
                sql = sql & " SET LIS_CODIGO=" & XN(TxtCodigo)
                sql = sql & " ,PTO_PRECIO= " & XN(GrdModulos.TextMatrix(j, 5))
                sql = sql & " ,PTO_PRECIOC= " & XN(GrdModulos.TextMatrix(j, 6))
                sql = sql & " WHERE PTO_CODIGO LIKE '" & GrdModulos.TextMatrix(j, 0) & "'"
                DBConn.Execute sql
            Next
        
         Screen.MousePointer = vbNormal
         DBConn.CommitTrans
         'Refrescar
         cmdPrecios.Enabled = True
         MsgBox "La Lista de Precios " & XS(TxtDescriB, True) & " fue grabada correctamente", vbInformation
    End If
    'cbodescri.Visible = True
    tabLista.Visible = True
    TxtDescriB.Text = ""
    TxtDescriB.Visible = False
    'cbodescri.Clear
    cargocboLista
    CargoCboListaPrecio
    CargoCboLPrecioRep
    cmdGrabar.Enabled = False
    CmdBorrar.Enabled = True
    CmdBuscAprox.Enabled = True
    freOpciones.Caption = ""
    freOpciones.Caption = "Opciones de Consulta"
    LimpiarOpciones
    cmdModificar.Enabled = True
    CmdBorrar.Enabled = True
    cmdImprimir.Enabled = True
    Cancelando
    Exit Sub
         
HayError:
         lblestado.Caption = ""
         DBConn.RollbackTrans
         Screen.MousePointer = vbNormal
         MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub cmdImportar_Click()
    frmTest.Show vbModal
    
End Sub

Private Sub cmdImprimir_Click()
    lblestado.Caption = "Buscando Listado..."
    Rep.Reset
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SIELECTROCENTRO"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    
     If TxtCodigo.Text <> "" Then
        Rep.SelectionFormula = "{LISTA_PRECIO.LIS_CODIGO}=" & TxtCodigo.Text
     Else
        MsgBox "Debe seleccionar una Lista de Precios", vbInformation
        Exit Sub
     End If
     If txtproducto.Text <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = "{PRODUCTO.PTO_CODIGO} LIKE '" & XS(txtproducto.Text) & "%' "
        Else
            Rep.SelectionFormula = Rep.SelectionFormula = " AND {PRODUCTO.PTO_CODIGO} LIKE  '%" & XS(txtproducto.Text) & "%' "
        End If
     End If
     
     If chkRepres.Value = Checked Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = "{PRODUCTO.TPRE_CODIGO}=" & XN(cboRep.ItemData(cboRep.ListIndex))
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {PRODUCTO.TPRE_CODIGO}=" & XN(cboRep.ItemData(cboRep.ListIndex))
        End If
     End If
     If chkRubro.Value = Checked Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = "{PRODUCTO.RUB_CODIGO}=" & XN(cboRubro.ItemData(cboRubro.ListIndex))
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {PRODUCTO.RUB_CODIGO}=" & XN(cboRubro.ItemData(cboRubro.ListIndex))
        End If
     End If
     If chkLinea.Value = Checked Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = "{PRODUCTO.LNA_CODIGO}=" & XN(cboLinea.ItemData(cboLinea.ListIndex))
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {PRODUCTO.LNA_CODIGO}=" & XN(cboLinea.ItemData(cboLinea.ListIndex))
        End If
     End If
     If txtObservaciones1.Text <> "" Then
        Rep.Formulas(0) = "OBSERVACION1=' 1 - Observación: " & Trim(txtObservaciones1.Text) & "'"
     End If
     If txtObservaciones2.Text <> "" Then
        Rep.Formulas(1) = "OBSERVACION2=' 2 - Observación: " & Trim(txtObservaciones2.Text) & "'"
     End If
     Rep.WindowTitle = "Lista de Precios..."
     Rep.ReportFileName = DRIVE & DirReport & "rptlistaprecio.rpt"
     Rep.Destination = crptToWindow
     Rep.WindowState = crptMaximized
     Rep.Action = 1
     
     Rep.SelectionFormula = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
     lblestado.Caption = ""
End Sub

Private Sub cmdModificar_Click()
    LimpiarOpciones
    cmdPrecios.Enabled = True
    If inicio > 0 Then
        cbodescri.Visible = False
        TxtDescriB.Visible = True
        TxtDescriB.Text = cbodescri.Text
        TxtDescriB.Enabled = True
        TxtDescriB.SetFocus
        CmdBuscAprox.Enabled = False
        cmdGrabar.Enabled = True
        Fecha1.Enabled = True
    End If
    
End Sub

Private Sub CmdNuevo_Click()
    cmdGrabar.Enabled = True
    CmdBorrar.Enabled = False
    Fecha1.Value = Date
    TxtCodigo.Text = ""
    TxtDescriB.Text = ""
    TxtDescriB.Visible = True
    cbodescri.Visible = False
    tabLista.Visible = False
    freOpciones.Caption = ""
    txtObservaciones1.Text = ""
    txtObservaciones2.Text = ""
    freOpciones.Caption = "Opciones de Carga"
    NuevaLista 'Carga los productos de la Tabla producto
    CmdBuscAprox.Enabled = False
    cmdPrecios.Enabled = False
    LimpiarOpciones
    TxtDescriB.SetFocus
    cmdModificar.Enabled = False
    CmdBorrar.Enabled = False
    cmdImprimir.Enabled = False
End Sub

Function NuevaLista()
    GrdModulos.Rows = 1
    Screen.MousePointer = vbHourglass
    
    sql = " SELECT P.PTO_DESCRI,L.LNA_DESCRI,R.RUB_DESCRI,TP.TPRE_DESCRI,"
    sql = sql & " P.PTO_PRECIO,P.PTO_PRECIOC,P.PTO_CODIGO "
    sql = sql & " FROM PRODUCTO P,LINEAS L,RUBROS R,TIPO_PRESENTACION TP"
    sql = sql & " WHERE P.LNA_CODIGO = L.LNA_CODIGO "
    sql = sql & " AND P.RUB_CODIGO = R.RUB_CODIGO "
    sql = sql & " AND P.TPRE_CODIGO = TP.TPRE_CODIGO "
    sql = sql & " AND P.LIS_CODIGO = " & 0 & " "
    sql = sql & " AND P.PTO_CODIGO <> 1 "
    sql = sql & "ORDER BY P.PTO_DESCRI"
    
    lblestado.Caption = " Creando Nueva Lista de Precios..."
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        Do While Not rec.EOF
           GrdModulos.AddItem rec!PTO_CODIGO & Chr(9) & rec.Fields(0) & Chr(9) & rec.Fields(1) & Chr(9) & _
                              rec.Fields(2) & Chr(9) & rec.Fields(3) & Chr(9) & _
                              Valido_Importe(rec.Fields(4)) & Chr(9) & _
                              IIf(IsNull(Valido_Importe(rec.Fields(5))), "0,00", Valido_Importe(rec.Fields(5)))
            rec.MoveNext
        Loop
      '  If GrdModulos.Enabled Then GrdModulos.SetFocus
    Else
        lblestado.Caption = ""
        MsgBox "Todos los Productos están asignados a una Lista de Precios", vbOKOnly + vbCritical, TIT_MSGBOX
        Me.TxtDescriB.SetFocus
    End If
    rec.Close
    Screen.MousePointer = vbNormal
    lblestado.Caption = ""
End Function

Private Sub cmdPrecios_Click()
    If TxtCodigo.Text <> "" Then
        If GrdModulos.Rows <> 1 Then
            frmModificoPrecios.Show vbModal
            Set frmModificoPrecios = Nothing
        Else
            MsgBox "Debe haber al menos un producto en la Lista de Precios"
        End If
        'Refrescar
    Else
        MsgBox "Debe seleccionar una Lista de Precios para poder modificar los precios", vbInformation
    End If
End Sub

Private Sub cmdPreCompra_Click()
    '''escondo o no el precio de compra
    If GrdModulos.ColWidth(6) <> 0 Then
        GrdModulos.ColWidth(6) = 0
    Else
        GrdModulos.ColWidth(6) = 1000
    End If
    
End Sub

Private Sub cmdQuitar_Click()
    If GrdModulos.Rows = 1 Then
        MsgBox "Debe seleccinar un producto de la Lista"
    Else
        resp = MsgBox("Seguro desea quitar el Producto " & GrdModulos.TextMatrix(GrdModulos.row, 0) & " de la Lista de Precios? ", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar")
        If resp <> 6 Then Exit Sub
            If TxtCodigo.Text = "" Then
                'CUANDO CARGO UNO NUEVO, SOLO ELIMINO EN LA GRILLA
                If GrdModulos.Rows = 2 Then
                    GrdModulos.Rows = 1
                    Else
                    GrdModulos.RemoveItem (GrdModulos.row)
                End If
            Else
            ' CUANDO ELIMINO UN ITEM DE LA LISTA DE PRECIO YA CARGADA
            DBConn.BeginTrans
            
            'sql = "DELETE FROM DETALLE_LISTA_PRECIO WHERE LIS_CODIGO = " & XN(TxtCodigo.Text) & " "
            'sql = sql & " AND PTO_CODIGO LIKE '" & GrdModulos.TextMatrix(GrdModulos.row, 0) & "' "
            
            sql = "UPDATE PRODUCTO "
            sql = sql & " SET LIS_CODIGO=" & XN("0")
            'sql = sql & " ,PTO_PRECIO= " & XN(GrdModulos.TextMatrix(j, 5))
            'sql = sql & " ,PTO_PRECIOC= " & XN(GrdModulos.TextMatrix(j, 6))
            sql = sql & " WHERE PTO_CODIGO LIKE '" & GrdModulos.TextMatrix(GrdModulos.RowSel, 0) & "'"
            DBConn.Execute sql
            
            If GrdModulos.Rows = 2 Then
                GrdModulos.Rows = 1
                Else
                GrdModulos.RemoveItem (GrdModulos.row)
            End If
            DBConn.Execute sql
            DBConn.CommitTrans
            End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Consulta = 4 'PARA QUE NO HAGA NADA
    Unload Me
    Set FrmListadePrecios = Nothing
End Sub

Private Sub cmdSalirP_Click()
    FrmListadePrecios.Enabled = True
    TabPrecios.Visible = False
    freLista.Enabled = True
    freOpciones.Enabled = True
    'frebotones.Enabled = True
    frebotopc.Enabled = True
End Sub



Private Sub Command1_Click()
'PROCESO ESPECIAL
        sql = " SELECT * FROM DETALLE_LISTA_PRECIO WHERE LIS_CODIGO=28"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While Not rec.EOF
                sql = "UPDATE PRODUCTO"
                sql = sql & " SET LIS_CODIGO = " & XN(rec!LIS_CODIGO)
                'sql = sql & " , PTO_PRECIO = " & XN(rec!LIS_PRECIO)
                'sql = sql & " , PTO_PRECIOC = " & XN(rec!LIS_PRECIOC)
                sql = sql & " WHERE PTO_CODIGO LIKE '" & rec!PTO_CODIGO & "'"
                DBConn.Execute sql

                rec.MoveNext
             Loop
        End If
       rec.Close
'***************
End Sub

Private Sub Form_Activate()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    
    ' actualizo la Factura proveedor
'    sql = "UPDATE FACTURA_PROVEEDOR "
'    sql = sql & "SET FPR_SALDO = 0"
'    sql = sql & " WHERE FPR_FECHA < 01/01/2007"
'    DBConn.Execute sql
    
    'SeteoInicial
    CmdBuscAprox.Visible = False
    inicio = 1
    If Consulta = 1 Or Consulta = 3 Then ' si se ingresa a esta pantalla en modo Consulta
       'cbodescri_Click
       cmdNuevo.Visible = False
       cmdGrabar.Visible = False
       CmdBorrar.Visible = False
       cmdModificar.Visible = False
       cmdPrecios.Enabled = False
       CmdAgregar.Enabled = False
       cmdQuitar.Enabled = False
       CmdBuscAprox.Visible = False
'       If Consulta = 3 Then
'            cargocboLista
'            Call BuscaCodigoProxItemData(frmRemitoCliente.nlista, cbodescri)
'       End If
    Else
       'CmdBuscAprox_Click
       cmdNuevo.Visible = True
       cmdGrabar.Visible = True
       CmdBorrar.Visible = True
       cmdModificar.Visible = True
       cmdPrecios.Enabled = True
       CmdAgregar.Enabled = True
       cmdQuitar.Enabled = True
       CmdBuscAprox.Visible = False
    End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    'Consulta = 1
    Dim nPrecioIVA As String
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    
'    sql = "SELECT * FROM PRODUCTO WHERE LNA_CODIGO = 7 AND PTO_PRECIVA = 0.00"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        MsgBox rec.RecordCount
'        Do While Not rec.EOF
''            If rec!LNA_CODIGO = 6 Then 'MAQUINARIA
''                NPRECIOIVA = (rec!PTO_PRECIO) * (1.105)
''            Else
'                NPRECIOIVA = (rec!PTO_PRECIO) * (1.21)
'            'End If
'            sql = "UPDATE PRODUCTO SET PTO_PRECIVA = " & XN(NPRECIOIVA) & " WHERE PTO_CODIGO LIKE '" & rec!PTO_CODIGO & "'"
'            DBConn.Execute sql
'            rec.MoveNext
'        Loop
'    End If
'    rec.Close
    
    SeteoInicial
    CmdBuscAprox.Visible = False
    inicio = 1
    If Consulta = 1 Or Consulta = 3 Then ' si se ingresa a esta pantalla en modo Consulta
       'cbodescri_Click
       cmdNuevo.Visible = False
       cmdGrabar.Visible = False
       CmdBorrar.Visible = False
       cmdModificar.Visible = False
       cmdPrecios.Enabled = False
       CmdAgregar.Enabled = False
       cmdQuitar.Enabled = False
       CmdBuscAprox.Visible = False
    Else
       'CmdBuscAprox_Click
       cmdNuevo.Visible = True
       cmdGrabar.Visible = True
       CmdBorrar.Visible = True
       cmdModificar.Visible = True
       cmdPrecios.Enabled = True
       CmdAgregar.Enabled = True
       cmdQuitar.Enabled = True '
       CmdBuscAprox.Visible = False
    End If
    
    

    
End Sub
Private Sub SeteoInicial()
    'cbodescri.Visible = True
    tabLista.Visible = True
    preparogrilla
    cargocboLinea
    cargocboRubro (-1)
    cargocboRepres -1, -1  ' Para Cargar Marcas sin Lineas y Rubros
    'cargocboLista
    CargoCboListaPrecio ' maquina
    CargoCboLPrecioRep ' repuesto
    CargoCboListaAdicionales ' adicionales
    TxtDescriB.Visible = False
    cmdGrabar.Enabled = False
    TabPrecios.Visible = False
    txtproducto.Enabled = False
    cboLinea.Enabled = False
    cboRubro.Enabled = False
    cboRep.Enabled = False
    Fecha1.Enabled = False
    lblestado.Caption = ""
End Sub
Function preparogrilla()
    GrdModulos.FormatString = "Código|Producto|Linea|Rubro|Marca|P Venta($)|P Compra($)|PrecIVA($)"
    GrdModulos.ColWidth(0) = 1300
    GrdModulos.ColWidth(1) = 2700
    GrdModulos.ColWidth(2) = 1600
    GrdModulos.ColWidth(3) = 1600
    GrdModulos.ColWidth(4) = 1200
    GrdModulos.ColWidth(5) = 1000
    GrdModulos.ColWidth(6) = 0
    GrdModulos.ColWidth(7) = 1000
    GrdModulos.Rows = 1

End Function


Function cargocboLinea()
    cboLinea.Clear
    sql = "SELECT * FROM LINEAS  ORDER BY LNA_DESCRI"
    If rec.State = 1 Then rec.Close
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboLinea.AddItem rec!LNA_DESCRI
            cboLinea.ItemData(cboLinea.NewIndex) = rec!LNA_CODIGO
            rec.MoveNext
        Loop
        cboLinea.ListIndex = -1
    End If
    rec.Close
End Function
Function cargocboRubro(cod As Integer)
    
    cboRubro.Clear
    sql = "SELECT * FROM RUBROS "
    If cod <> -1 Then
        sql = sql & " WHERE LNA_CODIGO= " & cod
    End If
    sql = sql & " ORDER BY RUB_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboRubro.AddItem rec!RUB_DESCRI
            cboRubro.ItemData(cboRubro.NewIndex) = rec!RUB_CODIGO
            rec.MoveNext
        Loop
        cboRubro.ListIndex = -1
    End If
    rec.Close
End Function
Function cargocboRepres(codL As Integer, codR As Integer)
    cboRep.Clear
    sql = "SELECT * FROM TIPO_PRESENTACION WHERE TPRE_CODIGO IS NOT NULL "
    If codL <> -1 Then
        sql = sql & " AND LNA_CODIGO = " & cboLinea.ItemData(cboLinea.ListIndex) & ""
    End If
    If codR <> -1 Then
        sql = sql & "AND RUB_CODIGO = " & cboRubro.ItemData(cboRubro.ListIndex) & ""
    End If
    sql = sql & " ORDER BY TPRE_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboRep.AddItem rec!TPRE_DESCRI
            cboRep.ItemData(cboRep.NewIndex) = rec!TPRE_CODIGO
            rec.MoveNext
        Loop
        cboRep.ListIndex = -1
    End If
    rec.Close
End Function
Private Sub CargoCboListaAdicionales() '' Lista de Precios de ADICIONALES
    sql = "SELECT DISTINCT LP.LIS_CODIGO, LP.LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO LP, PRODUCTO P"
    sql = sql & " WHERE LP.LIS_CODIGO = P.LIS_CODIGO"
    sql = sql & " AND LP.LIS_DESCRI LIKE '" & "ADICIONALES" & "'"
    sql = sql & " ORDER BY LIS_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    cboAdicionales.Clear
    If rec.EOF = False Then
        rec.MoveFirst
        Do While rec.EOF = False
            cboAdicionales.AddItem rec!LIS_DESCRI
            cboAdicionales.ItemData(cboAdicionales.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cboAdicionales.ListIndex = 0
    End If
    rec.Close
End Sub


Private Sub CargoCboLPrecioRep() '' Lista de Precios de Repuestos
    cboLPrecioRep.Clear
    sql = "SELECT DISTINCT LP.LIS_CODIGO, LP.LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO LP, PRODUCTO P"
    sql = sql & " WHERE LP.LIS_CODIGO = P.LIS_CODIGO"
    sql = sql & " AND P.LNA_CODIGO = 1"   '1: Repuestos
    sql = sql & " ORDER BY LIS_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
    If rec.EOF = False Then
        rec.MoveFirst
        Do While rec.EOF = False
            cboLPrecioRep.AddItem rec!LIS_DESCRI
            cboLPrecioRep.ItemData(cboLPrecioRep.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cboLPrecioRep.ListIndex = 0
    End If
    rec.Close
End Sub
Private Sub CargoCboListaPrecio() '' Lista de Precios de Repuestos
    cboListaPrecio.Clear
    sql = "SELECT DISTINCT LP.LIS_CODIGO, LP.LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO LP, PRODUCTO P"
    sql = sql & " WHERE LP.LIS_CODIGO = P.LIS_CODIGO"
    sql = sql & " AND P.LNA_CODIGO = 6"   '6: Maquinaria
    sql = sql & " ORDER BY LP.LIS_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
    If rec.EOF = False Then
        rec.MoveFirst
        Do While rec.EOF = False
            cboListaPrecio.AddItem rec!LIS_DESCRI
            cboListaPrecio.ItemData(cboListaPrecio.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cboListaPrecio.ListIndex = -1
    End If
    rec.Close
End Sub
Function cargocboLista()
    cbodescri.Clear
    sql = "SELECT DISTINCT LIS_CODIGO,LIS_DESCRI,LIS_FECHA FROM LISTA_PRECIO "
    sql = sql & "ORDER BY LIS_DESCRI"
    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    'cbodescri.AddItem "<TODOS>"
    If Rec2.EOF = False Then
        Do While Rec2.EOF = False
            cbodescri.AddItem Rec2!LIS_DESCRI
            cbodescri.ItemData(cbodescri.NewIndex) = Rec2!LIS_CODIGO
            Rec2.MoveNext
        Loop
        
        cbodescri.ListIndex = -1
    End If
'    Rec2.MoveFirst
'    txtcodigo.Text = Rec2!LIS_CODIGO
'    Fecha1.Value = Rec2!LIS_FECHA
'    Rec2.MoveFirst
    TxtCodigo.Text = ""
    Fecha1.Value = Date
    Rec2.Close
    
End Function



Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub GrdModulos_DblClick()
    If Consulta = 3 Then
        GrdModulos.Col = 0
        Me.Hide
    Else
        If Consulta = 2 Then
            If TxtCodigo.Text <> "" Then
                TabPrecios.Visible = True
                If optPVenta.Value = True Then
                    txtAnterior.Text = GrdModulos.TextMatrix(GrdModulos.row, 5)
                    txtActual.Text = "0,00"
                    txtActual.SetFocus
                Else
                    txtAnterior.Text = GrdModulos.TextMatrix(GrdModulos.row, 6)
                    txtActual.Text = "0,00"
                    txtActual.SetFocus
                End If
                
                
                freLista.Enabled = False
                freOpciones.Enabled = False
                'frebotones.Enabled = False
                frebotopc.Enabled = False
           Else
                MsgBox "Debe seleccionar una Lista de Precios", vbInformation
           End If
        End If
    End If
End Sub

Private Sub optPCompra_Click()
    txtAnterior.Text = GrdModulos.TextMatrix(GrdModulos.row, 6)
    txtActual.Text = "0,00"
   'txtActual.SetFocus
End Sub

Private Sub optPVenta_Click()
    txtAnterior.Text = GrdModulos.TextMatrix(GrdModulos.row, 5)
    txtActual.Text = "0,00"
'    txtActual.SetFocus
End Sub

Private Sub txtActual_GotFocus()
    SelecTexto txtActual
End Sub

Private Sub txtActual_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtActual, KeyAscii)
End Sub

Private Sub txtActual_LostFocus()
    txtActual.Text = Valido_Importe(txtActual)
End Sub

Private Sub txtAnterior_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtAnterior, KeyAscii)
End Sub

Private Sub txtAnterior_LostFocus()
    txtAnterior.Text = Valido_Importe(txtAnterior)
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtDescriB_GotFocus()
    SelecTexto TxtDescriB
End Sub

Private Sub TxtDescriB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtObservaciones1_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtObservaciones2_KeyPress(KeyAscii As Integer)
     'KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtproducto_Change()
    If txtproducto.Text <> "" Then
        chkProducto.Value = 1
    End If
End Sub

Private Sub txtproducto_GotFocus()
    SelecTexto txtproducto
End Sub

Private Sub txtproducto_KeyPress(KeyAscii As Integer)
'    KeyAscii = CarTexto(KeyAscii)
End Sub
Private Sub LimpiarOpciones()
    chkLinea.Value = Unchecked
    cboLinea.ListIndex = -1
    chkProducto.Value = Unchecked
    txtproducto.Text = ""
    chkRubro.Value = Unchecked
    cboRubro.ListIndex = -1
    chkRepres.Value = Unchecked
    cboRep.ListIndex = -1
    
End Sub
