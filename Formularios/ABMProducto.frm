VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form ABMProducto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABM de Productos"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMProducto.frx":0000
      Height          =   750
      Left            =   5445
      Picture         =   "ABMProducto.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6200
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   6015
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   529
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Datos"
      TabPicture(0)   =   "ABMProducto.frx":0614
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMProducto.frx":0630
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraDatos 
         Caption         =   " Datos del Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5490
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   6630
         Begin VB.CheckBox chkServicio 
            Caption         =   "Servicio (no lleva stock)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3840
            TabIndex        =   38
            Top             =   4800
            Width           =   2460
         End
         Begin VB.ComboBox cboListaPrecio 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   4320
            Width           =   4185
         End
         Begin VB.CommandButton cmdLisPrecio 
            Height          =   315
            Left            =   5835
            MaskColor       =   &H000000FF&
            Picture         =   "ABMProducto.frx":064C
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Agregar Presentacion"
            Top             =   4320
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CheckBox chkBaja 
            Caption         =   "Dar de Baja"
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
            Left            =   1560
            TabIndex        =   26
            Top             =   4800
            Width           =   1620
         End
         Begin VB.TextBox txtpCpra 
            Height          =   300
            Left            =   1530
            MaxLength       =   10
            TabIndex        =   5
            Top             =   2850
            Width           =   1400
         End
         Begin VB.CommandButton cmdNuevaPresentacion 
            Height          =   315
            Left            =   5835
            MaskColor       =   &H000000FF&
            Picture         =   "ABMProducto.frx":09D6
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Agregar Presentacion"
            Top             =   2325
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevaLinea 
            Height          =   315
            Left            =   5835
            MaskColor       =   &H000000FF&
            Picture         =   "ABMProducto.frx":0D60
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Agregar Línea"
            Top             =   1515
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevoRubro 
            Height          =   315
            Left            =   5835
            MaskColor       =   &H000000FF&
            Picture         =   "ABMProducto.frx":10EA
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Agregar Rubro"
            Top             =   1920
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtStock 
            Height          =   300
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   7
            Top             =   3810
            Width           =   1400
         End
         Begin VB.TextBox txtPrecio 
            Height          =   300
            Left            =   1530
            MaxLength       =   10
            TabIndex        =   6
            Top             =   3330
            Width           =   1400
         End
         Begin VB.ComboBox cboPres 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2340
            Width           =   4230
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   300
            Left            =   1530
            MaxLength       =   40
            TabIndex        =   0
            Top             =   360
            Width           =   1400
         End
         Begin VB.TextBox txtDescri 
            Height          =   660
            Left            =   1530
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   1
            Tag             =   "Descripción"
            Top             =   720
            Width           =   4710
         End
         Begin VB.ComboBox cboLinea 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1515
            Width           =   4230
         End
         Begin VB.ComboBox cboRubro 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1920
            Width           =   4230
         End
         Begin VB.TextBox txtSFisico 
            Height          =   300
            Left            =   4840
            MaxLength       =   50
            TabIndex        =   8
            Top             =   3840
            Width           =   1400
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Stock Fisico:"
            Height          =   195
            Left            =   3840
            TabIndex        =   37
            Top             =   3900
            Width           =   915
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Lista de Precios:"
            Height          =   195
            Left            =   270
            TabIndex        =   36
            Top             =   4350
            Width           =   1170
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Precio Compra ($):"
            Height          =   195
            Left            =   180
            TabIndex        =   34
            Top             =   2940
            Width           =   1305
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Marca:"
            Height          =   195
            Left            =   945
            TabIndex        =   33
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   900
            TabIndex        =   32
            Top             =   360
            Width           =   660
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   555
            TabIndex        =   31
            Top             =   780
            Width           =   885
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Familia:"
            Height          =   195
            Left            =   915
            TabIndex        =   30
            Top             =   1560
            Width           =   525
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Rubro:"
            Height          =   195
            Left            =   960
            TabIndex        =   29
            Top             =   1980
            Width           =   480
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Precio Vta ($):"
            Height          =   195
            Left            =   480
            TabIndex        =   28
            Top             =   3420
            Width           =   1005
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Stock Mínimo:"
            Height          =   195
            Left            =   405
            TabIndex        =   27
            Top             =   3870
            Width           =   1035
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74865
         TabIndex        =   18
         Top             =   600
         Width           =   6780
         Begin VB.TextBox TxtDescriB 
            Height          =   330
            Left            =   1260
            MaxLength       =   55
            TabIndex        =   10
            Top             =   240
            Width           =   4770
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   360
            Left            =   6240
            MaskColor       =   &H000000FF&
            Picture         =   "ABMProducto.frx":1474
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Buscar"
            Top             =   210
            UseMaskColor    =   -1  'True
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Producto:"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   19
            Top             =   270
            Width           =   690
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4485
         Left            =   -74880
         TabIndex        =   12
         Top             =   1410
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   7911
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   20
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMProducto.frx":177E
      Height          =   750
      Left            =   4560
      Picture         =   "ABMProducto.frx":1A88
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6200
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMProducto.frx":1D92
      Height          =   750
      Left            =   6330
      Picture         =   "ABMProducto.frx":209C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6200
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      DisabledPicture =   "ABMProducto.frx":23A6
      Height          =   750
      Left            =   3675
      Picture         =   "ABMProducto.frx":26B0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6200
      Width           =   870
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
      Left            =   240
      TabIndex        =   21
      Top             =   6200
      Width           =   750
   End
End
Attribute VB_Name = "ABMProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rec As ADODB.Recordset
Dim Rec1 As ADODB.Recordset
Dim Consulta As Boolean
Dim sql As String
Dim resp As Integer
Public CODIGOLISTA As Integer


Private Sub cboLinea_Click()
    cboRubro.Clear
    'cboPres.Clear
End Sub

Private Sub cboLinea_LostFocus()
    'If Consulta = False Then
        cargocboRubro
    'End If
End Sub
Private Sub chkkitcon_Click()

End Sub

Private Sub CmdBorrar_Click()
    On Error GoTo CLAVOSE
    If Trim(TxtCodigo) <> "" Then
        resp = MsgBox("Seguro desea eliminar el Producto: " & Trim(txtdescri.Text) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        lblestado.Caption = "Eliminando ..."
        
        DBConn.Execute "DELETE FROM PRODUCTO WHERE PTO_CODIGO LIKE '" & TxtCodigo & "'"
        ' ,PTO_CODIGO
        DBConn.Execute "DELETE FROM DETALLE_STOCK WHERE PTO_CODIGO LIKE '" & TxtCodigo & "' AND STK_CODIGO =1"
        lblestado.Caption = ""
        Screen.MousePointer = vbNormal
        CmdNuevo_Click
    End If
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblestado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    MousePointer = vbHourglass
    TxtDescriB.Text = Replace(TxtDescriB, "'", "´")
    sql = "SELECT top 1000 P.PTO_CODIGO, P.PTO_DESCRI, L.LNA_DESCRI, R.RUB_DESCRI"
    sql = sql & " FROM PRODUCTO P, LINEAS L, RUBROS R"
    sql = sql & " WHERE"
    sql = sql & " P.RUB_CODIGO = R.RUB_CODIGO AND P.LNA_CODIGO=L.LNA_CODIGO"
    sql = sql & " AND R.LNA_CODIGO=L.LNA_CODIGO"
    sql = sql & " AND (PTO_DESCRI LIKE '" & TxtDescriB.Text & "%' "
    sql = sql & " OR PTO_CODIGO LIKE '" & TxtDescriB.Text & "%' )"
    'sql = sql & " AND P.PTO_ESTADO <> 1"
    sql = sql & " ORDER BY PTO_DESCRI"
        
    lblestado.Caption = "Buscando..."
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        Do While Not rec.EOF
           GrdModulos.AddItem rec!PTO_CODIGO & Chr(9) & Trim(rec!PTO_DESCRI) & Chr(9) & _
                              rec!RUB_DESCRI & Chr(9) & rec!LNA_DESCRI
           rec.MoveNext
        Loop
        If GrdModulos.Enabled Then GrdModulos.SetFocus
    Else
        lblestado.Caption = ""
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
        TxtDescriB.SetFocus
    End If
    rec.Close
    MousePointer = vbNormal
    lblestado.Caption = ""
End Sub

Private Sub CmdGrabar_Click()
    If validarProuducto = False Then Exit Sub
    
    On Error GoTo HayError
    Screen.MousePointer = vbHourglass
    lblestado.Caption = "Guardando ..."
    DBConn.BeginTrans
    sql = "SELECT * FROM PRODUCTO WHERE PTO_CODIGO LIKE '" & TxtCodigo.Text & "'"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        sql = "UPDATE PRODUCTO "
        sql = sql & " SET PTO_DESCRI=" & XS(txtdescri, True)
        sql = sql & " , LNA_CODIGO=" & cboLinea.ItemData(cboLinea.ListIndex)
        sql = sql & " , RUB_CODIGO=" & cboRubro.ItemData(cboRubro.ListIndex)
        sql = sql & " , TPRE_CODIGO=" & cboPres.ItemData(cboPres.ListIndex)
        sql = sql & " , PTO_PRECIO=" & XN(txtPrecio)
        sql = sql & " , PTO_STKMIN=" & XN(txtStock)
        'sql = sql & " , PTO_SFISICO=" & XN(txtSFisico)
        If cboListaPrecio.ListIndex > 0 Then
            sql = sql & " , LIS_CODIGO=" & cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
        Else
            sql = sql & " , LIS_CODIGO=" & 0
        End If
        If chkBaja.Value = Checked Then
            sql = sql & " ,PTO_ESTADO=1"
        Else
            sql = sql & " ,PTO_ESTADO=0"
        End If
        If chkServicio.Value = Checked Then
            sql = sql & " ,PTO_SERVICIO=1"
        Else
            sql = sql & " ,PTO_SERVICIO=0"
        End If
        sql = sql & " , PTO_PRECIOC=" & XN(txtpCpra)
        
        sql = sql & " WHERE PTO_CODIGO LIKE '" & TxtCodigo.Text & "'"
        DBConn.Execute sql
        
        'UPDATE DETALLE STOCK
        sql = "UPDATE DETALLE_STOCK "
        sql = sql & " SET DST_STKFIS =" & XN(txtSFisico.Text)
        sql = sql & " WHERE PTO_CODIGO = " & XN(TxtCodigo.Text)
        DBConn.Execute sql
        
    Else
        TxtCodigo = "1"
        sql = "SELECT MAX(PTO_CODIGO) as Maximo FROM PRODUCTO WHERE PTO_CODIGO <> 99999999"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(Rec1!Maximo) Then TxtCodigo = XN(Rec1!Maximo) + 1
        Rec1.Close
                
        sql = "INSERT INTO PRODUCTO(PTO_CODIGO,PTO_DESCRI,LNA_CODIGO,"
        sql = sql & "RUB_CODIGO,TPRE_CODIGO,PTO_PRECIO,PTO_STKMIN,PTO_ESTADO,LIS_CODIGO,PTO_PRECIOC,PTO_SERVICIO)" ',PTO_SFISICO)"
        'sql = sql & "PTO_TIPO,PTO_TIPMOD,PTO_TRACCI,PTO_CABINA,PTO_MOTMAR,PTO_MOTMOD,"
        'sql = sql & "PTO_ASPIRA,PTO_MOTNRO,PTO_CHASIS,PTO_SERIE,PTO_NEUMDE,PTO_NEDECA,"
        'sql = sql & "PTO_NEUMTR,PTO_NETRCA,PTO_KITCON,PTO_SALHID,PTO_POSARA,PTO_CERFAB,"
        'sql = sql & "PTO_OPCION1,PTO_OPCION2,PTO_PRECIOC,PTO_CODCLA)"
        sql = sql & " VALUES ("
        sql = sql & XS(TxtCodigo, True) & ","
        sql = sql & XS(Trim(txtdescri), True) & ","
        sql = sql & cboLinea.ItemData(cboLinea.ListIndex) & ","
        sql = sql & cboRubro.ItemData(cboRubro.ListIndex) & ","
        sql = sql & cboPres.ItemData(cboPres.ListIndex) & ","
        sql = sql & XN(txtPrecio) & ","
        sql = sql & XN(txtStock) & ","
        If chkBaja.Value = True Then
            sql = sql & "1" & "," 'DADO DE BAJA
        Else
            sql = sql & "0" & "," 'NORMAL
        End If
        If cboListaPrecio.ListIndex > 0 Then
            sql = sql & cboListaPrecio.ItemData(cboListaPrecio.ListIndex) & ","
        Else
            sql = sql & 0 & ","
        End If
        sql = sql & XN(txtpCpra) & ","
        'sql = sql & XN(txtSFisico) & ")"
        If chkServicio.Value = Checked Then
            sql = sql & "1" & ")" 'SERVICIO
        Else
            sql = sql & "0" & ")" 'NORMAL
        End If
        DBConn.Execute sql
                
        'Aca Inserto el producto en la tabla DETALLE_STOCK
        
        sql = "SELECT * FROM STOCK ORDER BY STK_CODIGO"
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec2.EOF = False Then
            Do While Rec2.EOF <> True
                sql = "INSERT INTO DETALLE_STOCK(STK_CODIGO,PTO_CODIGO,DST_STKFIS)"
                sql = sql & " VALUES ("
                sql = sql & XN(Rec2!STK_CODIGO) & ","
                sql = sql & XN(TxtCodigo) & ","
                sql = sql & XN(txtSFisico) & ") "
                DBConn.Execute sql
                Rec2.MoveNext
            Loop
        End If
        Rec2.Close
        
        
        'INSERTO EN LISTA DE PRECIOS
        If CODIGOLISTA <> 0 Then
'            sql = "INSERT INTO DETALLE_LISTA_PRECIO(LIS_CODIGO,"
'            sql = sql & "PTO_CODIGO,LIS_PRECIO,LIS_PRECIOC)"
'            sql = sql & " VALUES ("
'            sql = sql & CODIGOLISTA & ","
'            sql = sql & XS(TxtCodigo) & ","
'            sql = sql & XN(txtPrecio.Text) & " ,"
'            sql = sql & XN(txtpCpra.Text) & " )"
'            DBConn.Execute sql
            sql = " UPDATE PRODUCTO "
            sql = sql & " SET LIS_CODIGO = " & CODIGOLISTA & " "
            sql = sql & " WHERE PTO_CODIGO LIKE '" & TxtCodigo.Text & "'"
            DBConn.Execute sql
        End If
    End If
    MsgBox "El producto se guardo correctamente", vbInformation, TIT_MSGBOX
    TxtCodigo.Text = ""
    txtdescri.SetFocus
    lblestado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.CommitTrans
    rec.Close
    
    If Consulta = 3 Then
        Me.Hide
    Else
        'CmdNuevo_Click
    End If
    
    Exit Sub
    
    
    
HayError:
    lblestado.Caption = ""
    DBConn.RollbackTrans
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX

End Sub

Private Sub cmdLisPrecio_Click()
    Consulta = 2
    FrmListadePrecios.Show vbModal
    Set FrmListadePrecios = Nothing
    cboListaPrecio.Clear
    CargoCboListaPrecio
End Sub

Private Sub cmdNuevaLinea_Click()
    ABMLinea.Show vbModal
    Set ABMLinea = Nothing
    cboLinea.Clear
    cargocboLinea
End Sub

Private Sub cmdNuevaPresentacion_Click()
    ABMPresentacion.Show vbModal
    Set ABMPresentacion = Nothing
    cboPres.Clear
    cargocbotiporep
End Sub
Private Sub CmdNuevo_Click()
    tabDatos.Tab = 0
    TxtCodigo.Text = ""
    txtdescri.Text = ""
    lblestado.Caption = ""
    txtPrecio.Text = "0,00"
    txtpCpra.Text = "0,00"
    txtStock.Text = ""
    GrdModulos.Rows = 1
    chkBaja.Value = Unchecked
    cboLinea.ListIndex = 0
    cboRubro.Clear
    'cboPres.ListIndex = 0
    TxtCodigo.SetFocus
    
    txtpCpra.Text = ""
    cboListaPrecio.ListIndex = 1
    txtSFisico.Text = ""
    Consulta = False
    chkServicio.Value = Unchecked
End Sub

Private Sub cmdNuevoRubro_Click()
    ABMRubro.cboLineas.Text = cboLinea.Text
    ABMRubro.Show vbModal
    Set ABMRubro = Nothing
    cboRubro.Clear
    cargocboRubro
End Sub

Private Sub cmdSalir_Click()
    If Consulta = 3 Then
        Consulta = 4
    End If
    Unload Me
    Set ABMProducto = Nothing
End Sub

Private Sub Form_Activate()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then tabDatos.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
Set rec = New ADODB.Recordset
Set Rec1 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)

    lblestado.Caption = ""
    GrdModulos.FormatString = "Código|Descripción|Rubro|Linea"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 4000
    GrdModulos.ColWidth(2) = 2500
    GrdModulos.ColWidth(3) = 2500
    GrdModulos.Rows = 1
    tabDatos.Tab = 0
    
    'TxtCodigo.SetFocus
    
    cargocboLinea
    'cargocboRubro
    'cargo combo Tipo de Presentación
    cargocbotiporep
    
    'cargocboLista de Precio
    CargoCboListaPrecio
    
    txtPrecio.Text = "0,00"
    txtpCpra.Text = "0,00"
    
    Consulta = False
'    sql = "UPDATE PRODUCTO SET TPRE_CODIGO = 0 WHERE PTO_CODIGO LIKE '1840'"
'    DBConn.Execute sql
    
    
'    DBConn.Execute "delete FROM  DETALLE_STOCK"
'    sql = "SELECT PTO_CODIGO FROM PRODUCTO"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    Do While rec.EOF = False
'        sql = "INSERT INTO DETALLE_STOCK (STK_CODIGO,PTO_CODIGO) VALUES("
'        sql = sql & "1,"
'        sql = sql & XN(rec!PTO_CODIGO) & ")"
'        DBConn.Execute sql
'        rec.MoveNext
'    Loop
'    rec.Close
'
    
    
End Sub
Private Sub CargoCboListaPrecio()
    sql = "SELECT LIS_CODIGO, LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO"
    sql = sql & " ORDER BY LIS_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
    If rec.EOF = False Then
        rec.MoveFirst
        cboListaPrecio.AddItem ""
        Do While rec.EOF = False
            cboListaPrecio.AddItem rec!LIS_DESCRI
            cboListaPrecio.ItemData(cboListaPrecio.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cboListaPrecio.ListIndex = 1
    End If
    rec.Close
End Sub
Private Sub cargocboLinea()
    sql = "SELECT * FROM LINEAS  ORDER BY LNA_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
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
    Dim Posicion As Integer
    Posicion = 0
    If cboRubro.ListIndex <> -1 Then
        Posicion = cboRubro.ItemData(cboRubro.ListIndex)
    End If
    
    If ActiveControl.Name = "cmdGrabar" Or ActiveControl.Name = "cmdBorrar" Then Exit Sub
    cboRubro.Clear
    sql = "SELECT * FROM RUBROS "
    sql = sql & " WHERE LNA_CODIGO= " & cboLinea.ItemData(cboLinea.ListIndex)
    sql = sql & " ORDER BY RUB_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboRubro.AddItem rec!RUB_DESCRI
            cboRubro.ItemData(cboRubro.NewIndex) = rec!RUB_CODIGO
            rec.MoveNext
        Loop
        If Posicion = 0 Then
            cboRubro.ListIndex = 0
        Else
            Call BuscaCodigoProxItemData(Posicion, cboRubro)
        End If
    End If
    rec.Close
End Sub

Private Sub cargocbotiporep()
    cboPres.Clear
    sql = "SELECT DISTINCT TPRE_CODIGO,TPRE_DESCRI  FROM TIPO_PRESENTACION "
    'sql = sql & " WHERE LNA_CODIGO = " & cboLinea.ItemData(cboLinea.ListIndex) & " "
    'If cborubro.ListIndex <> -1 Then
   '     sql = sql & " AND RUB_CODIGO = " & cborubro.ItemData(cborubro.ListIndex) & " "
   ' End If
    sql = sql & " ORDER BY TPRE_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboPres.AddItem rec!TPRE_DESCRI
            cboPres.ItemData(cboPres.NewIndex) = rec!TPRE_CODIGO
            rec.MoveNext
        Loop
        cboPres.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.row > 0 Then
           TxtCodigo = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
           cmdGrabar.Enabled = True
           CmdBorrar.Enabled = True
           Consulta = True
           TxtCodigo_LostFocus
           tabDatos.Tab = 0
           
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

'Private Sub mtxttipo_GotFocus()
'    SelecTexto mtxttipo
'End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
    If tabDatos.Tab = 0 And Me.Visible Then
        txtdescri.SetFocus
        cmdGrabar.Enabled = True
        CmdBorrar.Enabled = True
    End If
    If tabDatos.Tab = 1 Then
        'TxtDescriB.Text = ""
        TxtDescriB.SetFocus
        cmdGrabar.Enabled = False
        CmdBorrar.Enabled = False
        If TxtDescriB.Text <> "" Then
            CmdBuscAprox_Click
        End If
    End If
End Sub

Private Sub TxtCodigo_Change()
    If Trim(TxtCodigo) = "" And CmdBorrar.Enabled Then
        CmdBorrar.Enabled = False
    ElseIf Trim(TxtCodigo) <> "" Then
        CmdBorrar.Enabled = True
    End If
    
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto TxtCodigo
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
'    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    TxtCodigo.Text = Replace(TxtCodigo.Text, "'", "´")
    
    If TxtCodigo.Text <> "" Then
        sql = "SELECT * FROM PRODUCTO"
        sql = sql & " WHERE PTO_CODIGO LIKE '" & TxtCodigo.Text & "'"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtdescri.Text = Rec1!PTO_DESCRI
            Call BuscaCodigoProxItemData(Rec1!LNA_CODIGO, cboLinea)
            cboLinea_LostFocus
            Call BuscaCodigoProxItemData(Rec1!RUB_CODIGO, cboRubro)
            'cboRubro_LostFocus
            Call BuscaCodigoProxItemData(Rec1!TPRE_CODIGO, cboPres)
            Call BuscaCodigoProxItemData(Rec1!LIS_CODIGO, cboListaPrecio)
            txtPrecio.Text = IIf(IsNull(Rec1!PTO_PRECIO), "", Format(Rec1!PTO_PRECIO, "0.00"))
            txtStock.Text = IIf(IsNull(Rec1!PTO_STKMIN), "", Rec1!PTO_STKMIN)
            If Rec1!PTO_ESTADO = 1 Then chkBaja.Value = Checked
            If Rec1!PTO_SERVICIO = 1 Then chkServicio.Value = Checked
            txtdescri.SetFocus
            txtpCpra.Text = IIf(IsNull(Rec1!PTO_PRECIOC), "", Format(Rec1!PTO_PRECIOC, "0.00"))
            
            'Busco Stock en Detalle Stock
            sql = "SELECT DST_STKFIS FROM DETALLE_STOCK WHERE PTO_CODIGO =" & XN(TxtCodigo.Text)
            Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec2.EOF = False Then
                txtSFisico.Text = IIf(IsNull(Rec2!DST_STKFIS), "", Rec2!DST_STKFIS)
            End If
            Rec2.Close
            Consulta = True
        Else
         'MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
         'TxtCodigo.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtcodrepu_Change()
'    SelecTexto txtcodrepu
End Sub

Private Sub txtcodrepu_KeyPress(KeyAscii As Integer)
 '   KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDescri_Change()
    If Trim(txtdescri) = "" And cmdGrabar.Enabled Then
        cmdGrabar.Enabled = False
    Else
        cmdGrabar.Enabled = True
    End If
End Sub

Private Sub txtdescri_GotFocus()
    'SelecTexto txtDescri
End Sub

Private Sub txtdescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDescri_LostFocus()
    txtdescri.Text = Replace(txtdescri, "'", "´")
End Sub

Private Sub txtpCpra_GotFocus()
    SelecTexto txtpCpra
End Sub

Private Sub txtpCpra_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtpCpra, KeyAscii)
End Sub

Private Sub txtpCpra_LostFocus()
    If txtpCpra.Text <> "" Then
        txtpCpra.Text = Valido_Importe(txtpCpra)
    Else
        txtpCpra.Text = "0,00"
    End If
End Sub

Private Sub txtPrecio_GotFocus()
    SelecTexto txtPrecio
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPrecio, KeyAscii)
End Sub
Function validarProuducto()
'    If TxtCodigo.Text = "" Then
'        MsgBox "No ha ingresado el código del Producto", vbExclamation, TIT_MSGBOX
'        TxtCodigo.SetFocus
'        validarProuducto = False
'        Exit Function
'    End If
    If txtdescri.Text = "" Then
        MsgBox "No ha ingresado la Descripción", vbExclamation, TIT_MSGBOX
        txtdescri.SetFocus
        validarProuducto = False
        Exit Function
    End If
    If cboLinea.ListIndex = -1 Then
        MsgBox "No ha seleccionado la Línea del Producto", vbExclamation, TIT_MSGBOX
        cboLinea.SetFocus
        validarProuducto = False
        Exit Function
    End If
    If cboRubro.Text = "" Then
        MsgBox "No ha seleccionado el Rubro del Producto", vbExclamation, TIT_MSGBOX
        cboRubro.SetFocus
        validarProuducto = False
        Exit Function
    End If
    If cboPres.ListIndex = -1 Then
        MsgBox "No ha seleccionado la Marca del Producto", vbExclamation, TIT_MSGBOX
        cboPres.SetFocus
        validarProuducto = False
        Exit Function
    End If
    If txtPrecio.Text = "" Then
        MsgBox "No ha ingresado el Precio de Venta", vbExclamation, TIT_MSGBOX
        txtPrecio.SetFocus
        validarProuducto = False
        Exit Function
    End If
    If txtpCpra.Text = "" Then
        MsgBox "No ha ingresado el Precio de Compra", vbExclamation, TIT_MSGBOX
        txtpCpra.SetFocus
        validarProuducto = False
        Exit Function
    End If
    If txtStock.Text = "" Then
        MsgBox "No ha ingresado el Stock Mínimo", vbExclamation, TIT_MSGBOX
        txtStock.SetFocus
        validarProuducto = False
        Exit Function
    End If
'    If cboLinea.ItemData(cboLinea.ListIndex) = 7 Then 'si es repuesto
'        If txtcodrepu.Text = "" Then
'        MsgBox "No ha ingresado el Código de Clasificación del Repuesto", vbExclamation, TIT_MSGBOX
'        txtcodrepu.SetFocus
'        validarProuducto = False
'        Exit Function
'    End If
'    End If
    validarProuducto = True
End Function

Private Sub txtPrecio_LostFocus()
    If txtPrecio.Text <> "" Then
        txtPrecio.Text = Valido_Importe(txtPrecio)
    Else
        txtPrecio.Text = "0,00"
    End If
End Sub

Private Sub txtSFisico_GotFocus()
    SelecTexto txtSFisico
End Sub


Private Sub txtSFisico_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtStock_GotFocus()
    SelecTexto txtStock
End Sub

Private Sub txtStock_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub
