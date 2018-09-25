VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ABMRepresentada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABM de Representada"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMRepresentada.frx":0000
      Height          =   750
      Left            =   4695
      Picture         =   "ABMRepresentada.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5895
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMRepresentada.frx":0614
      Height          =   750
      Left            =   5580
      Picture         =   "ABMRepresentada.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5895
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "ABMRepresentada.frx":0C28
      Height          =   750
      Left            =   2940
      Picture         =   "ABMRepresentada.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5895
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMRepresentada.frx":123C
      Height          =   750
      Left            =   3825
      Picture         =   "ABMRepresentada.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5895
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   5760
      Left            =   60
      TabIndex        =   14
      Top             =   90
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   10160
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
      TabPicture(0)   =   "ABMRepresentada.frx":1850
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMRepresentada.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraDatos 
         Caption         =   "Datos de la Representada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4980
         Left            =   150
         TabIndex        =   22
         Top             =   555
         Width           =   6105
         Begin MSMask.MaskEdBox txtCUIT 
            Height          =   375
            Left            =   1320
            TabIndex        =   42
            Top             =   1560
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.CommandButton cmdNuevoPais 
            Height          =   315
            Left            =   4170
            MaskColor       =   &H000000FF&
            Picture         =   "ABMRepresentada.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Agregar Pa�s"
            Top             =   2295
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevaLocalidad 
            Height          =   315
            Left            =   4680
            MaskColor       =   &H000000FF&
            Picture         =   "ABMRepresentada.frx":1C12
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Agregar Localidad"
            Top             =   3060
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevaProvincia 
            Height          =   315
            Left            =   4170
            MaskColor       =   &H000000FF&
            Picture         =   "ABMRepresentada.frx":1F9C
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Agregar Provincia"
            Top             =   2685
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtIngBrutos 
            Height          =   285
            Left            =   4005
            MaxLength       =   10
            TabIndex        =   3
            Top             =   1575
            Width           =   1005
         End
         Begin VB.ComboBox cboIva 
            Height          =   315
            Left            =   1245
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1200
            Width           =   3375
         End
         Begin VB.ComboBox cboProv 
            Height          =   315
            Left            =   1245
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2685
            Width           =   2880
         End
         Begin VB.ComboBox cboPais 
            Height          =   315
            Left            =   1245
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2295
            Width           =   2880
         End
         Begin VB.ComboBox cboLocal 
            Height          =   315
            Left            =   1245
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   3060
            Width           =   3375
         End
         Begin VB.TextBox txtmail 
            Height          =   300
            Left            =   1245
            LinkTimeout     =   0
            MaxLength       =   40
            TabIndex        =   10
            Top             =   4560
            Width           =   4335
         End
         Begin VB.TextBox txtfax 
            Height          =   300
            Left            =   4005
            MaxLength       =   25
            TabIndex        =   9
            Top             =   4200
            Width           =   1575
         End
         Begin VB.TextBox txtdomici 
            Height          =   300
            Left            =   1245
            MaxLength       =   50
            TabIndex        =   7
            Top             =   3450
            Width           =   4335
         End
         Begin VB.TextBox txtRazSoc 
            Height          =   315
            Left            =   1245
            MaxLength       =   50
            TabIndex        =   1
            Tag             =   "Descripci�n"
            Top             =   825
            Width           =   4335
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   285
            Left            =   1245
            MaxLength       =   40
            TabIndex        =   0
            Top             =   435
            Width           =   975
         End
         Begin VB.TextBox txttelefono 
            Height          =   300
            Left            =   1245
            MaxLength       =   25
            TabIndex        =   8
            Top             =   4200
            Width           =   1575
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Left            =   495
            TabIndex        =   38
            Top             =   1620
            Width           =   600
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Ing. Brutos:"
            Height          =   195
            Left            =   3030
            TabIndex        =   37
            Top             =   1620
            Width           =   810
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cond. I.V.A.:"
            Height          =   195
            Left            =   195
            TabIndex        =   36
            Top             =   1230
            Width           =   900
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Vias de Comunicaci�n"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   420
            TabIndex        =   35
            Top             =   3855
            Width           =   1575
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FF0000&
            X1              =   630
            X2              =   5580
            Y1              =   4005
            Y2              =   4005
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   420
            TabIndex        =   34
            Top             =   2010
            Width           =   630
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            X1              =   510
            X2              =   5580
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Pa�s:"
            Height          =   195
            Left            =   720
            TabIndex        =   33
            Top             =   2340
            Width           =   375
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "E-Mail:"
            Height          =   195
            Left            =   615
            TabIndex        =   32
            Top             =   4605
            Width           =   480
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
            Height          =   210
            Left            =   3540
            TabIndex        =   31
            Top             =   4230
            Width           =   300
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tel�fono:"
            Height          =   195
            Left            =   420
            TabIndex        =   30
            Top             =   4245
            Width           =   675
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   390
            TabIndex        =   29
            Top             =   2715
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   360
            TabIndex        =   28
            Top             =   3105
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   420
            TabIndex        =   27
            Top             =   3480
            Width           =   675
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Raz. Soc.:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   345
            TabIndex        =   24
            Top             =   870
            Width           =   750
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "C�digo:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   555
            TabIndex        =   23
            Top             =   480
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74880
         TabIndex        =   19
         Top             =   600
         Width           =   6135
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   360
            Left            =   5595
            MaskColor       =   &H000000FF&
            Picture         =   "ABMRepresentada.frx":2326
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Buscar Representada"
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   420
         End
         Begin VB.TextBox TxtDescriB 
            Height          =   330
            Left            =   1290
            MaxLength       =   15
            TabIndex        =   15
            Top             =   225
            Width           =   4215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Codigo:"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1665
            TabIndex        =   21
            Top             =   315
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripci�n:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   20
            Top             =   270
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4020
         Left            =   -74895
         TabIndex        =   17
         Top             =   1440
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   7091
         _Version        =   393216
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
         TabIndex        =   25
         Top             =   570
         Width           =   1065
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
      Height          =   240
      Left            =   105
      TabIndex        =   26
      Top             =   6075
      Width           =   750
   End
End
Attribute VB_Name = "ABMRepresentada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim resp As Integer
Dim Consulta As Boolean
Dim Pais As String
Dim Provincia As String

Private Sub CboPais_LostFocus()
    If ActiveControl.Name = "cmdGrabar" Or ActiveControl.Name = "cmdBorrar" Then Exit Sub
    
    If Consulta = True And cboPais.Text = Pais Then
        Exit Sub
    Else
        Pais = ""
    End If
    'cargo combo PROVINCIA
    sql = "SELECT * FROM PROVINCIA"
    sql = sql & " WHERE PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
    sql = sql & " ORDER BY PRO_DESCRI "
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboProv.AddItem rec!PRO_DESCRI
            cboProv.ItemData(cboProv.NewIndex) = rec!PRO_CODIGO
            rec.MoveNext
        Loop
        cboProv.ListIndex = 0
    End If
    rec.Close
    cboProv.SetFocus
End Sub

Private Sub cboProv_LostFocus()
   If ActiveControl.Name = "cmdGrabar" Or ActiveControl.Name = "cmdBorrar" Then Exit Sub
    If cboProv.ListIndex <> -1 Then
        If Consulta = True And cboProv.Text = Provincia Then
            Exit Sub
        Else
            Provincia = ""
        End If
        'cargo combo Localidad
        cboLocal.Clear
        sql = "SELECT * FROM LOCALIDAD"
        sql = sql & " WHERE PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
        sql = sql & " AND PRO_CODIGO=" & cboProv.ItemData(cboProv.ListIndex)
        sql = sql & " ORDER BY LOC_DESCRI "
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                cboLocal.AddItem rec!LOC_DESCRI
                cboLocal.ItemData(cboLocal.NewIndex) = rec!LOC_CODIGO
                rec.MoveNext
            Loop
            cboLocal.ListIndex = 0
        End If
        rec.Close
     End If
End Sub

Private Sub CmdBorrar_Click()
    On Error GoTo CLAVOSE
    If Trim(TxtCodigo) <> "" Then
        resp = MsgBox("Seguro desea eliminar la Representada: " & Trim(txtRazSoc) & "? ", 36, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = 11
        lblEstado.Caption = "Eliminando ..."
        DBConn.BeginTrans
        
        sql = "DELETE FROM REPRESENTADA"
        sql = sql & " WHERE REP_CODIGO = " & XN(TxtCodigo)
        DBConn.Execute sql
         
        DBConn.CommitTrans
        lblEstado.Caption = ""
        Screen.MousePointer = 1
        CmdNuevo_Click
    End If
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    DBConn.RollbackTrans
    Screen.MousePointer = 1
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    Screen.MousePointer = vbHourglass
    
    sql = "SELECT REP_CODIGO,REP_RAZSOC"
    sql = sql & " FROM REPRESENTADA"
    sql = sql & " WHERE REP_RAZSOC"
    sql = sql & " LIKE '" & Me.TxtDescriB.Text & "%' ORDER BY REP_RAZSOC"
        
    lblEstado.Caption = "Buscando..."
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        Do While Not rec.EOF
           GrdModulos.AddItem rec.Fields(0) & Chr(9) & rec.Fields(1)
           rec.MoveNext
        Loop
        If GrdModulos.Enabled Then GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
        Me.TxtDescriB.SetFocus
    End If
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Private Sub cmdGrabar_Click()
    
    If ValidarRepresentada = False Then Exit Sub
    
    On Error GoTo HayError
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando ..."
    DBConn.BeginTrans
    If TxtCodigo.Text <> "" Then
        sql = "UPDATE REPRESENTADA "
        sql = sql & " SET REP_RAZSOC=" & XS(txtRazSoc)
        sql = sql & " , IVA_CODIGO=" & cboIva.ItemData(cboIva.ListIndex)
        sql = sql & " , REP_CUIT=" & XS(txtCUIT)
        sql = sql & " , REP_INGBRU=" & XS(txtIngBrutos)
        sql = sql & " , REP_DOMICI=" & XS(txtdomici)
        sql = sql & " , PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
        sql = sql & " , PRO_CODIGO=" & cboProv.ItemData(cboProv.ListIndex)
        sql = sql & " , LOC_CODIGO=" & cboLocal.ItemData(cboLocal.ListIndex)
        sql = sql & " , REP_TELEFONO=" & XS(txttelefono)
        sql = sql & " , REP_FAX=" & XS(txtfax)
        sql = sql & " , REP_MAIL=" & XS(txtmail)
        sql = sql & " WHERE REP_CODIGO=" & XN(TxtCodigo)
        DBConn.Execute sql
        
    Else
        TxtCodigo = "1"
        sql = "SELECT MAX(REP_CODIGO) as maximo FROM REPRESENTADA"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(rec.Fields!Maximo) Then TxtCodigo = XN(rec.Fields!Maximo) + 1
        rec.Close
        
        sql = "INSERT INTO REPRESENTADA(REP_CODIGO,REP_RAZSOC,REP_DOMICI,"
        sql = sql & "REP_CUIT,REP_INGBRU,REP_TELEFONO,REP_FAX,REP_MAIL,"
        sql = sql & "IVA_CODIGO,PAI_CODIGO,PRO_CODIGO,LOC_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & XN(TxtCodigo) & ","
        sql = sql & XS(txtRazSoc) & ","
        sql = sql & XS(txtdomici) & ","
        sql = sql & XS(txtCUIT) & ","
        sql = sql & XS(txtIngBrutos) & ","
        sql = sql & XS(txttelefono) & ","
        sql = sql & XS(txtfax) & ","
        sql = sql & XS(txtmail) & ","
        sql = sql & cboIva.ItemData(cboIva.ListIndex) & ","
        sql = sql & cboPais.ItemData(cboPais.ListIndex) & ","
        sql = sql & cboProv.ItemData(cboProv.ListIndex) & ","
        sql = sql & cboLocal.ItemData(cboLocal.ListIndex) & ")"

        DBConn.Execute sql
    End If
    Screen.MousePointer = vbNormal
    DBConn.CommitTrans
    CmdNuevo_Click
    Exit Sub
    
HayError:
    lblEstado.Caption = ""
    DBConn.RollbackTrans
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub cmdNuevaLocalidad_Click()
    ABMLocalidad.Show vbModal
    cboLocal.Clear
    cboProv_LostFocus
End Sub

Private Sub cmdNuevaProvincia_Click()
    ABMProvincia.Show vbModal
    cboProv.Clear
    Provincia = ""
    CboPais_LostFocus
End Sub

Private Sub CmdNuevo_Click()
    tabDatos.Tab = 0
    TxtCodigo.Text = ""
    txtRazSoc.Text = ""
    txttelefono.Text = ""
    txtfax.Text = ""
    txtmail.Text = ""
    lblEstado.Caption = ""
    txtdomici.Text = ""
    txtIngBrutos.Text = ""
    txtCUIT.Text = ""
    'para la consulta
    Consulta = False 'no consulta true consulta
    Pais = ""
    Provincia = ""
    GrdModulos.Rows = 1
    cboProv.Clear
    cboLocal.Clear
    cboPais.ListIndex = 0
    cboIva.ListIndex = 0
    TxtCodigo.SetFocus
End Sub

Private Sub cmdNuevoPais_Click()
    ABMPais.Show vbModal
    cboPais.Clear
    Pais = ""
    LlenarComboPais
    cboPais.SetFocus
End Sub

Private Sub LlenarComboPais()
    sql = "SELECT * FROM PAIS ORDER BY PAI_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboPais.AddItem rec!PAI_DESCRI
            cboPais.ItemData(cboPais.NewIndex) = rec!PAI_CODIGO
            rec.MoveNext
        Loop
        cboPais.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set ABMRepresentada = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then tabDatos.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)

    lblEstado.Caption = ""
    GrdModulos.FormatString = "C�digo|Nombre"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 4000
    GrdModulos.Rows = 1
    tabDatos.Tab = 0
    
    'cargo combo pais
    LlenarComboPais
    'CARGO COMBO IVA
    LlenarComboIva
    'para la consulta
    Consulta = False 'no consulta true consulta
    Pais = ""
    Provincia = ""
End Sub

Private Sub LlenarComboIva()
    sql = "SELECT * FROM CONDICION_IVA ORDER BY IVA_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboIva.AddItem rec!IVA_DESCRI
            cboIva.ItemData(cboIva.NewIndex) = rec!IVA_CODIGO
            rec.MoveNext
        Loop
        cboIva.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.row > 0 Then
        TxtCodigo = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
        cmdGrabar.Enabled = True
        cmdBorrar.Enabled = True
        TxtCodigo_LostFocus
        tabDatos.Tab = 0
    End If
End Sub

Private Function ValidarRepresentada() As Boolean
    If txtRazSoc.Text = "" Then
        MsgBox "No ha ingresado la Raz�n Social", vbExclamation, TIT_MSGBOX
        txtRazSoc.SetFocus
        ValidarRepresentada = False
        Exit Function
    End If
    If cboIva.ListIndex = -1 Then
        MsgBox "No ha seleccionado condici�n de I.V.A.", vbExclamation, TIT_MSGBOX
        cboIva.SetFocus
        ValidarRepresentada = False
        Exit Function
    End If
    If txtCUIT.Text = "" Then
        MsgBox "No ha ingresado el N�mero de C.U.I.T.", vbExclamation, TIT_MSGBOX
        txtCUIT.SetFocus
        ValidarRepresentada = False
        Exit Function
    End If
    If cboPais.ListIndex = -1 Then
        MsgBox "No ha seleccionado Pais", vbExclamation, TIT_MSGBOX
        cboPais.SetFocus
        ValidarRepresentada = False
        Exit Function
    End If
    If cboProv.ListIndex = -1 Then
        MsgBox "No ha seleccionado Provincia", vbExclamation, TIT_MSGBOX
        cboProv.SetFocus
        ValidarRepresentada = False
        Exit Function
    End If
    If cboLocal.ListIndex = -1 Then
        MsgBox "No ha seleccionado Localidad", vbExclamation, TIT_MSGBOX
        cboLocal.SetFocus
        ValidarRepresentada = False
        Exit Function
    End If
    If txtdomici.Text = "" Then
        MsgBox "No ha ingresado el Domicilio", vbExclamation, TIT_MSGBOX
        txtdomici.SetFocus
        ValidarRepresentada = False
        Exit Function
    End If
    ValidarRepresentada = True
End Function

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
    If tabDatos.Tab = 0 And Me.Visible Then
        txtRazSoc.SetFocus
        cmdGrabar.Enabled = True
        cmdBorrar.Enabled = True
    End If
    If tabDatos.Tab = 1 Then
        TxtDescriB.Text = ""
        TxtDescriB.SetFocus
        cmdGrabar.Enabled = False
        cmdBorrar.Enabled = False
    End If
End Sub

Private Sub TxtCodigo_Change()
    If Trim(TxtCodigo) = "" And cmdBorrar.Enabled Then
        cmdBorrar.Enabled = False
    ElseIf Trim(TxtCodigo) <> "" Then
        cmdBorrar.Enabled = True
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto TxtCodigo
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCUIT_LostFocus()
    If txtCUIT.Text <> "" Then
        If ValidoCuit(txtCUIT.Text) = False Then
         txtCUIT.SetFocus
        End If
    End If
End Sub

Private Sub TxtDescriB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtdomici_GotFocus()
    SelecTexto txtdomici
End Sub

Private Sub txtdomici_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtfax_GotFocus()
    SelecTexto txtfax
End Sub

Private Sub txtfax_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtIngBrutos_GotFocus()
    SelecTexto txtIngBrutos
End Sub

Private Sub txtIngBrutos_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtmail_GotFocus()
    SelecTexto txtmail
End Sub

Private Sub txtmail_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtRazSoc_Change()
If Trim(txtRazSoc) = "" And cmdGrabar.Enabled Then
        cmdGrabar.Enabled = False
    Else
        cmdGrabar.Enabled = True
    End If
End Sub

Private Sub TxtCodigo_LostFocus()
    If TxtCodigo.Text <> "" Then
        sql = "SELECT * FROM REPRESENTADA"
        sql = sql & " WHERE REP_CODIGO=" & XN(TxtCodigo)
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtRazSoc.Text = Rec1!REP_RAZSOC
            Call BuscaCodigoProxItemData(Rec1!IVA_CODIGO, cboIva)
            txtCUIT.Text = Rec1!REP_CUIT
            txtIngBrutos.Text = IIf(IsNull(Rec1!REP_INGBRU), "", Rec1!REP_INGBRU)
            Call BuscaCodigoProxItemData(Rec1!PAI_CODIGO, cboPais)
            CboPais_LostFocus
            Pais = cboPais.Text
            
            Call BuscaCodigoProxItemData(Rec1!PRO_CODIGO, cboProv)
            cboProv_LostFocus
            Provincia = cboPais.Text
            
            Call BuscaCodigoProxItemData(Rec1!LOC_CODIGO, cboLocal)
            txtdomici.Text = Rec1!REP_DOMICI
            txttelefono.Text = IIf(IsNull(Rec1!REP_TELEFONO), "", Rec1!REP_TELEFONO)
            txtfax.Text = IIf(IsNull(Rec1!REP_FAX), "", Rec1!REP_FAX)
            txtmail.Text = IIf(IsNull(Rec1!REP_MAIL), "", Rec1!REP_MAIL)
            txtRazSoc.SetFocus
        Else
             MsgBox "El C�digo no existe", vbExclamation, TIT_MSGBOX
             TxtCodigo.Text = ""
             'para la consulta
            Consulta = False 'no consulta true consulta
            Pais = ""
            Provincia = ""
            TxtCodigo.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtRazSoc_GotFocus()
    SelecTexto txtRazSoc
End Sub

Private Sub txtRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txttelefono_GotFocus()
    SelecTexto txttelefono
End Sub

Private Sub txttelefono_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub
