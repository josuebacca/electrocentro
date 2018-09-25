VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ABMCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABM de Clientes"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "ABMCliente.frx":0000
      Height          =   705
      Left            =   4875
      Picture         =   "ABMCliente.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6915
      Width           =   840
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMCliente.frx":0614
      Height          =   705
      Left            =   5730
      Picture         =   "ABMCliente.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6915
      Width           =   840
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "ABMCliente.frx":0C28
      Height          =   705
      Left            =   3165
      Picture         =   "ABMCliente.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6915
      Width           =   840
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMCliente.frx":123C
      Height          =   705
      Left            =   4020
      Picture         =   "ABMCliente.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6915
      Width           =   840
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   6810
      Left            =   80
      TabIndex        =   24
      Top             =   80
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   12012
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
      TabPicture(0)   =   "ABMCliente.frx":1850
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "ABMCliente.frx":186C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraDatos 
         Caption         =   "Crédito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6165
         Left            =   150
         TabIndex        =   28
         Top             =   540
         Width           =   6225
         Begin MSMask.MaskEdBox txtCUIT 
            Height          =   375
            Left            =   1350
            TabIndex        =   56
            Top             =   1560
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.CheckBox chkMorDias 
            Caption         =   "Si supera"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   2543
            Width           =   975
         End
         Begin VB.CheckBox chkMorPor 
            Caption         =   "Porcentaje"
            Height          =   255
            Left            =   3480
            TabIndex        =   8
            Top             =   2543
            Width           =   1095
         End
         Begin VB.TextBox txtMorPor 
            Enabled         =   0   'False
            Height          =   300
            Left            =   4590
            MaxLength       =   10
            TabIndex        =   9
            Top             =   2520
            Width           =   1140
         End
         Begin VB.TextBox txtMorDias 
            Height          =   300
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   7
            Top             =   2520
            Width           =   1260
         End
         Begin VB.TextBox txtcodpostal 
            Enabled         =   0   'False
            Height          =   300
            Left            =   4800
            MaxLength       =   10
            TabIndex        =   53
            Top             =   3960
            Width           =   1020
         End
         Begin VB.CheckBox chkBaja 
            Caption         =   "Dar de Baja Cliente"
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
            Left            =   1350
            TabIndex        =   51
            Top             =   5850
            Width           =   2415
         End
         Begin VB.CommandButton cmdNuevoCanal 
            Height          =   315
            Left            =   3015
            MaskColor       =   &H000000FF&
            Picture         =   "ABMCliente.frx":1888
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Agregar Canal"
            Top             =   1905
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevaProvincia 
            Height          =   315
            Left            =   4290
            MaskColor       =   &H000000FF&
            Picture         =   "ABMCliente.frx":1C12
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Agregar Provincia"
            Top             =   3585
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevaLocalidad 
            Height          =   315
            Left            =   4305
            MaskColor       =   &H000000FF&
            Picture         =   "ABMCliente.frx":1F9C
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Agregar Localidad"
            Top             =   3960
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevoPais 
            Height          =   315
            Left            =   4290
            MaskColor       =   &H000000FF&
            Picture         =   "ABMCliente.frx":2326
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Agregar País"
            Top             =   3210
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtCredito 
            Height          =   300
            Left            =   4590
            MaxLength       =   10
            TabIndex        =   5
            Top             =   1920
            Width           =   1140
         End
         Begin VB.ComboBox cboCanal 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1913
            Width           =   1605
         End
         Begin VB.ComboBox cboIva 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1166
            Width           =   3375
         End
         Begin VB.TextBox txtIngBrutos 
            Height          =   285
            Left            =   4590
            MaxLength       =   10
            TabIndex        =   3
            Top             =   1545
            Width           =   1125
         End
         Begin VB.ComboBox cboProv 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   3585
            Width           =   2895
         End
         Begin VB.ComboBox cboPais 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   3210
            Width           =   2895
         End
         Begin VB.ComboBox cboLocal 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   3960
            Width           =   2895
         End
         Begin VB.TextBox txtMail 
            Height          =   285
            Left            =   1350
            LinkTimeout     =   0
            MaxLength       =   40
            TabIndex        =   16
            Top             =   5430
            Width           =   4455
         End
         Begin VB.TextBox txtFax 
            Height          =   285
            Left            =   3975
            MaxLength       =   25
            TabIndex        =   15
            Top             =   5085
            Width           =   1815
         End
         Begin VB.TextBox txtDomici 
            Height          =   285
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   13
            Top             =   4350
            Width           =   4455
         End
         Begin VB.TextBox txtRazSoc 
            Height          =   285
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   1
            Tag             =   "Descripción"
            Top             =   825
            Width           =   4335
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   285
            Left            =   1350
            MaxLength       =   5
            TabIndex        =   0
            Top             =   420
            Width           =   975
         End
         Begin VB.TextBox txtTelefono 
            Height          =   285
            Left            =   1350
            MaxLength       =   25
            TabIndex        =   14
            Top             =   5085
            Width           =   1815
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Dias"
            Height          =   195
            Left            =   2640
            TabIndex        =   55
            Top             =   2580
            Width           =   315
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Morosidad"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   480
            TabIndex        =   54
            Top             =   2280
            Width           =   735
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FF0000&
            X1              =   540
            X2              =   5760
            Y1              =   2430
            Y2              =   2430
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "C.P."
            Height          =   195
            Left            =   5160
            TabIndex        =   52
            Top             =   3720
            Width           =   300
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Vias de Comunicación"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   480
            TabIndex        =   46
            Top             =   4740
            Width           =   1575
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FF0000&
            X1              =   540
            X2              =   5760
            Y1              =   4890
            Y2              =   4890
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   480
            TabIndex        =   45
            Top             =   2895
            Width           =   630
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            X1              =   540
            X2              =   5760
            Y1              =   3045
            Y2              =   3045
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Crédito:"
            Height          =   195
            Left            =   3930
            TabIndex        =   44
            Top             =   1950
            Width           =   540
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Canal:"
            Height          =   195
            Left            =   780
            TabIndex        =   43
            Top             =   1950
            Width           =   450
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cond. I.V.A.:"
            Height          =   195
            Left            =   330
            TabIndex        =   42
            Top             =   1200
            Width           =   900
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Ing. Brutos:"
            Height          =   195
            Left            =   3645
            TabIndex        =   41
            Top             =   1575
            Width           =   810
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Left            =   630
            TabIndex        =   40
            Top             =   1590
            Width           =   600
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "País:"
            Height          =   195
            Left            =   855
            TabIndex        =   39
            Top             =   3255
            Width           =   375
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "E-Mail:"
            Height          =   195
            Left            =   750
            TabIndex        =   38
            Top             =   5490
            Width           =   480
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
            Height          =   195
            Left            =   3555
            TabIndex        =   37
            Top             =   5130
            Width           =   300
         End
         Begin VB.Label Label7 
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   555
            TabIndex        =   36
            Top             =   5145
            Width           =   675
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   525
            TabIndex        =   35
            Top             =   3630
            Width           =   705
         End
         Begin VB.Label Label4 
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   495
            TabIndex        =   34
            Top             =   4005
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   555
            TabIndex        =   33
            Top             =   4380
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
            Left            =   480
            TabIndex        =   30
            Top             =   840
            Width           =   750
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
            Left            =   690
            TabIndex        =   29
            Top             =   465
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74865
         TabIndex        =   25
         Top             =   600
         Width           =   6135
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   360
            Left            =   5580
            MaskColor       =   &H000000FF&
            Picture         =   "ABMCliente.frx":26B0
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Buscar"
            Top             =   210
            UseMaskColor    =   -1  'True
            Width           =   420
         End
         Begin VB.TextBox TxtDescriB 
            Height          =   330
            Left            =   1275
            MaxLength       =   15
            TabIndex        =   21
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
            TabIndex        =   27
            Top             =   315
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   26
            Top             =   270
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4875
         Left            =   -74880
         TabIndex        =   23
         Top             =   1440
         Width           =   6210
         _ExtentX        =   10954
         _ExtentY        =   8599
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
         TabIndex        =   31
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
      TabIndex        =   32
      Top             =   7035
      Width           =   750
   End
End
Attribute VB_Name = "ABMCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rec As ADODB.Recordset
Dim sql As String
Dim resp As Integer
Dim Consulta As Boolean
Dim Pais As String
Dim Provincia As String



Private Sub cboLocal_Change()
    BuscoCodPostal cboLocal.ItemData(cboLocal.ListIndex)
End Sub

Private Sub cboLocal_Click()
    BuscoCodPostal cboLocal.ItemData(cboLocal.ListIndex)
End Sub

Private Sub cboLocal_LostFocus()
  '  BuscoCodPostal XN(cboLocal.ItemData(cboLocal.ListIndex))
End Sub

Private Sub CboPais_LostFocus()
   If ActiveControl.Name = "cmdGrabar" Or ActiveControl.Name = "cmdBorrar" Then Exit Sub
    If Consulta = True And cboPais.Text = Pais Then
        Exit Sub
    Else
        Pais = ""
    End If
    'cargo combo PROVINCIA
    
    cboProv.Clear
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

Private Sub Check1_Click()

End Sub

Private Sub chkMorDias_Click()
   'Dim Porcentaje As Double
    If chkMorDias.Value = 1 Then
        'Porcentaje = txtPorc.Text
        txtMorDias.Enabled = True
        txtMorDias.SetFocus
    Else
         If TxtCodigo.Text <> "" Then
            QuitoPorcRemitosCliente
        End If
        txtMorDias.Text = ""
    End If

End Sub

Private Sub chkMorPor_Click()
    Dim Porcentaje As Double
    If chkMorPor.Value = 1 Then
        'Porcentaje = txtPorc.Text
        txtMorPor.Enabled = True
        txtMorPor.SetFocus
    Else
        If TxtCodigo.Text <> "" Then
            QuitoPorcRemitosCliente
        End If
        txtMorPor.Text = ""
    End If

End Sub
Private Function QuitoPorcRemitosCliente()
    Dim TOTAL As String
    Dim subtotal As String
    Dim DIAS As Double
    If txtMorPor.Text <> "" Then
        sql = "SELECT * FROM REMITO_CLIENTE "
        sql = sql & "WHERE RCL_SALDO > 0"
        sql = sql & "AND CLI_CODIGO = " & XN(TxtCodigo)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                If txtMorDias.Text <> "" Then
                    DIAS = Date - rec!RCL_FECHA
                    If DIAS > txtMorDias.Text Then
                        TOTAL = Valido_Importe(rec!RCL_TOTAL - (rec!RCL_TOTAL * XN(txtMorPor)) / 100)
                        subtotal = Valido_Importe(rec!RCL_SUBTOTAL - (rec!RCL_SUBTOTAL * XN(txtMorPor)) / 100)
                    End If
                Else
                    TOTAL = Valido_Importe(rec!RCL_TOTAL - (rec!RCL_TOTAL * XN(txtMorPor)) / 100)
                    subtotal = Valido_Importe(rec!RCL_SUBTOTAL - (rec!RCL_SUBTOTAL * XN(txtMorPor)) / 100)
                End If
                
                sql = "UPDATE REMITO_CLIENTE SET"
                sql = sql & " RCL_TOTAL = " & XN(TOTAL)
                sql = sql & " , RCL_SUBTOTAL = " & XN(subtotal)
                sql = sql & " , RCL_PORCEN = " & XN(txtMorPor)
                sql = sql & " WHERE RCL_NUMERO = " & rec!RCL_NUMERO
                sql = sql & " AND RCL_SUCURSAL = " & rec!RCL_SUCURSAL
                DBConn.Execute sql
                rec.MoveNext
            Loop
        End If
    End If
    rec.Close
End Function

Private Sub CmdBorrar_Click()
    On Error GoTo CLAVOSE
    If Trim(TxtCodigo) <> "" Then
        resp = MsgBox("Seguro desea eliminar el Cliente: " & Trim(txtRazSoc.Text) & "? ", vbYesNo, "Eliminar:")
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Eliminando ..."
        
        DBConn.Execute "DELETE FROM CLIENTE WHERE CLI_CODIGO = " & XN(TxtCodigo)
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        CmdNuevo_Click
    End If
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    MousePointer = vbHourglass
    
    sql = "SELECT * FROM CLIENTE"
    sql = sql & " WHERE CLI_RAZSOC"
    sql = sql & " LIKE '" & TxtDescriB.Text & "%' ORDER BY CLI_RAZSOC"
        
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
        TxtDescriB.SetFocus
    End If
    rec.Close
    MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Private Sub cmdGrabar_Click()

    If ValidarCliente = False Then Exit Sub
    
    On Error GoTo HayError
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando ..."
    DBConn.BeginTrans
    If TxtCodigo.Text <> "" Then
        sql = "UPDATE CLIENTE "
        sql = sql & " SET CLI_RAZSOC=" & XS(txtRazSoc)
        sql = sql & " , IVA_CODIGO=" & cboIva.ItemData(cboIva.ListIndex)
        sql = sql & " , CLI_CUIT=" & XS(txtCUIT)
        sql = sql & " , CLI_INGBRU=" & XS(txtIngBrutos)
        sql = sql & " , CLI_DOMICI=" & XS(txtDomici)
        sql = sql & " , PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
        sql = sql & " , PRO_CODIGO=" & cboProv.ItemData(cboProv.ListIndex)
        sql = sql & " , LOC_CODIGO=" & cboLocal.ItemData(cboLocal.ListIndex)
        sql = sql & " , CLI_TELEFONO=" & XS(txtTelefono)
        sql = sql & " , CLI_FAX=" & XS(txtFax)
        sql = sql & " , CLI_MAIL=" & XS(txtMail)
        sql = sql & " , CLI_CREDITO=" & XN(txtCredito)
        sql = sql & " , CAN_CODIGO=" & cboCanal.ItemData(cboCanal.ListIndex)
        If chkBaja.Value = Checked Then
            sql = sql & " , CLI_ESTADO=2"
        Else
            sql = sql & " , CLI_ESTADO=1"
        End If
        sql = sql & " , CLI_MORDIAS=" & XN(txtMorDias.Text)
        sql = sql & " , CLI_MORPOR=" & XN(txtMorPor.Text)
        sql = sql & " WHERE CLI_CODIGO=" & XN(TxtCodigo)
        DBConn.Execute sql
        
        ActualizarRemitosCliente
    Else
        TxtCodigo = "1"
        sql = "SELECT MAX(CLI_CODIGO) as maximo FROM CLIENTE"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(rec.Fields!Maximo) Then TxtCodigo = XN(rec.Fields!Maximo) + 1
        rec.Close
        
        sql = "INSERT INTO CLIENTE(CLI_CODIGO,CLI_RAZSOC,CLI_DOMICI,"
        sql = sql & "CLI_CUIT,CLI_INGBRU,CLI_TELEFONO,CLI_FAX,CLI_MAIL,CLI_CREDITO,"
        sql = sql & "CAN_CODIGO,IVA_CODIGO,PAI_CODIGO,PRO_CODIGO,LOC_CODIGO,CLI_ESTADO,"
        sql = sql & "CLI_MORDIAS,CLI_MORPOR)"
        sql = sql & " VALUES ("
        sql = sql & XN(TxtCodigo) & ","
        sql = sql & XS(txtRazSoc) & ","
        sql = sql & XS(txtDomici) & ","
        sql = sql & XS(txtCUIT) & ","
        sql = sql & XS(txtIngBrutos) & ","
        sql = sql & XS(txtTelefono) & ","
        sql = sql & XS(txtFax) & ","
        sql = sql & XS(txtMail) & ","
        sql = sql & XN(txtCredito) & ","
        sql = sql & cboCanal.ItemData(cboCanal.ListIndex) & ","
        sql = sql & cboIva.ItemData(cboIva.ListIndex) & ","
        sql = sql & cboPais.ItemData(cboPais.ListIndex) & ","
        sql = sql & cboProv.ItemData(cboProv.ListIndex) & ","
        sql = sql & cboLocal.ItemData(cboLocal.ListIndex) & ","
        If chkBaja.Value = Checked Then
            sql = sql & "2," 'DADO DE BAJA
        Else
            sql = sql & "1," 'NORMAL
        End If
        sql = sql & XN(txtMorDias) & ","
        sql = sql & XN(txtMorPor) & ")"
        DBConn.Execute sql
        
        ActualizarRemitosCliente
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
Private Function ActualizarRemitosCliente()
    Dim TOTAL As String
    Dim subtotal As String
    Dim Porcentaje As Double
    Dim DIAS As Double
    If txtMorPor.Text <> "" Then
        sql = "SELECT * FROM REMITO_CLIENTE "
        sql = sql & "WHERE RCL_SALDO > 0"
        sql = sql & "AND CLI_CODIGO = " & XN(TxtCodigo)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                If txtMorDias.Text <> "" Then
                    DIAS = Date - rec!RCL_FECHA
                    If DIAS > txtMorDias.Text Then
                        If Not IsNull(rec!RCL_PORCEN) Then 'Lo saca si existe Y LO AGREGA
                            TOTAL = Valido_Importe(rec!RCL_TOTAL - (rec!RCL_TOTAL * rec!RCL_PORCEN) / 100)
                            subtotal = Valido_Importe(rec!RCL_SUBTOTAL - (rec!RCL_SUBTOTAL * rec!RCL_PORCEN) / 100)
                        End If
                            'Aplica el porcentaje
                        TOTAL = Valido_Importe(rec!RCL_TOTAL + (rec!RCL_TOTAL * XN(txtMorPor)) / 100)
                        subtotal = Valido_Importe(rec!RCL_SUBTOTAL + (rec!RCL_SUBTOTAL * XN(txtMorPor)) / 100)
                        Porcentaje = txtMorPor
                        
                    Else
                        'aca ver si ya fue aplicado antes o no
                        'If Not IsNull(rec!RCL_PORCEN) Or rec!RCL_PORCEN <> 0 Then
                        If rec!RCL_PORCEN <> 0 Then
                            TOTAL = Valido_Importe(rec!RCL_TOTAL - (rec!RCL_TOTAL * rec!RCL_PORCEN) / 100)
                            subtotal = Valido_Importe(rec!RCL_SUBTOTAL - (rec!RCL_SUBTOTAL * rec!RCL_PORCEN) / 100)
                            Porcentaje = 0
                        Else
                            TOTAL = Valido_Importe(rec!RCL_TOTAL)
                            subtotal = Valido_Importe(rec!RCL_SUBTOTAL)
                            Porcentaje = 0
                        End If
                    End If
                Else
                    If Not IsNull(rec!RCL_PORCEN) Then
                        TOTAL = Valido_Importe(rec!RCL_TOTAL - (rec!RCL_TOTAL * rec!RCL_PORCEN) / 100)
                        subtotal = Valido_Importe(rec!RCL_SUBTOTAL - (rec!RCL_SUBTOTAL * rec!RCL_PORCEN) / 100)
                    End If
                    TOTAL = Valido_Importe(rec!RCL_TOTAL + (rec!RCL_TOTAL * XN(txtMorPor)) / 100)
                    subtotal = Valido_Importe(rec!RCL_SUBTOTAL + (rec!RCL_SUBTOTAL * XN(txtMorPor)) / 100)
                    Porcentaje = txtMorPor
                    
                End If
                
                sql = "UPDATE REMITO_CLIENTE SET"
                sql = sql & " RCL_TOTAL = " & XN(TOTAL)
                sql = sql & " , RCL_SUBTOTAL = " & XN(subtotal)
                sql = sql & " , RCL_PORCEN = " & Porcentaje
                sql = sql & " WHERE RCL_NUMERO = " & rec!RCL_NUMERO
                sql = sql & " AND RCL_SUCURSAL = " & rec!RCL_SUCURSAL
                DBConn.Execute sql
                rec.MoveNext
            Loop
        End If
    rec.Close
    End If
    
End Function
Private Function ValidarCliente() As Boolean
    If txtRazSoc.Text = "" Then
        MsgBox "No ha ingresado la Razón Social", vbExclamation, TIT_MSGBOX
        txtRazSoc.SetFocus
        ValidarCliente = False
        Exit Function
    End If
    
        If cboIva.ListIndex = -1 Then
            MsgBox "No ha seleccionado condición de I.V.A.", vbExclamation, TIT_MSGBOX
            cboIva.SetFocus
            ValidarCliente = False
            Exit Function
        End If
    
    If cboCanal.ListIndex = -1 Then
        MsgBox "No ha seleccionado Canal", vbExclamation, TIT_MSGBOX
        cboCanal.SetFocus
        ValidarCliente = False
        Exit Function
    End If
    If (cboIva.ItemData(cboIva.ListIndex) <> 3) Then ' CONSUMIDOR FINAL
        If txtCUIT.Text = "" Then
            MsgBox "No ha ingresado el Número de C.U.I.T.", vbExclamation, TIT_MSGBOX
            txtCUIT.SetFocus
            ValidarCliente = False
            Exit Function
        End If
    End If
    If cboPais.ListIndex = -1 Then
        MsgBox "No ha seleccionado Pais", vbExclamation, TIT_MSGBOX
        cboPais.SetFocus
        ValidarCliente = False
        Exit Function
    End If
    If cboProv.ListIndex = -1 Then
        MsgBox "No ha seleccionado Provincia", vbExclamation, TIT_MSGBOX
        cboProv.SetFocus
        ValidarCliente = False
        Exit Function
    End If
    If cboLocal.ListIndex = -1 Then
        MsgBox "No ha seleccionado Localidad", vbExclamation, TIT_MSGBOX
        cboLocal.SetFocus
        ValidarCliente = False
        Exit Function
    End If
    If txtMorDias.Text <> "" Then
        If txtMorPor = "" Then
            MsgBox "No ha ingresado el Porcentaje", vbExclamation, TIT_MSGBOX
            txtMorPor.SetFocus
            ValidarCliente = False
            Exit Function
        End If
        
    End If
    ValidarCliente = True
End Function

Private Sub cmdNuevaLocalidad_Click()
    Dim Localidad As Integer
    If cboLocal.ListIndex <> -1 Then
        Localidad = cboProv.ItemData(cboProv.ListIndex)
    End If
    ABMLocalidad.Show vbModal
    Consulta = False
    cboLocal.Clear
    cboProv_LostFocus
    'If cboLocal <> "" Then
    '    Call BuscaCodigoProxItemData(Localidad, cboLocal)
    'End If
End Sub

Private Sub cmdNuevaProvincia_Click()
    Dim Provincia As Integer
    Provincia = cboPais.ItemData(cboPais.ListIndex)
    ABMProvincia.Show vbModal
    Consulta = False
    cboProv.Clear
    'Provincia = ""
    CboPais_LostFocus
    'Call BuscaCodigoProxItemData(Provincia, cboProv)
End Sub

Private Sub CmdNuevo_Click()
    TxtCodigo.Text = ""
    txtRazSoc.Text = ""
    txtTelefono.Text = ""
    txtFax.Text = ""
    txtMail.Text = ""
    lblEstado.Caption = ""
    txtCredito.Text = ""
    txtDomici.Text = ""
    chkBaja.Value = Unchecked
    txtIngBrutos.Text = ""
    txtCUIT.Text = ""
    GrdModulos.Rows = 1
    cboProv.Clear
    cboLocal.Clear
    cboPais.ListIndex = 0
    cboIva.ListIndex = 0
    cboCanal.ListIndex = 0
    tabDatos.Tab = 0
    Consulta = False
    Pais = ""
    Provincia = ""
    txtcodpostal.Text = ""
    TxtCodigo.SetFocus
    txtMorDias.Text = ""
    txtMorPor.Text = ""
    chkMorDias.Value = 0
    chkMorPor.Value = 0
    
End Sub

Private Sub cmdNuevoCanal_Click()
    ABMCanal.Show vbModal
    cboCanal.Clear
    LlenarComboCanal
    cboCanal.SetFocus
End Sub

Private Sub cmdNuevoPais_Click()
    ABMPais.Show vbModal
    Consulta = False
    cboPais.Clear
    Pais = ""
    LlenarComboPais
    cboPais.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set ABMCliente = Nothing
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
    Set Rec2 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)

    lblEstado.Caption = ""
    GrdModulos.FormatString = "Código|Razón Social"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 5000
    GrdModulos.Rows = 1
    tabDatos.Tab = 0
    
    'cargo combo pais
    LlenarComboPais
    'cargo combo canal
    LlenarComboCanal
    'cargo combo IVA
    LlenarComboIva
    'para la consulta
    Consulta = False 'no consulta true consulta
    Pais = ""
    Provincia = ""
    
    'sql = "UPDATE REMITO_CLIENTE SET "
    'sql = sql & " EST_CODIGO = 4 "
    'sql = sql & " WHERE CLI_CODIGO = 3 "
    'DBConn.Execute sql
    
'    sql = "UPDATE PERMISOS SET "
'    sql = sql & " PRM_SISTEMA = " & XS("ELECTROCENTRO")
'    sql = sql & " IVA_CODIGO = 3 "
'    sql = sql & " WHERE CLI_CUIT IS NULL "
'    sql = sql & ", PRO_CODIGO = 1 "
'    sql = sql & ", LOC_CODIGO = 127 "
'    DBConn.Execute sql

' ACA ARMO LO DE CTA CTES
'     sql = "SELECT C.RCL_NUMERO, C.FACIMPO FROM RC R, CT_RC C "
'     sql = sql & " WHERE R.RCL_NUMERO = C.RCL_NUMERO "
'     sql = sql & " AND C.FAPAGAS = 'N'"
'     rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'     If rec.EOF = False Then
'        Do While rec.EOF = False
'
'            sql = "UPDATE RC SET"
'            sql = sql & " RCL_SALDO = " & XN(rec!FACIMPO)
'            sql = sql & " WHERE RCL_NUMERO = " & XN(rec!RCL_NUMERO)
'            DBConn.Execute sql
'            rec.MoveNext
'        Loop
'     End If
'     rec.Close
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

Private Sub LlenarComboCanal()
    sql = "SELECT * FROM CANALES ORDER BY CAN_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboCanal.AddItem rec!CAN_DESCRI
            cboCanal.ItemData(cboCanal.NewIndex) = rec!CAN_CODIGO
            rec.MoveNext
        Loop
        cboCanal.ListIndex = 0
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

Private Sub txtCredito_GotFocus()
    SelecTexto txtCredito
End Sub

Private Sub txtCredito_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtCredito, KeyAscii)
End Sub

Private Sub txtCredito_LostFocus()
    txtCredito.Text = Valido_Importe(txtCredito.Text)
End Sub

Private Sub txtCUIT_LostFocus()
    If txtCUIT.Text <> "" Then
        If ValidoCuit(txtCUIT.Text) = False Then
         txtCUIT.SetFocus
        End If
    End If
End Sub

Private Sub txtdomici_GotFocus()
    SelecTexto txtDomici
End Sub

Private Sub txtdomici_KeyPress(KeyAscii As Integer)
     KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtfax_GotFocus()
    SelecTexto txtFax
End Sub

Private Sub txtIngBrutos_GotFocus()
    SelecTexto txtIngBrutos
End Sub

Private Sub txtIngBrutos_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtMorDias_GotFocus()
    SelecTexto txtMorDias
End Sub

Private Sub txtMorDias_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtMorPor_GotFocus()
    SelecTexto txtMorPor
End Sub

Private Sub txtMorPor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtMorPor, KeyAscii)
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
        sql = "SELECT * FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(TxtCodigo)
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            Consulta = True
            txtRazSoc.Text = Rec1!CLI_RAZSOC
            Call BuscaCodigoProxItemData(Rec1!IVA_CODIGO, cboIva)
            txtCUIT.Text = IIf(IsNull(Rec1!CLI_CUIT), "", Rec1!CLI_CUIT)
            txtIngBrutos.Text = IIf(IsNull(Rec1!CLI_INGBRU), "", Rec1!CLI_INGBRU)
            Call BuscaCodigoProxItemData(Rec1!CAN_CODIGO, cboCanal)
            
            If IsNull(Rec1!CLI_CREDITO) Then
                txtCredito.Text = ""
            Else
                txtCredito.Text = Valido_Importe(Rec1!CLI_CREDITO)
            End If
            
            Call BuscaCodigoProxItemData(Rec1!PAI_CODIGO, cboPais)
            CboPais_LostFocus
            Pais = cboPais.Text
            
            Call BuscaCodigoProxItemData(Rec1!PRO_CODIGO, cboProv)
            cboProv_LostFocus
            Provincia = cboProv.Text
            
            Call BuscaCodigoProxItemData(Rec1!LOC_CODIGO, cboLocal)
            BuscoCodPostal Rec1!LOC_CODIGO
            txtDomici.Text = IIf(IsNull(Rec1!CLI_DOMICI), "", Rec1!CLI_DOMICI)
            txtTelefono.Text = IIf(IsNull(Rec1!CLI_TELEFONO), "", Rec1!CLI_TELEFONO)
            txtFax.Text = IIf(IsNull(Rec1!CLI_FAX), "", Rec1!CLI_FAX)
            txtMail.Text = IIf(IsNull(Rec1!CLI_MAIL), "", Rec1!CLI_MAIL)
            If Rec1!CLI_ESTADO = 2 Then chkBaja.Value = Checked
            txtMorDias.Text = IIf(IsNull(Rec1!CLI_MORDIAS), 0, Rec1!CLI_MORDIAS)
            If XN(txtMorDias.Text) > 0 Then
                chkMorDias.Value = 1
            End If
            
            txtMorPor.Text = IIf(IsNull(Rec1!CLI_MORPOR), 0, Rec1!CLI_MORPOR)
            If txtMorPor.Text <> "" Or txtMorPor.Text > 0 Then
                chkMorPor.Value = 1
            End If
            txtRazSoc.SetFocus
        Else
         MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
         CmdNuevo_Click
        End If
        Rec1.Close
    End If
End Sub
Function BuscoCodPostal(Codigo As Integer) As String
    sql = "SELECT LOC_CODPOS FROM LOCALIDAD "
    sql = sql & "WHERE PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex) & ""
    sql = sql & " AND PRO_CODIGO = " & cboProv.ItemData(cboProv.ListIndex) & ""
    sql = sql & " AND LOC_CODIGO = " & Codigo
    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec2.EOF = False Then
        txtcodpostal.Text = IIf(IsNull(Rec2!LOC_CODPOS), "", Rec2!LOC_CODPOS)
    End If
    Rec2.Close
End Function

Private Sub txtmail_GotFocus()
    SelecTexto txtMail
End Sub

Private Sub txtmail_KeyPress(KeyAscii As Integer)
     KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtRazSoc_GotFocus()
    SelecTexto txtRazSoc
End Sub

Private Sub txtRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txttelefono_GotFocus()
    SelecTexto txtTelefono
End Sub
