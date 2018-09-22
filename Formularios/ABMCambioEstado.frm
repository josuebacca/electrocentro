VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Begin VB.Form ABMCambioEstado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ABM de Estado de Cheques"
   ClientHeight    =   6600
   ClientLeft      =   2280
   ClientTop       =   435
   ClientWidth     =   7200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7200
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "ABMCambioEstado.frx":0000
      Height          =   750
      Left            =   6120
      Picture         =   "ABMCambioEstado.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5805
      Width           =   915
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "ABMCambioEstado.frx":0614
      Height          =   750
      Left            =   4260
      Picture         =   "ABMCambioEstado.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5805
      Width           =   915
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      DisabledPicture =   "ABMCambioEstado.frx":0C28
      Height          =   750
      Left            =   5190
      Picture         =   "ABMCambioEstado.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5805
      Width           =   915
   End
   Begin VB.Frame Frame2 
      Height          =   2565
      Left            =   105
      TabIndex        =   14
      Top             =   0
      Width           =   6975
      Begin VB.CommandButton cmdBuscaCheque 
         Height          =   315
         Left            =   3000
         MaskColor       =   &H000000FF&
         Picture         =   "ABMCambioEstado.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Buscar Cheques"
         Top             =   210
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.TextBox TxtCodInt 
         BackColor       =   &H80000018&
         Height          =   345
         Left            =   5985
         TabIndex        =   21
         Top             =   180
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Frame Frame3 
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   420
         TabIndex        =   15
         Top             =   570
         Width           =   6135
         Begin VB.TextBox TxtSUCURSAL 
            Height          =   285
            Left            =   3720
            MaxLength       =   3
            TabIndex        =   3
            Top             =   285
            Width           =   540
         End
         Begin VB.TextBox TxtBANCO 
            Height          =   285
            Left            =   780
            MaxLength       =   3
            TabIndex        =   1
            Top             =   285
            Width           =   540
         End
         Begin VB.TextBox TxtLOCALIDAD 
            Height          =   285
            Left            =   2250
            MaxLength       =   3
            TabIndex        =   2
            Top             =   285
            Width           =   540
         End
         Begin VB.TextBox TxtCODIGO 
            Height          =   285
            Left            =   5145
            MaxLength       =   6
            TabIndex        =   4
            Top             =   285
            Width           =   795
         End
         Begin VB.TextBox TxtBanDescri 
            BackColor       =   &H00C0C0C0&
            Height          =   330
            Left            =   165
            TabIndex        =   16
            Top             =   645
            Width           =   5820
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Localidad:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   11
            Left            =   1470
            TabIndex        =   20
            Top             =   315
            Width           =   735
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Banco:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   10
            Left            =   210
            TabIndex        =   19
            Top             =   330
            Width           =   510
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   3000
            TabIndex        =   18
            Top             =   315
            Width           =   645
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
            Left            =   4515
            TabIndex        =   17
            Top             =   315
            Width           =   540
         End
      End
      Begin VB.TextBox TxtCheNumero 
         Height          =   315
         Left            =   1635
         MaxLength       =   10
         TabIndex        =   0
         Top             =   210
         Width           =   1230
      End
      Begin VB.TextBox TxtCheImport 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1635
         TabIndex        =   7
         Top             =   2160
         Width           =   1140
      End
      Begin FechaCtl.Fecha TxtCheFecVto 
         Height          =   300
         Left            =   5370
         TabIndex        =   6
         Top             =   1830
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el m�nimo permitido"
      End
      Begin FechaCtl.Fecha TxtCheFecEmi 
         Height          =   345
         Left            =   1650
         TabIndex        =   5
         Top             =   1830
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   609
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el m�nimo permitido"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro Cheque:"
         Height          =   195
         Index           =   7
         Left            =   615
         TabIndex        =   25
         Top             =   255
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Index           =   2
         Left            =   945
         TabIndex        =   24
         Top             =   2220
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  Vencimiento:"
         Height          =   195
         Index           =   5
         Left            =   3825
         TabIndex        =   23
         Top             =   1860
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  Emisi�n:"
         Height          =   195
         Index           =   3
         Left            =   390
         TabIndex        =   22
         Top             =   1860
         Width           =   1125
      End
   End
   Begin VB.TextBox TxtCheObserv 
      Height          =   660
      Left            =   120
      TabIndex        =   10
      Top             =   5010
      Width           =   6900
   End
   Begin VB.ComboBox CboEstado 
      Height          =   315
      Left            =   3915
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4380
      Width           =   3165
   End
   Begin FechaCtl.Fecha TxtCesFecha 
      Height          =   315
      Left            =   1965
      TabIndex        =   8
      Top             =   4380
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Separador       =   "/"
      Text            =   ""
      MensajeErrMin   =   "La fecha ingresada no alcanza el m�nimo permitido"
   End
   Begin MSFlexGridLib.MSFlexGrid Grd1 
      Height          =   1500
      Left            =   90
      TabIndex        =   26
      Top             =   2700
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   2646
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorSel    =   16777215
      ForeColorSel    =   -2147483624
      AllowBigSelection=   -1  'True
      Enabled         =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones:"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   30
      Top             =   4785
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Cambio de Estado:"
      Height          =   270
      Index           =   1
      Left            =   90
      TabIndex        =   29
      Top             =   4395
      Width           =   1905
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Estado:"
      Height          =   195
      Index           =   0
      Left            =   3180
      TabIndex        =   28
      Top             =   4410
      Width           =   690
      WordWrap        =   -1  'True
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
      Left            =   150
      TabIndex        =   27
      Top             =   5970
      Width           =   750
   End
End
Attribute VB_Name = "ABMCambioEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuscaCheque_Click()
    Dim codint As Integer
    frmBuscar.TipoBusqueda = 6
    frmBuscar.Show vbModal
    'TxtCheNumero.Text = frmBuscar.grdBuscar.Col
    TxtCheNumero.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
    TxtBanco.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 5)
    txtlocalidad.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 6)
    TxtSucursal.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 7)
    TxtCodigo.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 8)
    TxtCheImport.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 3)
    TxtCheFecVto.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 2)
    'TxtBanDescri.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0)
    
     If Trim(Me.TxtCheNumero.Text) <> "" And _
       Trim(Me.TxtBanco.Text) <> "" And _
       Trim(Me.txtlocalidad.Text) <> "" And _
       Trim(Me.TxtSucursal.Text) <> "" And _
       Trim(Me.TxtCodigo.Text) <> "" Then
       
       If Len(Me.TxtCodigo.Text) < 6 Then Me.TxtCodigo.Text = CompletarConCeros(Me.TxtCodigo.Text, 6)
           
       Dim MtrObjetos As Variant
    
       Set rec = New ADODB.Recordset
       Set Rec1 = New ADODB.Recordset
       
       'BUSCO EL CODIGO INTERNO
       sql = "SELECT BAN_CODINT,BAN_DESCRI FROM BANCO WHERE BAN_BANCO = " & _
       XS(TxtBanco.Text) & " AND BAN_LOCALIDAD = " & _
       XS(Me.txtlocalidad.Text) & " AND BAN_SUCURSAL = " & _
       XS(Me.TxtSucursal.Text) & " AND BAN_CODIGO = " & XS(Me.TxtCodigo.Text)
       rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
       If rec.RecordCount > 0 Then 'EXITE
            Me.TxtCodInt.Text = rec!BAN_CODINT
            TxtBanDescri.Text = rec!BAN_DESCRI
       Else
          MsgBox "NO ESTA REGISTRADO EL BANCO.", 16, TIT_MSGBOX
          Me.TxtCodigo.Text = ""
          Me.TxtCodigo.SetFocus
          Exit Sub
       End If
       rec.Close
       
       'CONSULTO SI EXISTE EL CHEQUE
        sql = "SELECT * FROM CHEQUE " & _
              " WHERE CHE_NUMERO = " & XS(TxtCheNumero.Text) & _
                " AND BAN_CODINT = " & XN(TxtCodInt.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then 'EXITE
            
            TxtCheNumero.Enabled = False
            TxtBanco.Enabled = False
            txtlocalidad.Enabled = False
            TxtSucursal.Enabled = False
            TxtCodigo.Enabled = False
            
            MtrObjetos = Array(TxtCheNumero, TxtBanco, txtlocalidad, TxtSucursal, TxtCodigo)
            Call CambiarColor(MtrObjetos, 5, &H80000018, "D")
            
            Me.TxtCheFecEmi.Text = rec!CHE_FECEMI
            Me.TxtCheFecVto.Text = rec!CHE_FECVTO
            Me.TxtCheImport.Text = Format(rec!che_import, "$ #0.00")

            'Cargo la Grilla
            sql = "SELECT CES_FECHA,EC.ECH_CODIGO,ECH_DESCRI,CES_DESCRI" & _
                  " FROM CHEQUE_ESTADOS CE, ESTADO_CHEQUE EC " & _
                  " WHERE CE.ECH_CODIGO = EC.ECH_CODIGO " & _
                    " AND CE.CHE_NUMERO = " & XS(TxtCheNumero.Text) & _
                    " AND CE.BAN_CODINT = " & XN(TxtCodInt.Text) & _
                    " ORDER BY CES_FECHA"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            Grd1.Rows = 1
            If Rec1.RecordCount > 0 Then
               Rec1.MoveFirst
               Do While Not Rec1.EOF
                 Grd1.AddItem Rec1!CES_FECHA & Chr(9) & Trim(Rec1.Fields(1)) & Chr(9) & Trim(Rec1.Fields(2))
                 Rec1.MoveNext
               Loop
            End If
            Rec1.MoveLast
            Call BuscaCodigoProxItemData(Rec1!ECH_CODIGO, CboEstado)
            Rec1.Close
            Me.TxtCesFecha.SetFocus
        End If

        rec.Close
    End If
    
End Sub

Private Sub CmdGrabar_Click()
 Dim Rec1 As New Recordset
 If Me.ActiveControl.Name <> "CmdNuevo" And Me.ActiveControl.Name <> "CmdSalir" Then

    'Verifico que NO graben dos veces el mismo estado en el mismo d�a
    sql = "SELECT ECH_CODIGO,MAX(CES_FECHA)as maximo"
    sql = sql & " FROM CHEQUE_ESTADOS "
    sql = sql & " WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text)
    sql = sql & " AND ECH_CODIGO = " & CboEstado.ItemData(CboEstado.ListIndex)
    sql = sql & " GROUP BY ECH_CODIGO,CES_FECHA"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.RecordCount > 0 Then
       If DMY(Rec1!Maximo) = DMY(TxtCesFecha.Text) Then
            MsgBox "NO se puede registrar el mismo car�cter en la misma fecha.", 16, TIT_MSGBOX
            Rec1.Close
            Exit Sub
       End If
    End If
    Rec1.Close
            
    If Trim(Me.TxtCheNumero.Text) = "" Or _
       Trim(Me.TxtBanco.Text) = "" Or _
       Trim(Me.txtlocalidad.Text) = "" Or _
       Trim(Me.TxtSucursal.Text) = "" Or _
       Trim(Me.TxtCodigo.Text) = "" Or _
       Trim(Me.TxtCesFecha.Text) = "" Then
       
        If Trim(Me.TxtCheNumero.Text) = "" Then
           MsgBox "Falta el N�mero de Cheque.", 16, TIT_MSGBOX
           TxtCheNumero.SetFocus
           Exit Sub
        End If
        
        If Trim(Me.TxtBanco.Text) = "" Then
           MsgBox "Falta el BANCO.", 16, TIT_MSGBOX
           TxtBanco.SetFocus
           Exit Sub
        End If
        
        If Trim(Me.txtlocalidad.Text) = "" Then
           MsgBox "Falta la LOCALIDAD.", 16, TIT_MSGBOX
           txtlocalidad.SetFocus
           Exit Sub
        End If
        
        If Trim(Me.TxtSucursal.Text) = "" Then
           MsgBox "Falta la SUCURSAL.", 16, TIT_MSGBOX
           TxtSucursal.SetFocus
           Exit Sub
        End If
        
        If Trim(Me.TxtCodigo.Text) = "" Then
           MsgBox "Falta el C�DIGO.", 16, TIT_MSGBOX
           TxtCodigo.SetFocus
           Exit Sub
        End If
        
        If Trim(Me.TxtCesFecha.Text) = "" Then
           MsgBox "Falta la Fecha.", 16, TIT_MSGBOX
           TxtCesFecha.SetFocus
           Exit Sub
        End If
 Else
        
        'Inserto en Cheque_Estados
         sql = "INSERT INTO CHEQUE_ESTADOS(ECH_CODIGO,BAN_CODINT,CHE_NUMERO,CES_FECHA,"
         sql = sql & " CES_DESCRI)VALUES ( " & CboEstado.ItemData(CboEstado.ListIndex)
         sql = sql & "," & XN(Me.TxtCodInt.Text) & "," & XS(Me.TxtCheNumero.Text)
         sql = sql & "," & XDQ(TxtCesFecha) & "," & XS(Me.TxtCheObserv.Text) & " )"
         DBConn.Execute sql
         
         CmdNuevo_Click
   End If
 End If
End Sub

Private Sub CmdNuevo_Click()
   Dim MtrObjetos As Variant
   
   Me.TxtCheNumero.Enabled = True
   Me.TxtBanco.Enabled = True
   Me.txtlocalidad.Enabled = True
   Me.TxtSucursal.Enabled = True
   Me.TxtCodigo.Enabled = True
   MtrObjetos = Array(TxtCheNumero, TxtBanco, txtlocalidad, TxtSucursal, TxtCodigo)
   Call CambiarColor(MtrObjetos, 5, &H80000005, "E")
            
   Me.TxtCheNumero.Text = ""
   Me.TxtBanco.Text = ""
   Me.txtlocalidad.Text = ""
   Me.TxtSucursal.Text = ""
   Me.TxtCodigo.Text = ""
   
   Me.TxtCodInt.Text = ""
   Me.TxtCheFecEmi.Text = ""
   Me.TxtCheFecVto.Text = ""
   Me.TxtCheImport.Text = ""
   Me.Grd1.Rows = 1
   Me.TxtCesFecha.Text = ""
   Me.CboEstado.ListIndex = 0
   Me.TxtCheObserv.Text = ""
   Me.TxtCheNumero.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    Set ABMCambioEstado = Nothing
End Sub

Private Sub Form_Activate()
    Call Centrar_pantalla(Me)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'si presiono ESCAPE salgo del form
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
    If KeyAscii = vbKeyReturn Then 'avanza de campo
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    
    lblestado.Caption = ""
    Grd1.FormatString = "^Fecha|Estado|Observaci�n"
    Grd1.ColWidth(0) = 1100
    Grd1.ColWidth(1) = 2500
    Grd1.ColWidth(2) = 4500
    Grd1.Rows = 1
    
    'Cargo el Combo de Estados
    Set rec = New ADODB.Recordset
    
    sql = "SELECT ECH_CODIGO,ECH_DESCRI FROM ESTADO_CHEQUE"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        Do While Not rec.EOF
            CboEstado.AddItem Trim(rec!ECH_DESCRI) '& Space(100 - Len(Trim(rec!ECH_DESCRI))) & Trim(rec!ech_codigo)
            CboEstado.ItemData(CboEstado.NewIndex) = rec!ECH_CODIGO
            rec.MoveNext
        Loop
        CboEstado.ListIndex = 0
    End If
    rec.Close
    
End Sub

Private Sub TxtBanco_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub TxtBANCO_LostFocus()
    If TxtBanco.Text <> "" Then TxtBanco.Text = Format(TxtBanco.Text, "000")
End Sub

Private Sub TxtCesFecha_LostFocus()
    If TxtCesFecha.Text = "" Then TxtCesFecha.Text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub TxtCheNumero_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCheNumero_LostFocus()
'If Me.ActiveControl.Name <> "CmdSalir" And Me.ActiveControl.Name <> "CmdNuevo" Then
   If Trim(Me.TxtCheNumero.Text) = "" Then
       Me.cmdBuscaCheque.SetFocus
    Else
        If Len(TxtCheNumero.Text) < 10 Then TxtCheNumero.Text = CompletarConCeros(TxtCheNumero.Text, 10)
        
        'busco el cheque
        If TxtCheNumero.Text <> "" Then
        If Len(TxtCheNumero.Text) < 8 Then TxtCheNumero.Text = CompletarConCeros(TxtCheNumero.Text, 8)
    'sql = "SELECT * FROM CHEQUE WHERE "
        sql = "SELECT DISTINCT CE.CHE_NUMERO, CH.CHE_IMPORT, CH.CHE_FECVTO, CE.BAN_CODINT, B.BAN_BANCO, B.BAN_LOCALIDAD,"
        sql = sql & " B.BAN_SUCURSAL, B.BAN_CODIGO, B.BAN_NOMCOR,CE.CES_DESCRI,B.BAN_DESCRI"
        sql = sql & " FROM CHEQUE_ESTADOS CE, CHEQUE CH, BANCO B,ESTADO_CHEQUE E"
        sql = sql & " Where "
        sql = sql & " CE.CHE_NUMERO = CH.CHE_NUMERO And "
        sql = sql & " CE.BAN_CODINT = CH.BAN_CODINT And "
        sql = sql & " CH.BAN_CODINT=B.BAN_CODINT  "
        'sql = sql & " CE.ECH_CODIGO= E.ECH_CODIGO AND" '
        'sql = sql & " E.ECH_CODIGO=7" ' 7-entregado
        sql = sql & " AND CH.CHE_NUMERO LIKE '%" & Trim(TxtCheNumero) & "%'"  'CODIGO (1) ES CHEQUE EN CARTERA
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            TxtCheNumero.Text = rec!CHE_NUMERO
            TxtBanco.Text = rec!BAN_BANCO
            txtlocalidad.Text = rec!BAN_LOCALIDAD
            TxtSucursal.Text = rec!BAN_SUCURSAL
            TxtCodigo.Text = rec!BAN_CODIGO
            TxtCheImport.Text = rec!che_import
            TxtCheFecVto.Text = rec!CHE_FECVTO
            TxtBanDescri.Text = rec!BAN_NOMCOR
            TxtCodInt.Text = rec!BAN_CODINT
        
        End If
        rec.Close
    End If
        
    
     If Trim(Me.TxtCheNumero.Text) <> "" And _
       Trim(Me.TxtBanco.Text) <> "" And _
       Trim(Me.txtlocalidad.Text) <> "" And _
       Trim(Me.TxtSucursal.Text) <> "" And _
       Trim(Me.TxtCodigo.Text) <> "" Then
       
       If Len(Me.TxtCodigo.Text) < 6 Then Me.TxtCodigo.Text = CompletarConCeros(Me.TxtCodigo.Text, 6)
           
       Dim MtrObjetos As Variant
    
       Set rec = New ADODB.Recordset
       Set Rec1 = New ADODB.Recordset
       
       'BUSCO EL CODIGO INTERNO
       sql = "SELECT BAN_CODINT,BAN_DESCRI FROM BANCO WHERE BAN_BANCO = " & _
       XS(TxtBanco.Text) & " AND BAN_LOCALIDAD = " & _
       XS(Me.txtlocalidad.Text) & " AND BAN_SUCURSAL = " & _
       XS(Me.TxtSucursal.Text) & " AND BAN_CODIGO = " & XS(Me.TxtCodigo.Text)
       rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
       If rec.RecordCount > 0 Then 'EXITE
            Me.TxtCodInt.Text = rec!BAN_CODINT
            TxtBanDescri.Text = rec!BAN_DESCRI
       Else
          MsgBox "NO ESTA REGISTRADO EL BANCO.", 16, TIT_MSGBOX
          Me.TxtCodigo.Text = ""
          Me.TxtCodigo.SetFocus
          Exit Sub
       End If
       rec.Close
       
       'CONSULTO SI EXISTE EL CHEQUE
        sql = "SELECT * FROM CHEQUE " & _
              " WHERE CHE_NUMERO = " & XS(TxtCheNumero.Text) & _
                " AND BAN_CODINT = " & XN(TxtCodInt.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then 'EXITE
            
            TxtCheNumero.Enabled = False
            TxtBanco.Enabled = False
            txtlocalidad.Enabled = False
            TxtSucursal.Enabled = False
            TxtCodigo.Enabled = False
            
            MtrObjetos = Array(TxtCheNumero, TxtBanco, txtlocalidad, TxtSucursal, TxtCodigo)
            Call CambiarColor(MtrObjetos, 5, &H80000018, "D")
            
            Me.TxtCheFecEmi.Text = rec!CHE_FECEMI
            Me.TxtCheFecVto.Text = rec!CHE_FECVTO
            Me.TxtCheImport.Text = Format(rec!che_import, "$ #0.00")

            'Cargo la Grilla
            sql = "SELECT CES_FECHA,EC.ECH_CODIGO,ECH_DESCRI,CES_DESCRI" & _
                  " FROM CHEQUE_ESTADOS CE, ESTADO_CHEQUE EC " & _
                  " WHERE CE.ECH_CODIGO = EC.ECH_CODIGO " & _
                    " AND CE.CHE_NUMERO = " & XS(TxtCheNumero.Text) & _
                    " AND CE.BAN_CODINT = " & XN(TxtCodInt.Text) & _
                    " ORDER BY CES_FECHA"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            Grd1.Rows = 1
            If Rec1.RecordCount > 0 Then
               Rec1.MoveFirst
               Do While Not Rec1.EOF
                 Grd1.AddItem Rec1!CES_FECHA & Chr(9) & Trim(Rec1.Fields(1)) & Chr(9) & Trim(Rec1.Fields(2))
                 Rec1.MoveNext
               Loop
            End If
            Rec1.MoveLast
            Call BuscaCodigoProxItemData(Rec1!ECH_CODIGO, CboEstado)
            Rec1.Close
            Me.TxtCesFecha.SetFocus
        End If

        rec.Close
    End If
        
        
        
    End If
' End If
End Sub

Private Sub TxtCheObserv_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCodigo_Change()
    If Trim(TxtCodigo) = "" And cmdNuevo.Enabled Then
        cmdNuevo.Enabled = False
    ElseIf Trim(TxtCodigo) <> "" Then
        cmdNuevo.Enabled = True
    End If
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    
'    If Trim(Me.TxtCheNumero.Text) <> "" And _
'       Trim(Me.TxtBanco.Text) <> "" And _
'       Trim(Me.txtlocalidad.Text) <> "" And _
'       Trim(Me.TxtSucursal.Text) <> "" And _
'       Trim(Me.TxtCodigo.Text) <> "" Then
'
'       If Len(Me.TxtCodigo.Text) < 6 Then Me.TxtCodigo.Text = CompletarConCeros(Me.TxtCodigo.Text, 6)
'
'       Dim MtrObjetos As Variant
'
'       Set rec = New ADODB.Recordset
'       Set Rec1 = New ADODB.Recordset
'
'       'BUSCO EL CODIGO INTERNO
'       sql = "SELECT BAN_CODINT,BAN_DESCRI FROM BANCO WHERE BAN_BANCO = " & _
'       XS(TxtBanco.Text) & " AND BAN_LOCALIDAD = " & _
'       XS(Me.txtlocalidad.Text) & " AND BAN_SUCURSAL = " & _
'       XS(Me.TxtSucursal.Text) & " AND BAN_CODIGO = " & XS(Me.TxtCodigo.Text)
'       rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'       If rec.RecordCount > 0 Then 'EXITE
'            Me.TxtCodInt.Text = rec!BAN_CODINT
'            TxtBanDescri.Text = rec!BAN_DESCRI
'       Else
'          MsgBox "NO ESTA REGISTRADO EL BANCO.", 16, TIT_MSGBOX
'          Me.TxtCodigo.Text = ""
'          Me.TxtCodigo.SetFocus
'          Exit Sub
'       End If
'       rec.Close
'
'       'CONSULTO SI EXISTE EL CHEQUE
'        sql = "SELECT * FROM CHEQUE " & _
'              " WHERE CHE_NUMERO = " & XS(TxtCheNumero.Text) & _
'                " AND BAN_CODINT = " & XN(TxtCodInt.Text)
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If rec.RecordCount > 0 Then 'EXITE
'
'            TxtCheNumero.Enabled = False
'            TxtBanco.Enabled = False
'            txtlocalidad.Enabled = False
'            TxtSucursal.Enabled = False
'            TxtCodigo.Enabled = False
'
'            MtrObjetos = Array(TxtCheNumero, TxtBanco, txtlocalidad, TxtSucursal, TxtCodigo)
'            Call CambiarColor(MtrObjetos, 5, &H80000018, "D")
'
'            Me.TxtCheFecEmi.Text = rec!CHE_FECEMI
'            Me.TxtCheFecVto.Text = rec!CHE_FECVTO
'            Me.TxtCheImport.Text = Format(rec!che_import, "$ #0.00")
'
'            'Cargo la Grilla
'            sql = "SELECT CES_FECHA,ECH_DESCRI,CES_DESCRI" & _
'                  " FROM CHEQUE_ESTADOS CE, ESTADO_CHEQUE EC " & _
'                  " WHERE CE.ECH_CODIGO = EC.ECH_CODIGO " & _
'                    " AND CE.CHE_NUMERO = " & XS(TxtCheNumero.Text) & _
'                    " AND CE.BAN_CODINT = " & XN(TxtCodInt.Text) & _
'                    " ORDER BY CES_FECHA"
'            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'            If Rec1.RecordCount > 0 Then
'               Rec1.MoveFirst
'               Do While Not Rec1.EOF
'                 Grd1.AddItem Rec1!CES_FECHA & Chr(9) & Trim(Rec1.Fields(1)) & Chr(9) & Trim(Rec1.Fields(2))
'                 Rec1.MoveNext
'               Loop
'            End If
'            Rec1.Close
'            Me.TxtCesFecha.SetFocus
'        End If
'        rec.Close
'    End If
End Sub

Private Sub Txtlocalidad_KeyPress(KeyAscii As Integer)
KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub TxtLOCALIDAD_LostFocus()
    If txtlocalidad.Text <> "" Then txtlocalidad.Text = Format(txtlocalidad.Text, "000")
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub txtSucursal_LostFocus()
    If TxtSucursal.Text <> "" Then TxtSucursal.Text = Format(TxtSucursal.Text, "000")
End Sub
