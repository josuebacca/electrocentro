VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form ConsultaBoleta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta"
   ClientHeight    =   5685
   ClientLeft      =   150
   ClientTop       =   1065
   ClientWidth     =   8820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5685
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmbSalirConsContra 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7170
      TabIndex        =   12
      Top             =   5175
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   150
      TabIndex        =   9
      Top             =   90
      Width           =   8490
      Begin FechaCtl.Fecha FechaBusq 
         Height          =   330
         Left            =   3060
         TabIndex        =   3
         Top             =   750
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el m�nimo permitido"
      End
      Begin VB.ComboBox cboOrden 
         Height          =   315
         ItemData        =   "ConsultaBol.frx":0000
         Left            =   5715
         List            =   "ConsultaBol.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   1845
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   6450
         TabIndex        =   4
         Top             =   765
         Width           =   1110
      End
      Begin VB.ComboBox cboBusqueda 
         Height          =   315
         ItemData        =   "ConsultaBol.frx":001D
         Left            =   1665
         List            =   "ConsultaBol.frx":0027
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   1845
      End
      Begin VB.TextBox txtCond_Busqueda 
         Height          =   285
         Left            =   3060
         TabIndex        =   2
         Top             =   750
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Ordenado por:"
         Height          =   255
         Left            =   4545
         TabIndex        =   13
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "B�squeda por:"
         Height          =   255
         Left            =   495
         TabIndex        =   11
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label lblDescripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese la condici�n de B�squeda:"
         Height          =   225
         Left            =   510
         TabIndex        =   10
         Top             =   780
         Width           =   2460
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5595
      TabIndex        =   6
      Top             =   5175
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid GrdCheques 
      Height          =   3645
      Left            =   180
      TabIndex        =   5
      Top             =   1440
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   6429
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorSel    =   8388736
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin VB.TextBox txtDes_Cons 
      Height          =   285
      Left            =   195
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4680
      Width           =   8430
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Seleccionado:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   330
      TabIndex        =   8
      Top             =   4410
      Width           =   1935
   End
End
Attribute VB_Name = "ConsultaBoleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CerrarSnapConsulta As Boolean
Dim snpConsulta As Recordset
Dim ventana As Form
Dim CrtlCodigo As CONTROL

Public Function Parametros(auxVentana As Form, auxCrtlCodigo As CONTROL, auxCrtlCodigo1 As CONTROL, auxCrtlCodigo2 As CONTROL, auxCrtlCodigo3 As CONTROL)
    Set ventana = auxVentana 'Objeto ventana que llama a la ayuda
    Set CrtlCodigo = auxCrtlCodigo   'Objeto Control del form ventana al que se asigna el codigo
End Function

Private Sub cboBusqueda_Change()
    txtCond_Busqueda.Text = ""
End Sub

Private Sub cboBusqueda_Click()
    txtCond_Busqueda.Text = ""
End Sub

Private Sub cboBusqueda_LostFocus()
   If Me.cboBusqueda.List(Me.cboBusqueda.ListIndex) = "N�mero" Then
        Me.txtCond_Busqueda.Visible = True
        Me.FechaBusq.Visible = False
        Me.txtCond_Busqueda.SetFocus
   ElseIf Me.cboBusqueda.List(Me.cboBusqueda.ListIndex) = "Fecha" Then
        Me.txtCond_Busqueda.Visible = False
        Me.FechaBusq.Visible = True
        Me.FechaBusq.SetFocus
   End If
End Sub

Private Sub cmbSalirConsContra_Click()
    CrtlCodigo = ""
    Unload Me
End Sub

Private Sub CmdAceptar_Click()
    Call ValidarIngreso
End Sub

Private Sub cmdBuscar_Click()
    Dim Rec As ADODB.Recordset
    Set Rec = New ADODB.Recordset
    
    Screen.MousePointer = 11
    
    Select Case cboBusqueda.Text
       Case "N�mero"
              'TODOS las Boletas de Dep�sito
              sql = "SELECT D.BOL_NUMERO,D.BOL_FECHA,B.BAN_descri,D.CTA_NROCTA"
              sql = sql & " FROM BOL_DEPOSITO D, BANCO B"
              sql = sql & " WHERE D.BAN_CODINT = B.BAN_CODINT"
              If Trim(txtCond_Busqueda.Text) <> "" Then
                 sql = sql & " AND D.BOL_NUMERO = " & XS(txtCond_Busqueda)
              End If
       Case "Fecha"
              sql = "SELECT D.BOL_NUMERO,D.BOL_FECHA,B.BAN_descri,D.CTA_NROCTA"
              sql = sql & " FROM BOL_DEPOSITO D, BANCO B"
              sql = sql & " WHERE D.BAN_CODINT = B.BAN_CODINT"
              If Trim(FechaBusq.Text) <> "" Then
                sql = sql & " AND D.BOL_FECHA = " & XDQ(Me.FechaBusq.Text)
              End If
    End Select
    
    Select Case cboOrden.Text
       Case "N�mero"
            sql = sql & " ORDER BY D.BOL_NUMERO"
       Case "Fecha de Vto"
            sql = sql & " ORDER BY D.BOL_FECHA "
    End Select
    
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.RecordCount > 0 Then
        GrdCheques.Rows = 1
        Rec.MoveFirst
        Do While Not Rec.EOF()
        
           GrdCheques.AddItem Trim(Rec.Fields(0)) & Chr(9) & _
                              Rec.Fields(1) & Chr(9) & _
                              Trim(Rec.Fields(2)) & Chr(9) & _
                              Trim(Rec.Fields(3))
            Rec.MoveNext
        Loop
        If Me.GrdCheques.Enabled Then
           Me.GrdCheques.row = 1
           Me.GrdCheques.SetFocus
        End If
    Else
        Me.GrdCheques.Rows = 1
        MsgBox "No se encontraron registros.", 16, TIT_MSGBOX
        txtDes_Cons.Text = ""
    End If
    Rec.Close
    Screen.MousePointer = 1
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    CerrarSnapConsulta = False
    Call Centrar_pantalla(Me)
    GrdCheques.FormatString = ">N�mero|^Fecha|<Banco|<Cuenta"
    GrdCheques.ColWidth(0) = 1500
    GrdCheques.ColWidth(1) = 1500
    GrdCheques.ColWidth(2) = 3900
    GrdCheques.ColWidth(3) = 1200
    GrdCheques.Rows = 1
    
    Me.cboBusqueda.Text = "N�mero"
    Me.cboOrden.Text = "N�mero"
    Me.Caption = "Consulta de Boletas de Dep�sito"
    Me.lblDescripcion.Caption = "Ingrese Condici�n de Busqueda"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ConsultaBoleta = Nothing
End Sub

Private Sub GrdCheques_Click()
    txtDes_Cons.Text = Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 0))) & _
               " - " & Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 1))) & _
               " - " & Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 2))) & _
               " - " & Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 3)))
End Sub

Private Sub GrdCheques_DblClick()
    CmdAceptar_Click
End Sub

Private Sub txtCond_Busqueda_KeyPress(KeyAscii As Integer)
    If cboBusqueda.Text = "N�mero" Then
        KeyAscii = NumeroEntero(KeyAscii)
    End If
End Sub

Private Function ValidarIngreso()
    If Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 0))) <> "" Then
        CrtlCodigo = Trim(GrdCheques.TextArray(GRIDINDEX(GrdCheques, GrdCheques.row, 0)))
    End If
    Unload Me
End Function

Private Sub txtCond_Busqueda_LostFocus()
   If Me.cboBusqueda.Text = "Fecha de Vto" Then
       txtCond_Busqueda.Text = ValidarIngresoFecha(txtCond_Busqueda)
   End If
End Sub

Private Sub txtdes_cons_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

