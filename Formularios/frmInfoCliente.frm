VERSION 5.00
Begin VB.Form frmInfoCliente 
   Caption         =   "Informacion Cliente"
   ClientHeight    =   1635
   ClientLeft      =   9690
   ClientTop       =   3585
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   3540
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txtGastado 
         Height          =   315
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox txtAutorizado 
         Height          =   315
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Gastado:"
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
         Left            =   795
         TabIndex        =   2
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Autorizado:"
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
         Left            =   600
         TabIndex        =   1
         Top             =   300
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmInfoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'Dim REC As ADODB.Recordset
'Set REC = New ADODB.Recordset
'AUTORIZADO
'sql = "SELECT CLI_CREDITO FROM CLIENTE "
'sql = sql & "WHERE CLI_CODIGO = " & XN(frmRemitoCliente.TxtCodigoCli.Text)
'REC.Open sql, DBConn, adOpenStatic, adLockOptimistic
'If REC.EOF = False Then
'    txtAutorizado.Text = REC!CLI_CREDITO
'End If
'REC.Close

'sql = "SELECT CLI_CODIGO,SUM(RCL_CLIENTE) AS FROM REMITO_CLIENTE "
'sql = sql & "WHERE CLI_CODIGO = " & XN(TxtCodigoCli.Text) & " "
'sql = sql & "GROUP BY CLI_CODIGO"
'REC.Open sql, DBConn, adOpenStatic, adLockOptimistic
'If REC.EOF = False Then
    
'End If

End Sub
