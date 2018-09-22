VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmListadoStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Stock - Productos Faltantes!!!"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "frmListadoStock.frx":0000
      Height          =   750
      Left            =   9990
      Picture         =   "frmListadoStock.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   870
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      DisabledPicture =   "frmListadoStock.frx":0614
      Height          =   750
      Left            =   9120
      Picture         =   "frmListadoStock.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   20
      TabIndex        =   4
      Top             =   0
      Width           =   10935
      Begin VB.TextBox TxtDescriB 
         Height          =   330
         Left            =   1275
         MaxLength       =   150
         TabIndex        =   0
         Top             =   225
         Width           =   8775
      End
      Begin VB.CommandButton CmdBuscAprox 
         Height          =   360
         Left            =   10260
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoStock.frx":30C0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Buscar"
         Top             =   210
         UseMaskColor    =   -1  'True
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   270
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   1665
         TabIndex        =   5
         Top             =   315
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   4815
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorSel    =   8388736
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin VB.Label lblestado 
      AutoSize        =   -1  'True
      Caption         =   "estado"
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
      Left            =   240
      TabIndex        =   8
      Top             =   6000
      Width           =   585
   End
End
Attribute VB_Name = "frmListadoStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBuscAprox_Click()
    sql = " SELECT P.PTO_DESCRI,L.LNA_DESCRI, "
    sql = sql & " R.RUB_DESCRI,RE.TPRE_DESCRI,P.PTO_PRECIO,P.PTO_PRECIOC,P.PTO_CODIGO,P.PTO_PRECIVA, "
    sql = sql & " P.PTO_STKMIN,DS.DST_STKFIS "
    sql = sql & " FROM PRODUCTO P,LINEAS L,RUBROS R,TIPO_PRESENTACION RE,LISTA_PRECIO LP,STOCK ST,DETALLE_STOCK DS"
    sql = sql & " WHERE P.LNA_CODIGO = L.LNA_CODIGO  AND "
    sql = sql & " LP.LIS_CODIGO = P.LIS_CODIGO"
    sql = sql & " AND P.LIS_CODIGO <> " & 0 & " "
    sql = sql & " AND P.RUB_CODIGO = R.RUB_CODIGO AND P.TPRE_CODIGO = RE.TPRE_CODIGO "
    sql = sql & " AND ST.STK_CODIGO = DS.STK_CODIGO AND P.PTO_CODIGO = DS.PTO_CODIGO"
    sql = sql & " AND DS.DST_STKFIS < P.PTO_STKMIN"
    If TxtDescriB.Text <> "" Then
        If IsNumeric(TxtDescriB.Text) Then
            sql = sql & " AND P.PTO_CODIGO = " & XN(Trim(TxtDescriB)) & ""
        Else
            sql = sql & " AND P.PTO_DESCRI LIKE '%" & Trim(TxtDescriB.Text) & "%'"
        End If
        
    End If
    
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    GrdModulos.Rows = 1
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!PTO_CODIGO & Chr(9) & rec!PTO_DESCRI & Chr(9) & _
                               rec!LNA_DESCRI & Chr(9) & rec!RUB_DESCRI & Chr(9) & _
                               rec!TPRE_DESCRI & Chr(9) & rec!PTO_STKMIN & Chr(9) & _
                               rec!DST_STKFIS
            rec.MoveNext
        Loop
    Else
        MsgBox "No se encontraron productos con faltantes", vbInformation, TIT_MSGBOX
    End If
    rec.Close
    lblestado.Caption = "Se encontraron " & GrdModulos.Rows - 1 & " productos debajo del stock mininmo"
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    Centrar_pantalla Me
    preparogrilla
    lblestado = ""
    CmdBuscAprox_Click
'    sql = "UPDATE DETALLE_STOCK SET "
'    sql = sql & " DST_STKFIS = 120"
'    sql = sql & " WHERE PTO_CODIGO < 4000"
'    DBConn.Execute sql
'
'    sql = "UPDATE PRODUCTO SET "
'    sql = sql & " PTO_STKMIN = 100"
'    'sql = sql & " WHERE PTO_CODIGO LIKE '" & "11" & "'%"
'    DBConn.Execute sql
    
    
    
End Sub
Private Function preparogrilla()
    GrdModulos.FormatString = "Código|Producto|Linea|Rubro|Marca|Stock Minimo|Stock Actual"
    GrdModulos.ColWidth(0) = 900 ' codigo
    GrdModulos.ColWidth(1) = 2500 ' Producto
    GrdModulos.ColWidth(2) = 1500 ' linea
    GrdModulos.ColWidth(3) = 1500 ' Rubro
    GrdModulos.ColWidth(4) = 1500 ' Marca
    GrdModulos.ColWidth(5) = 1300 ' Stock minimo
    GrdModulos.ColWidth(6) = 1200 ' Stock actual
    GrdModulos.Rows = 1
End Function

