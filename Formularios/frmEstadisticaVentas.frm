VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEstadisticaVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estadisticas de Ventas"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTotPorGan 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   16
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtTotCtaCte 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtTotCont 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmEstadisticaVentas.frx":0000
      Height          =   750
      Left            =   7830
      Picture         =   "frmEstadisticaVentas.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   870
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   8715
      Picture         =   "frmEstadisticaVentas.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   840
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   6960
      Picture         =   "frmEstadisticaVentas.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtTotDif 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtTotTot 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Frame fraPeriodo 
      Caption         =   "Seleccionar Periodo"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   9615
      Begin VB.CommandButton CmdBuscAprox 
         Height          =   330
         Left            =   8640
         MaskColor       =   &H000000FF&
         Picture         =   "frmEstadisticaVentas.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Buscar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   675
      End
      Begin VB.ComboBox cboAño 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
         Width           =   1215
      End
      Begin VB.OptionButton optPeriodo 
         Caption         =   "Periodo"
         Height          =   255
         Left            =   3360
         TabIndex        =   2
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton optAnio 
         Caption         =   "Año"
         Height          =   375
         Left            =   600
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   375
         Left            =   5040
         TabIndex        =   20
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   58982401
         CurrentDate     =   43367
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   375
         Left            =   7080
         TabIndex        =   21
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   58982401
         CurrentDate     =   43367
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   6480
         TabIndex        =   14
         Top             =   330
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   4440
         TabIndex        =   13
         Top             =   330
         Width           =   465
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   3435
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   6059
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorSel    =   8388736
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Rep 
      Height          =   480
      Left            =   480
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   19
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Label lblPromedio 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Promedio mensual: "
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
      Left            =   2160
      TabIndex        =   17
      Top             =   5040
      Width           =   1665
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Totales"
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
      Left            =   1440
      TabIndex        =   15
      Top             =   4380
      Width           =   810
   End
End
Attribute VB_Name = "frmEstadisticaVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboAño_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub cboAño_LostFocus()
    If cboAño.Text <> "" Then
        cboAño.Text = Mid(cboAño.Text, 1, 4)
        CmdBuscAprox_Click
    End If
End Sub

Private Sub CmdBuscAprox_Click()
    Dim I As Integer
    Dim MesAño As String
    Dim Contado As Double
    Dim CtaCTe As Double
    Dim TOTAL As Double
    Dim TOTCont As Double
    Dim TOTCtaCte As Double
    Dim TOTTot As Double
    Dim TOTTotDif As Double
    Dim TOTPorGan As Double
    Dim PORGAN As Double
    Dim ProMen As Double
    Dim j As Integer
    Set Rec1 = New ADODB.Recordset
    
    ' CONTADO
    GrdModulos.Rows = 1
    For I = 1 To 12
        MesAño = MES(I)
        GrdModulos.AddItem MesAño & " " & cboAño.Text & Chr(9) & "0,00" _
                           & Chr(9) & "0,00" & Chr(9) & "0,00" & Chr(9) & "0,00"
    Next I
    'CONTADO
    sql = "SELECT MONTH(RCL_FECHA)AS MES,SUM(RCL_TOTAL)AS TOTAL FROM REMITO_CLIENTE "
    sql = sql & " WHERE FPG_CODIGO = 1"
    
    If optAnio.Value = True Then
        sql = sql & " AND RCL_FECHA >= " & XDQ("01/01/" & cboAño.Text & "")
        sql = sql & " AND RCL_FECHA <= " & XDQ("12/31/" & cboAño.Text & "")
    End If
    sql = sql & " GROUP BY MONTH(RCL_FECHA)"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    'GrdModulos.Rows = 1
    If rec.EOF = False Then
        Do While rec.EOF = False
            'MesAño = mes(rec!mes)
            'GrdModulos.AddItem MesAño & " " & cboAño.Text & Chr(9) & rec!TOTAL
            'contado
             GrdModulos.TextMatrix(rec!MES, 1) = Valido_Importe(rec!TOTAL)
            
            rec.MoveNext
        Loop
    End If
    rec.Close
    
    'CTA CTE
    
    sql = "SELECT MONTH(RCL_FECHA)AS MES,SUM(RCL_TOTAL)AS TOTAL FROM REMITO_CLIENTE "
    sql = sql & " WHERE FPG_CODIGO = 2"
    If optAnio.Value = True Then
        sql = sql & " AND RCL_FECHA >= " & XDQ("01/01/" & cboAño.Text & "")
        sql = sql & " AND RCL_FECHA <= " & XDQ("12/31/" & cboAño.Text & "")
    End If
    sql = sql & " GROUP BY MONTH(RCL_FECHA)"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    'GrdModulos.Rows = 1
    If rec.EOF = False Then
        Do While rec.EOF = False
            'CTA CTE
             GrdModulos.TextMatrix(rec!MES, 2) = Valido_Importe(rec!TOTAL)
            
            rec.MoveNext
        Loop
    End If
    rec.Close
    ' busco el porcentaje de ganancia del empleado en tabla parametros
    Rec1.Open "SELECT PORGAN FROM PARAMETROS", DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
       PORGAN = Rec1!PORGAN
    End If
    Rec1.Close
    For I = 1 To 12
        Contado = Valido_Importe(GrdModulos.TextMatrix(I, 1))
        CtaCTe = Valido_Importe(GrdModulos.TextMatrix(I, 2))
        TOTAL = Contado + CtaCTe
        GrdModulos.TextMatrix(I, 3) = Format(TOTAL, "0.00")
        If GrdModulos.TextMatrix(I, 3) = 0 Then
            GrdModulos.TextMatrix(I, 3) = "0,00"
        End If
        'calculo porcentaje ganancia empleado
        TOTPorGan = Valido_Importe(GrdModulos.TextMatrix(I, 3) * PORGAN) / 100
        GrdModulos.TextMatrix(I, 5) = Format(TOTPorGan, "0.00")
        If GrdModulos.TextMatrix(I, 5) = 0 Then
            GrdModulos.TextMatrix(I, 5) = "0,00"
        End If
    Next I
    
    ' Calculo Diferencia Año Interior
    'Todo - Contado y Cta cte
    sql = "SELECT MONTH(RCL_FECHA)AS MES,SUM(RCL_TOTAL)AS TOTAL FROM REMITO_CLIENTE "
    sql = sql & " WHERE (FPG_CODIGO = 1 OR FPG_CODIGO = 2) "
    
    If optAnio.Value = True Then
        sql = sql & " AND RCL_FECHA >= " & XDQ("01/01/" & XN(cboAño.Text) - 1 & "")
        sql = sql & " AND RCL_FECHA <= " & XDQ("12/31/" & XN(cboAño.Text) - 1 & "")
    End If
    sql = sql & " GROUP BY MONTH(RCL_FECHA)"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    'GrdModulos.Rows = 1
    If rec.EOF = False Then
        Do While rec.EOF = False
            'MesAño = mes(rec!mes)
            'GrdModulos.AddItem MesAño & " " & cboAño.Text & Chr(9) & rec!TOTAL
            'contado
             GrdModulos.TextMatrix(rec!MES, 4) = Valido_Importe(GrdModulos.TextMatrix(rec!MES, 3) - rec!TOTAL)
            
            rec.MoveNext
        Loop
    End If
    rec.Close
    
    
    
    'Calculo totales
    For I = 1 To 12
        TOTCont = TOTCont + Valido_Importe(GrdModulos.TextMatrix(I, 1))
        TOTCtaCte = TOTCtaCte + Valido_Importe(GrdModulos.TextMatrix(I, 2))
        TOTTot = TOTTot + Valido_Importe(GrdModulos.TextMatrix(I, 3))
        TOTTotDif = TOTTotDif + Valido_Importe(GrdModulos.TextMatrix(I, 4))
        TOTPorGan = TOTPorGan + Valido_Importe(GrdModulos.TextMatrix(I, 5))
    Next I
    txtTotCont.Text = Format(TOTCont, "0.00")
    txtTotCtaCte.Text = Format(TOTCtaCte, "0.00")
    txtTotTot.Text = Format(TOTTot, "0.00")
    txtTotDif.Text = Format(TOTTotDif, "0.00")
    txtTotPorGan = Format(TOTPorGan, "0.00")
    'contar meses con total
    j = 0
    For I = 1 To 12
        If GrdModulos.TextMatrix(I, 3) <> "0,00" Then
            j = j + 1
        End If
    Next I
    If j <> 0 Then
        ProMen = TOTTot / j
    Else
        ProMen = 0
    End If
    lblPromedio.Caption = Format(ProMen, "0.00")
End Sub
Private Function MES(nromes As Integer) As String
    Select Case nromes
        Case 1: MES = "Enero"
        Case 2: MES = "Febrero"
        Case 3: MES = "Marzo"
        Case 4: MES = "Abril"
        Case 5: MES = "Mayo"
        Case 6: MES = "Junio"
        Case 7: MES = "Julio"
        Case 8: MES = "Agosto"
        Case 9: MES = "Septiembre"
        Case 10: MES = "Octubre"
        Case 11: MES = "Noviembre"
        Case 12: MES = "Diciembre"
    End Select
End Function

Private Sub cmdListar_Click()
    Dim I As Integer
    If GrdModulos.Rows > 2 Then
            
        DBConn.Execute "DELETE FROM TMP_VENTAS"
        For I = 1 To 12
            sql = "INSERT INTO TMP_VENTAS (TMP_MESANIO,TMP_CONTADO,TMP_CTACTE,TMP_TOTAL,TMP_DIFAANT) VALUES ("
            sql = sql & XS(Format(I, "00") & "- " & GrdModulos.TextMatrix(I, 0)) & ","
            sql = sql & XN(GrdModulos.TextMatrix(I, 1)) & ","
            sql = sql & XN(GrdModulos.TextMatrix(I, 2)) & ","
            sql = sql & XN(GrdModulos.TextMatrix(I, 3)) & ","
            sql = sql & XN(GrdModulos.TextMatrix(I, 4)) & ")"
            DBConn.Execute sql
        Next I
        ListarVentas
    End If
    
End Sub
Private Sub ListarVentas()
    'lblestado.Caption = "Buscando Listado..."
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SIELECTROCENTRO"
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
    Rep.WindowTitle = "Estadisticas de Ventas"
    Rep.ReportFileName = DRIVE & DirReport & "rptgraficoventas.rpt"
    Rep.Destination = crptToWindow
    Rep.Action = 1
    'lblestado.Caption = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
End Sub

Private Sub CmdNuevo_Click()
    GrdModulos.Rows = 1
    txtTotCont.Text = "0,00"
    txtTotCtaCte.Text = "0,00"
    txtTotTot.Text = "0,00"
    txtTotDif.Text = "0,00"
    txtTotPorGan.Text = "0,00"
    
    cboAño.Text = Year(Date)
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    FechaDesde.Value = Date
    FechaHasta.Value = Date
    optAnio.Value = True
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub FechaDesde_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)
    preparogrilla
    cargocboanio
    txtTotCont.Text = "0,00"
    txtTotCtaCte.Text = "0,00"
    txtTotTot.Text = "0,00"
    txtTotDif.Text = "0,00"
    txtTotPorGan.Text = "0,00"
    cboAño.Text = Year(Date)
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    FechaDesde.Value = Date
    FechaHasta.Value = Date
    CmdBuscAprox_Click
End Sub
Private Function preparogrilla()
    GrdModulos.FormatString = "Periodo|Contado|Cta. Cte.|Total|Dif Año Anterior|% Empleado"
    GrdModulos.ColWidth(0) = 2300 ' Periodo
    GrdModulos.ColWidth(1) = 1450 ' Contado
    GrdModulos.ColWidth(2) = 1450 ' Cta Cte
    GrdModulos.ColWidth(3) = 1450 ' Total
    GrdModulos.ColWidth(4) = 1450 ' Dif Año Anterior
    GrdModulos.ColWidth(5) = 1400 ' % Empleado
    GrdModulos.Rows = 1
    
End Function
Private Function cargocboanio()
    For I = 1980 To 2099
        cboAño.AddItem I
    Next I
End Function

Private Sub optAnio_Click()
    If optPeriodo.Value = True Then
        cboAño.Enabled = True
    Else
        cboAño.Enabled = False
    End If
End Sub

Private Sub optPeriodo_Click()
    If optPeriodo.Value = True Then
        FechaDesde.Enabled = True
        FechaHasta.Enabled = True
    Else
        FechaDesde.Enabled = False
        FechaHasta.Enabled = False
    End If
End Sub
