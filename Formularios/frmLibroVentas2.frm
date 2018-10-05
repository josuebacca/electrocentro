VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLibroVentas2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libro IVA Ventas"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Rep 
      Left            =   960
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame FrameImpresora 
      Caption         =   "impresora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   11
      Top             =   2040
      Width           =   5610
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   14
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   13
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   375
         Left            =   3810
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   3045
      Picture         =   "frmLibroVentas2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2850
      Width           =   840
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   4755
      Picture         =   "frmLibroVentas2.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2850
      Width           =   840
   End
   Begin VB.Frame Frame2 
      Height          =   2040
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5595
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   210
         TabIndex        =   2
         Top             =   1395
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   397
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   54657025
         CurrentDate     =   43367
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   54657025
         CurrentDate     =   43367
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   405
         TabIndex        =   6
         Top             =   870
         Width           =   960
      End
      Begin VB.Label lblPeriodo1 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2865
         TabIndex        =   5
         Top             =   510
         Width           =   1785
      End
      Begin VB.Label lblPeriodo2 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2865
         TabIndex        =   4
         Top             =   840
         Width           =   1785
      End
      Begin VB.Label lblPor 
         AutoSize        =   -1  'True
         Caption         =   "100 %"
         Height          =   195
         Left            =   5085
         TabIndex        =   3
         Top             =   1425
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmLibroVentas2.frx":0614
      Height          =   735
      Left            =   3900
      Picture         =   "frmLibroVentas2.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2850
      Width           =   840
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2130
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   90
      TabIndex        =   10
      Top             =   3015
      Width           =   750
   End
End
Attribute VB_Name = "frmLibroVentas2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Registro As Long
Dim Tamanio As Long
Dim TotIva As Double

Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub CmdAceptar_Click()
     Registro = 0
     Tamanio = 0
     TotIva = 0
     
     If FechaDesde.Value = Date Then
        MsgBox "Debe ingresar el periodo", vbExclamation, TIT_MSGBOX
        FechaDesde.SetFocus
        Exit Sub
     End If
     If FechaHasta.Value = Date Then
        MsgBox "Debe ingresar el periodo", vbExclamation, TIT_MSGBOX
        FechaHasta.SetFocus
        Exit Sub
     End If
     
     On Error GoTo CLAVO
     Screen.MousePointer = vbHourglass
     DBConn.BeginTrans
     lblEstado.Caption = "Buscando Datos..."
     
        'BORRO LA TABLA TEMPORAL DE IVA VENTAS
        sql = "DELETE FROM TMP_LIBRO_IVA_VENTAS"
        DBConn.Execute sql
        
        'BUSCO FACTURAS LO CAMBIE PARA ELECTROCENTRO
        sql = "SELECT FC.RCL_NUMEROTXT, FC.RCL_SUCURSAL, FC.RCL_FECHA, "
        sql = sql & " FC.RCL_SUBTOTAL, FC.RCL_TOTAL,"
        sql = sql & " FC.EST_CODIGO, C.CLI_CUIT, C.CLI_INGBRU,"
        sql = sql & " C.CLI_RAZSOC, TC.TCO_ABREVIA"
        sql = sql & " FROM REMITO_CLIENTE FC,CLIENTE C,"
        sql = sql & " TIPO_COMPROBANTE TC"
        sql = sql & " WHERE"
        'sql = sql & " FC.RCL_NUMERO=RC.RCL_NUMERO"
        'sql = sql & " AND FC.RCL_SUCURSAL=RC.RCL_SUCURSAL"
        sql = sql & " FC.TCO_CODIGO=TC.TCO_CODIGO"
        sql = sql & " AND FC.CLI_CODIGO = C.CLI_CODIGO "
        sql = sql & " AND FC.EST_CODIGO <> 2" 'ESTADO DEFINITIVO Y ANULADO
        If FechaDesde <> "" Then sql = sql & " AND FC.RCL_FECHA>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND FC.RCL_FECHA<=" & XDQ(FechaHasta)
        sql = sql & " ORDER BY FC.RCL_NUMEROTXT,FC.RCL_FECHA"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_VENTAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "CLIENTE,CUIT,INGBRUTOS,SUBTOTAL,TOTAL)"
                sql = sql & "VALUES ("
                sql = sql & XDQ(rec!RCL_FECHA) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(Format(rec!RCL_SUCURSAL, "0000") & "-" & rec!RCL_NUMEROTXT) & ","
                sql = sql & XS(rec!CLI_RAZSOC) & ","
                sql = sql & XS(Format(rec!CLI_CUIT, "##-########-#")) & ","
                sql = sql & "NULL" & ","
                'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
                If rec!EST_CODIGO = 2 Then
                    sql = sql & "0" & ","
                    'sql = sql & XN(rec!RCL_IVA) & ","
                    'sql = sql & "0" & ","
                    sql = sql & "0" & ")"
                Else
                    sql = sql & XN(rec!RCL_SUBTOTAL) & ","
                    'sql = sql & XN(rec!RCL_IVA) & ","
                    'TotIva = (CDbl(rec!RCL_SUBTOTAL) * CDbl(rec!RCL_IVA)) / 100
                    'sql = sql & XN(CStr(TotIva)) & ","
                    sql = sql & XN(rec!RCL_TOTAL) & ")"
                End If
                DBConn.Execute sql
                rec.MoveNext
                
                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            Loop
        End If
        rec.Close
        
        'BUSCO NOTA DE CREDITO------------------------------------
         sql = "SELECT NC.NCC_NUMEROTXT, NC.NCC_SUCURSAL, NC.NCC_FECHA, NC.NCC_IVA,"
         sql = sql & " NC.NCC_SUBTOTAL, NC.NCC_TOTAL,"
         sql = sql & " NC.EST_CODIGO,C.CLI_CUIT,C.CLI_INGBRU,"
         sql = sql & " C.CLI_RAZSOC,TC.TCO_ABREVIA"
         sql = sql & " FROM NOTA_CREDITO_CLIENTE NC"
         sql = sql & ",TIPO_COMPROBANTE TC , CLIENTE C"
         sql = sql & " WHERE"
         sql = sql & " NC.TCO_CODIGO=TC.TCO_CODIGO"
         sql = sql & " AND NC.CLI_CODIGO=C.CLI_CODIGO"
         If FechaDesde <> "" Then sql = sql & " AND NC.NCC_FECHA>=" & XDQ(FechaDesde)
         If FechaHasta <> "" Then sql = sql & " AND NC.NCC_FECHA<=" & XDQ(FechaHasta)
         sql = sql & " ORDER BY NC.NCC_FECHA"
         
         rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_VENTAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "CLIENTE,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,TOTAL)"
                sql = sql & "VALUES ("
                sql = sql & XDQ(rec!NCC_FECHA) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(Format(rec!NCC_SUCURSAL, "0000") & "-" & rec!NCC_NUMEROTXT) & ","
                sql = sql & XS(rec!CLI_RAZSOC) & ","
                sql = sql & XS(Format(rec!CLI_CUIT, "##-########-#")) & ","
                sql = sql & "NULL" & ","
                'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
                If rec!EST_CODIGO = 2 Then
                    sql = sql & "0" & ","
                    sql = sql & XN(rec!NCC_IVA) & ","
                    sql = sql & "0" & ","
                    sql = sql & "0" & ")"
                Else
                    sql = sql & XN(CStr((-1) * CDbl(IIf(IsNull(rec!NCC_SUBTOTAL), 0, rec!NCC_SUBTOTAL)))) & ","
                    sql = sql & XN(rec!NCC_IVA) & ","
                    TotIva = (CDbl(IIf(IsNull(rec!NCC_SUBTOTAL), 0, rec!NCC_SUBTOTAL)) * CDbl(rec!NCC_IVA)) / 100
                    sql = sql & XN(CStr((-1) * CDbl(TotIva))) & ","
                    sql = sql & XN(CStr((-1) * CDbl(IIf(IsNull(rec!NCC_TOTAL), 0, rec!NCC_TOTAL)))) & ")"
                End If
                DBConn.Execute sql
                rec.MoveNext
                
                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            Loop
        End If
        rec.Close
        
        'BUSCO NOTA DE DEBITO SERVICIOS Y CHEQUES DEVUELTOS-----
        sql = "SELECT ND.NDC_NUMEROTXT, ND.NDC_SUCURSAL, ND.NDC_FECHA, ND.NDC_IVA,"
        sql = sql & " ND.NDC_SUBTOTAL, ND.NDC_TOTAL,"
        sql = sql & " ND.EST_CODIGO, C.CLI_CUIT, C.CLI_INGBRU,"
        sql = sql & " C.CLI_RAZSOC,TC.TCO_ABREVIA"
        sql = sql & " FROM NOTA_DEBITO_CLIENTE ND,"
        sql = sql & " TIPO_COMPROBANTE TC , CLIENTE C"
        sql = sql & " WHERE ND.TCO_CODIGO=TC.TCO_CODIGO"
        sql = sql & " AND ND.CLI_CODIGO=C.CLI_CODIGO"
        If FechaDesde <> "" Then sql = sql & " AND ND.NDC_FECHA>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND ND.NDC_FECHA<=" & XDQ(FechaHasta)
        sql = sql & " ORDER BY ND.NDC_FECHA"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_VENTAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "CLIENTE,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,TOTAL)"
                sql = sql & "VALUES ("
                sql = sql & XDQ(rec!NDC_FECHA) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(Format(rec!NDC_SUCURSAL, "0000") & "-" & rec!NDC_NUMEROTXT) & ","
                sql = sql & XS(rec!CLI_RAZSOC) & ","
                sql = sql & XS(Format(rec!CLI_CUIT, "##-########-#")) & ","
                sql = sql & "NULL" & ","
                'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
                If rec!EST_CODIGO = 2 Then
                    sql = sql & "0" & ","
                    sql = sql & XN(rec!NDC_IVA) & ","
                    sql = sql & "0" & ","
                    sql = sql & "0" & ")"
                Else
                    sql = sql & XN(rec!NDC_SUBTOTAL) & ","
                    sql = sql & XN(rec!NDC_IVA) & ","
                    TotIva = (CDbl(rec!NDC_SUBTOTAL) * CDbl(rec!NDC_IVA)) / 100
                    sql = sql & XN(CStr(TotIva)) & ","
                    sql = sql & XN(rec!NDC_TOTAL) & ")"
                End If
                DBConn.Execute sql
                rec.MoveNext
                
                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            Loop
        End If
        rec.Close
        
    lblEstado.Caption = ""
    DBConn.CommitTrans
    'cargo el reporte
    ListarLibroIVA
        
    Screen.MousePointer = vbNormal
    
    Exit Sub

CLAVO:
 Screen.MousePointer = vbNormal
 lblEstado.Caption = ""
 DBConn.RollbackTrans
 If rec.State = 1 Then rec.Close
 MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub ListarLibroIVA()
    lblEstado.Caption = "Buscando Listado..."
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SIELECTROCENTRO"
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
        
    sql = "SELECT CUIT,ING_BRUTOS,RAZ_SOCIAL FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Rep.Formulas(0) = "EMPRESA='     Empresa:  " & Trim(rec!RAZ_SOCIAL) & "'"
        Rep.Formulas(1) = "CUIT='       C.U.I.T.:  " & Format(rec!cuit, "##-########-#") & "'"
        Rep.Formulas(2) = "INGBRUTOS='Ing. Brutos:  " & Format(rec!ING_BRUTOS, "###-#####-##") & "'"
    End If
    rec.Close
    
    Rep.WindowTitle = "Libro I.V.A. Ventas"
    Rep.ReportFileName = DRIVE & DirReport & "rptlibroivaventas.rpt"
    
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
     Rep.Action = 1
     
     lblEstado.Caption = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
     Rep.Formulas(2) = ""
End Sub

Private Sub CmdNuevo_Click()
    FechaDesde.Value = Date
    lblPeriodo1.Caption = ""
    FechaHasta.Value = Date
    lblPeriodo2.Caption = ""
    FechaDesde.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmLibroVentas2 = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    lblEstado.Caption = ""
    lblPor.Caption = "100 %"
    Call Centrar_pantalla(Me)
    Set rec = New ADODB.Recordset
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub FechaDesde_LostFocus()
    If Trim(FechaDesde.Value) <> "" Then
        FechaHasta.Value = FechaDesde.Value
        lblPeriodo1.Caption = UCase(Format(FechaDesde.Value, "mmmm/yyyy"))
    Else
        lblPeriodo1.Caption = ""
    End If
End Sub

Private Sub FechaHasta_LostFocus()
    If Trim(FechaHasta.Value) <> "" Then
        lblPeriodo2.Caption = UCase(Format(FechaHasta.Value, "mmmm/yyyy"))
    Else
        lblPeriodo2.Caption = ""
    End If
End Sub

