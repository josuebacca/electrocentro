VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmImputarNCaFacturaProveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imputar Nota de Cr�dito Proveedores a Facturas..."
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11715
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   10725
      TabIndex        =   13
      Top             =   5880
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   8955
      TabIndex        =   11
      Top             =   5880
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   9840
      TabIndex        =   12
      Top             =   5880
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   5820
      Left            =   60
      TabIndex        =   25
      Top             =   15
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   10266
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   512
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
      TabPicture(0)   =   "frmImputarNCaFacturaProveedores.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameNotaCredito"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameProveedor"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmImputarNCaFacturaProveedores.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameBuscar"
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).ControlCount=   2
      Begin VB.Frame FrameProveedor 
         Caption         =   "Proveedor..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   105
         TabIndex        =   45
         Top             =   435
         Width           =   6735
         Begin VB.CommandButton cmdBuscarProveedor1 
            Height          =   300
            Left            =   1995
            MaskColor       =   &H000000FF&
            Picture         =   "frmImputarNCaFacturaProveedores.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Buscar Proveedor"
            Top             =   765
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.ComboBox cboTipoProveedor 
            Height          =   315
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   405
            Width           =   3375
         End
         Begin VB.TextBox txtCodProveedor 
            Height          =   300
            Left            =   990
            MaxLength       =   40
            TabIndex        =   1
            Top             =   765
            Width           =   975
         End
         Begin VB.TextBox txtProvRazSoc 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2430
            MaxLength       =   50
            TabIndex        =   2
            Tag             =   "Descripci�n"
            Top             =   765
            Width           =   4170
         End
         Begin VB.TextBox txtProvLoc 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
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
            Left            =   990
            TabIndex        =   47
            Top             =   1110
            Width           =   5610
         End
         Begin VB.TextBox txtProvDom 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
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
            Left            =   990
            MaxLength       =   50
            TabIndex        =   46
            Top             =   1425
            Width           =   5610
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Prov.:"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   435
            Width           =   780
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
            Left            =   360
            TabIndex        =   50
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Loc.:"
            Height          =   180
            Left            =   540
            TabIndex        =   49
            Top             =   1155
            Width           =   360
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Dom.:"
            Height          =   195
            Left            =   480
            TabIndex        =   48
            Top             =   1455
            Width           =   420
         End
      End
      Begin VB.Frame frameBuscar 
         Caption         =   "Buscar por..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1890
         Left            =   -74610
         TabIndex        =   30
         Top             =   525
         Width           =   11025
         Begin VB.CommandButton cmdBuscarProveedor 
            Height          =   300
            Left            =   4080
            MaskColor       =   &H000000FF&
            Picture         =   "frmImputarNCaFacturaProveedores.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Buscar Proveedor"
            Top             =   765
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtProveedor 
            Height          =   300
            Left            =   3060
            MaxLength       =   40
            TabIndex        =   19
            Top             =   765
            Width           =   975
         End
         Begin VB.TextBox txtDesProv 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4515
            MaxLength       =   50
            TabIndex        =   52
            Tag             =   "Descripci�n"
            Top             =   765
            Width           =   4995
         End
         Begin VB.CheckBox chkProveedor 
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   300
            TabIndex        =   15
            Top             =   765
            Width           =   1125
         End
         Begin VB.CheckBox chkTipoProveedor 
            Caption         =   "Tipo Prov"
            Height          =   195
            Left            =   300
            TabIndex        =   14
            Top             =   495
            Width           =   1050
         End
         Begin VB.ComboBox cboBuscaTipoProveedor 
            Height          =   315
            Left            =   3060
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   420
            Width           =   3750
         End
         Begin VB.ComboBox cboNotaCredito1 
            Height          =   315
            Left            =   3060
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1440
            Width           =   2400
         End
         Begin VB.CheckBox chkTipoFactura 
            Caption         =   "Tipo"
            Height          =   195
            Left            =   300
            TabIndex        =   17
            Top             =   1290
            Width           =   720
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1425
            Left            =   9885
            MaskColor       =   &H000000FF&
            Picture         =   "frmImputarNCaFacturaProveedores.frx":064C
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Buscar "
            Top             =   315
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin VB.PictureBox FechaHasta 
            Height          =   285
            Left            =   5565
            ScaleHeight     =   225
            ScaleWidth      =   1125
            TabIndex        =   21
            Top             =   1110
            Width           =   1185
         End
         Begin VB.PictureBox FechaDesde 
            Height          =   330
            Left            =   3060
            ScaleHeight     =   270
            ScaleWidth      =   1110
            TabIndex        =   20
            Top             =   1110
            Width           =   1170
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   300
            TabIndex        =   16
            Top             =   1020
            Width           =   810
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   2190
            TabIndex        =   55
            Top             =   795
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Prov.:"
            Height          =   195
            Left            =   2190
            TabIndex        =   54
            Top             =   465
            Width           =   780
         End
         Begin VB.Label lbltipoFac 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   2610
            TabIndex        =   41
            Top             =   1485
            Width           =   360
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4515
            TabIndex        =   32
            Top             =   1155
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   1965
            TabIndex        =   31
            Top             =   1140
            Width           =   1005
         End
      End
      Begin VB.Frame FrameNotaCredito 
         Caption         =   "Nota de Cr�dito..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   6840
         TabIndex        =   27
         Top             =   435
         Width           =   4635
         Begin MSComCtl2.DTPicker FechaNotaCredito 
            Height          =   375
            Left            =   1320
            TabIndex        =   57
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   53936129
            CurrentDate     =   43367
         End
         Begin VB.TextBox txtNroNotaCredito 
            Height          =   300
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   5
            Top             =   840
            Width           =   960
         End
         Begin VB.TextBox txtNroSucursal 
            Height          =   300
            Left            =   1260
            MaxLength       =   4
            TabIndex        =   4
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton cmdBuscarNotaCRedito 
            Caption         =   "&Buscar"
            Height          =   315
            Left            =   3750
            TabIndex        =   6
            ToolTipText     =   "Buscar Factura"
            Top             =   1500
            UseMaskColor    =   -1  'True
            Width           =   750
         End
         Begin VB.ComboBox cboNotaCredito 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   465
            Width           =   2400
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   840
            TabIndex        =   36
            Top             =   480
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   705
            TabIndex        =   35
            Top             =   1230
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "N�mero:"
            Height          =   195
            Left            =   600
            TabIndex        =   34
            Top             =   855
            Width           =   600
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   3105
         Left            =   -74625
         TabIndex        =   24
         Top             =   2490
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   5477
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Frame Frame3 
         Height          =   3330
         Left            =   105
         TabIndex        =   28
         Top             =   2295
         Width           =   11400
         Begin VB.CommandButton cmdQuitar 
            Height          =   555
            Left            =   5445
            Picture         =   "frmImputarNCaFacturaProveedores.frx":2DEE
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Quitar Factura de la Imputaci�n"
            Top             =   1455
            Width           =   540
         End
         Begin VB.CommandButton CmdAgregar 
            Height          =   555
            Left            =   5445
            Picture         =   "frmImputarNCaFacturaProveedores.frx":3230
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Agregar Factura a la Imputaci�n"
            Top             =   885
            Width           =   540
         End
         Begin VB.TextBox txtSaldoNC 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9750
            TabIndex        =   38
            Top             =   2895
            Width           =   1350
         End
         Begin VB.TextBox txtTotalNC 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9750
            TabIndex        =   37
            Top             =   2550
            Width           =   1350
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   6495
            TabIndex        =   29
            Top             =   990
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   1965
            Left            =   6000
            TabIndex        =   10
            Top             =   495
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   3466
            _Version        =   393216
            Rows            =   3
            Cols            =   9
            FixedCols       =   0
            BackColorSel    =   12648447
            ForeColorSel    =   0
            BackColorBkg    =   -2147483633
            GridColor       =   -2147483633
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            ScrollBars      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid grillaFactura 
            Height          =   1965
            Left            =   60
            TabIndex        =   7
            Top             =   495
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   3466
            _Version        =   393216
            Rows            =   3
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   260
            BackColor       =   16777215
            BackColorSel    =   8388736
            BackColorBkg    =   -2147483633
            FocusRect       =   0
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Facturas asignadas a la Nota de Cr�dito"
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
            Left            =   6060
            TabIndex        =   44
            Top             =   210
            Width           =   4215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Facturas con Saldo"
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
            Left            =   120
            TabIndex        =   43
            Top             =   210
            Width           =   2025
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nota de Cr�dito"
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
            Left            =   7335
            TabIndex        =   42
            Top             =   2535
            Width           =   1650
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   9270
            TabIndex        =   40
            Top             =   2580
            Width           =   405
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Saldo:"
            Height          =   195
            Left            =   9225
            TabIndex        =   39
            Top             =   2940
            Width           =   450
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   26
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
      Left            =   225
      TabIndex        =   33
      Top             =   5940
      Width           =   750
   End
End
Attribute VB_Name = "frmImputarNCaFacturaProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim w As Integer
Dim VBonificacion As Double
Dim VTotal As Double

Private Sub chkProveedor_Click()
    If chkProveedor.Value = Checked Then
        txtProveedor.Enabled = True
        cmdBuscarProveedor.Enabled = True
    Else
        txtProveedor.Text = ""
        txtProveedor.Enabled = False
        cmdBuscarProveedor.Enabled = False
    End If
End Sub

Private Sub chkTipoProveedor_Click()
    If chkTipoProveedor.Value = Checked Then
        cboBuscaTipoProveedor.Enabled = True
        cboBuscaTipoProveedor.ListIndex = 0
    Else
        cboBuscaTipoProveedor.Enabled = False
        cboBuscaTipoProveedor.ListIndex = -1
    End If
End Sub

Private Sub CmdAgregar_Click()
    If grillaFactura.Rows > 1 Then
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 6) = grillaFactura.TextMatrix(grillaFactura.RowSel, 5) And _
               grdGrilla.TextMatrix(I, 1) = grillaFactura.TextMatrix(grillaFactura.RowSel, 1) And _
               grdGrilla.TextMatrix(I, 2) = grillaFactura.TextMatrix(grillaFactura.RowSel, 2) Then
               
               MsgBox "La Factura ya fue elegida", vbExclamation, TIT_MSGBOX
               grillaFactura.SetFocus
               Exit Sub
            End If
        Next
        grdGrilla.AddItem grillaFactura.TextMatrix(grillaFactura.RowSel, 0) & Chr(9) & _
                          grillaFactura.TextMatrix(grillaFactura.RowSel, 1) & Chr(9) & _
                          grillaFactura.TextMatrix(grillaFactura.RowSel, 2) & Chr(9) & _
                          grillaFactura.TextMatrix(grillaFactura.RowSel, 3) & Chr(9) & _
                          grillaFactura.TextMatrix(grillaFactura.RowSel, 4) & Chr(9) & _
                          "" & Chr(9) & _
                          grillaFactura.TextMatrix(grillaFactura.RowSel, 5)
    End If
End Sub

Private Sub cmdBuscarNotaCRedito_Click()
    If txtCodProveedor.Text <> "" Then
        tabDatos.Tab = 1
    Else
        MsgBox "Debe elegir un Proveedor", vbExclamation, TIT_MSGBOX
        txtCodProveedor.SetFocus
        Exit Sub
    End If
End Sub

Private Sub QuitoFacturaImputadaAntes()
  Dim SaldoFac As Double
  SaldoFac = 0
  
    If MsgBox("�Seguro que desea quitar la Factura Nro.: " _
               & grdGrilla.TextMatrix(grdGrilla.RowSel, 1) & "?" _
               & Chr(13) & "La misma ya fue imputada con anterioridad.", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        
        On Error GoTo QuePaso
        DBConn.BeginTrans
        
        SaldoFac = BuscoSaldoFactura(grdGrilla.TextMatrix(grdGrilla.RowSel, 6), grdGrilla.TextMatrix(grdGrilla.RowSel, 1) _
                                        , grdGrilla.TextMatrix(grdGrilla.RowSel, 2))
                                        
        If grdGrilla.Rows > 2 Then
            For I = 1 To grillaFactura.Rows - 1
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grillaFactura.TextMatrix(I, 5) And _
                   grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = grillaFactura.TextMatrix(I, 1) And _
                   grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = grillaFactura.TextMatrix(I, 2) Then
    
                    grillaFactura.TextMatrix(I, 4) = Valido_Importe(CStr(SaldoFac + CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 5))))
                   Exit For
                End If
            Next
            txtSaldoNC.Text = Valido_Importe(CStr(CDbl(txtSaldoNC.Text) + CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 5))))
            
            Call ActualizoSaldoFacturas(grdGrilla.TextMatrix(grdGrilla.RowSel, 6), grdGrilla.TextMatrix(grdGrilla.RowSel, 1) _
                                        , grdGrilla.TextMatrix(grdGrilla.RowSel, 2), CStr(SaldoFac + CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 5))))
                                        
            'FACTURAS POR NOTA DE CREDITO
            Call QuitoLaFacturaDeLaTabla(grdGrilla.TextMatrix(grdGrilla.RowSel, 6), grdGrilla.TextMatrix(grdGrilla.RowSel, 1) _
                                        , grdGrilla.TextMatrix(grdGrilla.RowSel, 2))
            
            grdGrilla.RemoveItem grdGrilla.RowSel
            
        Else
            
            For I = 1 To grillaFactura.Rows - 1
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grillaFactura.TextMatrix(I, 5) And _
                   grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = grillaFactura.TextMatrix(I, 1) And _
                   grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = grillaFactura.TextMatrix(I, 2) Then
    
                    grillaFactura.TextMatrix(I, 4) = Valido_Importe(CStr(SaldoFac + CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 5))))
                   Exit For
                End If
            Next
            txtSaldoNC.Text = Valido_Importe(CStr(CDbl(txtSaldoNC.Text) + CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 5))))
            
            Call ActualizoSaldoFacturas(grdGrilla.TextMatrix(grdGrilla.RowSel, 6), grdGrilla.TextMatrix(grdGrilla.RowSel, 1) _
                              , grdGrilla.TextMatrix(grdGrilla.RowSel, 2), CStr(SaldoFac + CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 5))))
                              
            'FACTURAS POR NOTA DE CREDITO
            Call QuitoLaFacturaDeLaTabla(grdGrilla.TextMatrix(grdGrilla.RowSel, 6), grdGrilla.TextMatrix(grdGrilla.RowSel, 1) _
                                        , grdGrilla.TextMatrix(grdGrilla.RowSel, 2))
                                        
            grdGrilla.Rows = 1
            'grdGrilla.HighLight = flexHighlightNever
        End If
            
            'ACTUALIZO EL SALDO DE LA NOTA DE CREDITO
            sql = "UPDATE NOTA_CREDITO_PROVEEDOR"
            sql = sql & " SET CPR_SALDO=" & XN(txtSaldoNC.Text)
            sql = sql & " WHERE"
            sql = sql & " TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
            sql = sql & " AND CPR_NUMERO=" & XN(Right(txtNroNotaCredito, 8))
            sql = sql & " AND CPR_NROSUC=" & XN(Left(txtNroNotaCredito, 4))
            sql = sql & " AND CPR_FECHA=" & XDQ(FechaNotaCredito)
            DBConn.Execute sql
            
        DBConn.CommitTrans
    End If
    Exit Sub
    
QuePaso:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX, vbCritical, TIT_MSGBOX
    CmdNuevo_Click
End Sub

Private Sub QuitoLaFacturaDeLaTabla(TIPO As String, Numero As String, Fecha As String)
    'FACTURAS POR NOTA DE CREDITO
    sql = "DELETE FROM FACTURAS_NOTA_CREDITO_PROVEEDOR"
    sql = sql & " WHERE"
    sql = sql & " TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
    sql = sql & " AND CPR_NUMERO=" & XN(Right(txtNroNotaCredito, 8))
    sql = sql & " AND CPR_NROSUC=" & XN(Left(txtNroNotaCredito, 4))
    sql = sql & " AND CPR_FECHA=" & XDQ(FechaNotaCredito)
    sql = sql & " AND FPR_TCO_CODIGO=" & XN(TIPO)
    sql = sql & " AND FPR_NUMERO=" & XN(Right(Numero, 8))
    sql = sql & " AND FPR_NROSUC=" & XN(Left(Numero, 4))
    sql = sql & " AND FPR_FECHA=" & XDQ(Fecha)
    
    DBConn.Execute sql
End Sub

Private Sub cmdBuscarProveedor_Click()
    frmBuscar.TipoBusqueda = 5
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 1
        txtProveedor.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 2
        txtDesProv.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 3
        Call BuscaCodigoProxItemData(CInt(frmBuscar.grdBuscar.Text), cboBuscaTipoProveedor)
    Else
        txtProveedor.SetFocus
    End If
End Sub

Private Sub cmdBuscarProveedor1_Click()
      frmBuscar.TipoBusqueda = 5
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 1
        txtCodProveedor.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 2
        txtProvRazSoc.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 3
        Call BuscaCodigoProxItemData(CInt(frmBuscar.grdBuscar.Text), cboTipoProveedor)
        txtCodProveedor_LostFocus
        txtProvRazSoc.SetFocus
    Else
        txtCodProveedor.SetFocus
    End If
End Sub

Private Sub cmdQuitar_Click()
    If grdGrilla.Rows > 1 Then
        If grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = "X" Then
            QuitoFacturaImputadaAntes
            Exit Sub
        End If
        If MsgBox("�Seguro que desea quitar la Factura Nro.: " _
                & grdGrilla.TextMatrix(grdGrilla.RowSel, 1) & "?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
                If grdGrilla.Rows > 2 Then
                    For I = 1 To grillaFactura.Rows - 1
                        If grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grillaFactura.TextMatrix(I, 5) And _
                           grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = grillaFactura.TextMatrix(I, 1) And _
                           grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = grillaFactura.TextMatrix(I, 2) Then
            
                            grillaFactura.TextMatrix(I, 4) = Valido_Importe(grdGrilla.TextMatrix(grdGrilla.RowSel, 4))
                            txtSaldoNC.Text = Valido_Importe(CStr(CDbl(txtSaldoNC.Text) + CDbl(IIf(grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = "", "0", grdGrilla.TextMatrix(grdGrilla.RowSel, 5)))))
                           Exit For
                        End If
                    Next

                    grdGrilla.RemoveItem grdGrilla.RowSel
                Else
                    For I = 1 To grillaFactura.Rows - 1
                        If grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grillaFactura.TextMatrix(I, 5) And _
                           grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = grillaFactura.TextMatrix(I, 1) And _
                           grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = grillaFactura.TextMatrix(I, 2) Then
            
                            grillaFactura.TextMatrix(I, 4) = Valido_Importe(grdGrilla.TextMatrix(grdGrilla.RowSel, 4))
                            txtSaldoNC.Text = Valido_Importe(CStr(CDbl(txtSaldoNC.Text) + CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 5))))
                           Exit For
                        End If
                    Next
                    grdGrilla.Rows = 1
                    'grdGrilla.HighLight = flexHighlightNever
                End If
        End If
        grdGrilla.SetFocus
    End If
End Sub

Private Function BuscoSaldoFactura(TIPO As String, Numero As String, Fecha As String) As Double
    sql = "SELECT FPR_SALDO"
    sql = sql & " FROM FACTURA_PROVEEDOR"
    sql = sql & " WHERE"
    sql = sql & " TCO_CODIGO=" & XN(TIPO)
    sql = sql & " AND FPR_NUMERO=" & XN(Right(Numero, 8))
    sql = sql & " AND FPR_NROSUC=" & XN(Left(Numero, 4))
    sql = sql & " AND FPR_FECHA=" & XDQ(Fecha)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        BuscoSaldoFactura = rec!FPR_SALDO
    Else
        BuscoSaldoFactura = 0
    End If
    rec.Close
End Function

Private Sub ActualizoSaldoFacturas(TIPO As String, Numero As String, Fecha As String, Saldo As String)
    'ACTUALIZO EL SALDO DE LAS FACTURAS ELEGIDAS
    sql = "UPDATE FACTURA_PROVEEDOR"
    sql = sql & " SET FPR_SALDO=" & XN(Saldo)
    sql = sql & " WHERE"
    sql = sql & " TCO_CODIGO=" & XN(TIPO)
    sql = sql & " AND FPR_NUMERO=" & XN(Right(Numero, 8))
    sql = sql & " AND FPR_NROSUC=" & XN(Left(Numero, 4))
    sql = sql & " AND FPR_FECHA=" & XDQ(Fecha)
    
    DBConn.Execute sql
End Sub

Private Sub BuscarFacturas(TipoProv As String, Proveedor As String)
    grillaFactura.Rows = 1
    grillaFactura.HighLight = flexHighlightNever
    
    sql = "SELECT TCO_CODIGO,TCO_ABREVIA,FPR_NROSUC,FPR_NUMERO,FPR_FECHA,FPR_TOTAL,FPR_SALDO"
    sql = sql & " FROM SALDO_FACTURAS_PROVEEDOR_V"
    sql = sql & " WHERE TPR_CODIGO=" & XN(TipoProv)
    sql = sql & " AND PROV_CODIGO=" & XN(Proveedor)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            grillaFactura.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!FPR_NROSUC, "0000") & "-" & Format(rec!FPR_NUMERO, "00000000") & Chr(9) & _
                                  rec!FPR_FECHA & Chr(9) & Valido_Importe(rec!FPR_TOTAL) & Chr(9) & _
                                  Valido_Importe(rec!FPR_SALDO) & Chr(9) & rec!TCO_CODIGO
                                  
            rec.MoveNext
        Loop
        grillaFactura.HighLight = flexHighlightAlways
    Else
        MsgBox "El Proveedor no tiene facturas con Saldo", vbExclamation, TIT_MSGBOX
    End If
    rec.Close
End Sub

Private Sub chkFecha_Click()
    If chkFecha.Value = Checked Then
        FechaDesde.Enabled = True
        FechaHasta.Enabled = True
    Else
        FechaDesde.Enabled = False
        FechaHasta.Enabled = False
    End If
End Sub

Private Sub chkTipoFactura_Click()
    If chkTipoFactura.Value = Checked Then
        cboNotaCredito1.Enabled = True
    Else
        cboNotaCredito1.Enabled = False
    End If
End Sub

Private Sub chkTipoFactura_LostFocus()
    If chkTipoFactura.Value = Checked And chkTipoProveedor.Value = Unchecked _
        And chkProveedor.Value = Unchecked _
        And chkFecha.Value = Unchecked Then cboNotaCredito1.SetFocus
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
        
     sql = "SELECT NP.*, P.PROV_RAZSOC, TP.TPR_DESCRI, TC.TCO_ABREVIA"
     sql = sql & " FROM NOTA_CREDITO_PROVEEDOR NP,"
     sql = sql & " TIPO_COMPROBANTE TC, PROVEEDOR P, TIPO_PROVEEDOR TP"
     sql = sql & " WHERE"
     sql = sql & " NP.TCO_CODIGO=TC.TCO_CODIGO"
     sql = sql & " AND NP.TPR_CODIGO=TP.TPR_CODIGO"
     sql = sql & " AND NP.PROV_CODIGO=P.PROV_CODIGO"
     sql = sql & " AND P.TPR_CODIGO=TP.TPR_CODIGO"
     sql = sql & " AND NP.CPR_SALDO <> " & XN("0,00")
    If chkTipoProveedor.Value = Checked Then sql = sql & " AND NP.TPR_CODIGO=" & cboBuscaTipoProveedor.ItemData(cboBuscaTipoProveedor.ListIndex)
    If txtProveedor.Text <> "" Then sql = sql & " AND NP.PROV_CODIGO=" & XN(txtProveedor)
    If FechaDesde <> "" Then sql = sql & " AND NP.CPR_FECHA>=" & XDQ(FechaDesde)
    If FechaHasta <> "" Then sql = sql & " AND NP.CPR_FECHA<=" & XDQ(FechaHasta)
    If chkTipoFactura.Value = Checked Then sql = sql & " AND NP.TCO_CODIGO=" & cboNotaCredito1.ItemData(cboNotaCredito1.ListIndex)
    sql = sql & " ORDER BY NP.CPR_NROSUC,NP.CPR_NUMERO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!CPR_NROSUC, "0000") & "-" & Format(rec!CPR_NUMERO, "00000000") _
                            & Chr(9) & rec!CPR_FECHA _
                            & Chr(9) & rec!PROV_RAZSOC & Chr(9) & rec!TPR_DESCRI _
                            & Chr(9) & rec!TPR_CODIGO & Chr(9) & rec!CPR_TOTAL _
                            & Chr(9) & rec!CPR_SALDO & Chr(9) & rec!TCO_CODIGO _
                            & Chr(9) & rec!PROV_CODIGO
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
    End If
    
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    rec.Close
End Sub

Private Sub cmdGrabar_Click()
    
    If ValidarNotaCredito = False Then Exit Sub
    If MsgBox("�Confirma la imputaci�n de la Nota de Cr�dito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo HayErrorFactura
    
    DBConn.BeginTrans
    sql = "SELECT PROV_CODIGO FROM FACTURAS_NOTA_CREDITO_PROVEEDOR"
    sql = sql & " WHERE TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
    sql = sql & " AND CPR_NROSUC=" & XN(txtNroSucursal)
    sql = sql & " AND CPR_NUMERO = " & XN(txtNroNotaCredito)
    sql = sql & " AND CPR_FECHA=" & XDQ(FechaNotaCredito)
    sql = sql & " AND TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
    sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = False Then
        If MsgBox("Seguro que modificar la imputaci�n de la Nota de Cr�dito Nro.: " & Trim(txtNroNotaCredito), vbQuestion + vbYesNo + vbDefaultButton2, TIT_MSGBOX) = vbYes Then
        
            'FACTURAS POR NOTA DE CREDITO
            sql = "DELETE FROM FACTURAS_NOTA_CREDITO_PROVEEDOR"
            sql = sql & " WHERE"
            sql = sql & " WHERE TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
            sql = sql & " AND CPR_NROSUC=" & XN(txtNroSucursal)
            sql = sql & " AND CPR_NUMERO = " & XN(txtNroNotaCredito)
            sql = sql & " AND CPR_FECHA=" & XDQ(FechaNotaCredito)
            sql = sql & " AND TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
            sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
            DBConn.Execute sql
            
            For I = 1 To grdGrilla.Rows - 1
                If grdGrilla.TextMatrix(I, 0) <> "" Then
                    sql = "INSERT INTO FACTURAS_NOTA_CREDITO_PROVEEDOR"
                    sql = sql & " (TPR_CODIGO,PROV_CODIGO,TCO_CODIGO,CPR_NROSUC,CPR_NUMERO,CPR_FECHA,"
                    sql = sql & " FPR_TCO_CODIGO,FPR_NROSUC,FPR_NUMERO,FPR_FECHA,FNC_IMPORTE)"
                    sql = sql & " VALUES ("
                    sql = sql & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex) & ","
                    sql = sql & XN(txtCodProveedor) & ","
                    sql = sql & cboNotaCredito.ItemData(cboNotaCredito.ListIndex) & ","
                    sql = sql & XN(txtNroSucursal) & ","
                    sql = sql & XN(txtNroNotaCredito) & ","
                    sql = sql & XDQ(FechaNotaCredito) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(I, 6)) & ","
                    sql = sql & XN(Left(grdGrilla.TextMatrix(I, 1), 4)) & ","
                    sql = sql & XN(Right(grdGrilla.TextMatrix(I, 1), 8)) & ","
                    sql = sql & XDQ(grdGrilla.TextMatrix(I, 2)) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(I, 5)) & ")"
                    DBConn.Execute sql
                End If
            Next
            
            'ACTUALIZO EL SALDO DE LA NOTA DE CREDITO
            sql = "UPDATE NOTA_CREDITO_PROVEEDOR"
            sql = sql & " SET CPR_SALDO=" & XN(txtSaldoNC.Text)
            sql = sql & " WHERE"
            sql = sql & " TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
            sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
            sql = sql & " AND TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
            sql = sql & " AND CPR_NROSUC=" & XN(txtNroSucursal)
            sql = sql & " AND CPR_NUMERO=" & XN(txtNroNotaCredito)
            DBConn.Execute sql
            
            'ACTUALIZO EL SALDO DE LAS FACTURAS ELEGIDAS
            For I = 1 To grdGrilla.Rows - 1
                If grdGrilla.TextMatrix(I, 7) <> "X" Then
                    sql = "UPDATE FACTURA_PROVEEDOR"
                    sql = sql & " SET FPR_SALDO=" & XN(CStr(CDbl(grdGrilla.TextMatrix(I, 4)) - CDbl(grdGrilla.TextMatrix(I, 5))))
                    sql = sql & " WHERE"
                    sql = sql & " TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
                    sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
                    sql = sql & " AND TCO_CODIGO=" & XN(grdGrilla.TextMatrix(I, 6))
                    sql = sql & " AND FPR_NROSUC=" & XN(Left(grdGrilla.TextMatrix(I, 1), 4))
                    sql = sql & " AND FPR_NUMERO=" & XN(Right(grdGrilla.TextMatrix(I, 1), 8))
                    DBConn.Execute sql
                End If
            Next
            
            DBConn.CommitTrans
        End If
        
    Else 'INSERTO UNA NUEVA NOTA DE CEDITO
               
        'FACTURAS POR NOTA DE CREDITO
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                sql = "INSERT INTO FACTURAS_NOTA_CREDITO_PROVEEDOR"
                sql = sql & " (TPR_CODIGO,PROV_CODIGO,TCO_CODIGO,CPR_NROSUC,CPR_NUMERO,CPR_FECHA,"
                sql = sql & " FPR_TCO_CODIGO,FPR_NROSUC,FPR_NUMERO,FPR_FECHA,FNC_IMPORTE)"
                sql = sql & " VALUES ("
                sql = sql & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex) & ","
                sql = sql & XN(txtCodProveedor) & ","
                sql = sql & cboNotaCredito.ItemData(cboNotaCredito.ListIndex) & ","
                sql = sql & XN(txtNroSucursal) & ","
                sql = sql & XN(txtNroNotaCredito) & ","
                sql = sql & XDQ(FechaNotaCredito) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 6)) & ","
                sql = sql & XN(Left(grdGrilla.TextMatrix(I, 1), 4)) & ","
                sql = sql & XN(Right(grdGrilla.TextMatrix(I, 1), 8)) & ","
                sql = sql & XDQ(grdGrilla.TextMatrix(I, 2)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 5)) & ")"
                DBConn.Execute sql
            End If
        Next
        
        'ACTUALIZO EL SALDO DE LA NOTA DE CREDITO
        sql = "UPDATE NOTA_CREDITO_PROVEEDOR"
        sql = sql & " SET CPR_SALDO=" & XN(txtSaldoNC.Text)
        sql = sql & " WHERE"
        sql = sql & " TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
        sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
        sql = sql & " AND TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
        sql = sql & " AND CPR_NROSUC=" & XN(txtNroSucursal)
        sql = sql & " AND CPR_NUMERO=" & XN(txtNroNotaCredito)
        DBConn.Execute sql
        
        'ACTUALIZO EL SALDO DE LAS FACTURAS ELEGIDAS
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 7) <> "X" Then
                sql = "UPDATE FACTURA_PROVEEDOR"
                sql = sql & " SET FPR_SALDO=" & XN(CStr(CDbl(grdGrilla.TextMatrix(I, 4)) - CDbl(grdGrilla.TextMatrix(I, 5))))
                sql = sql & " WHERE"
                sql = sql & " TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
                sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
                sql = sql & " AND TCO_CODIGO=" & XN(grdGrilla.TextMatrix(I, 6))
                sql = sql & " AND FPR_NROSUC=" & XN(Left(grdGrilla.TextMatrix(I, 1), 4))
                sql = sql & " AND FPR_NUMERO=" & XN(Right(grdGrilla.TextMatrix(I, 1), 8))
                DBConn.Execute sql
            End If
        Next
        DBConn.CommitTrans
    End If
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    CmdNuevo_Click
    Exit Sub
    
HayErrorFactura:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX, vbCritical, TIT_MSGBOX
End Sub

Private Function ValidarNotaCredito() As Boolean
    If txtCodProveedor.Text = "" Then
        MsgBox "Debe ingresar un proveedor", vbExclamation, TIT_MSGBOX
        txtCodProveedor.SetFocus
        ValidarNotaCredito = False
        Exit Function
    End If
    If txtNroNotaCredito.Text = "" Then
        MsgBox "La Nota de Cr�dito es requerida", vbExclamation, TIT_MSGBOX
        txtNroNotaCredito.SetFocus
        ValidarNotaCredito = False
        Exit Function
    End If
    If FechaNotaCredito.Value = "" Then
        MsgBox "La Fecha de la Nota de Cr�dito es requerida", vbExclamation, TIT_MSGBOX
        FechaNotaCredito.SetFocus
        ValidarNotaCredito = False
        Exit Function
    End If
    If grillaFactura.Rows = 1 Then
        MsgBox "El Proveedor no tiene Facturas con Saldo", vbExclamation, TIT_MSGBOX
        txtCodProveedor.SetFocus
        ValidarNotaCredito = False
        Exit Function
    End If
    ValidarNotaCredito = True
End Function

Private Sub CmdNuevo_Click()
   grdGrilla.Rows = 1
   txtCodProveedor.Text = ""
   txtNroNotaCredito.Text = ""
   FechaNotaCredito.Value = ""
   lblEstado.Caption = ""
   txtTotalNC.Text = ""
   txtSaldoNC.Text = ""
   cmdGrabar.Enabled = True
'   'CARGO ESTADO
'    Call BuscoEstado(1, lblEstadoNotaCredito) 'ESTADO PENDIENTE
'    VEstadoNotaCredito = 1
    '--------------
    FrameNotaCredito.Enabled = True
    FrameProveedor.Enabled = True
    tabDatos.Tab = 0
    cboNotaCredito.ListIndex = 0
    cboTipoProveedor.ListIndex = 0
    cboTipoProveedor.SetFocus
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmImputarNCaFacturaProveedores = Nothing
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
       tabDatos.Tab = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl.Name <> "grdGrilla" And _
        Me.ActiveControl.Name <> "txtEdit" And _
        KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)
    
    grdGrilla.FormatString = "^Comp.|>N�mero|^Fecha|>Total|>Saldo|>Importe|Tipo Comp.|MARCA"
    grdGrilla.ColWidth(0) = 900  'COMPROBANTE
    grdGrilla.ColWidth(1) = 1200 'NUMERO
    grdGrilla.ColWidth(2) = 1100 'FECHA
    grdGrilla.ColWidth(3) = 1000 'TOTAL
    grdGrilla.ColWidth(4) = 0    'SALDO
    grdGrilla.ColWidth(5) = 1000 'IMPORTE A ASIGNAR
    grdGrilla.ColWidth(6) = 0    'TIPO COMPROBANTE
    grdGrilla.ColWidth(7) = 0    'MARCA
    grdGrilla.Cols = 8
    grdGrilla.Rows = 1
    
    'GRILLA (GrdModulos) PARA LA BUSQUEDA
    GrdModulos.FormatString = "^Tipo|>N�mero|^Fecha|Proveedor|Tipo Proveedor|Cod Tipo Proveedor|" _
                              & "TOTAL|SALDO|COD TIPO COMPROBANTE FACTURA O NC|COD proveedor"
    GrdModulos.ColWidth(0) = 1000 'TIPO NOTA CREDITO O TIPO FACTURA
    GrdModulos.ColWidth(1) = 1200 'NUMERO
    GrdModulos.ColWidth(2) = 1000 'FECHA
    GrdModulos.ColWidth(3) = 4000 'PROVEEDOR
    GrdModulos.ColWidth(4) = 4000 'TIPO PROVEEDOR
    GrdModulos.ColWidth(5) = 0    'COD TIPO PROVEEDOR
    GrdModulos.ColWidth(6) = 0    'TOTAL
    GrdModulos.ColWidth(7) = 0    'SALDO
    GrdModulos.ColWidth(8) = 0    'COD TIPO COMPROBANTE NOTA CREDITO O FACTURA
    GrdModulos.ColWidth(9) = 0    'CODIGO PROVEEDOR
    GrdModulos.Cols = 10
    GrdModulos.Rows = 1
    'GRILLA PARA MOSTRAR LAS FACTURAS PENDIENTES DE LOS CLIENTES
    grillaFactura.FormatString = "^Comp.|^N�mero|^Fecha|>Total|>Saldo|Tipo Comp."
    grillaFactura.ColWidth(0) = 900  'COMPROBANTE
    grillaFactura.ColWidth(1) = 1200 'NUMERO
    grillaFactura.ColWidth(2) = 1100 'FECHA
    grillaFactura.ColWidth(3) = 1000 'TOTAL
    grillaFactura.ColWidth(4) = 1000 'SALDO
    grillaFactura.ColWidth(5) = 0    'TIPO COMPROBANTE
    grillaFactura.Rows = 1
    grillaFactura.Cols = 6
     grillaFactura.HighLight = flexHighlightNever
    '------------------------------------
    lblEstado.Caption = ""
    'CARGO COMBO CON LOS TIPOS DE NOTA DE CREDITO
    LlenarComboNotaCredito
    'CARGO COMBO CON LOS TIPOS PROVEEDOR
    LlenarComboTipoProv
    'LlenarComboFactura
    'CARGO ESTADO
    'Call BuscoEstado(1, lblEstadoNotaCredito) 'ESTADO PENDIENTE
   ' VEstadoNotaCredito = 1
    tabDatos.Tab = 0
End Sub

Private Sub LlenarComboTipoProv()
    sql = "SELECT * FROM TIPO_PROVEEDOR ORDER BY TPR_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboTipoProveedor.AddItem "TODOS"
        Do While rec.EOF = False
            cboTipoProveedor.AddItem rec!TPR_DESCRI
            cboTipoProveedor.ItemData(cboTipoProveedor.NewIndex) = rec!TPR_CODIGO
            cboBuscaTipoProveedor.AddItem rec!TPR_DESCRI
            cboBuscaTipoProveedor.ItemData(cboBuscaTipoProveedor.NewIndex) = rec!TPR_CODIGO
            rec.MoveNext
        Loop
        cboTipoProveedor.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboNotaCredito()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'NOTA DE CRE%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboNotaCredito.AddItem rec!TCO_DESCRI
            cboNotaCredito.ItemData(cboNotaCredito.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboNotaCredito.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub txtCodProveedor_Change()
    If txtCodProveedor.Text = "" Then
        txtProvRazSoc.Text = ""
        txtProvLoc.Text = ""
        txtProvDom.Text = ""
        txtNroSucursal.Text = ""
        txtNroNotaCredito.Text = ""
        FechaNotaCredito.Value = ""
        cboNotaCredito.ListIndex = 0
        grillaFactura.Rows = 1
        grillaFactura.HighLight = flexHighlightNever
    End If
End Sub

Private Sub txtCodProveedor_GotFocus()
    SelecTexto txtCodProveedor
End Sub

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodProveedor_LostFocus()
    If txtCodProveedor.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        Rec1.Open BuscoProveedor(txtCodProveedor), DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtProvRazSoc.Text = Rec1!PROV_RAZSOC
            txtProvLoc.Text = Rec1!LOC_DESCRI
            txtProvDom.Text = Rec1!PROV_DOMICI
            Call BuscaCodigoProxItemData(CInt(Rec1!TPR_CODIGO), cboTipoProveedor)
            Call BuscarFacturas(CStr(cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)), txtCodProveedor)
        Else
            MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
            txtCodProveedor.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtNroNotaCredito_Change()
    If txtNroNotaCredito.Text = "" Then
        FechaNotaCredito.Value = ""
    End If
End Sub

Private Sub txtNroNotaCredito_GotFocus()
    SelecTexto txtNroNotaCredito
End Sub

Private Sub txtNroSucursal_Change()
    If txtNroSucursal.Text = "" Then
        txtNroNotaCredito.Text = ""
        FechaNotaCredito.Value = ""
    End If
End Sub

Private Sub txtNroSucursal_GotFocus()
    SelecTexto txtNroSucursal
End Sub

Private Sub txtNroSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroSucursal_LostFocus()
    If txtNroSucursal.Text = "" Then
        txtNroSucursal.Text = "1"
        txtNroSucursal.Text = Format(txtNroSucursal.Text, "0000")
    Else
        txtNroSucursal.Text = Format(txtNroSucursal.Text, "0000")
    End If
End Sub

Private Sub txtProveedor_Change()
    If txtProveedor.Text = "" Then
        txtDesProv.Text = ""
    End If
End Sub

Private Sub txtProveedor_GotFocus()
    SelecTexto txtProveedor
End Sub

Private Sub txtProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtProveedor_LostFocus()
    If txtProveedor.Text <> "" Then
        sql = "SELECT TPR_CODIGO,PROV_CODIGO,PROV_RAZSOC,"
        sql = sql & " FROM PROVEEDOR"
        sql = sql & " WHERE"
        sql = sql & " PROV_CODIGO=" & XN(txtProveedor)
        
        Rec1.Open BuscoProveedor(txtCodProveedor), DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtDesProv.Text = Rec1!PROV_RAZSOC
            Call BuscaCodigoProxItemData(CInt(Rec1!TPR_CODIGO), cboBuscaTipoProveedor)
        Else
            MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
            txtProveedor.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtProvRazSoc_Change()
    If txtProvRazSoc.Text = "" Then
        txtCodProveedor.Text = ""
        txtProvLoc.Text = ""
        txtProvDom.Text = ""
    End If
End Sub

Private Sub txtProvRazSoc_GotFocus()
    SelecTexto txtProvRazSoc
End Sub

Private Sub txtProvRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtProvRazSoc_LostFocus()
    If txtCodProveedor.Text = "" And txtProvRazSoc.Text <> "" Then
        rec.Open BuscoProveedor(txtProvRazSoc), DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                frmBuscar.TipoBusqueda = 5
                frmBuscar.TxtDescriB.Text = txtProvRazSoc.Text
                frmBuscar.Show vbModal
                If frmBuscar.grdBuscar.Text <> "" Then
                    frmBuscar.grdBuscar.Col = 1
                    txtCodProveedor.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 2
                    txtProvRazSoc.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 3
                    Call BuscaCodigoProxItemData(CInt(frmBuscar.grdBuscar.Text), cboTipoProveedor)
                    txtCodProveedor_LostFocus
                Else
                    txtCodProveedor.SetFocus
                End If
            Else
                txtCodProveedor.Text = rec!PROV_CODIGO
                txtProvRazSoc.Text = rec!PROV_RAZSOC
                txtCodProveedor_LostFocus
            End If
        Else
            MsgBox "No se encontro el Proveedor", vbExclamation, TIT_MSGBOX
            txtCodProveedor.SetFocus
        End If
        rec.Close
    ElseIf txtCodProveedor.Text = "" And txtProvRazSoc.Text = "" Then
        MsgBox "Debe elegir un Proveedor", vbExclamation, TIT_MSGBOX
        txtCodProveedor.SetFocus
    End If
End Sub

Private Function BuscoProveedor(Pro As String) As String
    sql = "SELECT PRO.TPR_CODIGO,PRO.PROV_CODIGO, PRO.PROV_RAZSOC,"
    sql = sql & " PRO.PROV_DOMICI, L.LOC_DESCRI"
    sql = sql & " FROM PROVEEDOR PRO,LOCALIDAD L"
    sql = sql & " WHERE"
    If txtCodProveedor.Text <> "" Then
        sql = sql & " PRO.PROV_CODIGO=" & XN(Pro)
    Else
        sql = sql & " PRO.PROV_RAZSOC LIKE '" & Pro & "%'"
    End If
    If cboTipoProveedor.List(cboTipoProveedor.ListIndex) <> "TODOS" Then
        sql = sql & " AND PRO.TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
    End If
    sql = sql & " AND PRO.LOC_CODIGO=L.LOC_CODIGO"

    BuscoProveedor = sql
End Function

Private Sub LlenarComboNCyFAC()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'NOTA DE CRE%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    cboNotaCredito1.Clear
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboNotaCredito1.AddItem rec!TCO_DESCRI
            cboNotaCredito1.ItemData(cboNotaCredito1.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboNotaCredito1.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If grdGrilla.Rows > 1 Then
        If KeyCode = vbKeyDelete Then
            Select Case grdGrilla.Col
            Case 5
                grdGrilla.Col = 5
                grdGrilla.Text = ""
                grdGrilla.Col = 5
            End Select
        End If
    End If
End Sub

Private Sub grdGrilla_KeyPress(KeyAscii As Integer)
    If grdGrilla.Rows > 1 Then
        If (grdGrilla.Col = 0) Or (grdGrilla.Col = 1) Or _
           (grdGrilla.Col = 2) Or (grdGrilla.Col = 3) Or _
           (grdGrilla.Col = 4) Or (grdGrilla.Col = 5) Then
            If KeyAscii = vbKeyReturn Then
                If grdGrilla.Col = 5 Then
                    If grdGrilla.row < grdGrilla.Rows - 1 Then
                        grdGrilla.row = grdGrilla.row + 1
                        grdGrilla.Col = 5
                    Else
                        SendKeys "{TAB}"
                    End If
                Else
                    grdGrilla.Col = grdGrilla.Col + 1
                End If
            Else
                If grdGrilla.Col = 5 Then
                    If KeyAscii > 47 And KeyAscii < 58 And grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = "" Then
                        EDITAR grdGrilla, txtEdit, KeyAscii
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub grdGrilla_LeaveCell()
    If txtEdit.Visible = False Then Exit Sub
    grdGrilla = txtEdit.Text
    txtEdit.Visible = False
End Sub

Private Sub grdGrilla_GotFocus()
    If grdGrilla.Rows > 1 Then
        If txtEdit.Visible = False Then
            grdGrilla.Col = 5
            Exit Sub
        End If
        grdGrilla = txtEdit.Text
        txtEdit.Visible = False
    End If
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.Rows > 1 Then
        Set Rec1 = New ADODB.Recordset
        
        lblEstado.Caption = "Buscando..."
        Screen.MousePointer = vbHourglass
        'CABEZA NOTA CREDITO
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 8)), cboNotaCredito)
        txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
        txtNroNotaCredito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
        FechaNotaCredito.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 5)), cboTipoProveedor)
        txtCodProveedor.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 9)
        txtCodProveedor_LostFocus
        txtTotalNC.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 6))
        txtSaldoNC.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 7))
        txtNroNotaCredito_LostFocus
        
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        '--------------
        FrameNotaCredito.Enabled = False
        FrameProveedor.Enabled = False
        grillaFactura.SetFocus
        '--------------
        tabDatos.Tab = 0
    End If
End Sub

Private Function BuscarTipoDocAbre(Codigo As String) As String
    sql = "SELECT TCO_ABREVIA"
    sql = sql & " FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_CODIGO=" & XN(Codigo)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        BuscarTipoDocAbre = rec!TCO_ABREVIA
    Else
        BuscarTipoDocAbre = ""
    End If
    rec.Close
End Function
Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
  If tabDatos.Tab = 1 Then
    GrdModulos.Rows = 2
    txtProveedor.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    cboNotaCredito1.Enabled = False
    cboBuscaTipoProveedor.Enabled = False
    cmdBuscarProveedor.Enabled = False
    cmdGrabar.Enabled = False
    LimpiarBusqueda
    LlenarComboNCyFAC
    frameBuscar.Caption = "Buscar Nota de Cr�dito por..."
    GrdModulos.ColWidth(0) = 2000 'TIPO NOTA CREDITO O TIPO FACTURA
    If Me.Visible = True Then chkTipoProveedor.SetFocus
  End If
End Sub

Private Sub LimpiarBusqueda()
    cboBuscaTipoProveedor.ListIndex = -1
    txtProveedor.Text = ""
    txtDesProv.Text = ""
    FechaDesde.Value = ""
    FechaHasta.Value = ""
    GrdModulos.Rows = 1
    chkFecha.Value = Unchecked
    chkTipoFactura.Value = Unchecked
End Sub

Private Function BuscoCondicionIVA(IVACodigo As String) As String
    sql = "SELECT * FROM CONDICION_IVA"
    sql = sql & " WHERE IVA_CODIGO=" & XN(IVACodigo)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        BuscoCondicionIVA = rec!IVA_DESCRI
    Else
        BuscoCondicionIVA = ""
    End If
    rec.Close
End Function

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    If grdGrilla.Col = 5 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
    CarTexto KeyAscii
End Sub

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    VBonificacion = 0
    
    If KeyCode = vbKeyF1 Then
        frmBuscar.TipoBusqueda = 2
        frmBuscar.CodListaPrecio = 0
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        frmBuscar.Show vbModal
    End If

    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
            Case 5 'IMPORTE
                If txtEdit.Text <> "" Then
                    txtEdit.Text = Valido_Importe(txtEdit)
                    If CDbl(Trim(txtEdit.Text)) <= CDbl(Trim(txtSaldoNC.Text)) Then
                        If CDbl(Trim(txtEdit.Text)) <= CDbl(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 4))) Then
                            For I = 1 To grillaFactura.Rows - 1
                                If grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grillaFactura.TextMatrix(I, 5) And _
                                   grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = grillaFactura.TextMatrix(I, 1) And _
                                   grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = grillaFactura.TextMatrix(I, 2) Then
                                   
                                    grillaFactura.TextMatrix(I, 4) = Valido_Importe(CStr(CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 4)) - CDbl(txtEdit.Text)))
                                    txtSaldoNC.Text = Valido_Importe(CStr(CDbl(txtSaldoNC.Text) - CDbl(txtEdit.Text)))
                                   Exit For
                                End If
                            Next
                        Else
                            MsgBox "No puede ingresar un importe Mayor al Saldo de la Factura", vbExclamation, TIT_MSGBOX
                            txtEdit.Text = ""
                        End If
                    Else
                        MsgBox "No puede ingresar un importe Mayor al Saldo de la Nota de Cr�dito", vbExclamation, TIT_MSGBOX
                        txtEdit.Text = ""
                    End If
                End If
        End Select
        grdGrilla.SetFocus
    End If
    If KeyCode = vbKeyEscape Then
       txtEdit.Visible = False
       grdGrilla.SetFocus
    End If
End Sub

Private Function BuscoRepetetidos(Codigo As Long, Linea As Integer) As Boolean
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" Then
            If Codigo = CLng(grdGrilla.TextMatrix(I, 0)) And (I <> Linea) Then
                MsgBox "El producto ya fue elegido anteriormente", vbExclamation, TIT_MSGBOX
                BuscoRepetetidos = False
                Exit Function
            End If
        End If
    Next
    BuscoRepetetidos = True
End Function

Private Function SumaTotal() As Double
    VTotal = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 6) <> "" Then
            VTotal = VTotal + (CInt(grdGrilla.TextMatrix(I, 2)) * CDbl(grdGrilla.TextMatrix(I, 3)))
        End If
    Next
    SumaTotal = Valido_Importe(CStr(VTotal))
End Function

Private Function SumaBonificacion() As Double
    VTotal = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 6) <> "" Then
            VTotal = VTotal + CDbl(grdGrilla.TextMatrix(I, 6))
        End If
    Next
    SumaBonificacion = Valido_Importe(CStr(VTotal))
End Function

Private Sub txtNroNotaCredito_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroNotaCredito_LostFocus()
    If txtCodProveedor.Text = "" Then
        MsgBox "Debe elegir un Proveedor", vbExclamation, TIT_MSGBOX
        txtNroNotaCredito.Text = ""
        txtCodProveedor.SetFocus
        Exit Sub
    End If
    If txtNroNotaCredito.Text <> "" Then
        txtNroNotaCredito.Text = Format(txtNroNotaCredito.Text, "00000000")
    End If
    grdGrilla.Rows = 1
    'BUSCO LAS FACTURAS RELACIONADAS A LA NOTA DE CREDITO
    sql = "SELECT FN.FPR_TCO_CODIGO,FN.FPR_NROSUC ,FN.FPR_NUMERO, FN.FPR_FECHA, FN.FNC_IMPORTE,TC.TCO_ABREVIA,"
    sql = sql & "FP.FPR_TOTAL,FP.FPR_SALDO"
    sql = sql & " FROM FACTURAS_NOTA_CREDITO_PROVEEDOR FN, FACTURA_PROVEEDOR FP, TIPO_COMPROBANTE TC"
    sql = sql & " WHERE"
    sql = sql & " FN.TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
    sql = sql & " AND FN.CPR_NROSUC=" & XN(txtNroSucursal)
    sql = sql & " AND FN.CPR_NUMERO=" & XN(txtNroNotaCredito)
    If FechaNotaCredito.Value <> "" Then
        sql = sql & " AND FN.CPR_FECHA=" & XDQ(FechaNotaCredito)
    End If
    sql = sql & " AND FN.TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
    sql = sql & " AND FN.PROV_CODIGO=" & XN(txtCodProveedor)
    sql = sql & " AND FN.TPR_CODIGO=FP.TPR_CODIGO"
    sql = sql & " AND FN.PROV_CODIGO=FP.PROV_CODIGO"
    sql = sql & " AND FN.FPR_TCO_CODIGO=FP.TCO_CODIGO"
    sql = sql & " AND FN.FPR_NROSUC=FP.FPR_NROSUC"
    sql = sql & " AND FN.FPR_NUMERO=FP.FPR_NUMERO"
    sql = sql & " AND FN.FPR_TCO_CODIGO=TC.TCO_CODIGO"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            grdGrilla.AddItem Rec1!TCO_ABREVIA & Chr(9) & Format(Rec1!FPR_NROSUC, "0000") & "-" & Format(Rec1!FPR_NUMERO, "00000000") & Chr(9) & _
                                  Rec1!FPR_FECHA & Chr(9) & Valido_Importe(Rec1!FPR_TOTAL) & Chr(9) & _
                                  "" & Chr(9) & Valido_Importe(Rec1!FNC_IMPORTE) & Chr(9) & _
                                  Rec1!FPR_TCO_CODIGO & Chr(9) & "X"
            Rec1.MoveNext
            CambiaColorAFilaDeGrilla grdGrilla, grdGrilla.Rows - 1, vbRed
        Loop
        grdGrilla.HighLight = flexHighlightAlways
    End If
    Rec1.Close
    
    'BUSCO EL TOTAL Y SALDO DE LA NOTA DE CREDITO
    sql = "SELECT CPR_FECHA,CPR_TOTAL,CPR_SALDO"
    sql = sql & " FROM NOTA_CREDITO_PROVEEDOR"
    sql = sql & " WHERE"
    sql = sql & " TCO_CODIGO=" & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
    sql = sql & " AND CPR_NROSUC=" & XN(txtNroSucursal)
    sql = sql & " AND CPR_NUMERO=" & XN(txtNroNotaCredito)
    If FechaNotaCredito.Value <> "" Then
        sql = sql & " AND CPR_FECHA=" & XDQ(FechaNotaCredito)
    End If
    sql = sql & " AND TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
    sql = sql & " AND PROV_CODIGO=" & XN(txtCodProveedor)
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        FechaNotaCredito.Value = rec!CPR_FECHA
        txtTotalNC.Text = Valido_Importe(rec!CPR_TOTAL)
        txtSaldoNC.Text = Valido_Importe(rec!CPR_SALDO)
        If CDbl(rec!CPR_SALDO) > 0 Then
            cmdGrabar.Enabled = True
        Else
            cmdGrabar.Enabled = False
        End If
    End If
    rec.Close
    
End Sub

