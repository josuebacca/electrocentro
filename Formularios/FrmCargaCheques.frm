VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Begin VB.Form FrmCargaCheques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de Cheques de Terceros"
   ClientHeight    =   5370
   ClientLeft      =   2535
   ClientTop       =   1005
   ClientWidth     =   7560
   Icon            =   "FrmCargaCheques.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "FrmCargaCheques.frx":08CA
      Height          =   735
      Left            =   6540
      Picture         =   "FrmCargaCheques.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4560
      Width           =   900
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
      DisabledPicture =   "FrmCargaCheques.frx":0EDE
      Height          =   735
      Left            =   5625
      Picture         =   "FrmCargaCheques.frx":11E8
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4560
      Width           =   900
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "FrmCargaCheques.frx":14F2
      Height          =   735
      Left            =   4710
      Picture         =   "FrmCargaCheques.frx":17FC
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4560
      Width           =   900
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      DisabledPicture =   "FrmCargaCheques.frx":1B06
      Height          =   735
      Left            =   3795
      Picture         =   "FrmCargaCheques.frx":1E10
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      Width           =   900
   End
   Begin TabDlg.SSTab TabDatos 
      Height          =   4455
      Left            =   50
      TabIndex        =   21
      Top             =   50
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
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
      TabPicture(0)   =   "FrmCargaCheques.frx":211A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "FrmCargaCheques.frx":2136
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdModulos"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         Caption         =   " Buscar por "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1185
         Left            =   -74850
         TabIndex        =   58
         Top             =   360
         Width           =   7110
         Begin VB.TextBox txtBusNroCheque 
            Height          =   315
            Left            =   1140
            MaxLength       =   10
            TabIndex        =   17
            Top             =   360
            Width           =   1380
         End
         Begin VB.CommandButton cmdBuscaAprox 
            Height          =   450
            Left            =   6345
            MaskColor       =   &H000000FF&
            Picture         =   "FrmCargaCheques.frx":2152
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Buscar"
            Top             =   420
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin VB.TextBox txtBusBanco 
            Height          =   315
            Left            =   1140
            TabIndex        =   18
            Top             =   750
            Width           =   4815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro Cheque:"
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   60
            Top             =   375
            Width           =   900
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Banco:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   570
            TabIndex        =   59
            Top             =   765
            Width           =   510
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos del Cheque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3870
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   7215
         Begin VB.CommandButton cmdAgregoBanco 
            Height          =   315
            Left            =   2760
            MaskColor       =   &H000000FF&
            Picture         =   "FrmCargaCheques.frx":245C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Agregar Banco"
            Top             =   1235
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtCodBanco 
            Height          =   315
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   2
            Top             =   1235
            Width           =   900
         End
         Begin VB.CommandButton cmdBusBanco 
            Height          =   315
            Left            =   2280
            MaskColor       =   &H000000FF&
            Picture         =   "FrmCargaCheques.frx":27E6
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Buscar Banco"
            Top             =   1235
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtDesBanco 
            Height          =   315
            Left            =   3240
            MaxLength       =   100
            TabIndex        =   5
            Top             =   1235
            Width           =   3780
         End
         Begin VB.TextBox TxtCheImport 
            Height          =   315
            Left            =   1350
            TabIndex        =   10
            Top             =   2865
            Width           =   1125
         End
         Begin VB.TextBox TxtCheObserv 
            Height          =   315
            Left            =   1350
            MaxLength       =   60
            TabIndex        =   11
            Top             =   3285
            Width           =   5640
         End
         Begin VB.TextBox TxtCheNombre 
            Height          =   315
            Left            =   1350
            MaxLength       =   40
            TabIndex        =   6
            Top             =   1650
            Width           =   5640
         End
         Begin VB.TextBox TxtCheMotivo 
            Height          =   315
            Left            =   1350
            MaxLength       =   40
            TabIndex        =   7
            Top             =   2065
            Width           =   5640
         End
         Begin VB.Frame Frame2 
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
            Height          =   240
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Visible         =   0   'False
            Width           =   480
            Begin VB.TextBox TxtBanDescri 
               BackColor       =   &H00C0C0C0&
               Height          =   330
               Left            =   210
               TabIndex        =   44
               Top             =   675
               Width           =   5820
            End
            Begin VB.TextBox TxtCodInt 
               BackColor       =   &H80000018&
               Height          =   300
               Left            =   5370
               TabIndex        =   43
               Top             =   660
               Visible         =   0   'False
               Width           =   420
            End
            Begin VB.TextBox TxtCODIGO 
               Height          =   285
               Left            =   4755
               MaxLength       =   6
               TabIndex        =   42
               Top             =   270
               Width           =   765
            End
            Begin VB.TextBox TxtLOCALIDAD 
               Height          =   285
               Left            =   2175
               MaxLength       =   3
               TabIndex        =   41
               Top             =   285
               Width           =   450
            End
            Begin VB.TextBox TxtBANCO 
               Height          =   285
               Left            =   780
               MaxLength       =   3
               TabIndex        =   40
               Top             =   285
               Width           =   450
            End
            Begin VB.TextBox TxtSUCURSAL 
               Height          =   285
               Left            =   3525
               MaxLength       =   3
               TabIndex        =   39
               Top             =   285
               Width           =   450
            End
            Begin VB.CommandButton CmdBanco 
               DisabledPicture =   "FrmCargaCheques.frx":2AF0
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   5655
               Picture         =   "FrmCargaCheques.frx":2DFA
               Style           =   1  'Graphical
               TabIndex        =   38
               Top             =   270
               Width           =   375
            End
            Begin VB.Label lbl 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Código:"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   4125
               TabIndex        =   48
               Top             =   315
               Width           =   540
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
               Left            =   2820
               TabIndex        =   47
               Top             =   315
               Width           =   645
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
               Left            =   225
               TabIndex        =   46
               Top             =   315
               Width           =   510
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
               Left            =   1395
               TabIndex        =   45
               Top             =   315
               Width           =   735
            End
         End
         Begin VB.TextBox TxtCheNumero 
            Height          =   315
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   1
            Top             =   820
            Width           =   1380
         End
         Begin VB.CommandButton cmdBuscaCheque 
            Height          =   315
            Left            =   6360
            MaskColor       =   &H000000FF&
            Picture         =   "FrmCargaCheques.frx":2F44
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Buscar Cheques"
            Top             =   600
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   405
         End
         Begin FechaCtl.Fecha TxtCheFecEnt 
            Height          =   315
            Left            =   1350
            TabIndex        =   0
            Top             =   390
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
            FechaMin        =   "01/01/1900"
         End
         Begin FechaCtl.Fecha TxtCheFecVto 
            Height          =   315
            Left            =   5895
            TabIndex        =   9
            Top             =   2445
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha TxtCheFecEmi 
            Height          =   315
            Left            =   1350
            TabIndex        =   8
            Top             =   2445
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Vto:"
            Height          =   195
            Index           =   3
            Left            =   4830
            TabIndex        =   57
            Top             =   2490
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha  Emisión:"
            Height          =   195
            Index           =   2
            Left            =   105
            TabIndex        =   56
            Top             =   2490
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Index           =   1
            Left            =   660
            TabIndex        =   55
            Top             =   2895
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Index           =   0
            Left            =   735
            TabIndex        =   54
            Top             =   435
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   53
            Top             =   3315
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro Cheque:"
            Height          =   195
            Index           =   7
            Left            =   330
            TabIndex        =   52
            Top             =   840
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Responsable:"
            Height          =   195
            Index           =   9
            Left            =   255
            TabIndex        =   51
            Top             =   1665
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Concepto:"
            Height          =   195
            Index           =   10
            Left            =   495
            TabIndex        =   50
            Top             =   2085
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Index           =   4
            Left            =   720
            TabIndex        =   49
            Top             =   1260
            Width           =   510
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Buscar por "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1185
         Left            =   -74865
         TabIndex        =   22
         Top             =   405
         Width           =   6750
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   330
            Left            =   5985
            MaskColor       =   &H000000FF&
            Picture         =   "FrmCargaCheques.frx":324E
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Buscar"
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   435
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   3615
            MaxLength       =   3
            TabIndex        =   27
            Top             =   330
            Width           =   540
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   660
            MaxLength       =   3
            TabIndex        =   26
            Top             =   330
            Width           =   540
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   2175
            MaxLength       =   3
            TabIndex        =   25
            Top             =   330
            Width           =   540
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   5025
            MaxLength       =   6
            TabIndex        =   24
            Top             =   330
            Width           =   795
         End
         Begin VB.TextBox TxtDescri 
            Height          =   285
            Left            =   1020
            TabIndex        =   23
            Top             =   795
            Width           =   4815
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Localidad:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   6
            Left            =   1380
            TabIndex        =   33
            Top             =   360
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
            Index           =   4
            Left            =   105
            TabIndex        =   32
            Top             =   360
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
            Index           =   3
            Left            =   2910
            TabIndex        =   31
            Top             =   360
            Width           =   645
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   4425
            TabIndex        =   30
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   29
            Top             =   825
            Width           =   870
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdBancos 
         Height          =   2610
         Left            =   -74895
         TabIndex        =   34
         Top             =   1635
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   4604
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorSel    =   8388736
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdModulos 
         Height          =   2610
         Left            =   -74880
         TabIndex        =   20
         Top             =   1590
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   4604
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         BackColorSel    =   8388736
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
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
      Left            =   180
      TabIndex        =   16
      Top             =   4065
      Width           =   750
   End
End
Attribute VB_Name = "FrmCargaCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function Validar() As Boolean
   If Trim(TxtCheNumero.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Número de Cheque.", 16, TIT_MSGBOX
        TxtCheNumero.SetFocus
        Exit Function
        
   ElseIf Trim(txtCodBanco.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Banco.", 16, TIT_MSGBOX
        txtCodBanco.SetFocus
        Exit Function
        
'   ElseIf Trim(txtlocalidad.Text) = "" Then
'        Validar = False
'        MsgBox "Ingrese la Localidad del Banco.", 16, TIT_MSGBOX
'        txtlocalidad.SetFocus
'        Exit Function
'
'   ElseIf Trim(TxtSucursal.Text) = "" Then
'        Validar = False
'        MsgBox "Ingrese la Sucursal del Banco.", 16, TIT_MSGBOX
'        TxtSucursal.SetFocus
'        Exit Function
'
'   ElseIf Trim(TxtCodigo.Text) = "" Then
'        Validar = False
'        MsgBox "Ingrese el Código del Banco.", 16, TIT_MSGBOX
'        TxtCodigo.SetFocus
'        Exit Function
'
'   ElseIf Trim(Me.TxtCodInt.Text) = "" Then
'        Validar = False
'        MsgBox "Verifique el Código de Banco.", 16, TIT_MSGBOX
'        TxtCodigo.SetFocus
'        Exit Function
        
   ElseIf Trim(TxtCheMotivo.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Concepto del Cheque.", 16, TIT_MSGBOX
        TxtCheMotivo.SetFocus
        Exit Function
        
   ElseIf Trim(TxtCheFecEmi.Text) = "" Then
        Validar = False
        MsgBox "Ingrese la Fecha de Emisión.", 16, TIT_MSGBOX
        TxtCheFecEmi.SetFocus
        Exit Function
        
   ElseIf Trim(TxtCheFecVto.Text) = "" Then
        Validar = False
        MsgBox "Ingrese la Fecha de Vencimiento.", 16, TIT_MSGBOX
        TxtCheFecVto.SetFocus
        Exit Function
        
   ElseIf Trim(TxtCheImport.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Importe del Cheque.", 16, TIT_MSGBOX
        TxtCheImport.SetFocus
        Exit Function
        
   End If
   
   Validar = True
End Function

Private Sub cmdAgregoBanco_Click()
    ABMBanco.Show vbModal
End Sub

Private Sub CmdBanco_Click()
    Viene_Cheque = True
    buscobanco = 1
    ABMBanco.Show vbModal
    Viene_Cheque = False
    buscobanco = 0
End Sub

Private Sub CmdBorrar_Click()
    On Error GoTo CLAVOSE
    
    If Trim(TxtCheNumero.Text) <> "" And Trim(Me.TxtCodInt.Text) <> "" Then
        resp = MsgBox("Seguro desea eliminar el Cheque Nº: " & Trim(Me.TxtCheNumero.Text) & "? ", 36, TIT_MSGBOX)
        If resp <> 6 Then Exit Sub
        
        
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Borrando..."
        
        sql = "SELECT BOL_NUMERO "
        sql = sql & " FROM ChequeEstadoVigente "
        sql = sql & " WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text)
        sql = sql & " AND BAN_CODINT = " & XN(Me.TxtCodInt.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If Not IsNull(rec!BOL_NUMERO) Then
               MsgBox "No se puede eliminar este Cheque porque fue depositado", vbExclamation, TIT_MSGBOX
               rec.Close
               Screen.MousePointer = vbNormal
               lblEstado.Caption = ""
               Exit Sub
             End If
        End If
        rec.Close

        DBConn.BeginTrans
        DBConn.Execute "DELETE FROM CHEQUE_ESTADOS WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text) & " AND BAN_CODINT = " & XN(Me.TxtCodInt.Text)
                       
        DBConn.Execute "DELETE FROM CHEQUE WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text) & " AND BAN_CODINT = " & XN(Me.TxtCodInt.Text)
        
        Screen.MousePointer = vbNormal
        lblEstado.Caption = ""
        DBConn.CommitTrans
        CmdNuevo_Click
    End If
    Exit Sub
    
CLAVOSE:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub cmdBusBanco_Click()
    frmBuscar.TipoBusqueda = 10
    frmBuscar.TxtDescriB.Text = ""
    frmBuscar.Show vbModal
    frmBuscar.grdBuscar.Col = 0
    txtCodBanco.Text = frmBuscar.grdBuscar.Text
    frmBuscar.grdBuscar.Col = 1
    txtDesBanco.Text = frmBuscar.grdBuscar.Text
    
End Sub

Private Sub cmdBuscaAprox_Click()
    sql = "SELECT C.*,B.BAN_DESCRI "
    sql = sql & " FROM CHEQUE C, BANCO B"
    sql = sql & " WHERE C.BAN_CODINT = B.BAN_CODINT"
    If txtBusNroCheque.Text <> "" Then
        sql = sql & " AND C.CHE_NUMERO LIKE '" & Trim(txtBusNroCheque.Text) & "%' "
    End If
    If txtBusBanco.Text <> "" Then
        sql = sql & " AND B.BAN_DESCRI LIKE '%" & XN(txtBusBanco.Text) & "%' "
    End If
    sql = sql & " ORDER BY C.CHE_FECVTO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    GrdModulos.Rows = 1
    If rec.EOF = False Then
        Do While rec.EOF = False
        'Nro Cheque|Banco|Responsable|Fecha Vto|Importe|CodBanco|Fecha|FechaEmision
            GrdModulos.AddItem rec!CHE_NUMERO & Chr(9) & rec!BAN_DESCRI & Chr(9) & _
                               rec!CHE_NOMBRE & Chr(9) & rec!CHE_FECVTO & Chr(9) & _
                               rec!che_import & Chr(9) & rec!BAN_CODINT & Chr(9) & _
                               rec!CHE_FECENT & Chr(9) & rec!CHE_FECEMI & Chr(9) & _
                               rec!CHE_MOTIVO & Chr(9) & rec!CHE_OBSERV
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
        txtBusNroCheque.SetFocus
    End If
    rec.Close
    
End Sub

Private Sub cmdBuscaCheque_Click()
    Dim codint As Integer
    frmBuscar.TipoBusqueda = 6
    frmBuscar.Show vbModal
    'TxtCheNumero.Text = frmBuscar.grdBuscar.Col
    TxtCheNumero.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
    
    
    
    
    sql = "SELECT C.*, E.ECH_CODIGO, E.ECH_DESCRI,B.BAN_DESCRI "
    sql = sql & " FROM CHEQUE C, ESTADO_CHEQUE E, CHEQUE_ESTADOS CE, BANCO B "
    sql = sql & "WHERE "
    sql = sql & " C.CHE_NUMERO = CE.CHE_NUMERO AND"
    sql = sql & " C.BAN_CODINT = CE.BAN_CODINT AND"
    sql = sql & " CE.ECH_CODIGO = E.ECH_CODIGO AND"
    sql = sql & " B.BAN_CODINT = C.BAN_CODINT AND"
    sql = sql & " C.CHE_NUMERO LIKE '" & TxtCheNumero & "'"
    sql = sql & " AND C.BAN_CODINT = " & XN(frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 4))
    sql = sql & " AND CE.CES_FECHA = (SELECT MAX(CES_FECHA) FROM CHEQUE_ESTADOS"
    sql = sql & " WHERE BAN_CODINT = " & XN(frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 4)) & " "
    sql = sql & " AND CHE_NUMERO LIKE '" & TxtCheNumero & "'" & ")"
    
    'sql = sql & " AND E.ECH_CODIGO <> 1 "
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        
        If rec!ECH_CODIGO = 1 Then 'EN CARTERA
            TxtBanDescri.Text = rec!BAN_DESCRI
            TxtBanco.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 5)
            TxtCodigo.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 8)
            txtlocalidad.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 6)
            TxtSucursal.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 7)
            TxtCodigo.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 8)
            TxtCheImport.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 3)
            TxtCheFecVto.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 2)
            TxtCodInt.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 4)
            TxtCheNombre.Text = IIf(IsNull(rec!CHE_NOMBRE), "", rec!CHE_NOMBRE)
            TxtCheMotivo.Text = IIf(IsNull(rec!CHE_MOTIVO), "", rec!CHE_MOTIVO)
            TxtCheFecEmi.Text = IIf(IsNull(rec!CHE_FECEMI), "", rec!CHE_FECEMI)
            TxtCheObserv.Text = IIf(IsNull(rec!CHE_OBSERV), "", rec!CHE_OBSERV)
            
        Else
            MsgBox "El Cheque " & TxtCheNumero.Text & " - " & Trim(rec!BAN_DESCRI) & " no puede MODIFICARSE porque está  en estado " & rec!ECH_DESCRI & " ", vbInformation, TIT_MSGBOX
            CmdNuevo_Click
        End If
    End If
    rec.Close
    
    
End Sub

Private Sub CmdGrabar_Click()
    
  If Validar = True Then
  
    On Error GoTo CLAVOSE
    
    DBConn.BeginTrans
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    Me.Refresh
    
    sql = "SELECT * FROM CHEQUE WHERE CHE_NUMERO LIKE '" & TxtCheNumero.Text & "' "
    sql = sql & " AND BAN_CODINT = " & XN(txtCodBanco.Text) & ""
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount = 0 Then
         sql = "INSERT INTO CHEQUE(CHE_NUMERO,BAN_CODINT,CHE_NOMBRE,CHE_IMPORT,CHE_FECEMI,"
         sql = sql & "CHE_FECVTO,CHE_FECENT,CHE_MOTIVO,CHE_OBSERV)"
         sql = sql & " VALUES (" & XS(Me.TxtCheNumero.Text) & ","
         sql = sql & XN(Me.txtCodBanco.Text) & "," & XS(Me.TxtCheNombre.Text) & ","
         sql = sql & XN(Me.TxtCheImport.Text) & "," & XDQ(Me.TxtCheFecEmi.Text) & ","
         sql = sql & XDQ(Me.TxtCheFecVto.Text) & "," & XDQ(Me.TxtCheFecEnt.Text) & ","
         sql = sql & XS(Me.TxtCheMotivo.Text) & "," & XS(Me.TxtCheObserv.Text) & " )"
         DBConn.Execute sql
         
         'Insert en la Tabla de Estados de Cheques
            sql = "INSERT INTO CHEQUE_ESTADOS (CHE_NUMERO,BAN_CODINT,ECH_CODIGO,CES_FECHA,CES_DESCRI)"
            sql = sql & " VALUES ("
            sql = sql & XS(Me.TxtCheNumero.Text) & ","
            sql = sql & XN(Me.txtCodBanco.Text) & "," & XN(1) & ","
            sql = sql & XDQ(Date) & ",'CHEQUE EN CARTERA')"
            DBConn.Execute sql
    Else
         sql = "UPDATE CHEQUE SET CHE_NOMBRE = " & XS(Me.TxtCheNombre.Text)
         sql = sql & ",CHE_IMPORT = " & XN(Me.TxtCheImport.Text)
         sql = sql & ",CHE_FECEMI =" & XDQ(Me.TxtCheFecEmi.Text)
         sql = sql & ",CHE_FECVTO =" & XDQ(Me.TxtCheFecVto.Text)
         sql = sql & ",CHE_FECENT = " & XDQ(Me.TxtCheFecEnt.Text)
         sql = sql & ",CHE_MOTIVO = " & XS(Me.TxtCheMotivo.Text)
         sql = sql & ",CHE_OBSERV = " & XS(Me.TxtCheObserv.Text)
         sql = sql & " WHERE CHE_NUMERO LIKE '" & Me.TxtCheNumero.Text & "'"
         sql = sql & " AND BAN_CODINT = " & XN(Me.txtCodBanco.Text)
         DBConn.Execute sql
    End If
    rec.Close
     
    
    
    '************* PREGUNTAR POR SI DESEA IMPRIMIR ***************
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.CommitTrans
    CmdNuevo_Click
 End If
 Exit Sub
      
CLAVOSE:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
End Sub

Private Sub CmdNuevo_Click()
    Me.TxtCheFecEnt.Text = ""
    Me.TxtCheNumero.Enabled = True
    Me.TxtBanco.Enabled = True
    Me.txtlocalidad.Enabled = True
    Me.TxtSucursal.Enabled = True
    Me.TxtCodigo.Enabled = True
    Me.TxtCheNombre.Enabled = True
   ' MtrObjetos = Array(TxtCheNumero, TxtBANCO, TxtLOCALIDAD, TxtSUCURSAL, TxtCODIGO, TxtCheNombre)
   ' Call CambiarColor(MtrObjetos, 6, &H80000005, "E")
    Me.TxtCheNumero.Text = ""
    Me.TxtBanco.Text = ""
    Me.txtlocalidad.Text = ""
    Me.TxtSucursal.Text = ""
    TxtBanDescri.Text = ""
    Me.TxtCodigo.Text = ""
    Me.TxtCodInt.Text = ""
    Me.TxtCheNombre.Text = ""
    Me.TxtCheMotivo.Text = ""
    Me.TxtCheFecEmi.Text = ""
    Me.TxtCheFecVto.Text = ""
    Me.TxtCheImport.Text = ""
    Me.TxtCheObserv.Text = ""
    Me.TxtCheFecEnt.SetFocus
    'TxtCheNombre.ForeColor = &H80000005
    lblEstado.Caption = ""
    Me.txtCodBanco = ""
    Me.txtDesBanco = ""
    tabDatos.Tab = 0
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    Set FrmCargaCheques = Nothing
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then tabDatos.Tab = 1
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    
    TxtCheFecEnt.Text = Date
    lblEstado.Caption = ""
    tabDatos.Tab = 0
    preparogrilla
End Sub
Private Function preparogrilla()
    lblEstado.Caption = ""
'ACA LLEGUE, CONFIGURAR LA GRILLA Y ARMAR LA SQL
    
    GrdModulos.FormatString = "Nro Cheque|Banco|Responsable|Fecha Vto|Importe|CODBANCO|FECHA ENT|FechaEmision|Motivo|Obser"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 1900
    GrdModulos.ColWidth(2) = 1900
    GrdModulos.ColWidth(3) = 1000
    GrdModulos.ColWidth(4) = 1000
    GrdModulos.ColWidth(5) = 0
    GrdModulos.ColWidth(6) = 0
    GrdModulos.ColWidth(7) = 0
    GrdModulos.ColWidth(8) = 0
    GrdModulos.ColWidth(9) = 0
    GrdModulos.Rows = 2
    
End Function

Private Sub TabTB_DblClick()

End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub
Private Sub GrdModulos_DblClick()
    If GrdModulos.Rows > 1 Then
        tabDatos.Tab = 0
        '"Nro Cheque|Banco|Responsable|Fecha Vto|Importe|CodBanco|Fecha|FechaEmision"
        TxtCheFecEnt.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 6)
        TxtCheNumero.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
        txtCodBanco.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 5)
        txtCodBanco_LostFocus
        TxtCheNombre.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
        TxtCheMotivo = GrdModulos.TextMatrix(GrdModulos.RowSel, 8)
        TxtCheFecEmi.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 7)
        TxtCheFecVto.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 3)
        TxtCheImport.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 4))
        TxtCheObserv.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 9)
    End If
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
    If tabDatos.Tab = 0 And Me.Visible Then
        TxtCheNumero.SetFocus
        cmdGrabar.Enabled = True
        cmdBorrar.Enabled = True
    End If
    If tabDatos.Tab = 1 Then
        txtBusNroCheque.Text = ""
        txtBusBanco.Text = ""
        txtBusNroCheque.SetFocus
        cmdGrabar.Enabled = False
        cmdBorrar.Enabled = False
    End If
End Sub

Private Sub TxtBANCO_LostFocus()
    If Len(TxtBanco.Text) < 3 Then TxtBanco.Text = CompletarConCeros(TxtBanco.Text, 3)
End Sub

Private Sub txtBusBanco_GotFocus()
    SelecTexto txtBusBanco
End Sub

Private Sub txtBusNroCheque_GotFocus()
    SelecTexto txtBusNroCheque
End Sub

Private Sub txtBusNroCheque_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtBusNroCheque_LostFocus()
    If txtBusNroCheque.Text <> "" Then
        If Len(txtBusNroCheque.Text) < 8 Then txtBusNroCheque.Text = CompletarConCeros(txtBusNroCheque.Text, 8)
    End If
End Sub

Private Sub TxtCheFecVto_LostFocus()
 If Trim(Me.TxtCheFecEmi.Text) <> "" And Trim(Me.TxtCheFecVto.Text) <> "" Then
 
   If IsDate(TxtCheFecEmi.Text) And IsDate(TxtCheFecVto.Text) Then
    
    If CVDate(TxtCheFecEmi.Text) > CVDate(TxtCheFecVto.Text) Then
        MsgBox "La Fecha de Vencimiento no puede ser anterior a la Fecha de Emisión del Cheque.! ", 16, TIT_MSGBOX
        Me.TxtCheFecVto.Text = ""
        Me.TxtCheFecVto.SetFocus
    Else
       If Me.TxtCheImport.Enabled = False Then 'PAGO EN CUOTAS
            Tasa = Trim(FrmComprobante.txtPmt_Tasa.Text)
            'Saco la Cantidad de Días del Cheque
            Cant_Dias = DateDiff("d", FrmComprobante.TxtFechaComprobante.Text, Me.TxtCheFecVto.Text)
            
            'Cálculo de Interes a Fecha del Cheque
            TxtCheImport.Text = Format(TxtCheImport.Text + (CDbl(TxtCheImport.Text) * CDbl(Chk0(Cant_Dias * Tasa)) / 100), "$ ##,##0.00")
            
        End If
    End If
  End If
 End If
 
End Sub

Private Sub TxtCheImport_GotFocus()
    With TxtCheImport
    .SelStart = 0
    .SelLength = Len(.Text)
    End With
End Sub

Private Sub TxtCheImport_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(TxtCheImport.Text, KeyAscii)
End Sub

Private Sub TxtCheImport_LostFocus()
   If Trim(TxtCheImport.Text) <> "" Then TxtCheImport.Text = Valido_Importe(TxtCheImport)
    
End Sub

Private Sub TxtCheNombre_LostFocus()
   If Me.TxtCheNombre.Text <> "" Then
      Me.TxtCheMotivo.SetFocus
   Else
      MsgBox "Debe ingresar la Persona responsable.!", 16, TIT_MSGBOX
   End If
End Sub

Private Sub TxtCheNumero_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCheFecEnt_LostFocus()
    If TxtCheFecEnt.Text = "" Then
        TxtCheFecEnt.Text = Format(Date, "dd/mm/yyyy")
    End If
End Sub

Private Sub TxtCheMotivo_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCheNombre_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCheNumero_LostFocus()
   If Len(TxtCheNumero.Text) < 8 Then TxtCheNumero.Text = CompletarConCeros(TxtCheNumero.Text, 8)
End Sub

Private Sub TxtCheObserv_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtBanco_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodBanco_GotFocus()
    SelecTexto txtCodBanco
End Sub

Private Sub txtCodBanco_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodBanco_LostFocus()
    If txtCodBanco.Text <> "" Then
        sql = "SELECT BAN_DESCRI FROM BANCO WHERE BAN_CODINT = " & XN(txtCodBanco.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesBanco.Text = IIf(IsNull(rec!BAN_DESCRI), "", rec!BAN_DESCRI)
        Else
            MsgBox "El Codigo ingresado no es valido!", vbExclamation, TIT_MSGBOX
            txtCodBanco.SetFocus
            txtDesBanco.Text = ""
        End If
        rec.Close
    End If
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    Dim MtrObjetos As Variant
    
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
        
    'ChequeRegistrado = False
    
    If Len(TxtCodigo.Text) < 6 Then TxtCodigo.Text = CompletarConCeros(TxtCodigo.Text, 6)
     
    If Trim(Me.TxtCheNumero.Text) <> "" And _
       Trim(Me.TxtBanco.Text) <> "" And _
       Trim(Me.txtlocalidad.Text) <> "" And _
       Trim(Me.TxtSucursal.Text) <> "" And _
       Trim(Me.TxtCodigo.Text) <> "" Then
       
       'BUSCO EL CODIGO INTERNO
       sql = "SELECT BAN_CODINT, BAN_DESCRI FROM BANCO WHERE BAN_BANCO = " & _
       XS(TxtBanco.Text) & " AND BAN_LOCALIDAD = " & _
       XS(Me.txtlocalidad.Text) & " AND BAN_SUCURSAL = " & _
       XS(Me.TxtSucursal.Text) & " AND BAN_CODIGO = " & XS(TxtCodigo.Text)
       rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
       If rec.RecordCount > 0 Then 'EXITE
          TxtCodInt.Text = rec!BAN_CODINT
          TxtBanDescri.Text = rec!BAN_DESCRI
          rec.Close
       Else
          If Me.ActiveControl.Name <> "CmdSalir" And Me.ActiveControl.Name <> "CmdNuevo" Then
            MsgBox "Banco NO Registrado.", 16, TIT_MSGBOX
            Me.CmdBanco.SetFocus
          End If
          rec.Close
          Exit Sub
       End If
       
       'CONSULTO SI EXISTE EL CHEQUE
        sql = "SELECT * FROM CHEQUE " & _
              " WHERE CHE_NUMERO = " & XS(TxtCheNumero.Text) & _
                " AND BAN_CODINT = " & XN(TxtCodInt.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then 'EXITE
            Me.TxtCheFecEnt.Text = rec!CHE_FECENT
            Me.TxtCheNumero.Text = Trim(rec!CHE_NUMERO)
            
            'BUSCO LOS ATRIBUTOS DE BANCO
            sql = "SELECT BAN_BANCO,BAN_LOCALIDAD,BAN_SUCURSAL,BAN_CODIGO FROM BANCO " & _
                   "WHERE BAN_CODINT = " & XN(Me.TxtCodInt.Text)
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.RecordCount > 0 Then 'EXITE
                Me.TxtBanco.Text = Rec1!BAN_BANCO
                Me.txtlocalidad.Text = Rec1!BAN_LOCALIDAD
                Me.TxtSucursal.Text = Rec1!BAN_SUCURSAL
                Me.TxtCodigo.Text = Rec1!BAN_CODIGO
            End If
            Rec1.Close
            Me.TxtCheNombre.Text = ChkNull(rec!CHE_NOMBRE)
            Me.TxtCheMotivo.Text = rec!CHE_MOTIVO
            Me.TxtCheFecEmi.Text = rec!CHE_FECEMI
            Me.TxtCheFecVto.Text = rec!CHE_FECVTO
            Me.TxtCheImport.Text = Valido_Importe(rec!che_import)
            Me.TxtCheObserv.Text = ChkNull(rec!CHE_OBSERV)
            
            TxtCheNumero.Enabled = False
            TxtBanco.Enabled = False
            txtlocalidad.Enabled = False
            TxtSucursal.Enabled = False
            TxtCodigo.Enabled = False
            
            MtrObjetos = Array(TxtCheNumero, TxtBanco, txtlocalidad, TxtSucursal, TxtCodigo)
            Call CambiarColor(MtrObjetos, 5, &H80000018, "D")
            
        Else
           'TxtCheNombre.ForeColor = &HC0FFFF
           rec.Close
           Exit Sub
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub Txtlocalidad_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtLOCALIDAD_LostFocus()
    If Len(txtlocalidad.Text) < 3 Then txtlocalidad.Text = CompletarConCeros(txtlocalidad.Text, 3)
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtSucursal_LostFocus()
    If Len(TxtSucursal.Text) < 3 Then TxtSucursal.Text = CompletarConCeros(TxtSucursal.Text, 3)
End Sub
