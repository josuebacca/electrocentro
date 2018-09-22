VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Begin VB.Form frmRemitoCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remito de Clientes..."
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7845
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   450
      Left            =   8082
      TabIndex        =   11
      Top             =   7380
      Width           =   990
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   450
      Left            =   7083
      TabIndex        =   10
      Top             =   7380
      Width           =   990
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Enabled         =   0   'False
      Height          =   450
      Left            =   10080
      TabIndex        =   13
      Top             =   7380
      Width           =   990
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Enabled         =   0   'False
      Height          =   450
      Left            =   9081
      TabIndex        =   12
      Top             =   7380
      Width           =   990
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7300
      Left            =   45
      TabIndex        =   27
      Top             =   45
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   12885
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
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
      TabPicture(0)   =   "frmRemitoCliente.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "freRemito"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "freCliente"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "freNotaPedido"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkNotaPedido"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraInfoCliente"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmRemitoCliente.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAum"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "GrdModulos"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "frameBuscar"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtAum"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdDesc"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.CommandButton cmdDesc 
         Caption         =   "- %"
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
         Left            =   -72480
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Porcentaje de descuento"
         Top             =   6960
         Width           =   555
      End
      Begin VB.TextBox txtAum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73680
         TabIndex        =   24
         Text            =   "0,00"
         Top             =   6960
         Width           =   555
      End
      Begin VB.Frame fraInfoCliente 
         Height          =   855
         Left            =   9000
         TabIndex        =   100
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
         Begin VB.TextBox txtAutorizado 
            Height          =   315
            Left            =   1155
            Locked          =   -1  'True
            TabIndex        =   102
            Text            =   "0,00"
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox txtGastado 
            Height          =   315
            Left            =   1155
            Locked          =   -1  'True
            TabIndex        =   101
            Text            =   "0,00"
            Top             =   420
            Width           =   855
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
            Index           =   2
            Left            =   120
            TabIndex        =   104
            Top             =   180
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
            Index           =   1
            Left            =   315
            TabIndex        =   103
            Top             =   480
            Width           =   780
         End
      End
      Begin VB.CheckBox chkNotaPedido 
         Caption         =   "Recupera datos del Presupuesto"
         Height          =   255
         Left            =   5040
         TabIndex        =   4
         Top             =   360
         Width           =   2775
      End
      Begin VB.Frame freNotaPedido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4440
         TabIndex        =   93
         Top             =   370
         Width           =   6630
         Begin VB.CommandButton cmdBuscarNotaPedido 
            Height          =   315
            Left            =   2085
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   95
            ToolTipText     =   "Buscar Nota de Pedido"
            Top             =   220
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtNroNotaPedido 
            Height          =   300
            Left            =   870
            TabIndex        =   94
            Top             =   230
            Width           =   1155
         End
         Begin FechaCtl.Fecha FechaNotaPedido 
            Height          =   285
            Left            =   3390
            TabIndex        =   96
            Top             =   195
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   2775
            TabIndex        =   98
            Top             =   225
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   150
            TabIndex        =   97
            Top             =   225
            Width           =   600
         End
      End
      Begin VB.Frame Frame5 
         Height          =   635
         Left            =   4440
         TabIndex        =   89
         Top             =   2400
         Width           =   6615
         Begin VB.ComboBox cboCondicion 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   91
            Top             =   195
            Width           =   3270
         End
         Begin VB.CommandButton cmdNuevoRubro 
            Height          =   315
            Left            =   5715
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   90
            ToolTipText     =   "Agregar Condición de Venta"
            Top             =   190
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Pago:"
            Height          =   195
            Left            =   1110
            TabIndex        =   92
            Top             =   210
            Width           =   1125
         End
      End
      Begin VB.Frame freCliente 
         Height          =   1455
         Left            =   4410
         TabIndex        =   51
         Top             =   960
         Width           =   6645
         Begin VB.TextBox txtcodpos 
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
            Height          =   285
            Left            =   930
            TabIndex        =   69
            Top             =   810
            Width           =   1215
         End
         Begin VB.TextBox txtlocalidad 
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
            Height          =   285
            Left            =   2250
            MaxLength       =   50
            TabIndex        =   68
            Top             =   810
            Width           =   4260
         End
         Begin VB.TextBox txtIngBrutos 
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
            Height          =   285
            Left            =   5490
            TabIndex        =   59
            Top             =   1130
            Width           =   975
         End
         Begin VB.TextBox txtDomici 
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
            Height          =   285
            Left            =   930
            MaxLength       =   50
            TabIndex        =   58
            Top             =   500
            Width           =   4380
         End
         Begin VB.TextBox txtRazSocCli 
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
            Left            =   2835
            MaxLength       =   50
            TabIndex        =   57
            Tag             =   "Descripción"
            Top             =   160
            Width           =   3630
         End
         Begin VB.TextBox txtCondicionIVA 
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
            Left            =   2415
            TabIndex        =   56
            Top             =   1130
            Width           =   2175
         End
         Begin VB.CommandButton cmdBuscarCliente 
            Height          =   315
            Left            =   1920
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":06CC
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Buscar Cliente"
            Top             =   160
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevoCliente 
            Height          =   315
            Left            =   2385
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":09D6
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Agregar Cliente"
            Top             =   160
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtCUIT 
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
            Left            =   930
            TabIndex        =   53
            Top             =   1130
            Width           =   1455
         End
         Begin VB.TextBox TxtCodigoCli 
            Height          =   300
            Left            =   930
            MaxLength       =   40
            TabIndex        =   5
            Top             =   160
            Width           =   975
         End
         Begin VB.TextBox txtprovincia 
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
            Height          =   285
            Left            =   930
            MaxLength       =   50
            TabIndex        =   52
            Top             =   780
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Ing. Brutos"
            Height          =   195
            Left            =   4650
            TabIndex        =   65
            Top             =   1150
            Width           =   765
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Left            =   255
            TabIndex        =   64
            Top             =   1150
            Width           =   600
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   180
            TabIndex        =   63
            Top             =   530
            Width           =   675
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   330
            TabIndex        =   62
            Top             =   200
            Width           =   525
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   860
            Width           =   735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   150
            TabIndex        =   60
            Top             =   825
            Visible         =   0   'False
            Width           =   705
         End
      End
      Begin VB.Frame freRemito 
         Caption         =   "Remito..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2650
         Left            =   120
         TabIndex        =   29
         Top             =   370
         Width           =   4275
         Begin TabDlg.SSTab tabLista 
            Height          =   1215
            Left            =   120
            TabIndex        =   75
            Top             =   1320
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   2143
            _Version        =   393216
            Tabs            =   2
            TabHeight       =   520
            TabCaption(0)   =   "Maquinarias"
            TabPicture(0)   =   "frmRemitoCliente.frx":0D60
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame2"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Repuestos"
            TabPicture(1)   =   "frmRemitoCliente.frx":0D7C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame4"
            Tab(1).ControlCount=   1
            Begin VB.Frame Frame4 
               Caption         =   "Lista de Precios"
               ForeColor       =   &H8000000D&
               Height          =   735
               Left            =   -74880
               TabIndex        =   78
               Top             =   360
               Width           =   3855
               Begin VB.ComboBox cboLPrecioRep 
                  Height          =   315
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   79
                  Top             =   240
                  Width           =   3585
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Lista de Precios"
               ForeColor       =   &H8000000D&
               Height          =   735
               Left            =   120
               TabIndex        =   76
               Top             =   360
               Width           =   3855
               Begin VB.ComboBox cboListaPrecio 
                  Height          =   315
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   77
                  Top             =   240
                  Width           =   3585
               End
            End
         End
         Begin VB.TextBox txtNroRemito 
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
            Height          =   330
            Left            =   1890
            MaxLength       =   8
            TabIndex        =   1
            Top             =   240
            Width           =   1005
         End
         Begin VB.ComboBox cboStock 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   2505
         End
         Begin VB.TextBox txtNroSucursal 
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
            Height          =   330
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   0
            Top             =   240
            Width           =   555
         End
         Begin FechaCtl.Fecha FechaRemito 
            Height          =   285
            Left            =   1320
            TabIndex        =   2
            Top             =   585
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label lblEstadoRemito 
            AutoSize        =   -1  'True
            Caption         =   "EST. REMITO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1320
            TabIndex        =   66
            Top             =   930
            Width           =   1215
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Stock:"
            Height          =   210
            Left            =   675
            TabIndex        =   49
            Top             =   120
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   765
            TabIndex        =   47
            Top             =   585
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   660
            TabIndex        =   46
            Top             =   285
            Width           =   600
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   720
            TabIndex        =   45
            Top             =   915
            Width           =   540
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
         Height          =   1950
         Left            =   -74600
         TabIndex        =   36
         Top             =   540
         Width           =   10410
         Begin FechaCtl.Fecha FechaHasta 
            Height          =   285
            Left            =   5745
            TabIndex        =   21
            Top             =   1080
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha FechaDesde 
            Height          =   330
            Left            =   3240
            TabIndex        =   20
            Top             =   1080
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Frame Frame1 
            Caption         =   "Estado Remito"
            Height          =   495
            Left            =   840
            TabIndex        =   70
            Top             =   1320
            Width           =   8535
            Begin VB.OptionButton optFac 
               Caption         =   "Facturados"
               Height          =   195
               Left            =   5400
               TabIndex        =   105
               Top             =   200
               Width           =   1335
            End
            Begin VB.OptionButton optTod 
               Caption         =   "Todos"
               Height          =   195
               Left            =   6840
               TabIndex        =   74
               Top             =   200
               Width           =   1215
            End
            Begin VB.OptionButton optAnu 
               Caption         =   "Anulados"
               Height          =   195
               Left            =   3960
               TabIndex        =   73
               Top             =   200
               Width           =   1455
            End
            Begin VB.OptionButton optDef 
               Caption         =   "Definitivos"
               Height          =   195
               Left            =   2640
               TabIndex        =   72
               Top             =   200
               Width           =   1455
            End
            Begin VB.OptionButton optPen 
               Caption         =   "Pendientes"
               Height          =   195
               Left            =   1200
               TabIndex        =   71
               Top             =   200
               Value           =   -1  'True
               Width           =   1575
            End
         End
         Begin VB.CommandButton cmdBuscarVendedor 
            Height          =   315
            Left            =   4290
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":0D98
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Buscar Vendedor"
            Top             =   660
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarCli 
            Height          =   315
            Left            =   4290
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":10A2
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Buscar Cliente"
            Top             =   255
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtVendedor 
            Height          =   300
            Left            =   3240
            TabIndex        =   19
            Top             =   667
            Width           =   990
         End
         Begin VB.TextBox txtDesVen 
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
            Height          =   285
            Left            =   4725
            TabIndex        =   41
            Top             =   675
            Width           =   4620
         End
         Begin VB.CheckBox chkVendedor 
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   900
            TabIndex        =   16
            Top             =   645
            Width           =   1035
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1395
            Left            =   9660
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":13AC
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Buscar "
            Top             =   350
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin VB.TextBox txtDesCli 
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
            Left            =   4725
            MaxLength       =   50
            TabIndex        =   37
            Tag             =   "Descripción"
            Top             =   255
            Width           =   4620
         End
         Begin VB.TextBox txtCliente 
            Height          =   300
            Left            =   3240
            MaxLength       =   40
            TabIndex        =   18
            Top             =   255
            Width           =   975
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   900
            TabIndex        =   17
            Top             =   975
            Width           =   810
         End
         Begin VB.CheckBox chkCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   878
            TabIndex        =   15
            Top             =   315
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Index           =   0
            Left            =   2415
            TabIndex        =   42
            Top             =   712
            Width           =   735
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4695
            TabIndex        =   40
            Top             =   1140
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   2145
            TabIndex        =   39
            Top             =   1125
            Width           =   1005
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   2625
            TabIndex        =   38
            Top             =   300
            Width           =   525
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4260
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "Actualiza los precios del Remito"
         Top             =   2970
         Width           =   10935
         Begin VB.CheckBox chkPorc 
            Caption         =   "Porcentaje"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   3525
            Width           =   1095
         End
         Begin VB.TextBox txtPorc 
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
            Height          =   285
            Left            =   1275
            TabIndex        =   8
            Top             =   3510
            Width           =   555
         End
         Begin VB.TextBox txtTotal 
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
            Height          =   285
            Left            =   8940
            TabIndex        =   84
            Top             =   3510
            Width           =   1350
         End
         Begin VB.TextBox txtPorcentajeIva 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5025
            TabIndex        =   83
            Top             =   3510
            Width           =   1155
         End
         Begin VB.TextBox txtImporteIva 
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
            Height          =   285
            Left            =   6990
            TabIndex        =   82
            Top             =   3510
            Width           =   1155
         End
         Begin VB.TextBox txtSubTotal 
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
            Height          =   285
            Left            =   2835
            TabIndex        =   81
            Top             =   3510
            Width           =   1155
         End
         Begin VB.CommandButton cmdredondeo 
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10410
            TabIndex        =   80
            Top             =   3780
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton cmdPrecio 
            Caption         =   "$"
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
            Left            =   10395
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Actualizar Precios"
            Top             =   1320
            Width           =   390
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   240
            TabIndex        =   31
            Top             =   525
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1455
            MaxLength       =   60
            TabIndex        =   14
            Top             =   3870
            Width           =   8850
         End
         Begin VB.CommandButton cmdBuscarProducto 
            Height          =   330
            Left            =   10395
            MaskColor       =   &H8000000F&
            Picture         =   "frmRemitoCliente.frx":3B4E
            Style           =   1  'Graphical
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Buscar Producto"
            Top             =   195
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdAgregarProducto 
            Height          =   330
            Left            =   10395
            MaskColor       =   &H8000000F&
            Picture         =   "frmRemitoCliente.frx":3E58
            Style           =   1  'Graphical
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "Agregar Producto"
            Top             =   570
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdQuitarProducto 
            Height          =   330
            Left            =   10395
            MaskColor       =   &H8000000F&
            Picture         =   "frmRemitoCliente.frx":4162
            Style           =   1  'Graphical
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Producto"
            Top             =   945
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   3375
            Left            =   90
            TabIndex        =   6
            Top             =   165
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   5953
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
            AllowUserResizing=   3
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
         Begin VB.Label Label16 
            Caption         =   "Total:"
            Height          =   195
            Left            =   8475
            TabIndex        =   88
            Top             =   3555
            Width           =   405
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "% I.V.A.:"
            Height          =   195
            Left            =   4320
            TabIndex        =   87
            Top             =   3555
            Width           =   600
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   6360
            TabIndex        =   86
            Top             =   3555
            Width           =   570
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Total:"
            Height          =   195
            Left            =   2040
            TabIndex        =   85
            Top             =   3555
            Width           =   735
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   210
            TabIndex        =   48
            Top             =   3915
            Width           =   1110
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4170
         Left            =   -74640
         TabIndex        =   23
         Top             =   2700
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7355
         _Version        =   393216
         Cols            =   14
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.CommandButton cmdAum 
         Caption         =   "+ %"
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
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Porcentaje de aumento"
         Top             =   6960
         Width           =   555
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "% Porcentaje"
         Height          =   195
         Left            =   -74640
         TabIndex        =   106
         Top             =   7005
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   28
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   6084
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7380
      Width           =   990
   End
   Begin VB.CommandButton cmdFacturar 
      Caption         =   "&Facturar"
      Height          =   450
      Left            =   5085
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   7380
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "<F1> Buscar Remito"
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
      Index           =   1
      Left            =   2640
      TabIndex        =   67
      Top             =   7440
      Width           =   2085
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
      Height          =   255
      Left            =   150
      TabIndex        =   44
      Top             =   7455
      Width           =   750
   End
End
Attribute VB_Name = "frmRemitoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim w As Integer
Dim TipoBusquedaDoc As Integer
Dim VEstadoRemito As Integer
Dim VCantidadBultos As Integer
Dim Rec1 As ADODB.Recordset
Public nlista As Integer
Private Sub chkCliente_Click()
    If chkCliente.Value = Checked Then
        txtCliente.Enabled = True
        cmdBuscarCli.Enabled = True
    Else
        txtCliente.Enabled = False
        cmdBuscarCli.Enabled = False
    End If
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

Private Sub chkRemitoSinFactura_Click()
   ' If chkRemitoSinFactura.Value = Checked Then
   '     txtConcepto.Enabled = True
   ' Else
   '     txtConcepto.Enabled = False
   ' End If
End Sub

Private Sub chkNotaPedido_Click()
    If chkNotaPedido.Value = 1 Then
        freNotaPedido.Enabled = True
        txtNroNotaPedido.Enabled = True
        txtNroNotaPedido.SetFocus
        freCliente.Enabled = False
    Else
        freNotaPedido.Enabled = False
        freCliente.Enabled = True
        TxtCodigoCli.SetFocus
    End If
End Sub

Private Sub chkPorc_Click()
    Dim Porcentaje As Double

    If chkPorc.Value = 1 Then
        'Porcentaje = txtPorc.Text
        txtPorc.Enabled = True
        txtPorc.SetFocus
    Else
        Porcentaje = IIf(txtPorc.Text = "", 0, XN(txtPorc.Text))
        txtPorc.Enabled = False
        txtPorc.Text = ""
        If Porcentaje <> 0 Then
            TotalPEDIDO
'            Dim TOTAL As Double
'            Dim subtotal As Double
'            Dim impIva As Double
'
'            txtSubtotal.Text = Valido_Importe(SumaTotal)
'            txtSubtotal.Text = txtSubtotal.Text - (txtSubtotal.Text * Porcentaje) / 100
'            'txtImporteIva.Text = (CDbl(txtSubTotal.Text) * CDbl(txtPorcentajeIva.Text)) / 100
'            txtImporteIva.Text = CDbl(txtSubtotal.Text) / CDbl("1," & Int(txtPorcentajeIva.Text))
'            txtSubtotal.Text = Valido_Importe(SumaTotal - txtImporteIva.Text)
'            txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)
'            subtotal = txtSubtotal
'            impIva = txtImporteIva
'            TOTAL = subtotal + impIva
'            txtTotal.Text = Format(TOTAL, "0.00")
        End If
    End If
End Sub

Private Sub chkVendedor_Click()
    If chkVendedor.Value = Checked Then
        txtVendedor.Enabled = True
        cmdBuscarVendedor.Enabled = True
    Else
        txtVendedor.Enabled = False
        cmdBuscarVendedor.Enabled = False
    End If
End Sub

Private Sub cmdAgregarProducto_Click()
    Consulta = 3
    'ABMProducto.CODIGOLISTA = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
    ABMProducto.Show vbModal
    If Consulta <> 4 Then
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        If Trim(ABMProducto.TxtCodigo) <> "" Then txtEdit.Text = ABMProducto.TxtCodigo
        TxtEdit_KeyDown vbKeyReturn, 0
    End If
    'grdGrilla.SetFocus
    'grdGrilla.row = 1
End Sub
Function BuscoImporte(nRemito As Integer, nSucursal As Integer) As Double
    Dim nsubtotal As Double
    Dim ntotal As Double
    Dim nIVA As Double
    ntotal = 0
    sql = "SELECT DR.DRC_PRECIO,DR.DRC_CANTIDAD,P.LNA_CODIGO "
    sql = sql & " FROM DETALLE_REMITO_CLIENTE DR,PRODUCTO P "
    sql = sql & " WHERE "
    sql = sql & " DR.PTO_CODIGO = P.PTO_CODIGO"
    sql = sql & " AND DR.RCL_NUMERO =" & nRemito
    sql = sql & " AND DR.RCL_SUCURSAL =" & nSucursal
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        ' me fijo si es maquinaria o repuesto
     '   If Rec1!LNA_CODIGO = 6 Then
      '      nIVA = "1,105"
      '  Else
      '      nIVA = "1,21"
      '  End If
        Do While Rec1.EOF = False
            
            nsubtotal = IIf(IsNull(Rec1!DRC_CANTIDAD), 1, Rec1!DRC_CANTIDAD) * Rec1!DRC_PRECIO
            ntotal = ntotal + nsubtotal
            Rec1.MoveNext
        Loop
    End If
    'ntotal = ntotal * nIVA
    Rec1.Close
    BuscoImporte = ntotal
End Function

Private Sub cmdAum_Click()
    Dim VTotal As String
    Dim VSaldo As String
    Dim vNuevoPrecio As Double
    If GrdModulos.Rows > 1 And txtAum.Text <> "0,00" Then
        If MsgBox("¿Confirma el aumento del % " & txtAum.Text & " en los Remitos seleccionados?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    Else
        Exit Sub
    End If
    
    For I = 1 To GrdModulos.Rows - 1
        'continuar aca
        sql = "SELECT PTO_CODIGO,DRC_PRECIO,DRC_CANTIDAD FROM DETALLE_REMITO_CLIENTE "
        sql = sql & " WHERE RCL_NUMERO = " & Right(GrdModulos.TextMatrix(I, 0), 8)
        sql = sql & " AND RCL_SUCURSAL = " & Left(GrdModulos.TextMatrix(I, 0), 4)
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        VTotal = 0
        VSaldo = 0
        If Rec1.EOF = False Then
            Do While Rec1.EOF = False
                sql = "UPDATE DETALLE_REMITO_CLIENTE SET "
                sql = sql & " DRC_PRECIO = " & XN(CDbl(Rec1!DRC_PRECIO) + (CDbl(Rec1!DRC_PRECIO) * CDbl(txtAum.Text)) / 100)
                sql = sql & " WHERE RCL_NUMERO = " & XN(Right(GrdModulos.TextMatrix(I, 0), 8))
                sql = sql & " AND RCL_SUCURSAL = " & XN(Left(GrdModulos.TextMatrix(I, 0), 4))
                sql = sql & " AND PTO_CODIGO = " & Rec1!PTO_CODIGO
                DBConn.Execute sql
                vNuevoPrecio = CDbl(Rec1!DRC_PRECIO) + ((CDbl(Rec1!DRC_PRECIO) * CDbl(txtAum.Text)) / 100)
                vNuevoPrecio = Format(vNuevoPrecio, "0.00")
                VTotal = CDbl(VTotal) + (Rec1!DRC_CANTIDAD * vNuevoPrecio)
                VSaldo = CDbl(VSaldo) + (Rec1!DRC_CANTIDAD * vNuevoPrecio)
                Rec1.MoveNext
            Loop
        End If
        Rec1.Close
                
        sql = "SELECT RCL_TOTAL,RCL_SALDO FROM REMITO_CLIENTE "
        sql = sql & " WHERE RCL_NUMERO = " & Right(GrdModulos.TextMatrix(I, 0), 8)
        sql = sql & " AND RCL_SUCURSAL = " & Left(GrdModulos.TextMatrix(I, 0), 4)
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            
            sql = "UPDATE REMITO_CLIENTE SET "
            If Rec1!RCL_SALDO = Rec1!RCL_TOTAL Then
                sql = sql & " RCL_SALDO = " & XN(VSaldo)
            Else
                sql = sql & "RCL_SALDO = " & XN(CDbl(Rec1!RCL_SALDO) + CDbl(Rec1!RCL_SALDO) * CDbl(txtAum.Text) / 100)
            End If
            sql = sql & ",RCL_TOTAL = " & XN(VTotal)
            
            sql = sql & " WHERE RCL_NUMERO = " & Right(GrdModulos.TextMatrix(I, 0), 8)
            sql = sql & " AND RCL_SUCURSAL = " & Left(GrdModulos.TextMatrix(I, 0), 4)
            DBConn.Execute sql
        End If
        Rec1.Close
        'txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 4)
        'txtNroRemito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8)
'        If grdGrilla.TextMatrix(I, 0) <> "" Then
'            grdGrilla.TextMatrix(I, 3) = CDbl(grdGrilla.TextMatrix(I, 3)) + (CDbl(grdGrilla.TextMatrix(I, 3)) * CDbl(txtAum)) / 100
'            grdGrilla.TextMatrix(I, 3) = Valido_Importe(grdGrilla.TextMatrix(I, 3))
'            grdGrilla.TextMatrix(I, 4) = CDbl(grdGrilla.TextMatrix(I, 3)) * CDbl(grdGrilla.TextMatrix(I, 2))
'            grdGrilla.TextMatrix(I, 4) = Valido_Importe(grdGrilla.TextMatrix(I, 4))
'        End If
    Next
    'TotalPEDIDO
    CmdBuscAprox_Click
End Sub

Private Sub CmdBuscAprox_Click()
    Dim TotalRemito As Double
    Dim TOTAL As Double
    GrdModulos.Rows = 1
    lblestado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
'    If (chkPend.Value = Unchecked) And (chkDef.Value = Unchecked) And (chkAnu.Value = Unchecked) Then
'        MsgBox "Debe Seleccionar un Estado del Remito", vbInformation
'        chkPend.SetFocus
'    End If
    Select Case TipoBusquedaDoc
    
    
    Case 1 'BUSCA REMITOS
        
        sql = "SELECT TOP 100 RC.RCL_NUMERO,RC.RCL_SUCURSAL,RC.RCL_FECHA,RC.RCL_TOTAL,RC.CLI_CODIGO,"
        sql = sql & " RC.NPE_NUMERO,RC.NPE_FECHA,RCL_OBSERVACION,RC.RCL_SINFAC,RC.STK_CODIGO,RC.EST_CODIGO,"
        sql = sql & " C.CLI_RAZSOC,C.CLI_DOMICI,L.LOC_DESCRI,P.PRO_DESCRI,RC.RCL_SALDO"
        sql = sql & " FROM REMITO_CLIENTE RC,CLIENTE C, LOCALIDAD L, PROVINCIA P"
        sql = sql & " WHERE"
        sql = sql & "  RC.CLI_CODIGO=C.CLI_CODIGO"
        sql = sql & "  AND C.LOC_CODIGO=L.LOC_CODIGO"
        sql = sql & "  AND C.PRO_CODIGO=P.PRO_CODIGO"
        sql = sql & "  AND L.PRO_CODIGO=P.PRO_CODIGO"
        
        If txtCliente.Text <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(txtCliente)
        If FechaDesde <> "" Then sql = sql & " AND RC.RCL_FECHA>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND RC.RCL_FECHA<=" & XDQ(FechaHasta)
        If optPen.Value = True Then
            sql = sql & " AND (RC.EST_CODIGO = 1 or RC.RCL_SALDO > 0)"
        End If
        If optDef.Value = True Then
            sql = sql & " AND RC.EST_CODIGO = 3 "
        End If
        If optAnu.Value = True Then
            sql = sql & " AND RC.EST_CODIGO = 2 "
        End If
        If optFac.Value = True Then
            sql = sql & " AND RC.EST_CODIGO = 4 "
        End If
        sql = sql & " ORDER BY RC.RCL_FECHA DESC"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        TOTAL = 0
        If rec.EOF = False Then
            Do While rec.EOF = False
                TotalRemito = BuscoImporte(rec!RCL_NUMERO, rec!RCL_SUCURSAL)
                GrdModulos.AddItem Format(rec!RCL_SUCURSAL, "0000") & "-" & Format(rec!RCL_NUMERO, "00000000") _
                                & Chr(9) & rec!RCL_FECHA _
                                & Chr(9) & Format(rec!RCL_TOTAL, "##0.00") _
                                & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!CLI_DOMICI _
                                & Chr(9) & rec!LOC_DESCRI & Chr(9) & rec!PRO_DESCRI _
                                & Chr(9) & rec!EST_CODIGO _
                                & Chr(9) & rec!NPE_NUMERO & Chr(9) & rec!NPE_FECHA _
                                & Chr(9) & rec!RCL_OBSERVACION & Chr(9) & rec!STK_CODIGO _
                                & Chr(9) & rec!RCL_SINFAC & Chr(9) & rec!CLI_CODIGO
                
                TOTAL = TOTAL + IIf(IsNull(rec!RCL_SALDO), 0, rec!RCL_SALDO)
                rec.MoveNext
            Loop
            GrdModulos.SetFocus
        Else
            lblestado.Caption = ""
            Screen.MousePointer = vbNormal
            MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
        End If
        
        GrdModulos.ToolTipText = Format(TOTAL, "00.00")
        
    Case 2 'BUSCA NOTA DE PEDIDO - PRESUPUESTO
        
        sql = "SELECT NP.NPE_NUMERO,NP.NPE_TOTAL, NP.NPE_FECHA, C.CLI_RAZSOC, "
        sql = sql & " C.CLI_CODIGO,C.CLI_DOMICI,L.LOC_DESCRI,P.PRO_DESCRI"
        sql = sql & " FROM NOTA_PEDIDO NP, CLIENTE C, LOCALIDAD L, PROVINCIA P"
        sql = sql & " WHERE"
        sql = sql & " NP.CLI_CODIGO=C.CLI_CODIGO"
        sql = sql & " AND L.LOC_CODIGO=C.LOC_CODIGO"
        sql = sql & " AND P.PRO_CODIGO=C.PRO_CODIGO"
        sql = sql & " AND P.PRO_CODIGO=L.PRO_CODIGO"
        'sql = sql & " AND NP.EST_CODIGO = 1"
        If txtCliente.Text <> "" Then sql = sql & " AND NP.CLI_CODIGO=" & XN(txtCliente)
        If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
        If FechaDesde <> "" Then sql = sql & " AND NP.NPE_FECHA>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND NP.NPE_FECHA<=" & XDQ(FechaHasta)
        If optPen.Value = True Then
            sql = sql & " AND NP.EST_CODIGO = 1 "
        End If
        If optDef.Value = True Then
            sql = sql & " AND NP.EST_CODIGO = 3 "
        End If
        If optAnu.Value = True Then
            sql = sql & " AND NP.EST_CODIGO = 2 "
        End If
        sql = sql & " ORDER BY NPE_FECHA"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                GrdModulos.AddItem rec!NPE_NUMERO & Chr(9) & rec!NPE_FECHA & Chr(9) & rec!NPE_TOTAL _
                                & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!CLI_DOMICI _
                                & Chr(9) & rec!LOC_DESCRI & Chr(9) & rec!PRO_DESCRI _
                                & Chr(9) & rec!CLI_CODIGO
                rec.MoveNext
            Loop
            GrdModulos.SetFocus
        Else
            lblestado.Caption = ""
            Screen.MousePointer = vbNormal
            MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
        End If
    End Select
    
    lblestado.Caption = ""
    Screen.MousePointer = vbNormal
    rec.Close
End Sub

Private Sub cmdBuscarCli_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtCliente.Text = frmBuscar.grdBuscar.Text
        txtCliente.SetFocus
        txtCliente_LostFocus
    Else
        txtCliente.SetFocus
    End If
End Sub



Private Sub cmdBuscarCliente_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB.Text = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        TxtCodigoCli.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 1
        txtRazSocCli.Text = frmBuscar.grdBuscar.Text
        TxtCodigoCli_LostFocus
    Else
        TxtCodigoCli.SetFocus
    End If
End Sub

Private Sub cmdBuscarNotaPedido_Click()
    TipoBusquedaDoc = 2
    tabDatos.Tab = 1
End Sub

Private Sub cmdBuscarProducto_Click()
'    grdGrilla.SetFocus
'    frmBuscar.TipoBusqueda = 2
'    frmBuscar.CodListaPrecio = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
'    frmBuscar.TxtDescriB.Text = ""
'    frmBuscar.Show vbModal
'    grdGrilla.Col = 0
'    EDITAR grdGrilla, txtEdit, 13
'    If Trim(frmBuscar.grdBuscar.Text) <> "" Then txtEdit.Text = frmBuscar.grdBuscar.Text
'    TxtEdit_KeyDown vbKeyReturn, 0
    Consulta = 3
    'FrmListadePrecios.cbodescri.ListIndex = cboListaPrecio.ListIndex
    
    If tabLista.Tab = 0 Then
        FrmListadePrecios.tabLista.Tab = 0
        FrmListadePrecios.cboListaPrecio.ListIndex = cboListaPrecio.ListIndex
        
    Else
        nlista = 1
        FrmListadePrecios.tabLista.Tab = 1
        FrmListadePrecios.cboLPrecioRep.ListIndex = cboLPrecioRep.ListIndex
    End If
    
    FrmListadePrecios.Show vbModal
    If Consulta <> 4 Then
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        If Trim(FrmListadePrecios.GrdModulos.Text) <> "" Then txtEdit.Text = FrmListadePrecios.GrdModulos.Text
        TxtEdit_KeyDown vbKeyReturn, 0
    End If
    Consulta = 2
End Sub

Private Sub cmdBuscarVendedor_Click()
    frmBuscar.TipoBusqueda = 4
    frmBuscar.TxtDescriB.Text = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtVendedor.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 1
        txtDesVen.Text = frmBuscar.grdBuscar.Text
        txtVendedor.SetFocus
    Else
        txtVendedor.SetFocus
    End If
End Sub

Private Sub CmdCancelar_Click()
    If MsgBox("Esta acccion Cancelara los cambios intorducidos ¿Confirma Cancelar?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    cmdNuevo.Enabled = True
    cmdImprimir.Enabled = True
    CmdSalir.Enabled = True
    CmdCancelar.Enabled = False
    txtNroRemito.Text = ""
End Sub

Private Sub cmdDesc_Click()
Dim VTotal As String
    Dim VSaldo As String
    Dim vNuevoPrecio As Double
    If GrdModulos.Rows > 1 And txtAum.Text <> "0,00" Then
        If MsgBox("¿Confirma el descuento del % " & txtAum.Text & " en los Remitos seleccionados?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    Else
        Exit Sub
    End If
    
    For I = 1 To GrdModulos.Rows - 1
        'continuar aca
        sql = "SELECT PTO_CODIGO,DRC_PRECIO,DRC_CANTIDAD FROM DETALLE_REMITO_CLIENTE "
        sql = sql & " WHERE RCL_NUMERO = " & Right(GrdModulos.TextMatrix(I, 0), 8)
        sql = sql & " AND RCL_SUCURSAL = " & Left(GrdModulos.TextMatrix(I, 0), 4)
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        VTotal = 0
        VSaldo = 0
        If Rec1.EOF = False Then
            Do While Rec1.EOF = False
                sql = "UPDATE DETALLE_REMITO_CLIENTE SET "
                sql = sql & " DRC_PRECIO = " & XN(CDbl(Rec1!DRC_PRECIO) - (CDbl(Rec1!DRC_PRECIO) * CDbl(txtAum.Text)) / 100)
                sql = sql & " WHERE RCL_NUMERO = " & XN(Right(GrdModulos.TextMatrix(I, 0), 8))
                sql = sql & " AND RCL_SUCURSAL = " & XN(Left(GrdModulos.TextMatrix(I, 0), 4))
                sql = sql & " AND PTO_CODIGO = " & Rec1!PTO_CODIGO
                DBConn.Execute sql
                vNuevoPrecio = CDbl(Rec1!DRC_PRECIO) - ((CDbl(Rec1!DRC_PRECIO) * CDbl(txtAum.Text)) / 100)
                vNuevoPrecio = Format(vNuevoPrecio, "0.00")
                VTotal = CDbl(VTotal) + (Rec1!DRC_CANTIDAD * vNuevoPrecio)
                VSaldo = CDbl(VSaldo) + (Rec1!DRC_CANTIDAD * vNuevoPrecio)
                Rec1.MoveNext
            Loop
        End If
        Rec1.Close
                
        sql = "SELECT RCL_TOTAL,RCL_SALDO FROM REMITO_CLIENTE "
        sql = sql & " WHERE RCL_NUMERO = " & Right(GrdModulos.TextMatrix(I, 0), 8)
        sql = sql & " AND RCL_SUCURSAL = " & Left(GrdModulos.TextMatrix(I, 0), 4)
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            
            sql = "UPDATE REMITO_CLIENTE SET "
            If Rec1!RCL_SALDO = Rec1!RCL_TOTAL Then
                sql = sql & " RCL_SALDO = " & XN(VSaldo)
            Else
                sql = sql & "RCL_SALDO = " & XN(CDbl(Rec1!RCL_SALDO) - CDbl(Rec1!RCL_SALDO) * CDbl(txtAum.Text) / 100)
            End If
            sql = sql & ",RCL_TOTAL = " & XN(VTotal)
            
            sql = sql & " WHERE RCL_NUMERO = " & Right(GrdModulos.TextMatrix(I, 0), 8)
            sql = sql & " AND RCL_SUCURSAL = " & Left(GrdModulos.TextMatrix(I, 0), 4)
            DBConn.Execute sql
        End If
        Rec1.Close
        'txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 4)
        'txtNroRemito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8)
'        If grdGrilla.TextMatrix(I, 0) <> "" Then
'            grdGrilla.TextMatrix(I, 3) = CDbl(grdGrilla.TextMatrix(I, 3)) + (CDbl(grdGrilla.TextMatrix(I, 3)) * CDbl(txtAum)) / 100
'            grdGrilla.TextMatrix(I, 3) = Valido_Importe(grdGrilla.TextMatrix(I, 3))
'            grdGrilla.TextMatrix(I, 4) = CDbl(grdGrilla.TextMatrix(I, 3)) * CDbl(grdGrilla.TextMatrix(I, 2))
'            grdGrilla.TextMatrix(I, 4) = Valido_Importe(grdGrilla.TextMatrix(I, 4))
'        End If
    Next
    'TotalPEDIDO
    CmdBuscAprox_Click
End Sub

Private Sub cmdFacturar_Click()
    If MsgBox("¿Confirma Facturar el Remito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    sql = "UPDATE REMITO_CLIENTE "
    sql = sql & "SET EST_CODIGO = 4 "
   ' sql = sql & ", RCL_SALDO = 0"
    sql = sql & " WHERE RCL_NUMERO = " & XN(txtNroRemito.Text)
    sql = sql & " AND RCL_SUCURSAL =" & XN(txtNroSucursal.Text)
    
    DBConn.Execute sql
    'AGREGAR PAGO EN CTA CTE
    'HABLAR ESTO CON NICOLAS!
 '   DBConn.Execute AgregoCtaCteCliente(TxtCodigoCli, 20 _
                                            , txtNroRemito, txtNroSucursal, _
                                            FechaRemito, txtTotal, "H", CStr(Date))
    
    CmdNuevo_Click
End Sub

Private Sub CmdGrabar_Click()
    
    On Error GoTo HayErrorRemito
    If ValidarRemito = False Then Exit Sub
    If MsgBox("¿Confirma Remito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
      
    DBConn.BeginTrans
    
    
    'sql = "SELECT * FROM REMITO_CLIENTE"
    'sql = sql & " WHERE RCL_NUMERO=" & XN(txtNroRemito)
    'sql = sql & " AND RCL_SUCURSAL=" & XN(txtNroSucursal)
    'rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            
    Screen.MousePointer = vbHourglass
    lblestado.Caption = "Guardando..."
    
    If txtNroRemito.Text = "" Then
        'BUSCO EL NUMERO DE REMITO QUE CORRESPONDE
        txtNroRemito.Text = BuscoUltimoRenito
        
        'If rec.EOF = True Then 'NUEVO REMITO
            sql = "INSERT INTO REMITO_CLIENTE"
            sql = sql & " (RCL_NUMERO,RCL_SUCURSAL,RCL_FECHA,NPE_NUMERO,"
            sql = sql & "NPE_FECHA,RCL_OBSERVACION,"
            sql = sql & "EST_CODIGO,RCL_NUMEROTXT,STK_CODIGO, CLI_CODIGO,FPG_CODIGO,"
            sql = sql & "TCO_CODIGO,RCL_TOTAL,RCL_SUBTOTAL,RCL_SALDO,RCL_PORCEN)"
            sql = sql & " VALUES ("
            sql = sql & XN(txtNroRemito) & ","
            sql = sql & XN(txtNroSucursal) & ","
            sql = sql & XDQ(FechaRemito) & ","
            sql = sql & XN(txtNroNotaPedido) & ","
            sql = sql & XDQ(FechaNotaPedido) & ","
            sql = sql & XS(txtObservaciones) & ","
            sql = sql & "1,"    'ESTADO PENDIENTE
            sql = sql & XS(Format(txtNroRemito.Text, "00000000")) & ","
            'sql = sql & cboStock.ItemData(cboStock.ListIndex) & ","
            sql = sql & 1 & ","
            sql = sql & XN(TxtCodigoCli.Text) & ","
            sql = sql & cboCondicion.ItemData(cboCondicion.ListIndex) & ","
            sql = sql & 20 & ","
            sql = sql & XN(txtTotal.Text) & ","
            sql = sql & XN(txtSubtotal.Text) & ","
                If (TxtCodigoCli.Text = "2") Or (cboCondicion.ItemData(cboCondicion.ListIndex) = 1) Then
                sql = sql & 0 & ","
            Else
                sql = sql & XN(txtTotal.Text) & ","
            End If
            If txtPorc.Text = "" Then
                sql = sql & 0 & ")"
            Else
                sql = sql & XN(txtPorc.Text) & ")"
            End If
            'sql = sql & XN(txtAum.Text) & ")"
            
            DBConn.Execute sql
               
            For I = 1 To grdGrilla.Rows - 1
                If grdGrilla.TextMatrix(I, 0) <> "" Then
                    sql = "INSERT INTO DETALLE_REMITO_CLIENTE"
                    sql = sql & " (RCL_NUMERO,RCL_SUCURSAL,RCL_FECHA,DRC_NROITEM,"
                    sql = sql & "PTO_CODIGO,DRC_CANTIDAD,DRC_PRECIO,DRC_DETALLE)"
                    sql = sql & " VALUES ("
                    sql = sql & XN(txtNroRemito) & ","
                    sql = sql & XN(txtNroSucursal) & ","
                    sql = sql & XDQ(FechaRemito) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(I, 7)) & ","
                    If grdGrilla.TextMatrix(I, 0) <> "----------" Then
                        sql = sql & XN(grdGrilla.TextMatrix(I, 0), True) & ","
                    Else
                        sql = sql & "99999999" & ","
                    End If
                    sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & ","
                    grdGrilla.TextMatrix(I, 1) = Replace(grdGrilla.TextMatrix(I, 1), "'", "´")
                    sql = sql & XS(grdGrilla.TextMatrix(I, 1)) & ")"
                    DBConn.Execute sql
                End If
            Next
            'ACTUALIZO EL STOCK CUANDO EL REMITO ES DEFINITIVO (STOCK PENDIENTE)
            'Y ES REMITO SIN FACTURAS
           ' If chkRemitoSinFactura.Value = Checked Then
                For I = 1 To grdGrilla.Rows - 1
                    If grdGrilla.TextMatrix(I, 0) <> "" Then
                            sql = "UPDATE DETALLE_STOCK"
                            sql = sql & " SET"
                            'sql = sql & " DST_STKPEN = DST_STKPEN + " & XN(grdGrilla.TextMatrix(I, 2))
                            sql = sql & " DST_STKFIS = DST_STKFIS - " & XN(grdGrilla.TextMatrix(I, 2))
                            sql = sql & " WHERE STK_CODIGO = 1 "
                            '& cboStock.ItemData(cboStock.ListIndex)
                            sql = sql & " AND PTO_CODIGO LIKE '" & grdGrilla.TextMatrix(I, 0) & "' "
                            DBConn.Execute sql
                    End If
                Next
           ' End If
            'CAMBIO ESTADO DE LA NOTA DE PEDIDO - PRESUPUESTO (LE PONGO DEFINITIVO)
            If txtNroNotaPedido.Text <> "" Then
                sql = "UPDATE NOTA_PEDIDO SET EST_CODIGO=3"
                sql = sql & " WHERE"
                sql = sql & " NPE_NUMERO=" & XN(txtNroNotaPedido)
                sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
                DBConn.Execute sql
            End If
            
            'ACTUALIZO LA TABLA PARAMENTROS Y LE SUMO UNO AL REMITO
            sql = "UPDATE PARAMETROS SET NRO_REMITO=" & XN(txtNroRemito)
            DBConn.Execute sql
            
            If TxtCodigoCli.Text <> 2 Then 'Cliente Consumidor Final
                'ACTUALIZO LA CUENTA CORRIENTE DEL CLIENTE
                DBConn.Execute AgregoCtaCteCliente(TxtCodigoCli, 20 _
                                                , txtNroRemito, txtNroSucursal, _
                                                FechaRemito, txtTotal, "D", CStr(Date))
            Else
                'ACTUALIZO SALDO DE LA REMITO DEL CLIENTE
                
            End If
            
            DBConn.CommitTrans
        Else
            ' modifico el Remito
            'If MsgBox("Confirma modificar el Remito Nro.: " & Trim(txtNroSucursal.Text) & "-" & Trim(txtNroNotaPedido.Text) & " ", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
                sql = "UPDATE REMITO_CLIENTE"
                sql = sql & " SET CLI_CODIGO=" & XN(TxtCodigoCli)
                sql = sql & " ,RCL_OBSERVACION=" & XS(txtObservaciones)
                sql = sql & " ,STK_CODIGO= 1"
                sql = sql & " ,RCL_NUMEROTXT=" & XS(Format(txtNroRemito.Text, "00000000"))
                sql = sql & " ,FPG_CODIGO =" & cboCondicion.ItemData(cboCondicion.ListIndex)
                sql = sql & " ,TCO_CODIGO= 20 "
                sql = sql & " ,RCL_TOTAL=" & XN(txtTotal.Text)
                sql = sql & " ,RCL_SUBTOTAL=" & XN(txtSubtotal.Text)
                If (TxtCodigoCli.Text = "2") Or (cboCondicion.ItemData(cboCondicion.ListIndex) = 1) Then
                    sql = sql & " ,RCL_SALDO = 0 "
                Else
                    sql = sql & " ,RCL_SALDO=" & XS(txtTotal.Text)
                End If
                If txtPorc.Text = "" Then
                    sql = sql & " ,RCL_PORCEN = 0 "
                Else
                    sql = sql & " ,RCL_PORCEN=" & XS(txtPorc.Text)
                End If
                'sql = sql & " ,RCL_AUM=" & XN(txtAum.Text)
                sql = sql & " WHERE"
                sql = sql & " RCL_NUMERO=" & XN(txtNroRemito)
                sql = sql & " AND RCL_FECHA=" & XDQ(FechaRemito)
                DBConn.Execute sql
                
                sql = "DELETE FROM DETALLE_REMITO_CLIENTE"
                sql = sql & " WHERE RCL_NUMERO=" & XN(txtNroRemito)
                sql = sql & " AND RCL_SUCURSAL=" & XN(txtNroSucursal)
                sql = sql & " AND RCL_FECHA=" & XDQ(FechaRemito)
                DBConn.Execute sql
                
                For I = 1 To grdGrilla.Rows - 1
                    If grdGrilla.TextMatrix(I, 0) <> "" Then
                        sql = "INSERT INTO DETALLE_REMITO_CLIENTE"
                        sql = sql & " (RCL_NUMERO,RCL_SUCURSAL,RCL_FECHA,DRC_NROITEM,"
                        sql = sql & "PTO_CODIGO,DRC_CANTIDAD,DRC_PRECIO,DRC_DETALLE)"
                        sql = sql & " VALUES ("
                        sql = sql & XN(txtNroRemito) & ","
                        sql = sql & XN(txtNroSucursal) & ","
                        sql = sql & XDQ(FechaRemito) & ","
                        sql = sql & XN(grdGrilla.TextMatrix(I, 7)) & ","
                        If grdGrilla.TextMatrix(I, 0) <> "----------" Then
                            sql = sql & XN(grdGrilla.TextMatrix(I, 0), True) & ","
                        Else
                            sql = sql & "99999999" & ","
                        End If
                        sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & ","
                        sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & ","
                        grdGrilla.TextMatrix(I, 1) = Replace(grdGrilla.TextMatrix(I, 1), "'", "´")
                        sql = sql & XS(grdGrilla.TextMatrix(I, 1)) & ")"
                        DBConn.Execute sql
                    End If
                Next
                
                'ACTUALIZO EL STOCK CUANDO EL REMITO ES DEFINITIVO (STOCK PENDIENTE)
                'Y ES REMITO SIN FACTURAS
                ' If chkRemitoSinFactura.Value = Checked Then
                For I = 1 To grdGrilla.Rows - 1
                    If grdGrilla.TextMatrix(I, 0) <> "" Then
                            sql = "UPDATE DETALLE_STOCK"
                            sql = sql & " SET"
                            sql = sql & " DST_STKFIS = DST_STKFIS + " & CDbl(XN(IIf(grdGrilla.TextMatrix(I, 8) = "", "0", grdGrilla.TextMatrix(I, 8)))) & " - " & CDbl(XN(grdGrilla.TextMatrix(I, 2)))
                            sql = sql & " WHERE STK_CODIGO = 1"
                            '& cboStock.ItemData(cboStock.ListIndex)
                            sql = sql & " AND PTO_CODIGO LIKE '" & grdGrilla.TextMatrix(I, 0) & "' "
                            DBConn.Execute sql
                    End If
                Next
                ' End If
                
                If TxtCodigoCli.Text <> 2 Then 'Cliente Consumidor Final
                'ACTUALIZO LA CUENTA CORRIENTE DEL CLIENTE
                DBConn.Execute QuitoCtaCteCliente(TxtCodigoCli, 20, txtNroRemito, txtNroSucursal)
                
                DBConn.Execute AgregoCtaCteCliente(TxtCodigoCli, 20 _
                                                , txtNroRemito, txtNroSucursal, _
                                                FechaRemito, txtTotal, "D", CStr(Date))
                End If
                
                DBConn.CommitTrans
                End If
            'End If
            
            
        'rec.Close
        Screen.MousePointer = vbNormal
        lblestado.Caption = ""
        
    '    If MsgBox("¿Desea Facturar el Remito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
    '
    '        frmFacturaCliente.txtRemSuc = Format(txtNroSucursal.Text, "0000")
    '        frmFacturaCliente.txtNroRemito = Format(txtNroRemito.Text, "00000000")
    '
    '        frmFacturaCliente.Show vbModal
    '
    '
    '    End If
    'End If
    CmdNuevo_Click
    CmdSalir.Enabled = True
    Exit Sub
    
HayErrorRemito:
    lblestado.Caption = ""
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Function ValidarRemito() As Boolean
    Dim BAND As Integer
    If FechaRemito.Text = "" Then
        MsgBox "La Fecha del Remito es requerida", vbExclamation, TIT_MSGBOX
        FechaRemito.SetFocus
        ValidarRemito = False
        Exit Function
    End If
    If chkNotaPedido.Value = 1 Then
        If txtNroNotaPedido.Text = "" Then
            MsgBox "El número de Presupuesto es requerido", vbExclamation, TIT_MSGBOX
            txtNroNotaPedido.SetFocus
            ValidarRemito = False
            Exit Function
        End If
        If FechaNotaPedido.Text = "" Then
            MsgBox "La Fecha del Presupuesto es requerida", vbExclamation, TIT_MSGBOX
            FechaNotaPedido.SetFocus
            ValidarRemito = False
            Exit Function
        End If
    End If
'    If chkRemitoSinFactura.Value = Checked Then
'        If txtConcepto.Text = "" Then
'            MsgBox "Debe ingresar un concepto", vbExclamation, TIT_MSGBOX
'            txtConcepto.SetFocus
'            ValidarRemito = False
'            Exit Function
'        End If
'    End If
    If TxtCodigoCli.Text = "" Then
        MsgBox "Debe ingresar un Cliente", vbExclamation, TIT_MSGBOX
        TxtCodigoCli.SetFocus
        ValidarRemito = False
        Exit Function
    End If
    BAND = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" Then
            BAND = 1
        End If
    Next
    If BAND = 0 Then
        MsgBox "Debe ingresar un item en detalle", vbExclamation, TIT_MSGBOX
        grdGrilla.SetFocus
        ValidarRemito = False
        Exit Function
    End If
    ValidarRemito = True
End Function

Private Sub cmdImprimir_Click()
    If MsgBox("¿Confirma Impresión Remito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    VCantidadBultos = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" Then
            VCantidadBultos = CInt(grdGrilla.TextMatrix(I, 2)) + VCantidadBultos
        End If
    Next I
'PONE A LA IMPRESORA  COMO PREDETERMINADA
    Dim X As Printer
    Dim mDriver As String
    mDriver = IMPRESORA
    For Each X In Printers
        If X.DeviceName = mDriver Then
            ' La define como predeterminada del sistema.
            Set Printer = X
            Exit For
        End If
    Next
'-----------------------------------
    Set_Impresora
    ImprimirRemito
End Sub


Public Sub ImprimirEncabezado()
 '---------------IMPRIME EL ENCABEZADO DEL REMITO-------------------
    'Imprimir 15.8, 2.7, False, Format(FechaRemito, "dd/mm/yyyy")
    Imprimir 15.7, 4.4, False, Format(Day(FechaRemito), "00")
    Imprimir 16.75, 4.4, False, Format(Month(FechaRemito), "00")
    Imprimir 17.8, 4.4, False, Format(Year(FechaRemito), "yy")
    
    'PROBAR IMPRESIÓN
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT C.CLI_RAZSOC,C.CLI_DOMICI,C.CLI_CUIT,C.CLI_INGBRU, L.LOC_DESCRI,CI.IVA_CODIGO"
    sql = sql & ", P.PRO_DESCRI,CI.IVA_DESCRI"
    sql = sql & " FROM CLIENTE C, LOCALIDAD L, REMITO_CLIENTE R, PROVINCIA P, CONDICION_IVA CI"
    sql = sql & " WHERE  R.RCL_NUMERO=" & XN(txtNroRemito)
    sql = sql & " AND R.RCL_FECHA=" & XDQ(FechaRemito)
    sql = sql & " AND R.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
    sql = sql & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Imprimir 2.5, 7, True, Trim(Rec1!CLI_RAZSOC)
        Imprimir 10.8, 6.8, False, Trim(Rec1!CLI_DOMICI)
        'nota de pedido
        'Imprimir 13, 5.3, True, "Nro.Pedido: " & Format(txtNroNotaPedido.Text, "00000000")
        Imprimir 10.8, 7.2, False, Trim(Rec1!LOC_DESCRI) & " - " & Trim(Rec1!PRO_DESCRI)
        'fecha nota pedido
        'If Rec1!IVA_CODIGO = 1 Then
        '    Imprimir 3.7, 7.8, False, "X"
        'Else
        '    Imprimir 9.5, 7.8, False, "X"
        'End If
        'Imprimir 13, 5.7, True, "Fecha: " & Format(FechaNotaPedido.Text, "dd/mm/yyyy")
        'Imprimir 1, 6.4, False, Trim(Rec1!IVA_DESCRI)
        Imprimir 13.7, 7.8, False, IIf(IsNull(Rec1!CLI_CUIT), "", Format(Rec1!CLI_CUIT, "##-########-#"))
        Imprimir 14.2, 8.5, False, IIf(IsNull(Rec1!CLI_INGBRU), "", Format(Rec1!CLI_INGBRU, "###-#####-##"))
    End If
    Rec1.Close
    'Imprimir 18.4, 7.9, False, CStr(VCantidadBultos)
    'Imprimir 0, 9.1, False, "Código"
    'Imprimir 3.1, 9.1, False, "Descripción"
    'Imprimir 12, 9.1, False, "Cantidad"
    'Imprimir 15, 9.1, False, "Rubro"
End Sub

Public Sub ImprimirRemito()
    Dim Renglon As Double
    Dim canttxt As Integer
    
    Screen.MousePointer = vbHourglass
    lblestado.Caption = "Imprimiendo..."
    
    For w = 1 To 2 'SE IMPRIME POR DUPLICADO
      '-----IMPRESION DEL ENCABEZADO------------------
        ImprimirEncabezado
        
      '---- IMPRESION DEL REINTEGRO ------------------
        Renglon = 9.8
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                Imprimir 1, Renglon, False, Format(grdGrilla.TextMatrix(I, 0), "000000000")  'codigo
                If Len(grdGrilla.TextMatrix(I, 1)) < 35 Then
                    Imprimir 3, Renglon, False, grdGrilla.TextMatrix(I, 1)  'descripcion
                Else
                    CortarCadena Renglon, grdGrilla.TextMatrix(I, 1)
'                    Imprimir 3, Renglon, False, Left(grdGrilla.TextMatrix(I, 1), 36) 'descripcion
'                    Imprimir 3, Renglon + 0.5, False, Mid(grdGrilla.TextMatrix(I, 1), 37, 38) 'descripcion
'                    Imprimir 3, Renglon + 1, False, Mid(grdGrilla.TextMatrix(I, 1), 75, 36) 'descripcion
'                    Imprimir 3, Renglon + 1.5, False, Mid(grdGrilla.TextMatrix(I, 1), 111, 36) 'descripcion
'                    Imprimir 3, Renglon + 2, False, Mid(grdGrilla.TextMatrix(I, 1), 147, 36) 'descripcion
'                    Imprimir 3, Renglon + 2.5, False, Mid(grdGrilla.TextMatrix(I, 1), 183, 36) 'descripcion
'                    Imprimir 3, Renglon + 3, False, Mid(grdGrilla.TextMatrix(I, 1), 219, 36) 'descripcion
                    canttxt = Len(grdGrilla.TextMatrix(I, 1))
                    canttxt = canttxt / 36 'es para sacar la cantidad de renglones
                    canttxt = Int(canttxt)
                End If
                Imprimir 13.2, Renglon, False, grdGrilla.TextMatrix(I, 2) 'cantidad
                'Imprimir 15, Renglon, False, Left(grdGrilla.TextMatrix(I, 4), 20) 'rubro
                Renglon = Renglon + (canttxt * 0.5) + 0.5
            End If
        Next I
        '-----OBSERVACIONES---------------------
        If txtObservaciones.Text <> "" Then
            Imprimir 0.5, Renglon + 1, False, "Observaciones: " & Trim(txtObservaciones.Text)
        End If
        'txtObservaciones
        '------------DATOS REPRESENTADA----------------------
'           Set Rec1 = New ADODB.Recordset
'           If chkFacturaTerceros.Value = Checked Then
'                sql = "SELECT REP_RAZSOC, REP_CUIT"
'                sql = sql & " FROM REPRESENTADA"
'                sql = sql & " WHERE REP_CODIGO=" & cboRep.ItemData(cboRep.ListIndex)
'                Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'                If Rec1.EOF = False Then
'                    Imprimir 0, 16.2, True, "Corresponde a Factura Nro.: " & Format(txtNroFactura.Text, "00000000")
'                    Imprimir 0, 16.6, True, "Por Cuenta y Orden de " & Trim(Rec1!REP_RAZSOC)
'                    Imprimir 0, 17, True, "CUIT: " & IIf(IsNull(Rec1!REP_CUIT), "NO INFORMADO", Format(Rec1!REP_CUIT, "##-########-#"))
'                End If
'                Rec1.Close
'           End If
'          '------------DATOS SUCURSAL-------------------------
'
'           sql = "SELECT S.SUC_DESCRI,S.SUC_DOMICI, L.LOC_DESCRI"
'           sql = sql & " FROM SUCURSAL S, NOTA_PEDIDO NP, LOCALIDAD L"
'           sql = sql & " WHERE NP.NPE_NUMERO=" & XN(txtNroNotaPedido)
'           sql = sql & " AND NP.NPE_FECHA=" & XDQ(FechaNotaPedido)
'           sql = sql & " AND NP.SUC_CODIGO=S.SUC_CODIGO"
'           sql = sql & " AND S.LOC_CODIGO=L.LOC_CODIGO"
'           sql = sql & " AND S.PRO_CODIGO=L.PRO_CODIGO"
'           sql = sql & " AND S.PAI_CODIGO=L.PAI_CODIGO"
'           Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'           If Rec1.EOF = False Then
'                Imprimir 0, 17.4, True, "Entregar: " & Left(Trim(Rec1!SUC_DESCRI), 25) & " -- " & Left(Trim(Rec1!SUC_DOMICI), 30) & " (" & Left(Trim(Rec1!LOC_DESCRI), 20) & ")"
'           End If
'           Rec1.Close
          '----------------------------------------------------
        Printer.EndDoc
    Next w
    Screen.MousePointer = vbNormal
    lblestado.Caption = ""
End Sub

Private Sub CmdNuevo_Click()
   For I = 1 To grdGrilla.Rows - 1
        grdGrilla.TextMatrix(I, 0) = ""
        grdGrilla.TextMatrix(I, 1) = ""
        grdGrilla.TextMatrix(I, 2) = ""
        grdGrilla.TextMatrix(I, 3) = ""
        grdGrilla.TextMatrix(I, 4) = ""
        grdGrilla.TextMatrix(I, 5) = ""
        grdGrilla.TextMatrix(I, 6) = ""
        grdGrilla.TextMatrix(I, 8) = I
   Next
   'grillaNotaPedido.TextMatrix(0, 1) = ""
   'grillaNotaPedido.TextMatrix(1, 1) = ""
   'grillaNotaPedido.TextMatrix(2, 1) = ""
   FechaNotaPedido.Text = ""
   txtNroNotaPedido.Text = ""
 '  chkRemitoSinFactura.Value = Unchecked
  ' txtConcepto.Text = ""
   lblEstadoRemito.Caption = ""
   txtObservaciones.Text = ""
   lblestado.Caption = ""
   cmdGrabar.Enabled = True
   freRemito.Enabled = True
   freCliente.Enabled = True
       
   txtNroRemito.Text = ""
   'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoRemito) 'ESTADO PENDIENTE
    VEstadoRemito = 1
    '--------------
    FechaRemito.Enabled = True
    txtNroNotaPedido.Enabled = True
    FechaNotaPedido.Enabled = True
    cmdBuscarNotaPedido.Enabled = True
    '--------------
    tabDatos.Tab = 0
'    TipoBusquedaDoc = 1
    FechaRemito.Text = Date
    'cboListaPrecio.ListIndex = 0
    cboListaPrecio.Enabled = True
    cboListaPrecio.SetFocus
    
    TxtCodigoCli.Text = ""
    TxtCodigoCli_Change
    
    
    chkNotaPedido.Enabled = True
    chkNotaPedido.Value = 0
    freNotaPedido.Enabled = False
    freCliente.Enabled = True
    
    BuscoIva
    txtImporteIva.Text = ""
    txtSubtotal.Text = ""
    txtTotal.Text = ""
    txtObservaciones.Text = ""
    
    Call BuscaCodigoProxItemData(1, cboCondicion)
    tabLista.Tab = 1
    'GrdModulos.Rows = 1
    cmdFacturar.Visible = False
    chkPorc.Value = 0
    txtPorc.Text = ""
    txtPorc.Enabled = False
    fraInfoCliente.Visible = False
        
    cmdNuevo.Enabled = False
    cmdImprimir.Enabled = False
    CmdSalir.Enabled = False
    CmdCancelar.Enabled = True
    txtGastado.Text = "0,00"
    txtAutorizado.Text = "0,00"
    
    txtAum.Text = "0,00"
End Sub

Private Sub cmdNuevoCliente_Click()
    ABMCliente.Show vbModal
    TxtCodigoCli.SetFocus
End Sub

Private Sub cmdPrecio_Click()
Dim BAND As Integer
    BAND = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" Then
            BAND = 1
        End If
    Next
    If BAND = 0 Then
        MsgBox "Para actualizar los precios debe haber al menos un Item el detalle", vbExclamation, TIT_MSGBOX
        grdGrilla.SetFocus
        Exit Sub
    End If
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" Then
            sql = "SELECT PTO_CODIGO,PTO_PRECIO FROM PRODUCTO"
            sql = sql & " WHERE PTO_CODIGO LIKE '" & grdGrilla.TextMatrix(I, 0) & "'"
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                grdGrilla.TextMatrix(I, 3) = Valido_Importe(rec!PTO_PRECIO)
            End If
            rec.Close
        End If
    Next
    grdGrilla_LeaveCell
    TotalPEDIDO
End Sub

Private Sub cmdQuitarProducto_Click()
    Dim TOTAL As Double
    Dim subtotal As Double
    Dim impIva As Double
    Dim BAND As Integer
    If MsgBox("Seguro que desea quitar el Detalle: " & grdGrilla.TextMatrix(grdGrilla.RowSel, 1), vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        'BORRO EN BD
'        sql = "DELETE FROM DETALLE_REMITO_CLIENTE WHERE RCL_NUMERO = " & XN(txtNroRemito)
'        sql = sql & " AND RCL_SUCURSAL = " & XN(txtNroSucursal)
'        sql = sql & " AND PTO_CODIGO LIKE '" & grdGrilla.TextMatrix(grdGrilla.RowSel, 0) & "'"
'
'        DBConn.Execute sql
        'BORRO EN PANTALLA
        grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grdGrilla.RowSel
    End If
    BAND = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" Then
            BAND = 1
        End If
    Next
    If BAND = 0 Then
        txtSubtotal.Text = "0,00"
        txtImporteIva.Text = "0,00"
        txtTotal.Text = "0,00"
    Else
        TotalPEDIDO
'        txtSubtotal.Text = Valido_Importe(SumaTotal)
'        'CAMBIO EL CALCULO
'        txtImporteIva.Text = CDbl(txtSubtotal.Text) / CDbl("1," & Int(txtPorcentajeIva.Text))
'        txtSubtotal.Text = Valido_Importe(SumaTotal - txtImporteIva.Text)
'        txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)
'        subtotal = txtSubtotal
'        impIva = txtImporteIva
'        TOTAL = subtotal + impIva
'        txtTotal.Text = Format(TOTAL, "0.00")
    End If
    'HABRIA QUE RECARGAR LA GRILLA
'    For I = 2 To grdGrilla.Rows - 1
'        'i = grdGrilla.Rows
'        grdGrilla.TextMatrix(I, 0) = ""
'        grdGrilla.TextMatrix(I, 1) = ""
'        grdGrilla.TextMatrix(I, 2) = ""
'        grdGrilla.TextMatrix(I, 3) = ""
'        grdGrilla.TextMatrix(I, 4) = ""
'        grdGrilla.TextMatrix(I, 5) = ""
'        grdGrilla.TextMatrix(I, 6) = ""
'        grdGrilla.TextMatrix(I, 7) = ""
'        I = I + 1
'    Next
'
'    sql = "SELECT DRC.*, P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI,DRC_DETALLE"
'    sql = sql & " FROM DETALLE_REMITO_CLIENTE DRC, PRODUCTO P, RUBROS R, LINEAS L"
'    sql = sql & " WHERE DRC.RCL_SUCURSAL=" & XN(txtNroSucursal.Text)
'    sql = sql & " AND DRC.RCL_NUMERO=" & XN(txtNroRemito.Text)
'    sql = sql & " AND DRC.RCL_FECHA=" & XDQ(FechaRemito)
'    sql = sql & " AND DRC.PTO_CODIGO=P.PTO_CODIGO"
'    sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
'    sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
'    sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
'    sql = sql & " ORDER BY DRC.DRC_NROITEM"
'    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If Rec2.EOF = False Then
'        I = 1
'        Do While Rec2.EOF = False
'            grdGrilla.TextMatrix(I, 0) = IIf(Rec2!PTO_CODIGO = 99999999, "----------", Rec2!PTO_CODIGO)
'            If (grdGrilla.TextMatrix(I, 0)) = "----------" Then
'                grdGrilla.TextMatrix(I, 1) = IIf(IsNull(Rec2!DRC_DETALLE), Rec2!PTO_DESCRI, Rec2!DRC_DETALLE)
'            Else
'                'grdGrilla.TextMatrix(i, 1) = Rec2!PTO_DESCRI
'                grdGrilla.TextMatrix(I, 1) = IIf(IsNull(Rec2!DRC_DETALLE), Rec2!PTO_DESCRI, Rec2!DRC_DETALLE)
'                grdGrilla.TextMatrix(I, 2) = IIf(IsNull(Rec2!DRC_CANTIDAD), "", Rec2!DRC_CANTIDAD)
'                grdGrilla.TextMatrix(I, 3) = IIf(IsNull(Rec2!DRC_PRECIO), "", Valido_Importe(Rec2!DRC_PRECIO))
'                grdGrilla.TextMatrix(I, 4) = IIf(IsNull(Rec2!RUB_DESCRI), "", Rec2!RUB_DESCRI)
'                grdGrilla.TextMatrix(I, 5) = IIf(IsNull(Rec2!LNA_DESCRI), "", Rec2!LNA_DESCRI)
'                grdGrilla.TextMatrix(I, 6) = I
'                grdGrilla.TextMatrix(I, 7) = IIf(IsNull(Rec2!DRC_CANTIDAD), "", Rec2!DRC_CANTIDAD)
'            End If
'            I = I + 1
'            Rec2.MoveNext
'        Loop
'    End If
'    Rec2.Close
End Sub
Private Sub cmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmRemitoCliente = Nothing
        Unload Me
        
        'Unload frmInfoCliente
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyF1 Then
        'TipoBusquedaDoc = 1
        'tabDatos.Tab = 1
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl.Name <> "grdGrilla" And _
        Me.ActiveControl.Name <> "txtEdit" And _
        KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)

    grdGrilla.FormatString = "Código|Descripción|Cantidad|Precio|Importe|Rubro|Linea|Orden|CANT"
    grdGrilla.ColWidth(0) = 1400 'CODIGO
    grdGrilla.ColWidth(1) = 5300 'DESCRIPCION
    grdGrilla.ColWidth(2) = 1000 'CANTIDAD
    grdGrilla.ColWidth(3) = 1100 'PRECIO
    grdGrilla.ColWidth(4) = 1100 'Importe
    grdGrilla.ColWidth(5) = 2100 'RUBRO
    grdGrilla.ColWidth(6) = 2100 'LINEA
    grdGrilla.ColWidth(7) = 0    'ORDEN
    grdGrilla.ColWidth(8) = 0    'CANTIDAD STOCK
    grdGrilla.Cols = 9
    grdGrilla.Rows = 1
    For I = 2 To 30
        grdGrilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & (I - 1)
    Next
    'GRILLA (GrdModulos) PARA LA BUSQUEDA
    GrdModulos.FormatString = "^Número|^Fecha|Importe|Cliente|Domicilio|Localidad|Provincia|Cod Estado|NP NUMERO|NP FECHA|OBSERVACIONES|" _
                              & "STOCK|REMITO SIN FACTURA|CODIGOCLIENTE"
    GrdModulos.ColWidth(0) = 1500 'NUMERO
    GrdModulos.ColWidth(1) = 1000 'FECHA
    GrdModulos.ColWidth(2) = 1000 'IMPORTE A COBRAR
    GrdModulos.ColWidth(3) = 3200 'CLIENTE
    GrdModulos.ColWidth(4) = 3200 'DOMICILIO
    GrdModulos.ColWidth(5) = 3200 'Localidad
    GrdModulos.ColWidth(6) = 3200 'Provincia
    GrdModulos.ColWidth(7) = 0    'COD ESTADO
    GrdModulos.ColWidth(8) = 0    'NOTA PEDIDO NUMERO
    GrdModulos.ColWidth(9) = 0    'NOTA PEDIDO FECHA
    GrdModulos.ColWidth(10) = 0   'OBSERVACIONES
    GrdModulos.ColWidth(11) = 0   'STOCK
    GrdModulos.ColWidth(12) = 0   'REMITO SIN FACTURAS
    GrdModulos.ColWidth(13) = 0   'CODIGOCLIENTE
    
    GrdModulos.Rows = 1
    '------------------------------------
    'grillaNotaPedido.ColWidth(0) = 950
    'grillaNotaPedido.ColWidth(1) = 5300
    'grillaNotaPedido.TextMatrix(0, 0) = "    Cliente:"
    'grillaNotaPedido.TextMatrix(1, 0) = " Sucursal:"
    'grillaNotaPedido.TextMatrix(2, 0) = "Vendedor:"
    '------------------------------------
    lblestado.Caption = ""
    'Llenar Combo forma de pago
    LlenarComboFormaPago
    'CARGO EL COMBO DE LISTA DE PRECIOS DE MAQUINARIAS
    CargoCboListaPrecio
    'CARGO EL COMBO DE LISTA DE PRECIOS DE REPUESTOS
    CargoCboLPrecioRep
    'BUSCO EL NUMERO DE REMITO QUE CORRESPONDE
    txtNroSucursal.Text = Sucursal
    
    'txtNroRemito.Text = BuscoUltimoRenito
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoRemito) 'ESTADO PENDIENTE
    VEstadoRemito = 1
    FechaRemito.Text = Date
    'CARGO COMBO STOCK
    'CargaCboStock
    'PONGO ENABLE LOS DATOS DE LA FACTURA DE TERCEROS

    'txtConcepto.Enabled = False
    TipoBusquedaDoc = 1 'ESTO ES PARA BUSCAR REMITOS(1), (2)PARA BUSCAR NOTA DE PEDIDO
    tabDatos.Tab = 0
    
    freNotaPedido.Enabled = False
    freCliente.Enabled = True
    BuscoIva
    
    tabLista.Tab = 1
    End Sub
Private Sub LlenarComboFormaPago()
    sql = "SELECT * FROM FORMA_PAGO"
    sql = sql & " ORDER BY FPG_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboCondicion.AddItem rec!FPG_DESCRI
            cboCondicion.ItemData(cboCondicion.NewIndex) = rec!FPG_CODIGO
            rec.MoveNext
        Loop
        cboCondicion.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub BuscoIva()
    sql = "SELECT IVA FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtPorcentajeIva.Text = IIf(IsNull(rec!IVA), "", Format(rec!IVA, "0.00"))
    End If
    rec.Close
End Sub

Private Sub CargaCboStock()
    sql = "SELECT S.STK_CODIGO,R.REP_RAZSOC"
    sql = sql & " FROM STOCK S, REPRESENTADA R"
    sql = sql & " WHERE S.REP_CODIGO=R.REP_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboStock.AddItem rec!REP_RAZSOC
            cboStock.ItemData(cboStock.NewIndex) = rec!STK_CODIGO
            rec.MoveNext
        Loop
        cboStock.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub CargoCboListaPrecio() '' Lista de Precios de Repuestos
    sql = "SELECT DISTINCT LP.LIS_CODIGO, LP.LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO LP, PRODUCTO P"
    sql = sql & " WHERE LP.LIS_CODIGO = P.LIS_CODIGO"
    sql = sql & " AND P.LNA_CODIGO = 6"   '6: Maquinaria
    sql = sql & " ORDER BY LIS_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
    If rec.EOF = False Then
        rec.MoveFirst
        Do While rec.EOF = False
            cboListaPrecio.AddItem rec!LIS_DESCRI
            cboListaPrecio.ItemData(cboListaPrecio.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cboListaPrecio.ListIndex = 0
    End If
    rec.Close
End Sub
Private Sub CargoCboLPrecioRep() '' Lista de Precios de Repuestos
    sql = "SELECT DISTINCT LP.LIS_CODIGO, LP.LIS_DESCRI"
    sql = sql & " FROM LISTA_PRECIO LP, PRODUCTO P"
    sql = sql & " WHERE LP.LIS_CODIGO = P.LIS_CODIGO"
    sql = sql & " AND P.LNA_CODIGO = 1"   '1: Repuestos
    sql = sql & " ORDER BY LIS_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
    If rec.EOF = False Then
        rec.MoveFirst
        Do While rec.EOF = False
            cboLPrecioRep.AddItem rec!LIS_DESCRI
            cboLPrecioRep.ItemData(cboLPrecioRep.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cboLPrecioRep.ListIndex = 0
    End If
    rec.Close
End Sub

Private Function BuscoUltimoRenito() As String
    'ACA BUSCA EL NUMERO DE REMITO SIGUIENTE AL ULTIMO CARGADO
    sql = "SELECT (NRO_REMITO) + 1 AS ULTIMO"
    sql = sql & " FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        txtNroSucursal.Text = Sucursal
        BuscoUltimoRenito = Format(rec!Ultimo, "00000000")
    End If
    rec.Close
End Function

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Select Case grdGrilla.Col
        Case 0, 1
            LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
            grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = grdGrilla.RowSel
            grdGrilla.Col = 0
        'Case Else
        '    grdGrilla.TextArray(GRIDINDEX(grdGrilla, grdGrilla.row, grdGrilla.Col)) = ""
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
        Case 1
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = "" Then
                txtObservaciones.SetFocus
            End If
        Case 2
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "" Then
                grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "1"
            End If
        End Select
    End If
End Sub

Private Sub grdGrilla_KeyPress(KeyAscii As Integer)
    If (grdGrilla.Col = 0) Or (grdGrilla.Col = 1) Or _
       (grdGrilla.Col = 2) Or (grdGrilla.Col = 3) Then
        If KeyAscii = vbKeyReturn Then
            If grdGrilla.Col = 3 Then
                If grdGrilla.row < grdGrilla.Rows - 1 Then
                    grdGrilla.row = grdGrilla.row + 1
                    grdGrilla.Col = 0
                Else
                    SendKeys "{TAB}"
                End If
            Else
                grdGrilla.Col = grdGrilla.Col + 1
            End If
        Else
            If (grdGrilla.Col <> 1) Then
                'If KeyAscii > 47 And KeyAscii < 58 Then
                    EDITAR grdGrilla, txtEdit, KeyAscii
               ' End If
            Else
                EDITAR grdGrilla, txtEdit, KeyAscii
            End If
        End If
    End If
End Sub

Private Sub grdGrilla_LeaveCell()
    If txtEdit.Visible = False Then Exit Sub
    'If Trim(TxtEdit) = "" Then TxtEdit = "0"
    grdGrilla = txtEdit.Text
    txtEdit.Visible = False
End Sub

Private Sub grdGrilla_GotFocus()
    If grdGrilla.Rows > 1 Then
        If txtEdit.Visible = False Then Exit Sub
        grdGrilla = txtEdit.Text
        txtEdit.Visible = False
    End If
End Sub

Private Sub GrdModulos_DblClick()
    If GrdModulos.Rows > 1 Then
        CmdNuevo_Click
        Select Case TipoBusquedaDoc
        Case 1 'BUSCA REMITOS
            lblestado.Caption = "Buscando..."
            Screen.MousePointer = vbHourglass
            Set Rec1 = New ADODB.Recordset
            Set Rec2 = New ADODB.Recordset
            txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 4)
            txtNroRemito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8)
            FechaRemito.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
            'CARGO EL ESTADO
            Call BuscoEstado(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 7)), lblEstadoRemito)
            VEstadoRemito = CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 7))
            
            
            
            If VEstadoRemito <> 1 Then
                cmdGrabar.Enabled = False
                freCliente.Enabled = False
                freNotaPedido.Enabled = False
                freRemito.Enabled = False
                chkNotaPedido.Enabled = False
                grdGrilla.SetFocus
            Else
                cmdGrabar.Enabled = True
                freCliente.Enabled = True
                freNotaPedido.Enabled = True
                freRemito.Enabled = True
                chkNotaPedido.Enabled = True
            End If
            
            
            txtNroNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 8)
            FechaNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 9)
            
            'BUSCO STOCK
            'If GrdModulos.TextMatrix(GrdModulos.RowSel, 10) <> "" Then
            '    Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 10)), cboStock)
            'Else
            '    cboStock.ListIndex = 0
            'End If
            txtObservaciones.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 9)
            
            TxtCodigoCli.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 13)
            TxtCodigoCli_LostFocus
            
            'BUSCO PORCENTAJE EN REMITO
            sql = "SELECT RCL_PORCEN,RCL_AUM FROM REMITO_CLIENTE "
            sql = sql & "WHERE RCL_NUMERO =" & XN(txtNroRemito.Text) & " "
            sql = sql & "AND RCL_SUCURSAL =" & XN(txtNroSucursal.Text) & ""
            Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec2.EOF = False Then
                If Rec2!RCL_PORCEN = 0 Or Rec2!RCL_PORCEN = "" Or IsNull(Rec2!RCL_PORCEN) Then
                    chkPorc.Value = 0
                Else
                    chkPorc.Value = 1
                    txtPorc.Text = IIf(IsNull(Rec2!RCL_PORCEN), "", Rec2!RCL_PORCEN)
                    txtPorc_LostFocus
                End If
                txtAum.Text = ChkNull(Rec2!RCL_AUM)
                txtAum.Text = Valido_Importe(txtAum.Text)
            End If
            Rec2.Close
        '----BUSCO DETALLE DEL REMITO------------------
            
            sql = "SELECT DRC.*, P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI,DRC_DETALLE"
            sql = sql & " FROM DETALLE_REMITO_CLIENTE DRC, PRODUCTO P, RUBROS R, LINEAS L"
            sql = sql & " WHERE DRC.RCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 4))
            sql = sql & " AND DRC.RCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8))
            sql = sql & " AND DRC.RCL_FECHA=" & XDQ(GrdModulos.TextMatrix(GrdModulos.RowSel, 1))
            sql = sql & " AND DRC.PTO_CODIGO=P.PTO_CODIGO"
            sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
            sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
            sql = sql & " ORDER BY DRC.DRC_NROITEM"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                I = 1
                Do While Rec1.EOF = False
                    grdGrilla.TextMatrix(I, 0) = IIf(Rec1!PTO_CODIGO = 99999999, "----------", Rec1!PTO_CODIGO)
                    If (grdGrilla.TextMatrix(I, 0)) = "----------" Then
                        grdGrilla.TextMatrix(I, 1) = IIf(IsNull(Rec1!DRC_DETALLE), Rec1!PTO_DESCRI, Rec1!DRC_DETALLE)
                    Else
                        'grdGrilla.TextMatrix(I, 1) = Rec1!PTO_DESCRI
                        grdGrilla.TextMatrix(I, 1) = IIf(IsNull(Rec1!DRC_DETALLE), Rec1!PTO_DESCRI, Rec1!DRC_DETALLE)
                        grdGrilla.TextMatrix(I, 2) = IIf(IsNull(Rec1!DRC_CANTIDAD), 0, Rec1!DRC_CANTIDAD)
                        grdGrilla.TextMatrix(I, 3) = IIf(IsNull(Rec1!DRC_PRECIO), "", Valido_Importe(Rec1!DRC_PRECIO))
                        grdGrilla.TextMatrix(I, 4) = Valido_Importe(grdGrilla.TextMatrix(I, 2) * grdGrilla.TextMatrix(I, 3))
                        grdGrilla.TextMatrix(I, 5) = IIf(IsNull(Rec1!RUB_DESCRI), "", Rec1!RUB_DESCRI)
                        grdGrilla.TextMatrix(I, 6) = IIf(IsNull(Rec1!LNA_DESCRI), "", Rec1!LNA_DESCRI)
                        grdGrilla.TextMatrix(I, 7) = IIf(IsNull(Rec1!DRC_NROITEM), "", Rec1!DRC_NROITEM)
                        grdGrilla.TextMatrix(I, 8) = IIf(IsNull(Rec1!DRC_CANTIDAD), "", Rec1!DRC_CANTIDAD)
                    End If
                    I = I + 1
                    Rec1.MoveNext
                Loop
            Else
                BuscoPtosBorrados
            End If
            lblestado.Caption = ""
            Screen.MousePointer = vbNormal
            '--------------
            FechaRemito.Enabled = False
            txtNroNotaPedido.Enabled = False
            FechaNotaPedido.Enabled = False
            cmdBuscarNotaPedido.Enabled = False
            '--------------
            tabDatos.Tab = 0
            grdGrilla.SetFocus
            grdGrilla.row = 1
            Rec1.Close
            
            'ACA PROGRAMO LO DEL BOTON FACTURAR
            If (TxtCodigoCli.Text <> "2") And (TxtCodigoCli.Text <> "2") Then
                If lblEstadoRemito = "PENDIENTE" Then
                    cmdFacturar.Visible = True
                Else
                    cmdFacturar.Visible = False
                End If
            End If
        '----------------------------------------------
        Case 2 'BUSCA NOTA PEDIDO
            txtNroNotaPedido.Text = Format(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), "00000000")
            FechaNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
            tabDatos.Tab = 0
            txtNroNotaPedido_LostFocus
            'TxtCodigoCli.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 7)
            'TxtCodigoCli_LostFocus
        End Select
        TotalPEDIDO
'        Dim TOTAL As Double
'        Dim subtotal As Double
'        Dim impIva As Double
''        txtSubtotal.Text = Valido_Importe(SumaTotal)
''        txtImporteIva.Text = (CDbl(txtSubtotal.Text) * CDbl(txtPorcentajeIva.Text)) / 100
''        txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)
''        subtotal = txtSubtotal
''        impIva = txtImporteIva
''        TOTAL = subtotal + impIva
''        txtTotal.Text = Format(TOTAL, "0.00")
'        txtSubtotal.Text = Valido_Importe(SumaTotal)
'        'txtImporteIva.Text = (CDbl(txtSubTotal.Text) * CDbl(txtPorcentajeIva.Text)) / 100
'        txtImporteIva.Text = CDbl(txtSubtotal.Text) / CDbl("1," & Int(txtPorcentajeIva.Text))
'        txtSubtotal.Text = Valido_Importe(SumaTotal - txtImporteIva.Text)
'        txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)
'        subtotal = txtSubtotal
'        impIva = txtImporteIva
'        TOTAL = subtotal + impIva
'        txtTotal.Text = Format(TOTAL, "0.00")
        
    End If
End Sub
Private Function BuscoPtosBorrados()
    Set Rec2 = New ADODB.Recordset
    sql = "SELECT *"
    sql = sql & " FROM DETALLE_REMITO_CLIENTE "
    sql = sql & " WHERE RCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 4))
    sql = sql & " AND RCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8))
    sql = sql & " AND RCL_FECHA=" & XDQ(GrdModulos.TextMatrix(GrdModulos.RowSel, 1))
    sql = sql & " ORDER BY DRC_NROITEM"
    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec2.EOF = False Then
        I = 1
        Do While Rec2.EOF = False
            grdGrilla.TextMatrix(I, 0) = IIf(Rec2!PTO_CODIGO = 99999999, "----------", Rec2!PTO_CODIGO)
            If (grdGrilla.TextMatrix(I, 0)) = "----------" Then
                grdGrilla.TextMatrix(I, 1) = IIf(IsNull(Rec2!DRC_DETALLE), Rec2!PTO_DESCRI, Rec2!DRC_DETALLE)
            Else
                'grdGrilla.TextMatrix(I, 1) = Rec2!PTO_DESCRI
                grdGrilla.TextMatrix(I, 1) = IIf(IsNull(Rec2!DRC_DETALLE), "", Rec2!DRC_DETALLE)
                grdGrilla.TextMatrix(I, 2) = IIf(IsNull(Rec2!DRC_CANTIDAD), "", Rec2!DRC_CANTIDAD)
                grdGrilla.TextMatrix(I, 3) = IIf(IsNull(Rec2!DRC_PRECIO), "", Valido_Importe(Rec2!DRC_PRECIO))
                grdGrilla.TextMatrix(I, 4) = Valido_Importe(grdGrilla.TextMatrix(I, 2) * grdGrilla.TextMatrix(I, 3))
                grdGrilla.TextMatrix(I, 5) = ""
                grdGrilla.TextMatrix(I, 6) = ""
                grdGrilla.TextMatrix(I, 7) = IIf(IsNull(Rec2!DRC_NROITEM), "", Rec2!DRC_NROITEM)
                grdGrilla.TextMatrix(I, 8) = IIf(IsNull(Rec2!DRC_CANTIDAD), "", Rec2!DRC_CANTIDAD)
            End If
            I = I + 1
            Rec2.MoveNext
        Loop
    
    End If
    Rec2.Close
End Function

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_DblClick
End Sub

Private Sub MaskEdBox1_Change()

End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
  If tabDatos.Tab = 1 Then
    txtCliente.Enabled = False
    FechaDesde.Enabled = False
    FechaHasta.Enabled = False
    txtVendedor.Enabled = False
    cmdGrabar.Enabled = False
    cmdBuscarCli.Enabled = False
    cmdBuscarVendedor.Enabled = False
    'LimpiarBusqueda
    If Me.Visible = True Then chkCliente.SetFocus
    If TipoBusquedaDoc = 1 Then
        frameBuscar.Caption = "Buscar Remito por..."
        chkVendedor.Enabled = False
        txtVendedor.Enabled = False
    Else
        frameBuscar.Caption = "Buscar Presupuestos Pendientes por..."
        chkVendedor.Enabled = True
        txtVendedor.Enabled = True
    End If
  Else
    If VEstadoRemito = 1 Then
        cmdGrabar.Enabled = True
    Else
        cmdGrabar.Enabled = False
    End If
  End If
End Sub

Private Sub LimpiarBusqueda()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    FechaDesde.Text = ""
    FechaHasta.Text = ""
    txtVendedor.Text = ""
    txtDesVen.Text = ""
    GrdModulos.Rows = 1
    chkCliente.Value = Unchecked
    chkFecha.Value = Unchecked
    chkVendedor.Value = Unchecked
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtAum_GotFocus()
    SelecTexto txtAum
End Sub

Private Sub txtAum_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtAum, KeyAscii)
End Sub

Private Sub txtAum_LostFocus()
    txtAum.Text = Valido_Importe(txtAum)
End Sub

Private Sub txtCliente_Change()
    If txtCliente.Text = "" Then
        txtDesCli.Text = ""
    End If
End Sub

Private Sub txtCliente_GotFocus()
    SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCliente_LostFocus()
    If txtCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_RAZSOC FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(txtCliente)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesCli.Text = rec!CLI_RAZSOC
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtDesCli.Text = ""
            txtCliente.SetFocus
        End If
        rec.Close
    End If
'    If chkSucursal.Value = Unchecked And chkFecha.Value = Unchecked _
'        And chkVendedor.Value = Unchecked And ActiveControl.Name <> "cmdBuscarCli" _
'        And ActiveControl.Name <> "cmdNuevo" And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub

Private Function BuscoCondicionIVA(IVACodigo As String) As String
    sql = "SELECT * FROM CONDICION_IVA"
    sql = sql & " WHERE IVA_CODIGO=" & XN(IVACodigo)
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        BuscoCondicionIVA = Rec1!IVA_DESCRI
    Else
        BuscoCondicionIVA = ""
    End If
    Rec1.Close
End Function

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
    CarTexto KeyAscii
End Sub

Private Sub TxtCodigoCli_Change()
    If TxtCodigoCli.Text = "" Then
        TxtCodigoCli.Text = ""
        txtRazSocCli.Text = ""
        txtCUIT.Text = ""
        txtIngBrutos.Text = ""
        txtCondicionIVA.Text = ""
        txtDomici.Text = ""
        txtlocalidad.Text = ""
        txtprovincia.Text = ""
        txtcodpos.Text = ""
    End If
End Sub

Private Sub TxtCodigoCli_GotFocus()
    SelecTexto TxtCodigoCli
End Sub

Private Sub txtCodigoCli_KeyPress(KeyAscii As Integer)
        KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigoCli_LostFocus()
Set Rec1 = New ADODB.Recordset
    If ActiveControl.Name = "cmdGrabar" Or ActiveControl.Name = "cmdBorrar" Or ActiveControl.Name = "CmdSalir" Then Exit Sub
    If TxtCodigoCli.Text <> "" Then
        sql = "SELECT C.CLI_RAZSOC,C.CLI_DOMICI,C.CLI_CUIT,C.IVA_CODIGO,C.CLI_INGBRU,"
        sql = sql & "L.LOC_DESCRI,P.PRO_DESCRI,L.LOC_CODPOS"
        sql = sql & " FROM CLIENTE C, LOCALIDAD L, PROVINCIA P"
        sql = sql & " WHERE "
        sql = sql & "C.LOC_CODIGO = L.LOC_CODIGO AND "
        sql = sql & "C.PRO_CODIGO = P.PRO_CODIGO AND "
        sql = sql & "L.PRO_CODIGO = P.PRO_CODIGO AND "
        sql = sql & "C.CLI_CODIGO =" & XN(TxtCodigoCli)
        'sql = sql & " AND CLI_ESTADO=1"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtRazSocCli.Text = rec!CLI_RAZSOC
            txtDomici.Text = IIf(IsNull(rec!CLI_DOMICI), "", rec!CLI_DOMICI)
            txtlocalidad.Text = IIf(IsNull(rec!LOC_DESCRI), "", rec!LOC_DESCRI)
            txtprovincia.Text = IIf(IsNull(rec!PRO_DESCRI), "", rec!PRO_DESCRI)
            txtCondicionIVA.Text = BuscoCondicionIVA(rec!IVA_CODIGO)
            txtCUIT.Text = IIf(IsNull(rec!CLI_CUIT), "", Format(rec!CLI_CUIT, "##-########-#"))
            txtIngBrutos.Text = IIf(IsNull(rec!CLI_INGBRU), "", Format(rec!CLI_INGBRU, "###-#####-##"))
            txtcodpos.Text = IIf(IsNull(rec!LOC_CODPOS), "", rec!LOC_CODPOS)
            If TxtCodigoCli.Text <> 2 Then
                fraInfoCliente.Visible = True
                'AUTORIZADO
                txtAutorizado.Text = "0,00"
                sql = "SELECT CLI_CREDITO FROM CLIENTE "
                sql = sql & "WHERE CLI_CODIGO = " & XN(frmRemitoCliente.TxtCodigoCli.Text)
                Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If Rec1.EOF = False Then
                    txtAutorizado.Text = IIf(IsNull(Rec1!CLI_CREDITO), "0,00", Format(Rec1!CLI_CREDITO, "0.00"))
                End If
                Rec1.Close
                'GASTADO
                txtGastado.Text = "0,00"
                sql = "SELECT CLI_CODIGO,SUM(RCL_SALDO) AS GASTADO FROM REMITO_CLIENTE "
                sql = sql & "WHERE CLI_CODIGO = " & XN(TxtCodigoCli.Text) & " "
                sql = sql & " AND EST_CODIGO <> 2 "
                sql = sql & "GROUP BY CLI_CODIGO"
                Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If Rec1.EOF = False Then
                    If Rec1!GASTADO = Null Or IsNull(Rec1!GASTADO) Then
                        txtGastado.Text = "0,00"
                    Else
                        txtGastado.Text = IIf(IsNull(Rec1!GASTADO), "", Valido_Importe(Rec1!GASTADO))
                    End If
'                    txtGastado.Text = IIf(IsNull(Rec1!GASTADO), "", Valido_Importe(Rec1!GASTADO))
                End If
                Rec1.Close
            Else
                fraInfoCliente.Visible = False
            End If
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtRazSocCli.Text = ""
            TxtCodigoCli.SetFocus
        End If
        rec.Close
    Else
        fraInfoCliente.Visible = False
    End If
    'verificar la deuda
        
    'End If
    If TxtCodigoCli.Text <> "" And TxtCodigoCli.Text <> "2" Then
        Call BuscaCodigoProxItemData(2, cboCondicion) ' Cuenta Corriente
    Else
        Call BuscaCodigoProxItemData(1, cboCondicion) ' Cuenta Corriente
    End If
End Sub

Private Sub txtEdit_GotFocus()
    'SelecTexto txtEdit
End Sub

'Private Sub txtEdit_Click()
'    If grdGrilla.Col = 2 Then
'        If txtEdit.Text <> "" Then
'            EDITAR grdGrilla, txtEdit, 1
'        End If
'    End If
'End Sub

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    'If grdGrilla.Col = 0 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 0 Then
        'CarTexto KeyAscii
        txtEdit.MaxLength = 10
    End If
    If grdGrilla.Col = 1 Then
        txtEdit.MaxLength = 50
    End If
    If grdGrilla.Col = 2 Then
        KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
        txtEdit.MaxLength = 10
    End If
    If grdGrilla.Col = 3 Then
        KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
        txtEdit.MaxLength = 10
    End If
    CarTexto KeyAscii
End Sub

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF1 Then
        frmBuscar.TipoBusqueda = 2
        frmBuscar.CodListaPrecio = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        frmBuscar.Show vbModal
    End If

    If KeyCode = vbKeyReturn Then
       If grdGrilla.Col = 0 Or grdGrilla.Col = 1 Or grdGrilla.Col = 2 Or grdGrilla.Col = 3 Then
            Select Case grdGrilla.Col
            Case 0, 1
            
                If Trim(txtEdit) = "" Then
                    txtEdit = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                txtEdit.Text = Replace(txtEdit.Text, "'", "´")
                If cboListaPrecio.ListIndex = 0 Then 'Busca en los Productos
                    If lblEstadoRemito.Caption = "PENDIENTE" Then
                        sql = "SELECT P.PTO_CODIGO,P.PTO_DESCRI, P.PTO_PRECIO, R.RUB_DESCRI, L.LNA_DESCRI, RE.TPRE_DESCRI"
                        sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L, TIPO_PRESENTACION RE"
                        sql = sql & " WHERE"
                        If grdGrilla.Col = 0 Then
                            sql = sql & " P.PTO_CODIGO LIKE '" & txtEdit.Text & "'"
                        Else
                            sql = sql & " P.PTO_DESCRI LIKE '" & Trim(txtEdit) & "%'"
                        End If
                        sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO AND P.RUB_CODIGO=R.RUB_CODIGO"
                        sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
                        sql = sql & " AND P.TPRE_CODIGO=RE.TPRE_CODIGO"
                        sql = sql & " AND P.PTO_ESTADO <> 1"
                    '*********
                    Else
                        sql = "SELECT DRC.PTO_CODIGO,DRC.DRC_DETALLE, DRC.DRC_PRECIO, R.RUB_DESCRI, L.LNA_DESCRI, RE.TPRE_DESCRI"
                        sql = sql & " FROM DETALLE_REMITO_CLIENTE DRC, PRODUCTO P,LINEAS L,RUBROS R,TIPO_PRESENTACION RE"
                        sql = sql & " WHERE"
                        sql = sql & " DRC.PTO_CODIGO = P.PTO_CODIGO AND "
                        If grdGrilla.Col = 0 Then
                            sql = sql & " DRC.PTO_CODIGO LIKE '" & txtEdit.Text & "'"
                        Else
                            sql = sql & " DRC.DRC_DETALLE LIKE '" & Trim(txtEdit) & "%'"
                        End If
                        sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO AND P.RUB_CODIGO=R.RUB_CODIGO"
                        sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
                        sql = sql & " AND P.TPRE_CODIGO=RE.TPRE_CODIGO"
                        
                    End If
                Else  ' Busca en un Lista de Precios
                    sql = "SELECT P.PTO_CODIGO, P.PTO_DESCRI, P.PTO_PRECIO, R.RUB_DESCRI, L.LNA_DESCRI"
                    sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L,TIPO_PRESENTACION RE"
                    sql = sql & " WHERE"
                    'sql = sql & " AND P.LIS_CODIGO=" & cboListaPrecio.ItemData(cboListaPrecio.ListIndex) & ""
                    'sql = sql & " AND P.PTO_CODIGO=D.PTO_CODIGO"
                    sql = sql & "  P.RUB_CODIGO=R.RUB_CODIGO"
                    sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
                    sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
                    sql = sql & " AND P.TPRE_CODIGO=RE.TPRE_CODIGO"
                    sql = sql & " AND P.PTO_ESTADO <> 1"
                    
                    If grdGrilla.Col = 0 Then
                         sql = sql & " AND P.PTO_CODIGO LIKE '" & txtEdit.Text & "'"
                    Else
                         If Not IsNumeric(txtEdit.Text) Then
                             sql = sql & " AND P.PTO_DESCRI LIKE '" & Trim(txtEdit) & "%'"
                         End If
                    End If
                End If
                rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If rec.EOF = False Then
                    If rec.RecordCount > 1 Then
                        grdGrilla.SetFocus
                        frmBuscar.TipoBusqueda = 2
                        'LE DIGO EN QUE LISTA DE PRECIO BUSCAR LOS PRECIOS
                        If cboListaPrecio.ListIndex <> 0 Then '<TODOS>
                            frmBuscar.CodListaPrecio = cboLPrecioRep.ItemData(cboLPrecioRep.ListIndex)
                        Else
                                frmBuscar.CodListaPrecio = 0 ' BUSCA EN LA TABLA PRODUCTOS
                        End If
                        If Not IsNumeric(txtEdit.Text) Then
                            frmBuscar.TxtDescriB.Text = txtEdit.Text
                        'Else
                            'frmBuscar.TxtDescriB.Text = ""
                        End If
                        
                        If frmBuscar.TxtDescriB.Text <> "" Then
                            frmBuscar.Show vbModal
                        
                            grdGrilla.Col = 0
                            grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0)
                            grdGrilla.Col = 1
                            grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
                            grdGrilla.Col = 3
                            grdGrilla.Text = Valido_Importe(frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 2))
                            grdGrilla.Col = 5
                            grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 3)
                            grdGrilla.Col = 6
                            grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 4)
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = grdGrilla.RowSel
                            grdGrilla.Col = 2
                        End If
                    Else
                        grdGrilla.Col = 0
                        grdGrilla.Text = Trim(rec!PTO_CODIGO)
                        If lblEstadoRemito.Caption = "PENDIENTE" Then
                            grdGrilla.Col = 1
                            grdGrilla.Text = Trim(rec!PTO_DESCRI)
                            grdGrilla.Col = 3
                            If cboListaPrecio.ListIndex = 0 Then
                                grdGrilla.Text = Valido_Importe(Trim(rec!PTO_PRECIO))
                            Else
                                grdGrilla.Text = Valido_Importe(Trim(rec!PTO_PRECIO))
                            End If
                        Else
                            grdGrilla.Col = 1
                            grdGrilla.Text = Trim(rec!DRC_DETALLE)
                            grdGrilla.Col = 3
                            grdGrilla.Text = Valido_Importe(Trim(rec!DRC_PRECIO))
                        End If
                        
                        grdGrilla.Col = 5
                        grdGrilla.Text = Trim(rec!RUB_DESCRI)
                        grdGrilla.Col = 6
                        grdGrilla.Text = Trim(rec!LNA_DESCRI)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = grdGrilla.RowSel
                        grdGrilla.Col = 2
                    End If
                        If BuscoRepetetidos(grdGrilla.TextMatrix(grdGrilla.RowSel, 0), grdGrilla.RowSel) = False Then
                         grdGrilla.Col = 0
                         grdGrilla_KeyDown vbKeyDelete, 0
                        End If
                Else
                        MsgBox "No se ha encontrado el Producto", vbExclamation, TIT_MSGBOX
                        txtEdit.Text = ""
                        LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                        grdGrilla.Col = 0
                End If
                rec.Close
                Screen.MousePointer = vbNormal
            Case 2
                If Trim(txtEdit) = "" Then
                    grdGrilla.Text = "1"
                    txtEdit.Text = "1"
                End If
                'ESTO ES IMPORTANTE REVISAR PARA EL STOCK
'                If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
'                    Set Rec1 = New ADODB.Recordset
'                    sql = "SELECT P.PTO_STKMIN, DS.DST_STKFIS, DS.DST_STKPEN"
'                    sql = sql & " FROM PRODUCTO P, DETALLE_STOCK DS"
'                    sql = sql & " WHERE P.PTO_CODIGO LIKE '" & grdGrilla.TextMatrix(grdGrilla.RowSel, 0) & "'"
'                    sql = sql & " AND P.PTO_CODIGO = DS.PTO_CODIGO"
'                    sql = sql & " AND STK_CODIGO=" & cboStock.ItemData(cboStock.ListIndex)
'
'                    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'                    If Rec1.EOF = False Then
'                        If (CInt(Rec1!DST_STKFIS) - CInt(txtEdit.Text)) < CInt(Rec1!PTO_STKMIN) Then
'                            MsgBox "El producto esta por debajo del Stock Minimo" & Chr(13) & Chr(13) & _
'                            " Stock Minimo=" & Rec1!PTO_STKMIN & Chr(13) & _
'                            " Stock Pendiente=" & Rec1!DST_STKPEN & Chr(13) & _
'                            " Stock Fisico=" & Rec1!DST_STKFIS & Chr(13) & _
'                            " Stock Fisico - Cantidad=" & (CInt(Rec1!DST_STKFIS) - CInt(txtEdit.Text)), vbExclamation, TIT_MSGBOX
'                        End If
'                    End If
'                    Rec1.Close
'                End If
                grdGrilla_LeaveCell
                TotalPEDIDO
                grdGrilla.SetFocus
            Case 3
                If Trim(txtEdit) <> "" Then
                    txtEdit.Text = Valido_Importe(txtEdit)
                    grdGrilla_LeaveCell
                    TotalPEDIDO
                Else
                    MsgBox "Debe ingresar el Importe", vbExclamation, TIT_MSGBOX
                    grdGrilla.Col = 3
                End If
            End Select
        End If
        grdGrilla.SetFocus
    End If
    If KeyCode = vbKeyEscape Then
       txtEdit.Visible = False
       grdGrilla.SetFocus
    End If
End Sub
Private Function TotalPEDIDO()
    Dim TOTAL As Double
    Dim subtotal As Double
    Dim impIva As Double
    Dim Cant As Double
    Dim Importe As Double
    grdGrilla_LeaveCell
    'Me.txtTotal.Text = 0
    
    If Me.grdGrilla.TextMatrix(Me.grdGrilla.RowSel, 2) = "" Then
        Cant = 0
    'Else
    '    Cant = 1
    End If
    If Me.grdGrilla.TextMatrix(Me.grdGrilla.RowSel, 3) = "" Then
        Importe = 0
    'Else
    '    Importe = 1
    End If
    
    'Cant = IIf(Me.grdGrilla.TextMatrix(Me.grdGrilla.RowSel, 2) = "", 0, CDbl(Me.grdGrilla.TextMatrix(Me.grdGrilla.RowSel, 2)))
    If Me.grdGrilla.TextMatrix(Me.grdGrilla.RowSel, 2) <> "" And Me.grdGrilla.TextMatrix(Me.grdGrilla.RowSel, 3) <> "" Then
        Me.grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = Format(CDbl(Me.grdGrilla.TextMatrix(Me.grdGrilla.RowSel, 2)) * CDbl(Me.grdGrilla.TextMatrix(Me.grdGrilla.RowSel, 3)), "#,##0.00")
    End If
    txtSubtotal.Text = Valido_Importe(SumaTotal)
'    txtTotal.Text = txtSubtotal.Text
    
'    txtImporteIva.Text = (CDbl(txtSubtotal.Text) * CDbl(txtPorcentajeIva.Text)) / 100
'    txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)

    'txtImporteIva.Text = (CDbl(txtSubTotal.Text) * CDbl(txtPorcentajeIva.Text)) / 100
    'txtImporteIva.Text = CDbl(txtSubtotal.Text) / CDbl("1," & Int(txtPorcentajeIva.Text))
    txtSubtotal.Text = Valido_Importe(CDbl(txtSubtotal.Text) / CDbl("1," & Int(txtPorcentajeIva.Text)))
    txtImporteIva.Text = Valido_Importe(SumaTotal - txtSubtotal.Text)
    subtotal = txtSubtotal
    impIva = txtImporteIva
    TOTAL = subtotal + impIva
    txtTotal.Text = Format(TOTAL, "0.00")

End Function
Private Function SumaTotal() As Double
    Dim VTotal As Double
    VTotal = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 4) <> "" Then
            VTotal = VTotal + (CDbl(grdGrilla.TextMatrix(I, 2)) * CDbl(grdGrilla.TextMatrix(I, 3)))
        End If
    Next
    SumaTotal = Valido_Importe(CStr(VTotal))
End Function
Private Function BuscoRepetetidos(Codigo As String, Linea As Integer) As Boolean
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" And grdGrilla.TextMatrix(I, 0) <> "----------" Then
            If Codigo = grdGrilla.TextMatrix(I, 0) And (I <> Linea) Then
                MsgBox "El producto ya fue elegido anteriormente", vbExclamation, TIT_MSGBOX
                BuscoRepetetidos = False
                Exit Function
            End If
        End If
    Next
    BuscoRepetetidos = True
End Function

Private Sub txtNroNotaPedido_GotFocus()
    SelecTexto txtNroNotaPedido
End Sub

Private Sub txtNroNotaPedido_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroNotaPedido_LostFocus()
    
    If txtNroNotaPedido.Text <> "" Then
        sql = "SELECT NP.*, E.EST_DESCRI"
        sql = sql & " FROM NOTA_PEDIDO NP, ESTADO_DOCUMENTO E"
        sql = sql & " WHERE NP.NPE_NUMERO=" & XN(txtNroNotaPedido)
        If FechaNotaPedido.Text <> "" Then
            sql = sql & " AND NP.NPE_FECHA=" & XDQ(FechaNotaPedido)
        End If
        sql = sql & " AND NP.EST_CODIGO=E.EST_CODIGO"
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If Rec2.EOF = False Then
            If Rec2.RecordCount > 1 Then
                MsgBox "Hay mas de una Presupuesto con el Número: " & txtNroNotaPedido.Text, vbInformation, TIT_MSGBOX
                Rec2.Close
                cmdBuscarNotaPedido_Click
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            lblestado.Caption = "Buscando..."
            
            'CARGO CABECERA DE LA NOTA DE PEDIDO
            FechaNotaPedido.Text = Rec2!NPE_FECHA
            'grillaNotaPedido.TextMatrix(0, 1) = BuscoCliente(Rec2!CLI_CODIGO)
            'grillaNotaPedido.TextMatrix(1, 1) = BuscoSucursal(Rec2!SUC_CODIGO, Rec2!CLI_CODIGO)
            'grillaNotaPedido.TextMatrix(2, 1) = BuscoVendedor(Rec2!VEN_CODIGO)
            'lblEstadoNotaPedido.Caption = "Estado: " & Rec2!EST_DESCRI
            TxtCodigoCli.Text = Rec2!CLI_CODIGO
            TxtCodigoCli_LostFocus
            If Rec2!EST_CODIGO <> 1 Then
                MsgBox "El Presupuesto número: " & txtNroNotaPedido.Text & Chr(13) & Chr(13) & _
                       "No puede ser asignado al Remito por su estado (" & Rec2!EST_DESCRI & ")", vbExclamation, TIT_MSGBOX
                LimpiarNotaPedido
                cmdGrabar.Enabled = False
                Screen.MousePointer = vbNormal
                lblestado.Caption = ""
                Rec2.Close
                Exit Sub
            Else
                cmdGrabar.Enabled = True
            End If
            Rec2.Close
            
        '-----BUSCO LOS DATOS DEL DETALLE DE LA NOTA DE PEDIDO - PRESUPUESTO---------
            sql = "SELECT DNP.*,P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI"
            sql = sql & " FROM DETALLE_NOTA_PEDIDO DNP, PRODUCTO P, RUBROS R, LINEAS L "
            sql = sql & " WHERE DNP.NPE_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND DNP.NPE_FECHA=" & XDQ(FechaNotaPedido)
            sql = sql & " AND DNP.PTO_CODIGO=P.PTO_CODIGO"
            sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
            sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO"
            sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
            sql = sql & " ORDER BY DNP.DNP_NROITEM"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                I = 1
                Do While Rec1.EOF = False
                    grdGrilla.TextMatrix(I, 0) = Rec1!PTO_CODIGO
                    grdGrilla.TextMatrix(I, 1) = Rec1!PTO_DESCRI
                    grdGrilla.TextMatrix(I, 2) = Rec1!DNP_CANTIDAD
                    grdGrilla.TextMatrix(I, 3) = Valido_Importe(Rec1!DNP_PRECIO)
                    grdGrilla.TextMatrix(I, 4) = Valido_Importe(Rec1!DNP_PRECIO * Rec1!DNP_CANTIDAD)
                    grdGrilla.TextMatrix(I, 5) = Rec1!RUB_DESCRI
                    grdGrilla.TextMatrix(I, 6) = Rec1!LNA_DESCRI
                    grdGrilla.TextMatrix(I, 7) = Rec1!DNP_NROITEM
                    I = I + 1
                    Rec1.MoveNext
                Loop
            End If
            Rec1.Close
            '--------------------------------------------------
            Screen.MousePointer = vbNormal
            lblestado.Caption = ""
            'chkRemitoSinFactura.SetFocus
        Else
            MsgBox "El Presupuesto no existe", vbExclamation, TIT_MSGBOX
            If Rec2.State = 1 Then Rec2.Close
            LimpiarNotaPedido
        End If
    End If
End Sub

Private Sub LimpiarNotaPedido()
    txtNroNotaPedido.Text = ""
    FechaNotaPedido.Text = ""
    'grillaNotaPedido.TextMatrix(0, 1) = ""
    'grillaNotaPedido.TextMatrix(1, 1) = ""
    'grillaNotaPedido.TextMatrix(2, 1) = ""
'    txtNroNotaPedido.SetFocus
End Sub
Private Function BuscoVendedor(Codigo As String) As String
    sql = "SELECT VEN_NOMBRE"
    sql = sql & " FROM VENDEDOR"
    sql = sql & " WHERE VEN_CODIGO=" & XN(Codigo)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        BuscoVendedor = Trim(rec!VEN_NOMBRE)
    Else
        BuscoVendedor = "No se encontro el Vendedor"
    End If
    rec.Close
End Function

Private Function BuscoCliente(Codigo As String) As String
        sql = "SELECT CLI_RAZSOC FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(Codigo)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            BuscoCliente = rec!CLI_RAZSOC
        Else
            BuscoCliente = "No se encontro el Cliente"
        End If
        rec.Close
End Function

Private Function BuscoSucursal(CodigoSuc As String, CodigoCli As String) As String
        sql = "SELECT * FROM SUCURSAL"
        sql = sql & " WHERE SUC_CODIGO=" & XN(CodigoSuc)
        sql = sql & " AND CLI_CODIGO=" & XN(CodigoCli)
        
        Set Rec1 = New ADODB.Recordset
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            BuscoSucursal = Rec1!SUC_DESCRI
        Else
            BuscoSucursal = "No se encontro la Sucursal"
        End If
        Rec1.Close
End Function

Private Sub txtObservaciones_GotFocus()
    SelecTexto txtObservaciones
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtPorc_GotFocus()
    SelecTexto txtPorc
End Sub

Private Sub txtPorc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPorc, KeyAscii)
End Sub

Private Sub txtPorc_LostFocus()
    If txtPorc.Text <> "" Then
        Dim TOTAL As Double
        Dim subtotal As Double
        Dim impIva As Double
        
        TotalPEDIDO
'        txtSubTotal.Text = Valido_Importe(SumaTotal)
'        txtSubTotal.Text = txtSubTotal.Text + (txtSubTotal.Text * txtPorc) / 100
'        'txtImporteIva.Text = (CDbl(txtSubTotal.Text) * CDbl(txtPorcentajeIva.Text)) / 100
'        txtImporteIva.Text = CDbl(txtSubTotal.Text) / CDbl("1," & Int(txtPorcentajeIva.Text))
'        txtSubTotal.Text = Valido_Importe(txtSubTotal.Text - txtImporteIva.Text)
'        txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)
'        subtotal = txtSubTotal
'        impIva = txtImporteIva
'        TOTAL = subtotal + impIva
'        txtTotal.Text = Format(TOTAL, "0.00")
    End If
End Sub

Private Sub txtVendedor_Change()
    If txtVendedor.Text = "" Then
        txtDesVen.Text = ""
    End If
End Sub

Private Sub txtVendedor_GotFocus()
    SelecTexto txtVendedor
End Sub

Private Sub txtVendedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtVendedor_LostFocus()
    If txtVendedor.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT VEN_NOMBRE"
        sql = sql & " FROM VENDEDOR"
        sql = sql & " WHERE VEN_CODIGO=" & XN(txtVendedor)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            txtDesVen.Text = Trim(rec!VEN_NOMBRE)
        Else
            MsgBox "El Vendedor no existe", vbExclamation, TIT_MSGBOX
            txtDesVen.Text = ""
            txtVendedor.SetFocus
            Exit Sub
        End If
        rec.Close
    End If
'    If chkFecha.Value = Unchecked And ActiveControl.Name <> "cmdNuevo" _
'        And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub
Private Sub CortarCadena(Renglon As Double, Cadena As String)
    Dim salto, Max, inf, I, leer, leerb As Integer
    Dim salto1, salto2, salto3, salto4, salto5, salto6, salto7 As Integer
    Dim salto1b, salto2b, salto3b, salto4b, salto5b, salto6b, salto7b As Integer
    Dim cadena1 As String
    Dim cadena2 As String
    Dim cadena3 As String
    Dim cadena4 As String
    Dim cadena5 As String
    Dim cadena6 As String
    Dim cadena7 As String
    Dim cadena8 As String
    
    cadena1 = ""
    cadena2 = ""
    cadena3 = ""
    cadena4 = ""
    cadena5 = ""
    cadena6 = ""
    cadena7 = ""
    
    salto = 1
    Max = 36 * salto
    inf = Max - 10
    'falta = 0
    'If Len(cadena) > 35 Then
        For I = 1 To Len(Cadena)
            If (Mid(Cadena, I, 1) = " ") And (I > inf) And (I < Max) Or (I > Max) Then
                
                    If salto = 1 Then
                    salto1 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        cadena1 = Left(Cadena, I)
                        cadena2 = Mid(Cadena, salto1, Max)
                        'Imprimir 3.2, renglon, False, Left(grdGrilla.TextMatrix(I, 1), 36) 'descripcion
                        'Imprimir 3.2, renglon + 0.5, False, Mid(grdGrilla.TextMatrix(I, 1), 37, 36) 'descripcion
                    
                    Else
                        cadena1 = Left(Cadena, I)
                    End If
                      'descripcion
                End If
                If salto = 2 Then
                    leer = I - salto1
                    salto2 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto1b = I
                        leerb = Len(Cadena) + 1 - salto1b
                        cadena2 = Mid(Cadena, salto1, leer)
                        cadena3 = Mid(Cadena, salto1b, leerb)  'descripcion
                    Else
                        cadena2 = Mid(Cadena, salto1, leer)  'descripcion
                    End If
                    
                    
                End If
                If salto = 3 Then
                    Max = 36 + I
                    inf = Max - 10
                    leer = I - salto2
                    salto3 = I
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto2b = I
                        leerb = Len(Cadena) + 1 - salto2b
                        cadena3 = Mid(Cadena, salto2, leer)
                        cadena4 = Mid(Cadena, salto2b, leerb)
                    Else
                        cadena3 = Mid(Cadena, salto2, leer)  'descripcion
                    End If
                    
                End If
                If salto = 4 Then
                    leer = I - salto3
                    salto4 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto3b = I
                        leerb = Len(Cadena) + 1 - salto3b
                        cadena4 = Mid(Cadena, salto3, leer)
                        cadena5 = Mid(Cadena, salto3b, leerb)  'descripcion
                    Else
                         cadena4 = Mid(Cadena, salto3, leer)  'descripcion
                    End If
                    
                    
                End If
                If salto = 5 Then
                    leer = I - salto4
                    salto5 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto4b = I
                        leerb = Len(Cadena) + 1 - salto4b
                        cadena5 = Mid(Cadena, salto4, leer)
                        cadena6 = Mid(Cadena, salto4b, leerb)  'descripcion
                    Else
                        cadena5 = Mid(Cadena, salto4, leer)  'descripcion
                    End If
                    
                    
                End If
                If salto = 6 Then
                    leer = I - salto5
                    salto6 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto5b = I
                        leerb = Len(Cadena) + 1 - salto5b
                        cadena6 = Mid(Cadena, salto5, leer)
                        cadena7 = Mid(Cadena, salto5b, leerb)  'descripcion
                        
                    Else
                        cadena6 = Mid(Cadena, salto5, leer)  'descripcion
                    End If
                    
                End If
                If salto = 7 Then
                    leer = I - salto6
                    salto7 = I
                    Max = 36 + I
                    inf = Max - 10
                    If Max > Len(Cadena) Then
                        inf = Len(Cadena)
                        Max = Len(Cadena)
                        salto6b = I
                        leerb = Len(Cadena) + 1 - salto6b
                        cadena7 = Mid(Cadena, salto6, leer)
                        cadena8 = Mid(Cadena, salto6b, leerb)  'descripcion
                        
                    Else
                        cadena7 = Mid(Cadena, salto6, leer)  'descripcion
                    End If
                    
                End If
                
                salto = salto + 1
                'Max = valor * salto
                'inf = Max - 10
                
            End If
        Next
    
        Imprimir 3.2, Renglon, False, cadena1
        Imprimir 3.2, Renglon + 0.5, False, cadena2
        Imprimir 3.2, Renglon + 1, False, cadena3
        Imprimir 3.2, Renglon + 1.5, False, cadena4
        Imprimir 3.2, Renglon + 2, False, cadena5
        Imprimir 3.2, Renglon + 2.5, False, cadena6
        Imprimir 3.2, Renglon + 3, False, cadena7
        Imprimir 3.2, Renglon + 3.5, False, cadena8
    'Else
    '    cadena1 = cadena
    '    MsgBox cadena1
    'End If
    
End Sub

