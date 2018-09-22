VERSION 5.00
Begin VB.Form frmAdmListados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrador de Listados"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraVentas 
      Caption         =   "Seleccione el Listado de Ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   5295
      Begin VB.OptionButton OptEstVentas 
         Caption         =   "Estadisticas de Ventas"
         Height          =   195
         Left            =   480
         TabIndex        =   33
         Top             =   480
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optRemitos 
         Caption         =   "de Remitos"
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   945
         Width           =   2175
      End
      Begin VB.OptionButton optVentasPorCli 
         Caption         =   "Ventas por Cliente"
         Height          =   195
         Left            =   480
         TabIndex        =   17
         Top             =   3495
         Width           =   2175
      End
      Begin VB.OptionButton optVentasxVend 
         Caption         =   "Ventas por Vendedor"
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   3975
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton optFacPendClie 
         Caption         =   "Facturas Pendientes por Cliente"
         Height          =   195
         Left            =   480
         TabIndex        =   14
         Top             =   2325
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.OptionButton optRecibos 
         Caption         =   "de Recibos"
         Height          =   195
         Left            =   480
         TabIndex        =   13
         Top             =   1425
         Width           =   2175
      End
      Begin VB.OptionButton optFacturas 
         Caption         =   "de Facturas"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   690
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton optCliente 
         Caption         =   "de Clientes"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   1920
         Width           =   2175
      End
      Begin VB.OptionButton optLibroDebito 
         Caption         =   "Libro Débito Fiscal"
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   1920
         Width           =   2175
      End
      Begin VB.OptionButton optPagosRealCli 
         Caption         =   "Pagos Realizados por Cliente"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   2520
         Width           =   3015
      End
      Begin VB.OptionButton optCantVendidas 
         Caption         =   "Cantidades Vendidas"
         Height          =   195
         Left            =   480
         TabIndex        =   16
         Top             =   3000
         Width           =   2895
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "frmAdmListados.frx":0000
      Height          =   720
      Left            =   6585
      Picture         =   "frmAdmListados.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5040
      Width           =   1020
   End
   Begin VB.Frame Frame1 
      Caption         =   "Módulos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton cmdFondos 
         Caption         =   "&Fondos"
         DownPicture     =   "frmAdmListados.frx":0614
         Height          =   1095
         Left            =   120
         Picture         =   "frmAdmListados.frx":22DE
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton cmdMercaderias 
         Caption         =   "&Mercaderías"
         DownPicture     =   "frmAdmListados.frx":3FA8
         Height          =   1095
         Left            =   120
         Picture         =   "frmAdmListados.frx":5C72
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2475
         Width           =   1935
      End
      Begin VB.CommandButton cmdCompras 
         Caption         =   "&Compras"
         DownPicture     =   "frmAdmListados.frx":793C
         Height          =   1095
         Left            =   120
         Picture         =   "frmAdmListados.frx":9606
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1365
         Width           =   1935
      End
      Begin VB.CommandButton cmdVentas 
         Caption         =   "&Ventas"
         DownPicture     =   "frmAdmListados.frx":B2D0
         Height          =   1095
         Left            =   120
         Picture         =   "frmAdmListados.frx":CF9A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraCompras 
      Caption         =   "Seleccione el Listado de Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   2400
      TabIndex        =   18
      Top             =   120
      Width           =   5295
      Begin VB.OptionButton optPagosRealProv 
         Caption         =   "Recibos de Proveedores"
         Height          =   195
         Left            =   480
         TabIndex        =   26
         Top             =   2808
         Width           =   3015
      End
      Begin VB.OptionButton optLibroFiscal 
         Caption         =   "Libro Créditos Fiscal"
         Height          =   195
         Left            =   480
         TabIndex        =   25
         Top             =   1704
         Width           =   2175
      End
      Begin VB.OptionButton optProveedores 
         Caption         =   "de Proveedores"
         Height          =   195
         Left            =   480
         TabIndex        =   24
         Top             =   2256
         Width           =   2175
      End
      Begin VB.OptionButton optRemitosProv 
         Caption         =   "de Remitos de Proveedores"
         Height          =   195
         Left            =   480
         TabIndex        =   23
         Top             =   600
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton optFacturasProv 
         Caption         =   "de Facturas de Proveedores"
         Height          =   195
         Left            =   480
         TabIndex        =   22
         Top             =   690
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.OptionButton optOrdenPago 
         Caption         =   "de Ordenes de Pagos"
         Height          =   195
         Left            =   480
         TabIndex        =   21
         Top             =   1152
         Width           =   2175
      End
      Begin VB.OptionButton optFacPendporProv 
         Caption         =   "Facturas Pendientes por Proveedor"
         Height          =   195
         Left            =   480
         TabIndex        =   20
         Top             =   1980
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.OptionButton optResumenCompra 
         Caption         =   "Resumen de Compras"
         Height          =   195
         Left            =   480
         TabIndex        =   19
         Top             =   3360
         Width           =   2895
      End
   End
   Begin VB.Frame fraFondos 
      Caption         =   "Seleccione el Listado de Fondos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   2400
      TabIndex        =   30
      Top             =   120
      Width           =   5295
      Begin VB.OptionButton optCheqTerceros 
         Caption         =   "de Cheques de Terceros"
         Height          =   195
         Left            =   480
         TabIndex        =   32
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton optCheqPropios 
         Caption         =   "de Cheques Propios"
         Height          =   195
         Left            =   480
         TabIndex        =   31
         Top             =   690
         Width           =   2535
      End
   End
   Begin VB.Frame fraMercaderias 
      Caption         =   "Seleccione el Listado de Mercaderías"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   2400
      TabIndex        =   27
      Top             =   120
      Width           =   5295
      Begin VB.OptionButton optListaPrecios 
         Caption         =   "Lista de Precios"
         Height          =   195
         Left            =   480
         TabIndex        =   29
         Top             =   690
         Width           =   2535
      End
      Begin VB.OptionButton optStock 
         Caption         =   "Listado de Stocks"
         Height          =   195
         Left            =   480
         TabIndex        =   28
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "frmAdmListados.frx":EC64
      Height          =   720
      Left            =   5520
      Picture         =   "frmAdmListados.frx":EF6E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5040
      Width           =   1020
   End
End
Attribute VB_Name = "frmAdmListados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCompras_Click()
    fraVentas.Visible = False
    fraCompras.Visible = True
    fraMercaderias.Visible = False
    fraFondos.Visible = False
    
    cmdVentas.Enabled = True
    cmdCompras.Enabled = False
    cmdMercaderias.Enabled = True
    cmdFondos.Enabled = True
    
    optRemitosProv.SetFocus
End Sub

Private Sub cmdFondos_Click()
    fraVentas.Visible = False
    fraCompras.Visible = False
    fraMercaderias.Visible = False
    fraFondos.Visible = True
    
    cmdVentas.Enabled = True
    cmdCompras.Enabled = True
    cmdMercaderias.Enabled = True
    cmdFondos.Enabled = False
    
    optCheqTerceros.SetFocus
End Sub

Private Sub CmdGrabar_Click()
    '************ VENTAS *****************
    If cmdVentas.Enabled = False Then
        'If optComposturas.Value = True Then
        '    MsgBox "Listado en Construccion", vbExclamation
        '    ' frmListadocomposturas
        'End If
        'If optRevelados.Value = True Then
        '    MsgBox "Listado en Construccion", vbExclamation
        '    ' frmListadocomposturas
        'End If
        If OptEstVentas.Value = True Then
            frmEstadisticaVentas.Show vbModal
        End If
        If optRemitos.Value = True Then
            frmListadoRemitoCliente.Show vbModal
        End If
        If optFacturas.Value = True Then
            MsgBox "Listado en Construccion", vbExclamation
            'frmListadoRemitoCliente.Show vbModal
        End If
        If optRecibos.Value = True Then
            frmListadoReciboCliente.Show vbModal
        End If
        If optLibroDebito.Value = True Then
            'frmLibroIvaVentas.Show 'vbModal
            frmLibroVentas2.Show vbModal
        End If
        If optCliente.Value = True Then
            frmListadoClientes.Show vbModal
        End If
        If optFacPendClie.Value = True Then
            frmListadoFacturasPendientePorCliente.Show vbModal
        End If
        If optPagosRealCli.Value = True Then
            frmListadoPagosPorCliente.Show vbModal
        End If
        If optCantVendidas.Value = True Then
            frmListadoCantidadesVendidas.Show vbModal
        End If
        If optVentasPorCli.Value = True Then
            frmListadoVentasPorCliente.Show vbModal
        End If
        If optVentasxVend.Value = True Then
            frmListadoVentasPorVendedor.Show vbModal
        End If
        'If optDestinoComp.Value = True Then
        '    MsgBox "Listado en Construccion", vbExclamation
        '    'frmListadoRemitoCliente.Show vbModal
        'End If
    End If
    
    '************** COMPRAS ********************
    If cmdCompras.Enabled = False Then
        If optRemitosProv.Value = True Then
            frmListadoRemitoProveedor.Show vbModal
        End If
        If optFacturasProv.Enabled = False Then
            MsgBox "Listado en Construccion", vbExclamation
        End If
        If optOrdenPago.Value = True Then
            MsgBox "Listado en Construccion", vbExclamation
        End If
        If optLibroFiscal.Value = True Then
'            frmLibroIvaCompras.Show vbModal
            frmLibroCompras2.Show vbModal
        End If
        If optProveedores.Value = True Then
            frmListadoProvedores.Show vbModal
        End If
        If optFacPendporProv.Value = True Then
            frmListadoFacturasPendientePorProveedor.Show vbModal
        End If
        If optPagosRealProv.Value = True Then
            frmListadoPagosPorProveedor.Show vbModal
        End If
        If optResumenCompra.Value = True Then
            frmListadoComprasProveedores.Show vbModal
        End If
    End If
    
    '************** MERCADERIAS *****************
    If cmdMercaderias.Enabled = False Then
        If optStock.Value = True Then
            Stock = 1 'Consulta Stock
            frmControlStock.Show vbModal
        End If
        If optListaPrecios.Value = True Then
            Consulta = 1
            FrmListadePrecios.Show vbModal
        End If
    End If
    
    '**************** FONDOS ********************
    If cmdFondos.Enabled = False Then
        If optCheqTerceros.Value = True Then
            FrmListCheques.Show vbModal
        End If
        If optCheqPropios.Value = True Then
            FrmListChequesPropios.Show vbModal
        End If
    End If
    
End Sub

Private Sub cmdMercaderias_Click()
    fraVentas.Visible = False
    fraCompras.Visible = False
    fraMercaderias.Visible = True
    fraFondos.Visible = False
    
    cmdVentas.Enabled = True
    cmdCompras.Enabled = True
    cmdMercaderias.Enabled = False
    cmdFondos.Enabled = True
    
    optStock.SetFocus
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdVentas_Click()
    fraVentas.Visible = True
    fraCompras.Visible = False
    fraMercaderias.Visible = False
    fraFondos.Visible = False
    
    cmdVentas.Enabled = False
    cmdCompras.Enabled = True
    cmdMercaderias.Enabled = True
    cmdFondos.Enabled = True
    'optComposturas.SetFocus
    
End Sub

Private Sub Form_Activate()
    'cmdVentas.SetFocus
    ' optComposturas.SetFocus
End Sub

Private Sub Form_Load()
    Centrar_pantalla Me
    
    fraVentas.Visible = True
    fraCompras.Visible = False
    fraMercaderias.Visible = False
    fraFondos.Visible = False
    cmdVentas.Enabled = False
    cmdCompras.Enabled = True
    cmdMercaderias.Enabled = True
    cmdFondos.Enabled = True
    
   
    
End Sub

