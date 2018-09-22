VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MENU 
   BackColor       =   &H80000010&
   ClientHeight    =   7590
   ClientLeft      =   435
   ClientTop       =   2745
   ClientWidth     =   11400
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   11760
      Left            =   0
      Picture         =   "Menu.frx":1CCA
      ScaleHeight     =   11700
      ScaleWidth      =   11340
      TabIndex        =   1
      Top             =   0
      Width           =   11400
      Begin ComctlLib.Toolbar tbrPrincipal 
         Height          =   390
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   17340
         _ExtentX        =   30586
         _ExtentY        =   688
         ButtonWidth     =   635
         ButtonHeight    =   582
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   16
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.ToolTipText     =   "Remitos de Clientes"
               Object.Tag             =   ""
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Visible         =   0   'False
               Key             =   ""
               Object.ToolTipText     =   "ABM de Clientes"
               Object.Tag             =   ""
               ImageIndex      =   1
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.ToolTipText     =   "Recibo de Clientes"
               Object.Tag             =   ""
               ImageIndex      =   13
            EndProperty
            BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Visible         =   0   'False
               Key             =   ""
               Object.ToolTipText     =   "Agenda de Clientes"
               Object.Tag             =   ""
               ImageIndex      =   9
            EndProperty
            BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Visible         =   0   'False
               Key             =   ""
               Object.ToolTipText     =   "Compras"
               Object.Tag             =   ""
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Visible         =   0   'False
               Key             =   ""
               Object.ToolTipText     =   "Proveedores"
               Object.Tag             =   ""
               ImageIndex      =   15
            EndProperty
            BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.ToolTipText     =   "Productos"
               Object.Tag             =   ""
               ImageIndex      =   4
            EndProperty
            BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Visible         =   0   'False
               Key             =   ""
               Object.ToolTipText     =   "Precios"
               Object.Tag             =   ""
               ImageIndex      =   3
            EndProperty
            BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Visible         =   0   'False
               Key             =   ""
               Object.ToolTipText     =   "Listado de Stock Faltantes"
               Object.Tag             =   ""
               ImageIndex      =   16
            EndProperty
            BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.ToolTipText     =   "Listados de Ventas"
               Object.Tag             =   ""
               ImageIndex      =   10
            EndProperty
            BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.ToolTipText     =   "Salir"
               Object.Tag             =   ""
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   240
         Top             =   720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   17
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":A25F
               Key             =   "Clientes"
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":A579
               Key             =   "Remitos"
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":A753
               Key             =   "Precios"
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":AA6D
               Key             =   "Productos"
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":AD87
               Key             =   "Compras"
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":B0A1
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":B3BB
               Key             =   "Salir"
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":B6D5
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":B9EF
               Key             =   "Agenda de Clientes"
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":BD09
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":C023
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":C33D
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":C657
               Key             =   "Cuenta Corriente"
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":C971
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":CC8B
               Key             =   "Proveedores"
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":CFA5
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":D2BF
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.StatusBar B1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   7245
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "NÚM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   1409
            MinWidth        =   1409
            TextSave        =   "MAYÚS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Alignment       =   1
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "INS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   9701
            MinWidth        =   9701
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "07/09/2011"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1500
            MinWidth        =   1500
            TextSave        =   "0:46"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      MousePointer    =   2
   End
   Begin VB.Menu mnuSistema 
      Caption         =   "&Sistema"
      Begin VB.Menu mnuConectar 
         Caption         =   "&Conectar"
      End
      Begin VB.Menu mnuDesconectar 
         Caption         =   "&Desconectar"
      End
      Begin VB.Menu mnuRaya23 
         Caption         =   "-"
      End
      Begin VB.Menu mnuusuario 
         Caption         =   "&Usuarios   "
         HelpContextID   =   1
      End
      Begin VB.Menu mnupermisos 
         Caption         =   "&Permisos   "
         HelpContextID   =   1
      End
      Begin VB.Menu MNURAYA1 
         Caption         =   "-"
         HelpContextID   =   1
      End
      Begin VB.Menu mnubackup 
         Caption         =   "Copia de Base de Datos"
      End
      Begin VB.Menu mnurayabackup 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParametros 
         Caption         =   "&Parametros"
      End
      Begin VB.Menu mnuCalendario 
         Caption         =   "C&alendario"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCalculadora 
         Caption         =   "Calc&uladora"
      End
      Begin VB.Menu MNURAYA8 
         Caption         =   "-"
      End
      Begin VB.Menu mnusalir 
         Caption         =   "Salir   "
         HelpContextID   =   1
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuComprasFacturacion 
      Caption         =   "   &Ventas"
      Begin VB.Menu mnuFacturaActualiza 
         Caption         =   "&Actualizaciones"
         Begin VB.Menu mnuABMClientes 
            Caption         =   "...ABM de &Clientes"
         End
         Begin VB.Menu mnuABMVendedores 
            Caption         =   "...ABM de &Vendedores"
         End
         Begin VB.Menu mnuABMPais 
            Caption         =   "...ABM de &País"
         End
         Begin VB.Menu mnuABMProvincias 
            Caption         =   "...ABM de &Provincias"
         End
         Begin VB.Menu mnuABMLocalidades 
            Caption         =   "...ABM de &Localidades"
         End
         Begin VB.Menu mnuFacturacionTipoComprobante 
            Caption         =   "...ABM Tipo de Comprobante"
         End
         Begin VB.Menu mnuABMInscIVA 
            Caption         =   "...ABM de Insc. &IVA"
         End
         Begin VB.Menu mnuABMCanales 
            Caption         =   "...ABM de Tipos de Clientes"
         End
         Begin VB.Menu mnuABMEstadoDocumento 
            Caption         =   "...ABM Estado de Documentos"
         End
         Begin VB.Menu mnuABMFormaPago 
            Caption         =   "...ABM de Forma de Pago"
         End
      End
      Begin VB.Menu mnuFacturacionPedidos 
         Caption         =   "&Presupuestos"
         Begin VB.Menu mnuIngNotaPedidoCliente 
            Caption         =   "Ingreso de Presupuestos"
         End
         Begin VB.Menu mnuListadoNotaPedidoCliente 
            Caption         =   "Listado de Presupuestos"
         End
      End
      Begin VB.Menu mnuFacturacionRemitos 
         Caption         =   "&Remitos"
         Begin VB.Menu mnuIngRemitosClientes 
            Caption         =   "Ingreso de Remitos de Clientes"
         End
         Begin VB.Menu mnuListadoRemitosClientes 
            Caption         =   "Listado de Remitos de Clientes"
         End
      End
      Begin VB.Menu mnuFacturacionRecibos 
         Caption         =   "R&ecibos"
         Begin VB.Menu mnuIngRecibosClientes 
            Caption         =   "Ingreso de Recibos de Clientes"
         End
         Begin VB.Menu mnuListadoRecibosClientes 
            Caption         =   "Listado de Recibos de Clientes"
         End
      End
      Begin VB.Menu mnuFacturacionCtaCte 
         Caption         =   "&Cuenta Corriente de Clientes"
      End
      Begin VB.Menu mnuConsultaAnulaciones 
         Caption         =   "C&onsulta - Anulaciones"
         Begin VB.Menu mnuConAnuPedidos 
            Caption         =   "... de Presupuestos"
         End
         Begin VB.Menu mnuConAnuRemitos 
            Caption         =   "... de Remitos"
         End
         Begin VB.Menu mnuConAnuRecibos 
            Caption         =   "... de Recibos"
         End
      End
      Begin VB.Menu mnuRaya19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListadoS 
         Caption         =   "&Listados"
         Begin VB.Menu mnuListadoClientes 
            Caption         =   "Agenda de Clientes"
         End
         Begin VB.Menu mnuEstVentas 
            Caption         =   "Estadisticas de Ventas"
         End
         Begin VB.Menu mnuRaya22 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFacturasPendientesCliente 
            Caption         =   "Remitos Pendientes por Cliente"
         End
         Begin VB.Menu mnuPagosRealizadosPorCliente 
            Caption         =   "Pagos realizados por Clientes"
         End
         Begin VB.Menu mnuRayaLis 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEstaCantidadVendida 
            Caption         =   "Listado de Cantidades Vendidas"
         End
         Begin VB.Menu mnuEstaVentaporCliente 
            Caption         =   "Ventas por Cliente"
         End
         Begin VB.Menu mnuListadoVentaPorVendedor 
            Caption         =   "Ventas por Vendedor"
         End
      End
      Begin VB.Menu mnurayal 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLibroIvaVentas 
         Caption         =   "Libro &Débito Fiscal"
      End
   End
   Begin VB.Menu mnuCompras 
      Caption         =   "   &Compras"
      Index           =   1
      Begin VB.Menu mnuComprasActualiza 
         Caption         =   "&Actualizaciones"
         Begin VB.Menu mnuABMTipoProveedores 
            Caption         =   "... ABM de Tipo de Proveedores"
         End
         Begin VB.Menu mnuABMProveedores 
            Caption         =   "... ABM de Proveedores"
         End
         Begin VB.Menu mnuABMTipoGastos 
            Caption         =   "... ABM de Tipo de Gastos"
         End
      End
      Begin VB.Menu mnuComprasOc 
         Caption         =   "&Compras"
         Begin VB.Menu mnuOrdendeCompra 
            Caption         =   "Ingreso de Ordenes de Compras"
         End
         Begin VB.Menu mnuListadoOrdCompra 
            Caption         =   "Listado de Ordenes de Compra"
         End
      End
      Begin VB.Menu mnuremito 
         Caption         =   "&Remitos"
         Begin VB.Menu mnuRemitoCompras 
            Caption         =   "Ingreso de Remitos de Proveedores"
         End
         Begin VB.Menu mnuListadoRemProv 
            Caption         =   "Listado de Remitos de Proveedores"
         End
         Begin VB.Menu MNURAYA12 
            Caption         =   "-"
         End
         Begin VB.Menu mnuImputarNCFacturasProveedores 
            Caption         =   "Imputar NC a Facturas"
         End
      End
      Begin VB.Menu mnuGastos 
         Caption         =   "&Gastos Generales"
         Begin VB.Menu mnuGastosGeneralesRegistro 
            Caption         =   "R&egistro de Gastos"
         End
      End
      Begin VB.Menu mnuPagoProveedores 
         Caption         =   "Re&cibo de Proveedores"
         Begin VB.Menu mnuPagoProveedoresOrdenPago 
            Caption         =   "Recibo de Proveedor"
         End
      End
      Begin VB.Menu mnuComprasCtaCte 
         Caption         =   "&Cuenta  Corriente de Proveedores"
      End
      Begin VB.Menu mnconsanu 
         Caption         =   "C&onsulta - Anulaciones"
         Begin VB.Menu mnuOrdenCompra 
            Caption         =   "... de Ordenes de Compra"
         End
         Begin VB.Menu mnuremitos 
            Caption         =   "... de Remitos"
         End
         Begin VB.Menu mnuOrdenesdePago 
            Caption         =   "...de Recibos de Proveedores"
         End
         Begin VB.Menu mnuAnularGastos 
            Caption         =   "... de Gastos Generales"
         End
      End
      Begin VB.Menu mnuraya10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComprasListado 
         Caption         =   "&Listado"
         Begin VB.Menu mnuListadoProveedores 
            Caption         =   "Listado de P&roveedores"
         End
         Begin VB.Menu mnuraya2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFacturasPendientesProveedor 
            Caption         =   "Facturas pendientes por Proveedor"
         End
         Begin VB.Menu mnuPagosRealizadosProveedores 
            Caption         =   "Recibos de Proveedores"
         End
         Begin VB.Menu mnuraya20 
            Caption         =   "-"
         End
         Begin VB.Menu mnuComprasListaResumen 
            Caption         =   "Resumen de Compras"
         End
      End
      Begin VB.Menu mnuraya11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComprasCreditoFiscal 
         Caption         =   "Libro &Crédito Fiscal"
      End
   End
   Begin VB.Menu mnuGestionStock 
      Caption         =   "   &Stock"
      Begin VB.Menu mnuStockActualiza 
         Caption         =   "&Actualizaciones"
         Begin VB.Menu mnuStockABMProductos 
            Caption         =   "...ABM de Productos"
         End
         Begin VB.Menu mnuRaya16 
            Caption         =   "-"
         End
         Begin VB.Menu mnuABMLineas 
            Caption         =   "...ABM de Familias"
         End
         Begin VB.Menu mnuABMRubros 
            Caption         =   "...ABM de Rubros"
         End
         Begin VB.Menu mnuABMPresentacion 
            Caption         =   "...ABM de Marcas"
         End
         Begin VB.Menu mnuraya 
            Caption         =   "-"
         End
         Begin VB.Menu mnuABMTransporte 
            Caption         =   "...ABM de Transporte"
         End
      End
      Begin VB.Menu mnuEntradaProductos 
         Caption         =   "&Recepción de Mercadería"
      End
      Begin VB.Menu mnuEntregaDeProductos 
         Caption         =   "&Salida de Mercadería"
      End
      Begin VB.Menu mnuRaya21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStockAjuste 
         Caption         =   "Control de &Stock"
      End
      Begin VB.Menu mnuConStock 
         Caption         =   "Consulta de Stock"
      End
      Begin VB.Menu MNURAYA5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListaPrecios 
         Caption         =   "Lista de &Precios"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuConsListaPrecios 
         Caption         =   "C&onsulta de Lista de Precios"
      End
      Begin VB.Menu mnuStockFaltantes 
         Caption         =   "Listado de Stock Faltantes"
      End
   End
   Begin VB.Menu mnuFondos 
      Caption         =   "   F&ondos"
      Begin VB.Menu mnuFondosActualizaciones 
         Caption         =   "&Actualizaciones"
         Begin VB.Menu mnuABMTipoCuentas 
            Caption         =   "...ABM Tipo de Cuentas"
         End
         Begin VB.Menu mnuFondosCuentas 
            Caption         =   "...ABM de Cuentas"
         End
         Begin VB.Menu mnuFondosBancos 
            Caption         =   "...ABM de Bancos"
         End
         Begin VB.Menu mnuFondosEstadoCheques 
            Caption         =   "...ABM de Estados de Cheques"
         End
         Begin VB.Menu mnuABMMoneda 
            Caption         =   "...ABM de Moneda"
         End
         Begin VB.Menu mnuABMTiposGastosBancarios 
            Caption         =   "...ABM Tipos de Gastos Bancarios"
         End
      End
      Begin VB.Menu mnuFondosGestonBancaria 
         Caption         =   "Gestión &Bancaria"
         Begin VB.Menu mnuFondosMovBancarios 
            Caption         =   "Movimientos Bancararios"
         End
         Begin VB.Menu mnuFondosResumenCuneta 
            Caption         =   "Resumen de Cuenta"
         End
      End
      Begin VB.Menu mnuFondosGestionCaja 
         Caption         =   "Gestión &Caja"
         Begin VB.Menu mnuFondosCargaIngresos 
            Caption         =   "Carga de Ingresos"
         End
         Begin VB.Menu mnuFondosCargaEgresos 
            Caption         =   "Carga de Egresos"
         End
         Begin VB.Menu mnuRaya15 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLiquidacionCobranza 
            Caption         =   "Liquidación de Cobranza"
         End
         Begin VB.Menu mnuFonfosCierreCaja 
            Caption         =   "Cierre de Caja"
         End
      End
      Begin VB.Menu mnuFondosValores 
         Caption         =   "&Valores"
         Begin VB.Menu mnuFondosCargaCheques 
            Caption         =   "Carga de Cheques de Terceros"
         End
         Begin VB.Menu mnuBoletaDeposito 
            Caption         =   "Boleta Déposito"
         End
         Begin VB.Menu mnuFondosCambioEstadoChques 
            Caption         =   "Cambio de Estado Cheques de Terceros"
         End
         Begin VB.Menu mnuFondosListadoCheques 
            Caption         =   "Listado de Cheques de Terceros"
         End
         Begin VB.Menu mnuraya25 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFondosCargaChequesPropios 
            Caption         =   "Carga de Cheques Propios"
         End
         Begin VB.Menu mnuCambioEstadoChequesPropios 
            Caption         =   "Cambio de Estado Cheques Propios"
         End
         Begin VB.Menu mnuListadoChequesPropios 
            Caption         =   "Listado de Cheques Propios"
         End
         Begin VB.Menu mnuraya26 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIngresoGastosBancarios 
            Caption         =   "Ingreso de Gastos Bancarios"
         End
      End
   End
   Begin VB.Menu mnuayuda 
      Caption         =   "  A&yuda"
   End
   Begin VB.Menu mnuacercade 
      Caption         =   "  &Acerca de... "
   End
End
Attribute VB_Name = "MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
Dim TituloPrincipal As String

Private Declare Function ShellAbout Lib "shell32.dll" Alias _
"ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, _
ByVal szOtherStuff As String, ByVal hIcon As Long) As Long



Private Sub Command1_Click()
    frmListadoStock.Show vbModal
End Sub

Private Sub mnubackup_Click()
    With frmRestaurarBD
        .Caption = "Backup de Archivos"
        .optCopiarA.Value = True
        .Label1 = "Guardar Backup en:"
        .Show
    End With
End Sub

Private Sub mnuEstVentas_Click()
    frmEstadisticaVentas.Show vbModal
End Sub

Private Sub mnuStockFaltantes_Click()
    frmListadoStock.Show vbModal
End Sub

Private Sub tbrPrincipal_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Index
        Case 2: Call mnuRemitosdirecto_Click
        Case 3: Call mnuClientesdirecto_Click
        Case 4: Call mnuIngRecibosClientes_Click
        Case 5: Call mnuListadoClientes_Click
        'Case 4: Separador
        Case 7: Call mnuRemitosProvdirecto_Click
        Case 8: Call mnuABMProveedores_Click
        'Case 9: Separador
        Case 10: Call mnuProductosdirecto_Click
        Case 11: Call mnuListaPrecios_Click
        Case 12: Call mnuStockFaltantes_Click
        'Case 9: Separador
        Case 14: Call mnuEstVentas_Click
        Case 16: Call mnusalir_Click
    End Select
End Sub
Private Sub mnuRemitosProvdirecto_Click()
    frmRemitoProveedor.Show vbModal
End Sub
Private Sub mnuClientesdirecto_Click()
    ABMCliente.Show vbModal
End Sub
Private Sub mnuProductosdirecto_Click()
    ABMProducto.Show vbModal
End Sub
Private Sub mnuRemitosdirecto_Click()
    frmRemitoCliente.Show
End Sub


Private Sub cmdFacCom_Click(Index As Integer)
    frmfacturaproveedor.Show vbModal
End Sub

Private Sub cmdFacturas_Click(Index As Integer)
    frmFacturaCliente.Show vbModal
End Sub

Private Sub cmdListados_Click()
    frmAdmListados.Show vbModal
End Sub

Private Sub cmdPrecio_Click(Index As Integer)
    Consulta = 2
    FrmListadePrecios.Show vbModal
End Sub

Private Sub cmdPto_Click()
    ABMProducto.Show vbModal
End Sub

Private Sub cmdRemCom_Click()
    frmRemitoProveedor.Show vbModal
End Sub

Private Sub cmdRemitos_Click()
    frmRemitoCliente.Show vbModal
End Sub

Private Sub mnuABMCanales_Click()
    ABMCanal.Show vbModal
End Sub

Private Sub mnuABMClientes_Click()
    ABMCliente.Show vbModal
End Sub

Private Sub mnuABMConceptoNotaCredito_Click()
'    ABMConceptoNotaCredito.Show vbModal
End Sub

Private Sub mnuABMEstadoDocumento_Click()
    ABMEstadoDocumento.Show vbModal
End Sub

Private Sub mnuABMFormaPago_Click()
    ABMFormaPago.Show vbModal
End Sub

Private Sub mnuABMInscIVA_Click()
    ABMInscIVA.Show vbModal
End Sub

Private Sub mnuABMLineas_Click()
    ABMLinea.Show vbModal
End Sub

Private Sub mnuABMLocalidades_Click()
    ABMLocalidad.Show vbModal
End Sub

Private Sub mnuABMMoneda_Click()
     ABMMoneda.Show vbModal
End Sub

Private Sub mnuABMPais_Click()
    ABMPais.Show vbModal
End Sub

Private Sub mnuABMPresentacion_Click()
    ABMPresentacion.Show vbModal
End Sub

Private Sub mnuABMProveedores_Click()
    ABMProveedor.Show vbModal
End Sub

Private Sub mnuABMProvincias_Click()
    ABMProvincia.Show vbModal
End Sub

Private Sub mnuABMRepresentada_Click()
    ABMRepresentada.Show vbModal
End Sub

Private Sub mnuABMRubros_Click()
    ABMRubro.Show vbModal
End Sub

Private Sub mnuABMServicios_Click()
'    ABMServicios.Show vbModal
End Sub

Private Sub mnuABMSucursales_Click()
    ABMSucursal.Show vbModal
End Sub

Private Sub mnuABMTipoCuentas_Click()
    ABMTipoCuenta.Show vbModal
End Sub

Private Sub mnuABMTipoGastos_Click()
    ABMTipoGasto.Show vbModal
End Sub

Private Sub mnuABMTipoProveedores_Click()
    ABMTipoProveedor.Show vbModal
End Sub

Private Sub mnuABMTiposGastosBancarios_Click()
    ABMTipoGastoBancario.Show vbModal
End Sub

Private Sub mnuABMTransporte_Click()
    ABMTransporte.Show vbModal
End Sub

Private Sub mnuABMVendedores_Click()
    ABMVendedor.Show vbModal
End Sub

Private Sub mnuacercade_Click()
    Call ShellAbout(Me.hwnd, "Mi Programa", "Copyright 2002, PMMF", Me.Icon)
End Sub

Private Sub mnuAnularGastos_Click()
    frmAnulaDocumentos.TipodeAnulacion = 9
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuayuda_Click()
    MsgBox "La Ayuda no está disponible en estos momentos", vbExclamation, TIT_MSGBOX
'    Call WinHelp(Me.hwnd, App.Path & "\Pecari.hlp", HelpFinder, 0&)
End Sub

Private Sub mnuBoletaDeposito_Click()
    FrmBoletaDeposito.Show vbModal
End Sub

Private Sub mnuCalculadora_Click()
    On Error Resume Next
    Shell "C:\WINDOWS\system32\calc.exe", vbNormalFocus
    'Form1.Show vbModal
End Sub

Private Sub mnuCalendario_Click()
    frmCalendario.Show
End Sub

Private Sub mnuCambioEstadoChequesPropios_Click()
    ABMCambioEstadoChPropio.Show vbModal
End Sub

Private Sub mnuCobranzaClientes_Click()
    frmListadoCobranzaCliente.Show vbModal
End Sub

Private Sub mnuComisionesVenderores_Click()
    ABMComision.Show vbModal
End Sub

Private Sub mnuComprasCreditoFiscal_Click()
'    frmLibroIvaCompras.Show
    frmLibroCompras2.Show vbModal
End Sub

Private Sub mnuComprasCtaCte_Click()
    frmCtaCteProveedores.Show vbModal
End Sub

Private Sub mnuComprasListaResumen_Click()
    frmListadoComprasProveedores.Show vbModal
End Sub

Private Sub mnuConAnuFactura_Click()
'    frmAnulaDocumentos.TipodeAnulacion = 3
'    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuConAnuPedidos_Click()
    frmAnulaDocumentos.TipodeAnulacion = 1
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuConAnuRecibos_Click()
    frmAnulaDocumentos.TipodeAnulacion = 4
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuConAnuRemitos_Click()
    frmAnulaDocumentos.TipodeAnulacion = 2
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuConectar_Click()
    'FrmInicio.Show vbModal
    inicio
    Me.Caption = TituloPrincipal & " - (Usuario " & UCase(mNomUser) & " conectado a " & UCase(SERVIDOR) & ")"
    Me.mnuConectar.Enabled = False
End Sub

Private Sub mnuConsListaPrecios_Click()
    Consulta = 1
    FrmListadePrecios.Show vbModal
End Sub

Private Sub mnuConsultaStock_Click()

End Sub

Private Sub mnuConStock_Click()
    Stock = 1 'cosnulta stock
    
    frmControlStock.Show
End Sub

Private Sub mnuDesconectar_Click()
    If DBConn.State = adStateOpen Then
        DBConn.Close
        
        DeshabilitarMenu Me
        
        Me.mnuSistema.Enabled = True
        Me.mnuConectar.Enabled = True
        Me.mnusalir.Enabled = True
        Me.mnuDesconectar.Enabled = False
        
        Me.Caption = TituloPrincipal & " - (No conectado)"
    End If
End Sub

Private Sub MDIForm_Load()
    TituloPrincipal = "Sistema de Gestión y Administración"
    Me.Caption = TituloPrincipal
    
    Me.Show
    
    inicio
    'FrmInicio.Show vbModal
       
    Me.Caption = TituloPrincipal '& " - (Usuario " & UCase(mNomUser) & " conectado a " & UCase(SERVIDOR) & " - " & UCase(BASEDATO) & ")"
    MENU.mnuConectar.Enabled = False
End Sub
Private Function inicio()
 Dim TxtUsuario As String
 Dim TxtClave As String
 Set rec = New ADODB.Recordset
    TxtUsuario = "a"
    TxtClave = "az"
    
    mNomUser = Trim(TxtUsuario)
    
    
    Conexion TxtUsuario, TxtClave
    
    
    
    If Not CONECCION Then
        If Err.Description <> "" Then
            MsgBox Err.Description
        End If
            
        'CUANTAS_VECES = CUANTAS_VECES + 1
        'If CUANTAS_VECES = 4 Then
        '    End
        'End If
        'txtusuario.SetFocus
        'Exit Sub
    End If


    sql = "SELECT * FROM USUARIO WHERE " & _
          "USU_NOMBRE = '" & Trim(TxtUsuario) & "' AND " & _
           "USU_CLAVE = '" & Trim(TxtClave) & "'"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.RecordCount <> 1 Then
'        sql = "La contraseña de usuario NO ES CORRECTA !" & Chr(13) & Chr(13)
'        If CUANTAS_VECES = 3 Then
'            sql = sql & "El sistema se cerrará."
'        Else
'            sql = sql & "Por favor intentelo nuevamente."
'        End If
'        MsgBox sql, vbCritical, "Error:"
'        If CUANTAS_VECES = 3 Then
'            'si ya pifió 3 veces salgo del Sistema
'            cmdSalir_Click
'        Else
'            TxtClave.SelStart = 0
'            TxtClave.SelLength = Len(TxtClave)
'            TxtUsuario.SetFocus
'            CUANTAS_VECES = CUANTAS_VECES + 1
'        End If
'    Else
        'Label1(1).FontBold = True
        'Label1(1).Caption = " Conectando ... "
        'Label1(1).Refresh
        'muestro un figureti de coneccion
        mNomUser = Trim(TxtUsuario)
        mPassword = Trim(TxtClave)
        
        'BUSCO SUCURSAL---
        BuscoNroSucursal
        '-----------------
        'Unload Me
        'Set FrmInicio = Nothing
    'End If
End Function
Public Function Conexion(TxtUsuario As String, TxtClave As String)
Dim DSN_DEF As String
    Screen.MousePointer = vbHourglass
    CONECCION = False

    On Error GoTo ErrorIni
    LeoIni
    
    On Error GoTo ErrorTrans
    'ME CONECTO !
    Set DBConn = New ADODB.Connection
'    mNomUser = TxtUsuario.Text
'    mPassword = TxtClave.Text
'    DSN_DEF = "ELECTROCENTRO"
   ' DBConn.ConnectionString = "ODBC;DATABASE=;UID=" & mNomUser & ";PWD=" & mPassword & ";DSN=" & DSN_DEF
    DBConn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SIELECTROCENTRO"
    'DBConn.ConnectionString = "driver={SQL Server}; server=DANIEL;database=ELECTROCENTRO"

    'DBConn.ConnectionTimeout = 30       'Default msado10.hlp => 15
    DBConn.CommandTimeout = 0          'Default msado10.hlp => 30
    'DBConn.Open , TxtUsuario, TxtClave
    DBConn.Open DBConn.ConnectionString, TxtUsuario, TxtClave
    
       
    If DBConn.State = adStateOpen Then CONECCION = True
        
    PERMISOS mNomUser
    Screen.MousePointer = vbNormal
    Exit Function
    
ErrorTrans:
        Screen.MousePointer = vbNormal
        MsgBox "No se pudo establecer la conección a la Base de Datos." & Chr(13) & Err.Description, vbExclamation, TIT_MSGBOX
        Exit Function
ErrorIni:
        Screen.MousePointer = vbNormal
        MsgBox "No se pudo leer el archivo de configuración del sistema." & Chr(13) & Err.Description, vbExclamation, TIT_MSGBOX
End Function
Public Sub PERMISOS(USUARIO As String)

    Dim sql As String
    Dim r As ADODB.Recordset
    Dim i As Integer

    Set r = New ADODB.Recordset

    On Error Resume Next

    If Trim(USUARIO) = "A" Then
        'Este usuario tiene derecho a todo
        For i = 0 To MENU.Controls.Count - 1
            If TypeName(MENU.Controls(i)) = "Menu" Then
               MENU.Controls(i).Enabled = True
            End If
        Next
    Else
        For i = 0 To MENU.Controls.Count - 1
            If TypeName(MENU.Controls(i)) = "Menu" Then
               MENU.Controls(i).Enabled = False
            End If
        Next
    
        On Error GoTo 0
    
        sql = "SELECT * FROM PERMISOS WHERE " & _
        "USU_NOMBRE = '" & Trim(USUARIO) & "' AND " & _
        "PRM_SISTEMA = '" & Trim(App.Title) & "'"
        r.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If r.RecordCount > 0 Then
            r.MoveFirst
            Do While Not r.EOF
                For i = 0 To MENU.Controls.Count - 1
                    If TypeName(MENU.Controls(i)) = "Menu" Then
                        If UCase(Trim(MENU.Controls(i).Name)) = UCase(Trim(r!PRM_OPMENU)) Then
                            MENU.Controls(i).Enabled = True
                        End If
                    End If
                Next
                r.MoveNext
            Loop
        End If
        r.Close
    End If
End Sub

Private Sub Cargar(formu As Form, Optional modalmente As Integer)
    On Error GoTo CLAVOSE
    formu.Top = 0
    formu.Left = 0
    Load formu
    formu.Show modalmente
    Exit Sub
    
CLAVOSE:
    MsgBox "Ha ocurrido un  error al tratar de cargar el formulario !" & Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub mnuEntradaProductos_Click()
    frmEntradaProductos.Show vbModal
End Sub

Private Sub mnuEntregaDeProductos_Click()
    frmEntregaProductos.Show vbModal
End Sub

Private Sub mnuEstaCantidadVendida_Click()
    frmListadoCantidadesVendidas.Show vbModal
End Sub

Private Sub mnuEstaVentaporCliente_Click()
    frmListadoVentasPorCliente.Show vbModal
End Sub

Private Sub mnuFacturacionCtaCte_Click()
    frmCtaCteCliente.Show vbModal
End Sub

Private Sub mnuFacturacionFactura_Click()
'    frmFacturaCliente.Show vbModal
End Sub

Private Sub mnuFacturacionTipoComprobante_Click()
    ABMTipoComprobante.Show vbModal
End Sub

Private Sub mnufacturas_Click()
'    frmAnulaDocumentos.TipodeAnulacion = 7
'    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuFacturasPendientesCliente_Click()
    frmListadoFacturasPendientePorCliente.Show vbModal
End Sub

Private Sub mnuFacturasPendientesProveedor_Click()
    frmListadoFacturasPendientePorProveedor.Show vbModal
End Sub

Private Sub mnuFondosBancos_Click()
    ABMBanco.Show vbModal
End Sub

Private Sub mnuFondosCambioEstadoChques_Click()
    ABMCambioEstado.Show vbModal
End Sub

Private Sub mnuFondosCargaCheques_Click()
    FrmCargaCheques.Show vbModal
End Sub

Private Sub mnuFondosCargaChequesPropios_Click()
    FrmCargaChequesPropios.Show vbModal
End Sub

Private Sub mnuFondosCargaEgresos_Click()
    ABMEgresos.Show vbModal
End Sub

Private Sub mnuFondosCargaIngresos_Click()
    ABMIngresos.Show vbModal
End Sub

Private Sub mnuFondosCuentas_Click()
    ABMCuentasBancarias.Show vbModal
End Sub

Private Sub mnuFondosEstadoCheques_Click()
    ABMEstadoCheques.Show vbModal
End Sub

Private Sub mnuFondosListadoCheques_Click()
    FrmListCheques.Show vbModal
End Sub

Private Sub mnuFondosResumenCuneta_Click()
    frmResumenCuentaBanco.Show vbModal
End Sub

Private Sub mnuFonfosCierreCaja_Click()
    frmCierreCaja.Show
End Sub

Private Sub mnuGastosGeneralesRegistro_Click()
    frmCargaGastosGenerales.Show vbModal
End Sub

Private Sub mnuImputarNCFacturas_Click()
'    frmImputarNCaFactura.Show vbModal
End Sub

Private Sub mnuimputaNCFAC_Click()
'    frmImputarNCaFactura.Show vbModal
End Sub

Private Sub mnuImputarNCFacturasProveedores_Click()
'    frmImputarNCaFacturaProveedores.Show vbModal
End Sub

Private Sub mnuIngNotaPedidoCliente_Click()
    frmNotaDePedido.Show vbModal
End Sub

Private Sub mnuIngRecibosClientes_Click()
    frmReciboCliente.Show vbModal
End Sub

Private Sub mnuIngRemitosClientes_Click()
    frmRemitoCliente.Show
End Sub

Private Sub mnuIngresoGastosBancarios_Click()
    frmIngresoGastosBancarios.Show vbModal
End Sub

Private Sub mnuLibroIvaVentas_Click()
    'frmLibroIvaVentas.Show
    frmLibroVentas2.Show vbModal
End Sub

Private Sub mnuLiquidacionCobranza_Click()
    frmLiquidacionCobranza.Show vbModal
End Sub

Private Sub mnuListadoChequesPropios_Click()
    FrmListChequesPropios.Show vbModal
End Sub

Private Sub mnuListadoClientes_Click()
    frmListadoClientes.Show vbModal
End Sub

Private Sub mnuListadoNotaPedidoCliente_Click()
    frmListadoNotaDePedido.Show vbModal
End Sub

Private Sub mnuListadoOrdCompra_Click()
    frmListadoOrdenCompra.Show vbModal
End Sub

Private Sub mnuListadoProveedores_Click()
    frmListadoProvedores.Show vbModal
End Sub

Private Sub mnuListadoRecibosClientes_Click()
    frmListadoReciboCliente.Show vbModal
End Sub

Private Sub mnuListadoRemitosClientes_Click()
    frmListadoRemitoCliente.Show vbModal
End Sub

Private Sub mnuListadoSucursalesCliente_Click()
    frmListadoSucursalesCliente.Show vbModal
End Sub

Private Sub mnuListadoRemProv_Click()
    frmListadoRemitoProveedor.Show vbModal
End Sub

Private Sub mnuListadoVentaPorVendedor_Click()
    frmListadoVentasPorVendedor.Show vbModal
End Sub

Private Sub mnuListaPrecios_Click()
    Consulta = 2
    FrmListadePrecios.Show vbModal
End Sub

Private Sub mnuMovimientoStock_Click()

End Sub

Private Sub mnuNotaCredito_Click()
'    frmNotaCreditoCliente.Show vbModal
End Sub

Private Sub mnuNotaCreditoProveedor_Click()
'    frmNotaCreditoProveedor.Show vbModal
End Sub

Private Sub mnuNotaDebitoPorChequeDevuelto_Click()
'    frmNotaDeditoClienteCheques.Show vbModal
End Sub

Private Sub mnuNotaDebitoPorServicios_Click()
'    frmNotaDeditoClienteServicio.Show vbModal
End Sub

Private Sub mnuNotaDebitoProveedor_Click()
'    frmNotaDebitoProveedores.Show vbModal
End Sub

Private Sub mnuOrdenCompra_Click()
    frmAnulaDocumentos.TipodeAnulacion = 5
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuOrdendeCompra_Click()
    frmOrdenesCompra.Show vbModal
End Sub

Private Sub mnuPagoProveedoresAnularPago_Click()
    frmAnulaOrdenPago.Show vbModal
End Sub

Private Sub mnuOrdenesdePago_Click()
    frmAnulaDocumentos.TipodeAnulacion = 8
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnuPagoProveedoresOrdenPago_Click()
    frmPagoProveedores.Show vbModal
End Sub

Private Sub mnuPagosRealizadosPorCliente_Click()
    frmListadoPagosPorCliente.Show vbModal
End Sub

Private Sub mnuPagosRealizadosProveedores_Click()
    frmListadoPagosPorProveedor.Show vbModal
End Sub

Private Sub mnuParametros_Click()
    frmParametros.Show vbModal
End Sub

Private Sub mnupermisos_Click()
    FrmPermisos.Show vbModal
End Sub

Private Sub mnuProveedoresFacturas_Click()
'    frmfacturaproveedor.Show vbModal
End Sub

Private Sub mnuRemitoCompras_Click()
    frmRemitoProveedor.Show vbModal
End Sub

Private Sub mnuRemitos_Click()
    frmAnulaDocumentos.TipodeAnulacion = 6
    frmAnulaDocumentos.Show vbModal
End Sub

Private Sub mnusalir_Click()
   Set MENU = Nothing
   End
End Sub

Private Sub mnuStockABMProductos_Click()
    Consulta = 1
    ABMProducto.CODIGOLISTA = 0
    ABMProducto.Show vbModal
End Sub

Private Sub mnuStockAjuste_Click()
    Stock = 0 'Actualiza el Stock
    frmControlStock.Show vbModal
End Sub

Private Sub mnuusuario_Click()
     Cargar FrmUsuarios, 1
End Sub

