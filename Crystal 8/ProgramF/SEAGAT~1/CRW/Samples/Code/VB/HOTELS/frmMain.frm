VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "First Class Hotels Reservation System"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About Sample"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1575
      TabIndex        =   5
      Top             =   4530
      Width           =   2070
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2970
      TabIndex        =   4
      Top             =   3810
      Width           =   2070
   End
   Begin VB.CommandButton cmdPricing 
      Caption         =   "&Prices"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2970
      TabIndex        =   3
      Top             =   3000
      Width           =   2070
   End
   Begin VB.CommandButton cmdCustomer 
      Caption         =   "&Customers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   240
      TabIndex        =   2
      Top             =   3810
      Width           =   2070
   End
   Begin VB.CommandButton cmdRooms 
      Caption         =   "R&ooms"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   2070
   End
   Begin VB.CommandButton cmdReservation 
      Caption         =   "&Reservations"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   1215
      TabIndex        =   0
      Top             =   2130
      Width           =   2790
   End
   Begin VB.Image Image1 
      Height          =   1650
      Left            =   240
      Picture         =   "frmMain.frx":030A
      Top             =   180
      Width           =   4800
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose:  The startup form for the application.
'           The user can choose from the four available options here --
'               Reservations, Rooms, Customers, Pricing and Exit
'
Option Explicit

' *************************************************************
' Show the Reservation form (the most interesting of the five forms
'
Private Sub cmdReservation_Click()
    Screen.MousePointer = vbHourglass           ' Set the mouse pointer to an hourglass
    frmReservation.Show vbModal                 ' Show the form
    ' After exiting the Reservation form, ensure that this form is visible
    ' over top of any other forms
    Me.ZOrder 0
End Sub

' *************************************************************
' Show the Customer form
'
Private Sub cmdCustomer_Click()
    Screen.MousePointer = vbHourglass           ' Set the mouse pointer to an hourglass
    frmCustomer.Show vbModal                    ' Show the form
End Sub

' *************************************************************
' Exit the application
'
Private Sub cmdExit_Click()
    Unload Me                                   ' Unload the form (and hopefully quit everything)
End Sub

' *************************************************************
' Show the Pricing form
'
Private Sub cmdPricing_Click()
    Screen.MousePointer = vbHourglass           ' Set the mouse pointer to an hourglass
    frmPrice.Show vbModal                       ' Show the form
End Sub

' *************************************************************
' Show the Rooms form
'
Private Sub cmdRooms_Click()
    Screen.MousePointer = vbHourglass           ' Set the mouse pointer to an hourglass
    frmRoom.Show vbModal                        ' Show the form
End Sub

' *************************************************************
' Show the "About Sample" Form
'
Private Sub cmdAbout_Click()
    frmAbout.Show vbModal
End Sub


' *************************************************************
' After the form has started, add a small visual enhancement
'
Private Sub Form_Activate()
    Dither Me                                   ' Make the nice teal background shading
End Sub


