VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   LinkTopic       =   "Form5"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox lvwTransactions 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   3000
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   20
      Top             =   6720
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   600
      TabIndex        =   6
      Top             =   1920
      Width           =   10215
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         DownPicture     =   "Form5.frx":0000
         Height          =   375
         Left            =   9360
         Picture         =   "Form5.frx":1E53
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dated"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   5880
         TabIndex        =   10
         Top             =   120
         Width           =   3375
         Begin VB.CommandButton cmdOk 
            Caption         =   "Proceed"
            DownPicture     =   "Form5.frx":895D
            Height          =   315
            Left            =   1320
            Picture         =   "Form5.frx":A7B0
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   840
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dtTo 
            Height          =   375
            Left            =   1800
            TabIndex        =   12
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   71892993
            CurrentDate     =   38816
         End
         Begin MSComCtl2.DTPicker dtFrom 
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   71892993
            CurrentDate     =   38718
         End
         Begin VB.Label dtToj 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "To:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            TabIndex        =   15
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "From:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.ComboBox cboCustomerID 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Text            =   "Select..."
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cboFirst 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Text            =   "Select..."
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cboAccNo 
         Height          =   315
         Left            =   3960
         TabIndex        =   7
         Text            =   "Select..."
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Account Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Customer ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Choose the View Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Width           =   5775
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Custom"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View All"
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      DownPicture     =   "Form5.frx":112BA
      Height          =   375
      Left            =   9960
      Picture         =   "Form5.frx":1310D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      DownPicture     =   "Form5.frx":19C17
      Height          =   375
      Left            =   600
      Picture         =   "Form5.frx":1BA6A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   1575
   End
   Begin VB.PictureBox ListView1 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   4200
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   0
      Top             =   6720
      Width           =   1000
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "Form5.frx":22574
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
