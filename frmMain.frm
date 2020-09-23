VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Knights IQ Game"
   ClientHeight    =   5790
   ClientLeft      =   990
   ClientTop       =   1455
   ClientWidth     =   8175
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5790
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.ListBox List1 
         BackColor       =   &H80000001&
         Enabled         =   0   'False
         ForeColor       =   &H00C0C0C0&
         Height          =   4545
         Left            =   4710
         TabIndex        =   41
         Top             =   240
         Width           =   3315
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00404040&
         Height          =   4545
         Left            =   150
         ScaleHeight     =   4485
         ScaleWidth      =   4515
         TabIndex        =   8
         Top             =   240
         Width           =   4575
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   63
            Left            =   3870
            Shape           =   3  'Circle
            Top             =   3840
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   62
            Left            =   3390
            Shape           =   3  'Circle
            Top             =   3840
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   61
            Left            =   2910
            Shape           =   3  'Circle
            Top             =   3840
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   60
            Left            =   2430
            Shape           =   3  'Circle
            Top             =   3840
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   59
            Left            =   1950
            Shape           =   3  'Circle
            Top             =   3840
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   58
            Left            =   1470
            Shape           =   3  'Circle
            Top             =   3840
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   57
            Left            =   990
            Shape           =   3  'Circle
            Top             =   3840
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   56
            Left            =   480
            Shape           =   3  'Circle
            Top             =   3840
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   55
            Left            =   3870
            Shape           =   3  'Circle
            Top             =   3360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   54
            Left            =   3390
            Shape           =   3  'Circle
            Top             =   3360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   53
            Left            =   2910
            Shape           =   3  'Circle
            Top             =   3360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   52
            Left            =   2430
            Shape           =   3  'Circle
            Top             =   3360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   51
            Left            =   1950
            Shape           =   3  'Circle
            Top             =   3360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   50
            Left            =   1470
            Shape           =   3  'Circle
            Top             =   3360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   49
            Left            =   990
            Shape           =   3  'Circle
            Top             =   3360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   48
            Left            =   510
            Shape           =   3  'Circle
            Top             =   3360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   47
            Left            =   3870
            Shape           =   3  'Circle
            Top             =   2880
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   46
            Left            =   3390
            Shape           =   3  'Circle
            Top             =   2880
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   45
            Left            =   2910
            Shape           =   3  'Circle
            Top             =   2880
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   44
            Left            =   2430
            Shape           =   3  'Circle
            Top             =   2880
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   43
            Left            =   1950
            Shape           =   3  'Circle
            Top             =   2880
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   42
            Left            =   1470
            Shape           =   3  'Circle
            Top             =   2880
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   41
            Left            =   990
            Shape           =   3  'Circle
            Top             =   2880
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   40
            Left            =   510
            Shape           =   3  'Circle
            Top             =   2880
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   39
            Left            =   3870
            Shape           =   3  'Circle
            Top             =   2400
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   38
            Left            =   3390
            Shape           =   3  'Circle
            Top             =   2400
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   37
            Left            =   2910
            Shape           =   3  'Circle
            Top             =   2400
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   36
            Left            =   2430
            Shape           =   3  'Circle
            Top             =   2400
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   35
            Left            =   1950
            Shape           =   3  'Circle
            Top             =   2400
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   34
            Left            =   1470
            Shape           =   3  'Circle
            Top             =   2400
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   33
            Left            =   990
            Shape           =   3  'Circle
            Top             =   2400
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   32
            Left            =   510
            Shape           =   3  'Circle
            Top             =   2370
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   31
            Left            =   3870
            Shape           =   3  'Circle
            Top             =   1920
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   30
            Left            =   3390
            Shape           =   3  'Circle
            Top             =   1920
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   29
            Left            =   2910
            Shape           =   3  'Circle
            Top             =   1920
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   28
            Left            =   2430
            Shape           =   3  'Circle
            Top             =   1920
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   27
            Left            =   1950
            Shape           =   3  'Circle
            Top             =   1920
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   26
            Left            =   1470
            Shape           =   3  'Circle
            Top             =   1920
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   25
            Left            =   990
            Shape           =   3  'Circle
            Top             =   1920
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   24
            Left            =   510
            Shape           =   3  'Circle
            Top             =   1920
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   23
            Left            =   3870
            Shape           =   3  'Circle
            Top             =   1440
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   22
            Left            =   3390
            Shape           =   3  'Circle
            Top             =   1440
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   21
            Left            =   2910
            Shape           =   3  'Circle
            Top             =   1440
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   20
            Left            =   2430
            Shape           =   3  'Circle
            Top             =   1440
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   19
            Left            =   1950
            Shape           =   3  'Circle
            Top             =   1440
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   18
            Left            =   1470
            Shape           =   3  'Circle
            Top             =   1440
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   17
            Left            =   990
            Shape           =   3  'Circle
            Top             =   1440
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   16
            Left            =   510
            Shape           =   3  'Circle
            Top             =   1440
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   15
            Left            =   3870
            Shape           =   3  'Circle
            Top             =   960
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   14
            Left            =   3390
            Shape           =   3  'Circle
            Top             =   960
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   13
            Left            =   2910
            Shape           =   3  'Circle
            Top             =   960
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   12
            Left            =   2430
            Shape           =   3  'Circle
            Top             =   960
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   11
            Left            =   1950
            Shape           =   3  'Circle
            Top             =   960
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   10
            Left            =   1470
            Shape           =   3  'Circle
            Top             =   960
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   9
            Left            =   990
            Shape           =   3  'Circle
            Top             =   960
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   8
            Left            =   510
            Shape           =   3  'Circle
            Top             =   960
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   7
            Left            =   3870
            Shape           =   3  'Circle
            Top             =   480
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   6
            Left            =   3390
            Shape           =   3  'Circle
            Top             =   480
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   5
            Left            =   2910
            Shape           =   3  'Circle
            Top             =   480
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   4
            Left            =   2430
            Shape           =   3  'Circle
            Top             =   480
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   3
            Left            =   1950
            Shape           =   3  'Circle
            Top             =   480
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   2
            Left            =   1470
            Shape           =   3  'Circle
            Top             =   480
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   1
            Left            =   990
            Shape           =   3  'Circle
            Top             =   480
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape RedDot 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   0
            Left            =   510
            Shape           =   3  'Circle
            Top             =   480
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape BlueDot 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   1
            Left            =   60
            Shape           =   3  'Circle
            Top             =   4590
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape BlueDot 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   2
            Left            =   330
            Shape           =   3  'Circle
            Top             =   4590
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape BlueDot 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   3
            Left            =   600
            Shape           =   3  'Circle
            Top             =   4590
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape BlueDot 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   4
            Left            =   870
            Shape           =   3  'Circle
            Top             =   4590
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape BlueDot 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   5
            Left            =   1140
            Shape           =   3  'Circle
            Top             =   4590
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape BlueDot 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   6
            Left            =   1410
            Shape           =   3  'Circle
            Top             =   4590
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape BlueDot 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   7
            Left            =   1680
            Shape           =   3  'Circle
            Top             =   4590
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape BlueDot 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   8
            Left            =   1950
            Shape           =   3  'Circle
            Top             =   4590
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "57"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   40
            Top             =   3780
            Width           =   195
         End
         Begin VB.Label lRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "49"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   39
            Top             =   3300
            Width           =   195
         End
         Begin VB.Label lRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "41"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   38
            Top             =   2820
            Width           =   195
         End
         Begin VB.Label lRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "33"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   37
            Top             =   2340
            Width           =   195
         End
         Begin VB.Label lRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "25"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   36
            Top             =   1860
            Width           =   195
         End
         Begin VB.Label lRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "17"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   35
            Top             =   1380
            Width           =   195
         End
         Begin VB.Label lRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "9"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   900
            Width           =   195
         End
         Begin VB.Label lRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "1"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   33
            Top             =   420
            Width           =   195
         End
         Begin VB.Label lCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "8"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   7
            Left            =   3810
            TabIndex        =   32
            Top             =   60
            Width           =   255
         End
         Begin VB.Label lCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "7"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   6
            Left            =   3330
            TabIndex        =   31
            Top             =   60
            Width           =   255
         End
         Begin VB.Label lCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "6"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   5
            Left            =   2850
            TabIndex        =   30
            Top             =   60
            Width           =   255
         End
         Begin VB.Label lCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "5"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   4
            Left            =   2370
            TabIndex        =   29
            Top             =   60
            Width           =   255
         End
         Begin VB.Label lCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "4"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   1890
            TabIndex        =   28
            Top             =   60
            Width           =   255
         End
         Begin VB.Label lCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "3"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   1410
            TabIndex        =   27
            Top             =   60
            Width           =   255
         End
         Begin VB.Label lCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "2"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   930
            TabIndex        =   26
            Top             =   60
            Width           =   255
         End
         Begin VB.Label lCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "1"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   450
            TabIndex        =   25
            Top             =   60
            Width           =   255
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   63
            Left            =   3690
            Top             =   3660
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   62
            Left            =   3210
            Top             =   3660
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   61
            Left            =   2730
            Top             =   3660
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   60
            Left            =   2250
            Top             =   3660
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   59
            Left            =   1770
            Top             =   3660
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   58
            Left            =   1290
            Top             =   3660
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   57
            Left            =   810
            Top             =   3660
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   56
            Left            =   330
            Top             =   3660
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   55
            Left            =   3690
            Top             =   3180
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   54
            Left            =   3210
            Top             =   3180
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   53
            Left            =   2730
            Top             =   3180
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   52
            Left            =   2250
            Top             =   3180
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   51
            Left            =   1770
            Top             =   3180
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   50
            Left            =   1290
            Top             =   3180
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   49
            Left            =   810
            Top             =   3180
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   48
            Left            =   330
            Top             =   3180
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   47
            Left            =   3690
            Top             =   2700
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   46
            Left            =   3210
            Top             =   2700
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   45
            Left            =   2730
            Top             =   2700
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   44
            Left            =   2250
            Top             =   2700
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   43
            Left            =   1770
            Top             =   2700
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   42
            Left            =   1290
            Top             =   2700
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   41
            Left            =   810
            Top             =   2700
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   40
            Left            =   330
            Top             =   2700
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   39
            Left            =   3690
            Top             =   2220
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   38
            Left            =   3210
            Top             =   2220
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   37
            Left            =   2730
            Top             =   2220
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   36
            Left            =   2250
            Top             =   2220
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   35
            Left            =   1770
            Top             =   2220
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   34
            Left            =   1290
            Top             =   2220
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   33
            Left            =   810
            Top             =   2220
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   32
            Left            =   330
            Top             =   2220
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   31
            Left            =   3690
            Top             =   1740
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   30
            Left            =   3210
            Top             =   1740
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   29
            Left            =   2730
            Top             =   1740
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   28
            Left            =   2250
            Top             =   1740
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   27
            Left            =   1770
            Top             =   1740
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   26
            Left            =   1290
            Top             =   1740
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   25
            Left            =   810
            Top             =   1740
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   24
            Left            =   330
            Top             =   1740
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   23
            Left            =   3690
            Top             =   1260
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   22
            Left            =   3210
            Top             =   1260
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   21
            Left            =   2730
            Top             =   1260
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   20
            Left            =   2250
            Top             =   1260
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   19
            Left            =   1770
            Top             =   1260
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   18
            Left            =   1290
            Top             =   1260
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   17
            Left            =   810
            Top             =   1260
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   16
            Left            =   330
            Top             =   1260
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   15
            Left            =   3690
            Top             =   780
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   14
            Left            =   3210
            Top             =   780
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   13
            Left            =   2730
            Top             =   780
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   12
            Left            =   2250
            Top             =   780
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   11
            Left            =   1770
            Top             =   780
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   10
            Left            =   1290
            Top             =   780
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   9
            Left            =   810
            Top             =   780
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   8
            Left            =   330
            Top             =   780
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   7
            Left            =   3690
            Top             =   300
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   6
            Left            =   3210
            Top             =   300
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   5
            Left            =   2730
            Top             =   300
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   4
            Left            =   2250
            Top             =   300
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   3
            Left            =   1770
            Top             =   300
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   2
            Left            =   1290
            Top             =   300
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   1
            Left            =   810
            Top             =   300
            Width           =   510
         End
         Begin VB.Image Board 
            Appearance      =   0  'Flat
            Height          =   510
            Index           =   0
            Left            =   330
            Top             =   300
            Width           =   510
         End
         Begin VB.Label lRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "64"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   8
            Left            =   4260
            TabIndex        =   24
            Top             =   3810
            Width           =   195
         End
         Begin VB.Label lRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "56"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   9
            Left            =   4260
            TabIndex        =   23
            Top             =   3330
            Width           =   195
         End
         Begin VB.Label lRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "48"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   10
            Left            =   4260
            TabIndex        =   22
            Top             =   2850
            Width           =   195
         End
         Begin VB.Label lRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "40"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   11
            Left            =   4260
            TabIndex        =   21
            Top             =   2370
            Width           =   195
         End
         Begin VB.Label lRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "32"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   12
            Left            =   4260
            TabIndex        =   20
            Top             =   1890
            Width           =   195
         End
         Begin VB.Label lRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "24"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   13
            Left            =   4260
            TabIndex        =   19
            Top             =   1410
            Width           =   195
         End
         Begin VB.Label lRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "16"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   14
            Left            =   4260
            TabIndex        =   18
            Top             =   930
            Width           =   195
         End
         Begin VB.Label lRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "8"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   15
            Left            =   4260
            TabIndex        =   17
            Top             =   450
            Width           =   195
         End
         Begin VB.Label lCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "64"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   8
            Left            =   3840
            TabIndex        =   16
            Top             =   4230
            Width           =   255
         End
         Begin VB.Label lCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "63"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   9
            Left            =   3360
            TabIndex        =   15
            Top             =   4230
            Width           =   255
         End
         Begin VB.Label lCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "62"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   10
            Left            =   2880
            TabIndex        =   14
            Top             =   4230
            Width           =   255
         End
         Begin VB.Label lCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "61"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   11
            Left            =   2400
            TabIndex        =   13
            Top             =   4230
            Width           =   255
         End
         Begin VB.Label lCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "60"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   12
            Left            =   1920
            TabIndex        =   12
            Top             =   4230
            Width           =   255
         End
         Begin VB.Label lCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "59"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   13
            Left            =   1440
            TabIndex        =   11
            Top             =   4230
            Width           =   255
         End
         Begin VB.Label lCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "58"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   14
            Left            =   960
            TabIndex        =   10
            Top             =   4230
            Width           =   255
         End
         Begin VB.Label lCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "57"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   15
            Left            =   480
            TabIndex        =   9
            Top             =   4230
            Width           =   255
         End
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   315
         Left            =   7080
         TabIndex        =   6
         Top             =   5100
         Width           =   915
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Undo"
         Height          =   315
         Left            =   5970
         TabIndex        =   5
         Top             =   5100
         Width           =   885
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "&Stop"
         Height          =   315
         Left            =   4710
         TabIndex        =   3
         Top             =   5100
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "The game objective is put all knights in board"
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   4800
         Width           =   4005
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright by "
         Height          =   255
         Left            =   4710
         TabIndex        =   1
         Top             =   4830
         Width           =   3285
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   5640
      Top             =   2610
   End
   Begin VB.Label lblKnights 
      Height          =   315
      Left            =   5640
      TabIndex        =   7
      Top             =   5580
      Width           =   2535
   End
   Begin VB.Label lblInf 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   5580
      Width           =   5625
   End
   Begin VB.Image imgKnightBlack 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   600
      Picture         =   "frmMain.frx":000C
      Top             =   5940
      Width           =   510
   End
   Begin VB.Image imgBlackBack 
      Height          =   480
      Left            =   1620
      Picture         =   "frmMain.frx":0C4E
      Top             =   5940
      Width           =   480
   End
   Begin VB.Image imgWhiteBack 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   1110
      Picture         =   "frmMain.frx":1890
      Top             =   5940
      Width           =   510
   End
   Begin VB.Image imgKnightWhite 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   90
      Picture         =   "frmMain.frx":24D2
      Top             =   5940
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Visible         =   0   'False
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuSolve 
         Caption         =   "&Solve"
      End
      Begin VB.Menu mnuHint 
         Caption         =   "&Hint"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************************************
'**Knights IQ Game 1.0
'**Copyright by Paulo Matoso
'**E-Mail: paulomt1@clix.pt
'**
'**                                     frmMain
'**Last Modification ---> 13/12/2002
'**********************************************************************************

Sub Solution1st()

                                         '    1  2  3  4  5  6  7  8
    ArraySolution(0) = 44  'First Knight ' A 01 02 03 04 05 06 07 08
    ArraySolution(1) = 59                ' B 09 10 11 12 13 14 15 16
    ArraySolution(2) = 49                ' C 17 18 19 20 21 22 23 24
    ArraySolution(3) = 34                ' D 25 26 27 28 29 30 31 32
    ArraySolution(4) = 17                ' E 33 34 35 36 37 38 39 40
    ArraySolution(5) = 2                 ' F 41 42 43 XX 45 46 47 48
    ArraySolution(6) = 12                ' G 49 50 51 52 53 54 55 56
    ArraySolution(7) = 6                 ' H 57 58 59 60 61 62 63 64
    ArraySolution(8) = 16
    ArraySolution(9) = 31    'Tis is my 1st Solution, but if u find another
    ArraySolution(10) = 48   'please send to paulomt1@clix.pt
    ArraySolution(11) = 63
    ArraySolution(12) = 53
    ArraySolution(13) = 47
    ArraySolution(14) = 64
    ArraySolution(15) = 54
    ArraySolution(16) = 60
    ArraySolution(17) = 50
    ArraySolution(18) = 33
    ArraySolution(19) = 18
    ArraySolution(20) = 1
    ArraySolution(21) = 11
    ArraySolution(22) = 5
    ArraySolution(23) = 15
    ArraySolution(24) = 32
    ArraySolution(25) = 38
    ArraySolution(26) = 55
    ArraySolution(27) = 61
    ArraySolution(28) = 51
    ArraySolution(29) = 57
    ArraySolution(30) = 42
    ArraySolution(31) = 25
    ArraySolution(32) = 10
    ArraySolution(33) = 4
    ArraySolution(34) = 14
    ArraySolution(35) = 8
    ArraySolution(36) = 23
    ArraySolution(37) = 40
    ArraySolution(38) = 46
    ArraySolution(39) = 56
    ArraySolution(40) = 62
    ArraySolution(41) = 52
    ArraySolution(42) = 58
    ArraySolution(43) = 41
    ArraySolution(44) = 26
    ArraySolution(45) = 9
    ArraySolution(46) = 3
    ArraySolution(47) = 13
    ArraySolution(48) = 7
    ArraySolution(49) = 24
    ArraySolution(50) = 30
    ArraySolution(51) = 20
    ArraySolution(52) = 35
    ArraySolution(53) = 45
    ArraySolution(54) = 28
    ArraySolution(55) = 43
    ArraySolution(56) = 37
    ArraySolution(57) = 27
    ArraySolution(58) = 21
    ArraySolution(59) = 36
    ArraySolution(60) = 19
    ArraySolution(61) = 29
    ArraySolution(62) = 39
    ArraySolution(63) = 22

End Sub
Sub Solution2st()

                                         '    1  2  3  4  5  6  7  8
    ArraySolution(0) = 16  'First Knight ' A 01 02 03 04 05 06 07 08
    ArraySolution(1) = 31                ' B 09 10 11 12 13 14 15 XX ---First
    ArraySolution(2) = 48                ' C 17 18 19 20 21 22 23 24
    ArraySolution(3) = 63                ' D 25 26 27 28 29 30 31 32
    ArraySolution(4) = 53                ' E 33 34 35 36 37 38 39 40
    ArraySolution(5) = 59                ' F 41 42 43 44 45 46 47 48
    ArraySolution(6) = 49                ' G 49 50 51 52 53 54 55 56
    ArraySolution(7) = 34                ' H 57 58 59 60 61 62 63 64
    ArraySolution(8) = 17
    ArraySolution(9) = 2     'Tis is my 2st Solution, but if u find another
    ArraySolution(10) = 12   'please send to paulomt1@clix.pt
    ArraySolution(11) = 6
    ArraySolution(12) = 23
    ArraySolution(13) = 8
    ArraySolution(14) = 14
    ArraySolution(15) = 4
    ArraySolution(16) = 10
    ArraySolution(17) = 25
    ArraySolution(18) = 42
    ArraySolution(19) = 57
    ArraySolution(20) = 51
    ArraySolution(21) = 61
    ArraySolution(22) = 55
    ArraySolution(23) = 40
    ArraySolution(24) = 30
    ArraySolution(25) = 15
    ArraySolution(26) = 5
    ArraySolution(27) = 11
    ArraySolution(28) = 1
    ArraySolution(29) = 18
    ArraySolution(30) = 33
    ArraySolution(31) = 50
    ArraySolution(32) = 60
    ArraySolution(33) = 54
    ArraySolution(34) = 64
    ArraySolution(35) = 47
    ArraySolution(36) = 32
    ArraySolution(37) = 22
    ArraySolution(38) = 7
    ArraySolution(39) = 13
    ArraySolution(40) = 3
    ArraySolution(41) = 9
    ArraySolution(42) = 26
    ArraySolution(43) = 41
    ArraySolution(44) = 58
    ArraySolution(45) = 62
    ArraySolution(46) = 52
    ArraySolution(47) = 56
    ArraySolution(48) = 46
    ArraySolution(49) = 29
    ArraySolution(50) = 19
    ArraySolution(51) = 36
    ArraySolution(52) = 21
    ArraySolution(53) = 27
    ArraySolution(54) = 44
    ArraySolution(55) = 38
    ArraySolution(56) = 28
    ArraySolution(57) = 43
    ArraySolution(58) = 37
    ArraySolution(59) = 20
    ArraySolution(60) = 35
    ArraySolution(61) = 45
    ArraySolution(62) = 39
    ArraySolution(63) = 24

End Sub

Sub UpdateHints(Square As Integer)
Dim Col As Integer
Dim Row As Integer
Dim FreeDest As Integer
Dim i As Integer
Dim LocalSquare As Integer

    Col = Square Mod 8 + 1
    Row = Square \ 8 + 1
                                                         ' _ '
    GlobalCol(1) = Row - 2 'Cordinates for Knights moves '|
    GlobalRow(1) = Col + 1 ' 8 possible moves            '|
    
    GlobalCol(2) = Row - 1 'You understand this?         ' __
    GlobalRow(2) = Col + 2 'I hope so:)                  '|
    
    GlobalCol(3) = Row + 1
    GlobalRow(3) = Col + 2                               '|__
    
    GlobalCol(4) = Row + 2                               '|
    GlobalRow(4) = Col + 1                               '|_
    
    GlobalCol(5) = Row + 2                               ' |
    GlobalRow(5) = Col - 1                               '_|
    
    GlobalCol(6) = Row + 1                               '__|
    GlobalRow(6) = Col - 2
    
    GlobalCol(7) = Row - 1                               '__
    GlobalRow(7) = Col - 2                               '  |
    
    GlobalCol(8) = Row - 2                               '_
    GlobalRow(8) = Col - 1                               ' |
                                                         ' |
    FreeDest = 0
    For i = 1 To 8
        If (GlobalCol(i) > 0 And GlobalCol(i) <= 8) And (GlobalRow(i) > 0 And GlobalRow(i) <= 8) Then
            If SquareOcupied((GlobalCol(i) - 1) * 8 + (GlobalRow(i) - 1)) = False Then
                LocalSquare = (GlobalCol(i) - 1) * 8 + (GlobalRow(i) - 1)
                FreeDest = FreeDest + 1
                KnightPlaces(i) = True
                BlueDot(i).Move Board(LocalSquare).Left + 170, _
                            Board(LocalSquare).Top + 170
                If ShowHints = True Then BlueDot(i).Visible = True
            Else
                KnightPlaces(i) = False
                BlueDot(i).Visible = False
            End If
        Else
            KnightPlaces(i) = False
            BlueDot(i).Visible = False
        End If
    Next i
    
End Sub

Sub Board_Click(Index As Integer)
Dim Col As Integer
Dim Row As Integer
Dim WrongSquare As Integer
Dim i As Integer
Dim FreeDest As Integer ' Free places in board
Dim FreeDestPlace As Integer
        
    If Timer1.Enabled = True Then 'if program play alone, exit
        Exit Sub
    End If

    Col = Index Mod 8 + 1
    Row = Index \ 8 + 1

    If nKnightsInSquares <> 0 Then
        WrongSquare = False
        For i = 1 To 8
            If (GlobalCol(i) = Row And GlobalRow(i) = Col) Then
                If KnightPlaces(i) = True Then WrongSquare = True
                Exit For
            End If
        Next i
    Else
        'first Knight
        WrongSquare = True
        cmdUndo.Visible = True
    End If

    If WrongSquare = False Then 'if not a good place for Knight then exit
        Beep
        Exit Sub
    End If
    
    RedDot(LastSquare).Visible = False
    LastSquare = Index


    List1.AddItem "Move Number: " & Format$(nKnightsInSquares + 1, "00") & _
                    " to Square: " & Index + 1
    List1.ListIndex = List1.ListCount - 1

    If Board(Index).Picture = imgWhiteBack.Picture Then
        Board(Index).Picture = imgKnightWhite.Picture
        RedDot(Index).Visible = True
    Else
        Board(Index).Picture = imgKnightBlack.Picture
        RedDot(Index).Visible = True
    End If



                                                         ' _ '
    GlobalCol(1) = Row - 2 'Cordinates for Knights moves '|
    GlobalRow(1) = Col + 1 ' 8 possible moves            '|
    
    GlobalCol(2) = Row - 1 'You understand this?         ' __
    GlobalRow(2) = Col + 2 'I hope so:)                  '|
    
    GlobalCol(3) = Row + 1
    GlobalRow(3) = Col + 2                               '|__
    
    GlobalCol(4) = Row + 2                               '|
    GlobalRow(4) = Col + 1                               '|_
    
    GlobalCol(5) = Row + 2                               ' |
    GlobalRow(5) = Col - 1                               '_|
    
    GlobalCol(6) = Row + 1                               '__|
    GlobalRow(6) = Col - 2
    
    GlobalCol(7) = Row - 1                               '__
    GlobalRow(7) = Col - 2                               '  |
    
    GlobalCol(8) = Row - 2                               '_
    GlobalRow(8) = Col - 1                               ' |
                                                         ' |
                                                        ' |

    FreeDest = 0
    For i = 1 To 8 'search for valid places and show hints
            'check for valid moves
        If (GlobalCol(i) > 0 And GlobalCol(i) <= 8) And (GlobalRow(i) > 0 And GlobalRow(i) <= 8) Then
            If SquareOcupied((GlobalCol(i) - 1) * 8 + (GlobalRow(i) - 1)) = False Then
                FreeDestPlace = (GlobalCol(i) - 1) * 8 + (GlobalRow(i) - 1)
                FreeDest = FreeDest + 1
                KnightPlaces(i) = True
                BlueDot(i).Move Board(FreeDestPlace).Left + 170, _
                            Board(FreeDestPlace).Top + 170 'display hint
                If ShowHints = True Then BlueDot(i).Visible = True
            Else
                KnightPlaces(i) = False 'don´t display hint
                BlueDot(i).Visible = False
            End If
        Else
            KnightPlaces(i) = False ' no valid moves found
            BlueDot(i).Visible = False
        End If
    Next i
    
    SquareOcupied(Index) = True
    nKnightsInSquares = nKnightsInSquares + 1

    If nKnightsInSquares = 64 Then
        MsgBox "Well done:)", vbInformation + vbOKOnly
    Else
        If FreeDest = 0 Then MsgBox "Bad luck, try again.", vbInformation + vbOKOnly
    End If
    lblKnights.Caption = "Knight " & nKnightsInSquares & " of 64, left " & _
                            64 - nKnightsInSquares

End Sub



Private Sub Board_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuGame 'Right mouse button show the menu
End Sub



Private Sub Board_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInf.Caption = "Right mouse button show menu"
End Sub

Private Sub cmdStop_Click()

    Timer1.Enabled = False
    cmdUndo.Visible = False
    cmdStop.Visible = False

End Sub




Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInf.Caption = ""
End Sub


Sub cmdUndo_Click()
Dim BoardSquare, LocalBoardSquareTemp As Integer
Dim i As Integer

    RedDot(LastSquare).Visible = False
    nKnightsInSquares = nKnightsInSquares - 1
    BoardSquare = Val(Right$(List1.List(List1.ListCount - 1), 2))

    If Board(BoardSquare - 1).Picture = imgKnightWhite.Picture Then
        Board(BoardSquare - 1).Picture = imgWhiteBack.Picture
    Else
        Board(BoardSquare - 1).Picture = imgBlackBack.Picture
    End If
    
    SquareOcupied(BoardSquare - 1) = False
    List1.RemoveItem List1.ListCount - 1
    
    If List1.ListCount <> 0 Then
        List1.Selected(List1.ListCount - 1) = True
        LocalBoardSquareTemp = Val(Right$(List1.List(List1.ListCount - 1), 2))
        RedDot(LocalBoardSquareTemp - 1).Visible = True
        LastSquare = LocalBoardSquareTemp - 1
        UpdateHints LocalBoardSquareTemp - 1
    Else 'no more undos
        For i = 1 To 8 'Hide hint circles
            BlueDot(i).Visible = False
        Next i
        
        For i = 1 To 8 'set vars to 0
            GlobalCol(i) = 0
            GlobalRow(i) = 0
            KnightPlaces(i) = 0
        Next i
        cmdUndo.Visible = False
    End If
    lblKnights.Caption = "Knight " & nKnightsInSquares & " of 64, left " & _
                            64 - nKnightsInSquares

End Sub

Sub NewGame()
Dim i, Square As Integer

    List1.Clear 'clear the contents in listbox

    RedDot(LastSquare).Visible = False
    
    For i = 1 To 8
        BlueDot(i).Visible = False
    Next i
    
    For Square = 0 To 63
        SquareOcupied(Square) = False
    Next Square
    
        For Square = 1 To 64 'Paint Board Squares
            Select Case Square ' paint with white color only these squares
                Case 1, 3, 5, 7, 10, 12, 14, 16, 17, 19, 21, 23, _
                     26, 28, 30, 32, 33, 35, 37, 39, 42, 44, 46, _
                     48, 49, 51, 53, 55, 58, 60, 62, 64
                    Board(Square - 1).Picture = imgWhiteBack
                Case Else
                    Board(Square - 1).Picture = imgBlackBack
           End Select
        Next Square
        
    
    For i = 1 To 8
        GlobalCol(i) = 0
        GlobalRow(i) = 0
        KnightPlaces(i) = 0
    Next i
    
    nKnightsInSquares = 0
    cmdUndo.Visible = False
    lblKnights.Caption = "Knight " & nKnightsInSquares & " of 64, left " & _
                            64 - nKnightsInSquares
   
End Sub


Sub cmdExit_Click()
    Unload Me
End Sub

Sub mnuHint_Click() 'Is more hard to solve game if this option is disable
Dim i As Integer

    ShowHints = Not ShowHints
    mnuHint.Checked = ShowHints
    
    For i = 1 To 8
        If ShowHints = True Then
            If KnightPlaces(i) = True Then BlueDot(i).Visible = True
        Else
            BlueDot(i).Visible = False
        End If
    Next i
    
End Sub



Sub mnuNew_Click()
    
    NewGame
    
End Sub



Sub mnuSolve_Click()
Dim lSolution As Integer
    Randomize
    NewGame
    cmdUndo.Visible = False
    cmdStop.Visible = True
    lSolution = Rnd(2) + 1 '
    Select Case lSolution
        Case 1
            Solution1st 'read the 1st solution for solve game
        Case 2
            Solution2st 'read the 2st solution for solve game
    End Select
    Timer1.Enabled = True

End Sub


Sub Timer1_Timer() 'Solve game. U can´t solve this game?
Dim Col As Integer
Dim Row As Integer
   
    Col = (ArraySolution(nKnightsInSquares) Mod 8 + 1) - 1
    Row = (ArraySolution(nKnightsInSquares) \ 8 + 1)

    If Board(ArraySolution(nKnightsInSquares) - 1).Picture = imgWhiteBack.Picture Then
        Board(ArraySolution(nKnightsInSquares) - 1).Picture = imgKnightWhite.Picture
        RedDot(ArraySolution(nKnightsInSquares) - 1).Visible = True
    Else
        Board(ArraySolution(nKnightsInSquares) - 1).Picture = imgKnightBlack.Picture
        RedDot(ArraySolution(nKnightsInSquares) - 1).Visible = True
    End If
    
    List1.AddItem "Move Number: " & Format$(nKnightsInSquares + 1, "00") & _
                            " to Square: " & ArraySolution(nKnightsInSquares)
    List1.ListIndex = List1.ListCount - 1

    RedDot(LastSquare).Visible = False 'remove the previous RedDot
    LastSquare = ArraySolution(nKnightsInSquares) - 1 'and save the actual _
                                                    position
    
    SquareOcupied(ArraySolution(nKnightsInSquares) - 1) = True
    nKnightsInSquares = nKnightsInSquares + 1
    
    If nKnightsInSquares = 64 Then 'check if all knights is in board
        Timer1.Enabled = False 'yes, very good :)
        cmdUndo.Visible = True
        cmdStop.Visible = False
        cmdUndo.Visible = False
    End If
    lblKnights.Caption = "Knight " & nKnightsInSquares & " of 64, left " & _
                            64 - nKnightsInSquares
End Sub

Sub Form_Load()
Dim Square As Integer

    cmdStop.Visible = False
    cmdUndo.Visible = False
    DisableFormButton Me 'Disable the X(Close button in main form)
    

    For Square = 1 To 64 'Paint Board Squares
        Select Case Square
            Case 1, 3, 5, 7, 10, 12, 14, 16, 17, 19, 21, 23, _
                26, 28, 30, 32, 33, 35, 37, 39, 42, 44, 46, _
                48, 49, 51, 53, 55, 58, 60, 62, 64
                Board(Square - 1).Picture = imgWhiteBack
            Case Else
                Board(Square - 1).Picture = imgBlackBack
        End Select
    Next Square
    
    ShowHints = True
    nKnightsInSquares = 0
    lblCopyright.Caption = lblCopyright.Caption & Copyright
End Sub


Sub Form_Unload(Cancel As Integer)
    MsgBox "You like this game? Don´t forget to vote:)", vbInformation + vbOKOnly, "Bye bye"
    Unload Me 'bye bye, Dont forget my vote:)
End Sub


