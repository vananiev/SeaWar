VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Морской Бой"
   ClientHeight    =   6600
   ClientLeft      =   3255
   ClientTop       =   2475
   ClientWidth     =   8445
   DrawMode        =   6  'Mask Pen Not
   DrawStyle       =   4  'Dash-Dot-Dot
   FillColor       =   &H0080FF80&
   ForeColor       =   &H000000C0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   6600
   ScaleLeft       =   1
   ScaleMode       =   0  'User
   ScaleTop        =   1
   ScaleWidth      =   8445
   Tag             =   "100"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3720
      Top             =   4320
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Выбор изображения"
      Filter          =   "(*.JPEG,.JPG,.JPE)|*.JPG"
      Flags           =   4100
      MaxFileSize     =   10000
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   3720
      Top             =   4800
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3240
      Top             =   4800
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   7320
      TabIndex        =   12
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   960
      Index           =   6
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   5349
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   960
      Index           =   5
      Left            =   0
      Stretch         =   -1  'True
      Top             =   5349
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   960
      Index           =   4
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   5349
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   960
      Index           =   3
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   5349
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   960
      Index           =   2
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   5349
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   960
      Index           =   1
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   5349
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   960
      Index           =   0
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   5349
      Width           =   1245
   End
   Begin VB.Label Label11 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   25
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label10 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   24
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   23
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Счет:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   22
      Top             =   480
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   199
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   198
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   197
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   196
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   195
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   194
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   193
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   192
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   191
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   190
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   189
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   188
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   187
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   186
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   185
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   184
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   183
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   182
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   181
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   180
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   179
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   178
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   177
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   176
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   175
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   174
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   173
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   172
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   171
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   170
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   169
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   168
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   167
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   166
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   165
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   164
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   163
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   162
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   161
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   160
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   159
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   158
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   157
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   156
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   155
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   154
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   153
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   152
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   151
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   150
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   149
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   148
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   147
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   146
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   145
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   144
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   143
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   142
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   141
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   140
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   139
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   138
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   137
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   136
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   135
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   134
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   133
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   132
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   131
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   130
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   129
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   128
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   127
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   126
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   125
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   124
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   123
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   122
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   121
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   120
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   119
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   118
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   117
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   116
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   115
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   114
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   113
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   112
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   111
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   110
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   109
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   108
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   107
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   106
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   105
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   104
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   103
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   102
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   101
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   100
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   99
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   98
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   97
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   96
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   95
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   94
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   93
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   92
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   91
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   90
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   89
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   88
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   87
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   86
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   85
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   84
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   83
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   82
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   81
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   80
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   79
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   78
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   77
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   76
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   75
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   74
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   73
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   72
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   71
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   70
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   69
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   68
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   67
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   66
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   65
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   64
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   63
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   62
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   61
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   60
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   59
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   58
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   57
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   56
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   55
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   54
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   53
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   52
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   51
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   50
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   49
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   48
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   47
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   46
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   45
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   44
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   43
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   42
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   41
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   40
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   39
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   38
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   37
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   36
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   35
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   34
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   33
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   32
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   31
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   30
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   29
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   28
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   27
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   26
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   25
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   24
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   23
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   22
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   21
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   20
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   19
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   18
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   17
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   16
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   15
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   14
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   13
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   12
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   11
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   10
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   9
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   8
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   7
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   6
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   5
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   4
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   3
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   2
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Label Label5 
      Height          =   375
      Index           =   8
      Left            =   6720
      TabIndex        =   21
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label5 
      Height          =   375
      Index           =   7
      Left            =   6720
      TabIndex        =   20
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label5 
      Height          =   375
      Index           =   6
      Left            =   6720
      TabIndex        =   19
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label5 
      Height          =   375
      Index           =   5
      Left            =   6720
      TabIndex        =   18
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "0"
      Height          =   255
      Left            =   7680
      TabIndex        =   17
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      Height          =   255
      Left            =   7680
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   6360
      Width           =   8415
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   7680
      TabIndex        =   14
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Боевых"
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label5 
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   11
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "1"
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   10
      Top             =   5040
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label5 
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   9
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label5 
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   8
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label5 
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   7
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Офицерских"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Генеральских"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Королевский"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Боевых"
      Height          =   375
      Index           =   9
      Left            =   5160
      TabIndex        =   3
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Генеральских"
      Height          =   375
      Index           =   8
      Left            =   5160
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Офицерских"
      Height          =   375
      Index           =   1
      Left            =   5160
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Королевский"
      Height          =   375
      Index           =   7
      Left            =   5160
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu sm 
         Caption         =   "Сменить рисунок"
      End
      Begin VB.Menu z 
         Caption         =   "Заставка"
         Begin VB.Menu vkl 
            Caption         =   "Вкючить"
         End
         Begin VB.Menu vk 
            Caption         =   "Выключить"
         End
      End
      Begin VB.Menu cld 
         Caption         =   "Слайд"
         Begin VB.Menu cvk 
            Caption         =   "Включить показ"
         End
         Begin VB.Menu cvkl 
            Caption         =   "Выключить показ"
         End
      End
   End
   Begin VB.Menu hel 
      Caption         =   "Help"
      Begin VB.Menu bnx 
         Caption         =   "Больше ничего не хочешь"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim b(39) As Integer: Dim f(39) As Integer: Dim e As Integer: Dim d As Integer: Dim c As Integer: Dim cx As Integer: Dim cy As Integer: Dim n0 As Integer: Dim n1 As Integer: Dim n2 As Integer: Dim h(199) As Integer: Dim i As Integer
 
Private Sub Command1_Click()
If Command1.Caption = "Играть" Then
GoTo 1
Else
End If
Label8.Caption = "0"
For n = 3 To 0 Step -1
If Label3(n).ForeColor = &H8000000D Then
If n = 3 Then
 Label3(3).ForeColor = &H80000012: Label5(3).ForeColor = &H80000012
  Command1.Caption = "Играть"
  Command1.Visible = False
  Label6.Caption = Str(4)
  Label7.Caption = "Ваш ход"
 GoTo 2
 Else
  Label3(n + 1).ForeColor = &H8000000D: Label5(n + 1).ForeColor = &H8000000D
 Label3(n).ForeColor = &H80000012: Label5(n).ForeColor = &H80000012
 Label6 = Str(n + 1)
 GoTo 2
 End If
Else
End If
Next n
1:
PSet (10000, 0)
For n = 0 To 39
PSet (11000, n * 256)
Print n, b(n), f(n)
Next n
2:
End Sub
Private Sub cvk_Click()
Timer3.Enabled = True
End Sub

Private Sub cvkl_Click()
Timer3.Enabled = False
End Sub

Private Sub Form_Load()
For n = 0 To 199
Shape3(n).FillColor = &HFFC0C0
Next n
  Command1.Visible = True
Label6.Caption = "0"
Label8.Caption = "0"
Label9.Caption = "0"
n2 = -1
i = o
For n = 0 To 6
Image1(n).Picture = LoadPicture("C:\Sea\0 100.jpg")
Next n

Command1.Caption = "Next Ship"
Label5(0).Caption = "0"
Label5(1).Caption = "0"
Label5(2).Caption = "0"
Label5(3).Caption = "0"
Label5(5).Caption = "0"
Label5(6).Caption = "0"
Label5(7).Caption = "0"
Label5(8).Caption = "0"
 Label3(0).ForeColor = &H8000000D: Label5(0).ForeColor = &H8000000D
For m = 0 To 1
For n = 0 To 99
Shape3(n + m * 100).Left = 300 + 4860 * m + 300 * (n - Int(n / 10) * 10)
Shape3(n + m * 100).Top = 300 + 300 * (Int(n / 10))
Next n
Next m
For n = 0 To 39
b(n) = -1
Next n
Rem"Корабли противника"
Randomize Timer
c = 0
For l = 0 To 20
4:
e = Int(Rnd * 99)
For m = 0 To 19
If b(m) = e Then GoTo 4
Next m
5:
d = Int(Rnd * 19)
If c + Abs((Int(d / 5)) - 4) > 25 Then GoTo 6
If b(d) <> -1 Then GoTo 5
c = c + Abs((Int(d / 5)) - 4)
b(d) = e
f(d) = Abs(Int(d / 5) - 4)
Label5(5 + Int(d / 5)) = Label5(5 + Int(d / 5)) + 1
Next l
6:
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label7.Caption = ""
If Command1.Caption = "Играть" Then
GoTo 3
Else
For n = 0 To 99
If x > 300 + 300 * (n - Int(n / 10) * 10) And x < 300 + 300 * (n + 1 - Int(n / 10) * 10) And y > 300 + 300 * (Int(n / 10)) And y < 300 + 300 * (1 + Int(n / 10)) Then
If Shape3(n).FillColor = &HFFC0FF Then
Label7.Caption = "Два корабля не могут стоять на одной позиции"
GoTo 7
End If
If Val(Label8.Caption) = 5 Or Val(Label9.Caption) + Abs(Val(Label6.Caption) - 4) > 25 Then Label7.Caption = "Больше кораблей данного типа нет": GoTo 7
b(20 + Val(Label6.Caption) * 5 + Val(Label8.Caption)) = n
f(20 + Val(Label6.Caption) * 5 + Val(Label8.Caption)) = Abs(Val(Label6.Caption) - 4)
Label8.Caption = Str(Val(Label8.Caption) + 1)
Shape3(n).FillColor = &HFFC0FF
Label5(Val(Label6.Caption)) = Label5(Val(Label6.Caption)) + 1
Label9.Caption = Val(Label9.Caption) + Abs(Val(Label6.Caption) - 4)
Label7.Caption = "Осталось снаряжения  " + Str(25 - Val(Label9.Caption))
GoTo 7
Else
End If
Next n
End If
Label7.Caption = "Что это такое?"
GoTo 7
3:
Rem"Игра началась"
If x > 5160 And x < 8175 And y > 300 And y < 3315 Then
cx = x - 5160
cy = y - 300
n0 = Int(cx / 300)
n1 = Int(cy / 300)
n = n1 * 10 + n0
If Shape3(100 + n).FillColor = &HFFC0C0 Or Shape3(100 + n).FillColor = &H80FF80 Then
Shape3(100 + n).FillColor = &H80FFFF
For m = 0 To 19
If b(m) = n Then
f(m) = f(m) - 1
If f(m) = 0 Then
Label7.Caption = "Убит": Label5(5 + Int(m / 5)) = Label5(5 + Int(m / 5)) - 1: b(m) = -1: Shape3(100 + n).FillColor = &HFF&
Else
Label7.Caption = ("Ранен,осталось жизней:  " + Str(f(m))): Shape3(100 + n).FillColor = &H80FF80
End If
Else
End If
Next m

g = 0
For l = 0 To 19
g = g + b(l)
Next l
If g = -20 Then
Label7.Caption = "Вы выиграли": For n = 1 To 10 ^ 7: Next n: Label7.Caption = "Новая игра": Command1.Visible = True
Label2.Caption = Str(Val(Label2.Caption) + 1): Form_Load: GoTo 7
End If
Else
Label7.Caption = "Что это такое?": GoTo 7
End If
Else
Label7.Caption = "Что это такое?": GoTo 7
End If
Rem"Ходит компьютер"
8:
If n2 <> -1 Then
n = n2
GoTo 9
Else
n = Int(Rnd * 99)
End If
For m = 0 To i
If h(m) = n Then GoTo 8
Next m
h(i) = n
i = i + 1
9:
Shape3(n).FillColor = &H80FFFF
For m = 20 To 39
If b(m) = n Then
f(m) = f(m) - 1
If f(m) = 0 Then
Label7.Caption = "Ваш корабль убит": Label5(Int(m / 5) - 4) = Label5(Int(m / 5) - 4) - 1: b(m) = -1: Shape3(n).FillColor = &HFF&: n2 = -1
Else
Label7.Caption = "Ваш корабль ранен,осталось жизней:  " + Str(f(m)): Shape3(n).FillColor = &H80FF80: n2 = n
End If
Else
End If
Next m
g = 0
For l = 20 To 39
g = g + b(l)
Next l
If g = -20 Then
Label7.Caption = "Вы проиграли": For n = 1 To 10 ^ 7:  Next n: Label7.Caption = "Новая игра": Command1.Visible = True: Form_Load
Label11.Caption = Str(Val(Label11.Caption) + 1)
End If
7:
End Sub

Private Sub sm_Click()
CommonDialog1.ShowOpen
For n = 0 To 6
Image1(n).Picture = LoadPicture(CommonDialog1.FileName)
Next n
End Sub

Private Sub Timer1_Timer()
For n = 0 To 9
PSet (Rnd * 10000, Rnd * 10000)
Next n
End Sub

Private Sub Timer2_Timer()
Cls
End Sub

Private Sub Timer3_Timer()
Randomize Timer
For n = 0 To 6
m0 = Int(Rnd * 1000)
If m0 > 0 And m0 <= 100 Then
Image1(n).Picture = LoadPicture("C:\Sea\0" + Str(m0) + ".jpg")
End If
Next n
End Sub

Private Sub vk_Click()
Timer1.Enabled = False
Timer1.Enabled = False
Cls
End Sub

Private Sub vkl_Click()
Timer1.Enabled = True
Timer1.Enabled = True
End Sub
