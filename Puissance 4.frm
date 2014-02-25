VERSION 5.00
Begin VB.Form Partie 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Puissance 4"
   ClientHeight    =   6735
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   9075
   FillColor       =   &H00FF0000&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton changerNoms 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Changer les noms des joueurs"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   6240
      Width           =   3735
   End
   Begin VB.CommandButton restart 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Recommencer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   6240
      Width           =   1695
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   59
      Left            =   7920
      Picture         =   "Puissance 4.frx":0000
      ScaleHeight     =   765
      ScaleWidth      =   810
      TabIndex        =   60
      Top             =   4800
      Width           =   810
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   58
      Left            =   7080
      Picture         =   "Puissance 4.frx":03D4
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   59
      Top             =   4800
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   57
      Left            =   6240
      Picture         =   "Puissance 4.frx":07A8
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   58
      Top             =   4800
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   56
      Left            =   5400
      Picture         =   "Puissance 4.frx":0B7C
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   57
      Top             =   4800
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   55
      Left            =   4560
      Picture         =   "Puissance 4.frx":0F50
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   56
      Top             =   4800
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   54
      Left            =   3720
      Picture         =   "Puissance 4.frx":1324
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   55
      Top             =   4800
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   53
      Left            =   2880
      Picture         =   "Puissance 4.frx":16F8
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   54
      Top             =   4800
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   52
      Left            =   2040
      Picture         =   "Puissance 4.frx":1ACC
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   53
      Top             =   4800
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   51
      Left            =   1200
      Picture         =   "Puissance 4.frx":1EA0
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   52
      Top             =   4800
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   50
      Left            =   360
      Picture         =   "Puissance 4.frx":2274
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   51
      Top             =   4800
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   49
      Left            =   7920
      Picture         =   "Puissance 4.frx":2648
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   50
      Top             =   3960
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   48
      Left            =   7080
      Picture         =   "Puissance 4.frx":2A1C
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   49
      Top             =   3960
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   47
      Left            =   6240
      Picture         =   "Puissance 4.frx":2DF0
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   48
      Top             =   3960
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   46
      Left            =   5400
      Picture         =   "Puissance 4.frx":31C4
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   47
      Top             =   3960
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   45
      Left            =   4560
      Picture         =   "Puissance 4.frx":3598
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   46
      Top             =   3960
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   44
      Left            =   3720
      Picture         =   "Puissance 4.frx":396C
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   45
      Top             =   3960
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   43
      Left            =   2880
      Picture         =   "Puissance 4.frx":3D40
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   44
      Top             =   3960
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   42
      Left            =   2040
      Picture         =   "Puissance 4.frx":4114
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   43
      Top             =   3960
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   41
      Left            =   1200
      Picture         =   "Puissance 4.frx":44E8
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   42
      Top             =   3960
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   40
      Left            =   360
      Picture         =   "Puissance 4.frx":48BC
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   41
      Top             =   3960
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   39
      Left            =   7920
      Picture         =   "Puissance 4.frx":4C90
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   40
      Top             =   3120
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   38
      Left            =   7080
      Picture         =   "Puissance 4.frx":5064
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   39
      Top             =   3120
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   37
      Left            =   6240
      Picture         =   "Puissance 4.frx":5438
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   38
      Top             =   3120
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   36
      Left            =   5400
      Picture         =   "Puissance 4.frx":580C
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   37
      Top             =   3120
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   35
      Left            =   4560
      Picture         =   "Puissance 4.frx":5BE0
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   36
      Top             =   3120
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   34
      Left            =   3720
      Picture         =   "Puissance 4.frx":5FB4
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   35
      Top             =   3120
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   33
      Left            =   2880
      Picture         =   "Puissance 4.frx":6388
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   34
      Top             =   3120
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   32
      Left            =   2040
      Picture         =   "Puissance 4.frx":675C
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   33
      Top             =   3120
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   31
      Left            =   1200
      Picture         =   "Puissance 4.frx":6B30
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   32
      Top             =   3120
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   30
      Left            =   360
      Picture         =   "Puissance 4.frx":6F04
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   31
      Top             =   3120
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   29
      Left            =   7920
      Picture         =   "Puissance 4.frx":72D8
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   30
      Top             =   2280
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   28
      Left            =   7080
      Picture         =   "Puissance 4.frx":76AC
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   29
      Top             =   2280
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   27
      Left            =   6240
      Picture         =   "Puissance 4.frx":7A80
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   28
      Top             =   2280
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   26
      Left            =   5400
      Picture         =   "Puissance 4.frx":7E54
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   27
      Top             =   2280
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   25
      Left            =   4560
      Picture         =   "Puissance 4.frx":8228
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   26
      Top             =   2280
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   24
      Left            =   3720
      Picture         =   "Puissance 4.frx":85FC
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   25
      Top             =   2280
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   23
      Left            =   2880
      Picture         =   "Puissance 4.frx":89D0
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   24
      Top             =   2280
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   810
      Index           =   22
      Left            =   2040
      Picture         =   "Puissance 4.frx":8DA4
      ScaleHeight     =   810
      ScaleWidth      =   765
      TabIndex        =   23
      Top             =   2280
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   21
      Left            =   1200
      Picture         =   "Puissance 4.frx":9178
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   22
      Top             =   2280
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   20
      Left            =   360
      Picture         =   "Puissance 4.frx":954C
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   21
      Top             =   2280
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   19
      Left            =   7920
      Picture         =   "Puissance 4.frx":9920
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   20
      Top             =   1440
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   810
      Index           =   18
      Left            =   7080
      Picture         =   "Puissance 4.frx":9CF4
      ScaleHeight     =   810
      ScaleWidth      =   810
      TabIndex        =   19
      Top             =   1440
      Width           =   810
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   17
      Left            =   6240
      Picture         =   "Puissance 4.frx":A0C8
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   18
      Top             =   1440
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   16
      Left            =   5400
      Picture         =   "Puissance 4.frx":A49C
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   17
      Top             =   1440
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   15
      Left            =   4560
      Picture         =   "Puissance 4.frx":A870
      ScaleHeight     =   765
      ScaleWidth      =   810
      TabIndex        =   16
      Top             =   1440
      Width           =   810
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   14
      Left            =   3720
      Picture         =   "Puissance 4.frx":AC44
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   15
      Top             =   1440
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   13
      Left            =   2880
      Picture         =   "Puissance 4.frx":B018
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   14
      Top             =   1440
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   12
      Left            =   2040
      Picture         =   "Puissance 4.frx":B3EC
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   13
      Top             =   1440
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   11
      Left            =   1200
      Picture         =   "Puissance 4.frx":B7C0
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   12
      Top             =   1440
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      Height          =   765
      Index           =   10
      Left            =   360
      Picture         =   "Puissance 4.frx":BB94
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   11
      Top             =   1440
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   9
      Left            =   7920
      Picture         =   "Puissance 4.frx":BF68
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   10
      Top             =   600
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   8
      Left            =   7080
      Picture         =   "Puissance 4.frx":C33C
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   9
      Top             =   600
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   7
      Left            =   6240
      Picture         =   "Puissance 4.frx":C710
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   8
      Top             =   600
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   6
      Left            =   5400
      Picture         =   "Puissance 4.frx":CAE4
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   7
      Top             =   600
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   5
      Left            =   4560
      Picture         =   "Puissance 4.frx":CEB8
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   6
      Top             =   600
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   4
      Left            =   3720
      Picture         =   "Puissance 4.frx":D28C
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   5
      Top             =   600
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   3
      Left            =   2880
      Picture         =   "Puissance 4.frx":D660
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   4
      Top             =   600
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   2
      Left            =   2040
      Picture         =   "Puissance 4.frx":DA34
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   3
      Top             =   600
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   795
      Index           =   1
      Left            =   1200
      Picture         =   "Puissance 4.frx":DE08
      ScaleHeight     =   795
      ScaleWidth      =   765
      TabIndex        =   1
      Top             =   600
      Width           =   765
   End
   Begin VB.PictureBox Position 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   0
      Left            =   360
      Picture         =   "Puissance 4.frx":E1DC
      ScaleHeight     =   765
      ScaleWidth      =   765
      TabIndex        =   0
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "PUISSANCE 4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3240
      TabIndex        =   61
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label etat 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   8895
   End
End
Attribute VB_Name = "Partie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Forcer la déclaration des variables

Dim tableau(5, 9) As Integer 'Tableau Puissance 4
Dim joueur As Integer 'Numéro du joueur en cours
Dim statut As Integer 'Statut de la partie (1: en cours, 2: un vainqueur, 3: nul)
Dim nomJoueurs(2) As String 'Nom des 2 joueurs

'Timer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Chargement formulaire
Private Sub Form_Load()
    Dim i, j As Integer 'Compteur pour l'initialisation du tableau
    
    'Demander les noms des joueurs à l'aide d'InputBoxes
    Do While nomJoueurs(0) = "" Or nomJoueurs(1) = ""
        nomJoueurs(0) = InputBox("Entrez le nom du joueur 1", "Puissance 4")
        
        'Si le joueur appuie sur le bouton Annuler, arrêter le programme
        If nomJoueurs(0) = "" Then
            Unload Partie
            Exit Sub
        End If
        
        nomJoueurs(1) = InputBox("Entrez le nom du joueur 2", "Puissance 4")
        
        'Si le joueur 2 appuie sur le bouton Annuler, redemander le nom du joueur 1
        If nomJoueurs(0) = "" Then
            MsgBox ("Veuillez entrer un nom pour le joueur 1")
        ElseIf nomJoueurs(1) = "" Then
            MsgBox ("Veuillez entrer un nom pour le joueur 2")
        'Vérifier si les 2 noms ne sont pas identiques
        ElseIf nomJoueurs(0) = nomJoueurs(1) Then
            MsgBox ("Veuillez entrer des noms de joueur différents")
            nomJoueurs(0) = ""
            nomJoueurs(1) = ""
        End If
    Loop
   
    'Initialisation du tableau (0: vide; 1: case remplie par joueur 1; 2: case remplie par joueur 2)
    For i = 0 To 5 Step 1
        For j = 0 To 9 Step 1
            tableau(i, j) = 0
        Next j
    Next i

    'Initialisation des variables joueur (valeurs: 1 ou 2) et gagne (0: personne n'a gagné | 1: un joueur a gagné)
    joueur = ChangeJoueur() 'joueur = 1
    statut = 0
End Sub

'Fonction Changer de joueur
Private Function ChangeJoueur() As Integer
    Dim indexNomJoueur As Integer

    If joueur = 1 Then
        joueur = 2
        etat.ForeColor = &HFF&
    Else
        joueur = 1
        etat.ForeColor = &HFFFF&
    End If
    
    'Afficher "@nomJoueur, à vous de jouer"
    etat = nomJoueurs(joueur - 1) & ", à vous de jouer"
    
    ChangeJoueur = joueur
End Function

'Animation: Pion qui tombe du haut vers le bas jusqu'à la ligne @ligne dans la colonne @colonne
Private Sub Animation(ByVal ligne As Integer, ByVal colonne As Integer)
    Dim posActuelle As Integer 'Case où il faut afficher le pion
    Dim posAnterieure As Integer 'Case où il faut faire disparâitre le pion
    Dim i As Integer 'Compteur

    'Faire tomber le pion jusqu'à la ligne
    For i = 0 To ligne Step 1
        'Veiller que l'on est pas hors du tableau (ligne >= 0)
        If i > 0 Then
            posAnterieure = (i - 1) * 10 + colonne
            Position(posAnterieure).Picture = LoadPicture(App.Path & "\images\fond.gif")
        End If
    
        posActuelle = i * 10 + colonne
        Position(posActuelle).Picture = LoadPicture(App.Path & "\images\j" & joueur & ".gif") 'A
    
        Call Sleep(50)
    Next i
    
End Sub

'On vérifie si 4 pions sont alignés sur la même ligne
Private Function VerificationLigne(ByVal ligne As Integer) As Boolean
    Dim fin As Integer 'Derniere case dans la ligne à verifier
    Dim i As Integer 'Compteur
    Dim retour As Boolean 'Valeur retour
    
    retour = False
    
    'On balaie toute la colonne
    For i = 0 To 9 Step 1
    
        'Pour ne pas être hors du tableau le tableau (max: 9 colonnes)
        If i + 3 > 9 Then
            fin = 9
        Else
            fin = i + 3
        End If
        
        'Si 4 pions se suivent, la fonction retourne 1
        If tableau(ligne, fin) = joueur And tableau(ligne, fin - 1) = joueur And tableau(ligne, fin - 2) = joueur And tableau(ligne, fin - 3) = joueur Then
            retour = True
        End If
    
    Next i
    
    VerificationLigne = retour

End Function

'On vérifie si 4 pions sont alignés sur la même colonne
Private Function VerificationColonne(ByVal colonne As Integer) As Boolean
    Dim fin As Integer 'Dernière case dans la colonne à vérifier
    Dim i As Integer 'Compteur
    Dim retour As Boolean

    retour = False
    
    'On balaie toute la colonne
    For i = 0 To 5 Step 1
    
        'Pour ne pas être hors du tableau (max: 5 lignes)
        If i + 3 > 5 Then
            fin = 5
        Else
            fin = i + 3
        End If
    
        'Si 4 pions se suivent, la fonction retourne 1
        If tableau(fin, colonne) = joueur And tableau(fin - 1, colonne) = joueur And tableau(fin - 2, colonne) = joueur And tableau(fin - 3, colonne) = joueur Then
            retour = True
        End If
        
    Next i
    
    VerificationColonne = retour

End Function

'On vérifie si 4 pions sont alignés en diagonale Gauche-Bas vers Haut-Droite
Private Function VerificationDiagonaleMonte(ByVal colonne As Integer, ByVal ligne As Integer) As Boolean
    Dim verifLigne, verifColonne As Integer 'Référentiel pour la vérification
    Dim i As Integer 'Compteur
    Dim retour As Boolean 'Valeur retour
    
    retour = False
    
    'On se place dans le tableau de manière à vérifier toute la diagonale
    verifColonne = colonne - (5 - ligne)
    verifLigne = 5
    
    'On balaie toute la diagonale
    Do While verifColonne >= 0 And verifColonne + 3 <= 9 And verifLigne - 3 >= 0
        If tableau(verifLigne, verifColonne) And tableau(verifLigne - 1, verifColonne + 1) And tableau(verifLigne - 2, verifColonne + 2) And tableau(verifLigne - 3, verifColonne + 3) Then
            retour = True
        End If

        verifLigne = verifLigne - 1
        verifColonne = verifColonne + 1
    Loop
    
    VerificationDiagonaleMonte = retour
    
End Function

'On vérifie si 4 pions sont alignés en diagonale Haut-Gauche vers Bas-Droite
Private Function VerificationDiagonaleDescend(ByVal colonne As Integer, ByVal ligne As Integer) As Boolean
    Dim verifLigne, verifColonne As Integer 'Référentiel pour la vérification
    Dim i As Integer 'Compteur
    Dim retour As Boolean 'Valeur retour
    
    retour = False
    
    'On se place de manière à vérifier toute la diagonale
    verifColonne = colonne + (5 - ligne)
    
    'Si l'on est pas hors du tableau, on balaie la diagonale
    If verifColonne <= 9 Then
        verifLigne = 5
        
        Do While verifColonne >= 0 And verifColonne - 3 >= 0 And verifLigne - 3 >= 0
            If tableau(verifLigne, verifColonne) And tableau(verifLigne - 1, verifColonne - 1) And tableau(verifLigne - 2, verifColonne - 2) And tableau(verifLigne - 3, verifColonne - 3) Then
                retour = True
            End If

            verifLigne = verifLigne - 1
            verifColonne = verifColonne - 1
        Loop
    End If
    
    VerificationDiagonaleDescend = retour
    
End Function

'Vérifier si toutes les cases sont remplies
Private Function ToutRempli() As Boolean
    Dim i, j As Integer 'Compteurs pour le tableau
    Dim compteur As Integer 'Compteurs de pions remplis
    Dim retour As Boolean 'Valeur retour
    
    compteur = 0
    retour = False
    
    For i = 0 To 5 Step 1
        For j = 0 To 9 Step 1
            If tableau(i, j) > 0 Then
                compteur = compteur + 1
            End If
        Next j
    Next i
    
    'Si toutes les cases sont remplis, mettre fin au jeu
    If compteur >= 60 Then
        retour = True
    End If
        
    ToutRempli = retour
    
End Function

'Bouton Recommencer
Private Sub restart_Click()
    Dim i As Integer
    
    'Vider la grille
    For i = 0 To 59 Step 1
        Position(i).Picture = LoadPicture(App.Path & "\images\fond.gif")
    Next i

    Call Form_Load
End Sub

'Bouton Changer les noms
Private Sub changerNoms_Click()
    Dim i As Integer
    
    nomJoueurs(0) = ""
    nomJoueurs(1) = ""
    
    'Vider la grille
    For i = 0 To 59 Step 1
        Position(i).Picture = LoadPicture(App.Path & "\images\fond.gif")
    Next i
    
    Call Form_Load
End Sub

Private Sub Position_Click(Index As Integer)
    Dim ligne, colonne As Integer 'Ligne et colonne cliquée
    Dim gagneLigne, gagneColonne, gagneDiagonaleMonte, gagneDiagonaleDescend As Boolean 'Si l'une de ces 4 variables devient vrai, un joueur a gagné
    
    'On voit si personne n'a encore gagné
    If statut = 0 Then
        
        If ToutRempli() = True Then
            etat.ForeColor = &HFFFFFF
            etat = "Fin. Match nul"
        Else
            colonne = Index Mod 10 'Récupérer le numéro de colonne
    
            'On voit si des points ont déja été ajoutés dans la colonne. Pour cela, on remonte dans le tableau si un pion est placé
            ligne = 5
            Do Until tableau(ligne, colonne) = 0 Or ligne = 0
                ligne = ligne - 1
            Loop
        
            'On rajoute un nouveau pion
            If tableau(ligne, colonne) = 0 Then
                tableau(ligne, colonne) = joueur
    
                Call Animation(ligne, colonne)
                
                'On voit si le joueur a 4 pions alignés dans tous les sens
                gagneLigne = VerificationLigne(ligne)
                gagneColonne = VerificationColonne(colonne)
                gagneDiagonaleMonte = VerificationDiagonaleMonte(colonne, ligne)
                gagneDiagonaleDescend = VerificationDiagonaleDescend(colonne, ligne)
                
                'Si une suite de 4 a été trouvé soit en ligne soit en colonne soit en diagonale, le joueur a gagné
                If gagneLigne = True Or gagneColonne = True Or gagneDiagonaleMonte = True Or gagneDiagonaleDescend = True Then
                    statut = 1
                    etat = nomJoueurs(joueur - 1) & " a gagné !"
                Else
                    joueur = ChangeJoueur()
                End If
            End If
        End If
    End If
End Sub
