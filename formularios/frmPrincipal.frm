VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrincipal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculadora Dias - ALT + F2 para abrir"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin TabDlg.SSTab SSTab1 
      CausesValidation=   0   'False
      Height          =   3030
      Left            =   165
      TabIndex        =   24
      Top             =   1065
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   5345
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Cálc. de datas"
      TabPicture(0)   =   "frmPrincipal.frx":9B92
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblResultado"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblDiaSemana"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtDias"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtData"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "btnCalcular"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Dif. Datas"
      TabPicture(1)   =   "frmPrincipal.frx":9BAE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(2)=   "Label10"
      Tab(1).Control(3)=   "Label11"
      Tab(1).Control(4)=   "lblDias"
      Tab(1).Control(5)=   "lblDias2"
      Tab(1).Control(6)=   "dtFinal"
      Tab(1).Control(7)=   "dtInicial"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Dif. Horas"
      TabPicture(2)   =   "frmPrincipal.frx":9BCA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "btnCalc2"
      Tab(2).Control(1)=   "txtHora1"
      Tab(2).Control(2)=   "txtHora2"
      Tab(2).Control(3)=   "Label15"
      Tab(2).Control(4)=   "Label14"
      Tab(2).Control(5)=   "lblResultado2"
      Tab(2).Control(6)=   "Label12"
      Tab(2).Control(7)=   "Label7"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Cálc. Horas"
      TabPicture(3)   =   "frmPrincipal.frx":9BE6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblHora1"
      Tab(3).Control(1)=   "Label17"
      Tab(3).Control(2)=   "Label18"
      Tab(3).Control(3)=   "Label19"
      Tab(3).Control(4)=   "lblResultadoCalc"
      Tab(3).Control(5)=   "lblResultadoCalc1"
      Tab(3).Control(6)=   "MaskEdBox1"
      Tab(3).Control(7)=   "Command1"
      Tab(3).Control(8)=   "List1"
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "Horas B.10"
      TabPicture(4)   =   "frmPrincipal.frx":9C02
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblCalcB10"
      Tab(4).Control(1)=   "Label21"
      Tab(4).Control(2)=   "Label22"
      Tab(4).Control(3)=   "Label23"
      Tab(4).Control(4)=   "lblHora2"
      Tab(4).Control(5)=   "lblCalcB101"
      Tab(4).Control(6)=   "M"
      Tab(4).Control(7)=   "List2"
      Tab(4).Control(8)=   "Command2"
      Tab(4).ControlCount=   9
      TabCaption(5)   =   "Percentual"
      TabPicture(5)   =   "frmPrincipal.frx":9C1E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label20"
      Tab(5).Control(1)=   "lblResultadoPercentual"
      Tab(5).Control(2)=   "lbl1"
      Tab(5).Control(3)=   "lbl2"
      Tab(5).Control(4)=   "Label28"
      Tab(5).Control(5)=   "lblAjuda"
      Tab(5).Control(6)=   "txtValor1"
      Tab(5).Control(7)=   "txtPercentual"
      Tab(5).Control(8)=   "btnCalc3"
      Tab(5).Control(9)=   "Option1"
      Tab(5).Control(10)=   "Option2"
      Tab(5).Control(11)=   "Option3"
      Tab(5).Control(12)=   "Option4"
      Tab(5).Control(13)=   "Option5"
      Tab(5).ControlCount=   14
      TabCaption(6)   =   "Desligar PC"
      TabPicture(6)   =   "frmPrincipal.frx":9C3A
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "btnDesligar"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "txtDesligar"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Label13"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).ControlCount=   3
      Begin VB.CommandButton btnDesligar 
         Caption         =   "Desligar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -67425
         TabIndex        =   26
         Top             =   1455
         Width           =   1560
      End
      Begin VB.OptionButton Option5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Soma"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -70635
         TabIndex        =   19
         Top             =   750
         Width           =   1200
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Equivalência"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -67515
         TabIndex        =   21
         Top             =   735
         Width           =   2235
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Subtração"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -69345
         TabIndex        =   20
         Top             =   735
         Width           =   1920
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Lucro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -71895
         TabIndex        =   18
         Top             =   735
         Width           =   1290
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Correspondencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -74745
         TabIndex        =   17
         Top             =   735
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.CommandButton btnCalc3 
         Caption         =   "Calcular"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -70440
         TabIndex        =   16
         Top             =   1590
         Width           =   1560
      End
      Begin VB.TextBox txtPercentual 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   -72135
         TabIndex        =   15
         Top             =   1560
         Width           =   1620
      End
      Begin VB.TextBox txtValor1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   -74310
         TabIndex        =   14
         Top             =   1560
         Width           =   2130
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Adicionar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -74655
         TabIndex        =   12
         Top             =   1695
         Width           =   1860
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         ItemData        =   "frmPrincipal.frx":9C56
         Left            =   -72285
         List            =   "frmPrincipal.frx":9C58
         TabIndex        =   13
         Top             =   435
         Width           =   3150
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         ItemData        =   "frmPrincipal.frx":9C5A
         Left            =   -72285
         List            =   "frmPrincipal.frx":9C5C
         TabIndex        =   10
         Top             =   435
         Width           =   3150
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Adicionar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -74655
         TabIndex        =   9
         Top             =   1695
         Width           =   1860
      End
      Begin VB.CommandButton btnCalc2 
         Caption         =   "Calcular"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -70920
         TabIndex        =   7
         Top             =   1680
         Width           =   1560
      End
      Begin MSMask.MaskEdBox txtHora1 
         Height          =   450
         Left            =   -74400
         TabIndex        =   5
         Top             =   1680
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   794
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtInicial 
         Height          =   450
         Left            =   -74805
         TabIndex        =   3
         Top             =   1785
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   794
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51052547
         CurrentDate     =   42767
      End
      Begin VB.CommandButton btnCalcular 
         Caption         =   "Calcular"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3015
         TabIndex        =   2
         Top             =   1695
         Width           =   1560
      End
      Begin MSMask.MaskEdBox txtData 
         Height          =   405
         Left            =   210
         TabIndex        =   0
         Top             =   1725
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDias 
         Height          =   405
         Left            =   2130
         TabIndex        =   1
         Top             =   1725
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtFinal 
         Height          =   450
         Left            =   -72720
         TabIndex        =   4
         Top             =   1770
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   794
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8421504
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51052547
         CurrentDate     =   42767
      End
      Begin MSMask.MaskEdBox txtHora2 
         Height          =   450
         Left            =   -72645
         TabIndex        =   6
         Top             =   1680
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   794
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   450
         Left            =   -74655
         TabIndex        =   8
         Top             =   1170
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
         _Version        =   393216
         Appearance      =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "9##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox M 
         Height          =   450
         Left            =   -74655
         TabIndex        =   11
         Top             =   1170
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "9##,#"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDesligar 
         Height          =   450
         Left            =   -69330
         TabIndex        =   25
         Top             =   1455
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
         _Version        =   393216
         Appearance      =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora que o computador desliga"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74205
         TabIndex        =   63
         Top             =   1485
         Width           =   4755
      End
      Begin VB.Label lblDias2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00 Dias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   420
         Left            =   -70380
         TabIndex        =   62
         Top             =   2235
         Width           =   5085
      End
      Begin VB.Label lblResultadoCalc1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00:00 Horas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   420
         Left            =   -68520
         TabIndex        =   61
         Top             =   2055
         Width           =   3060
      End
      Begin VB.Label lblCalcB101 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00 Horas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   420
         Left            =   -68520
         TabIndex        =   60
         Top             =   2055
         Width           =   3060
      End
      Begin VB.Label lblAjuda 
         Height          =   225
         Left            =   -68460
         TabIndex        =   59
         Top             =   2115
         Width           =   2955
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -68775
         TabIndex        =   58
         Top             =   1635
         Width           =   180
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Percentual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72135
         TabIndex        =   57
         Top             =   1185
         Width           =   1620
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74340
         TabIndex        =   56
         Top             =   1185
         Width           =   765
      End
      Begin VB.Label lblResultadoPercentual 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   420
         Left            =   -68520
         TabIndex        =   55
         Top             =   1605
         Width           =   3060
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   -68505
         TabIndex        =   54
         Top             =   1140
         Width           =   1530
      End
      Begin VB.Label lblHora2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74655
         TabIndex        =   53
         Top             =   780
         Width           =   705
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+ Soma hora"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   -74640
         TabIndex        =   52
         Top             =   2265
         Width           =   1650
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- Subtrai hora"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   -74640
         TabIndex        =   51
         Top             =   2550
         Width           =   2100
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   -68505
         TabIndex        =   50
         Top             =   1140
         Width           =   1530
      End
      Begin VB.Label lblCalcB10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00:00 Horas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   420
         Left            =   -68520
         TabIndex        =   49
         Top             =   1620
         Width           =   3060
      End
      Begin VB.Label lblResultadoCalc 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00 Horas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   420
         Left            =   -68520
         TabIndex        =   48
         Top             =   1605
         Width           =   3060
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   -68505
         TabIndex        =   47
         Top             =   1140
         Width           =   1530
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- Subtrai hora"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   -74640
         TabIndex        =   46
         Top             =   2550
         Width           =   2100
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+ Soma hora"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   -74640
         TabIndex        =   45
         Top             =   2265
         Width           =   1650
      End
      Begin VB.Label lblHora1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74655
         TabIndex        =   44
         Top             =   780
         Width           =   705
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -69255
         TabIndex        =   43
         Top             =   1725
         Width           =   180
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   -69000
         TabIndex        =   42
         Top             =   1230
         Width           =   1530
      End
      Begin VB.Label lblResultado2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00 Horas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   420
         Left            =   -69015
         TabIndex        =   41
         Top             =   1695
         Width           =   3060
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora Final"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72675
         TabIndex        =   40
         Top             =   1230
         Width           =   1515
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora Inicial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74400
         TabIndex        =   39
         Top             =   1230
         Width           =   1650
      End
      Begin VB.Label lblDias 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00 Dias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   420
         Left            =   -70380
         TabIndex        =   38
         Top             =   1785
         Width           =   5085
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   -70365
         TabIndex        =   37
         Top             =   1335
         Width           =   1530
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -70620
         TabIndex        =   36
         Top             =   1815
         Width           =   180
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Final"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72735
         TabIndex        =   35
         Top             =   1335
         Width           =   1500
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inicial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74805
         TabIndex        =   34
         Top             =   1335
         Width           =   1635
      End
      Begin VB.Label lblDiaSemana 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   420
         Left            =   6720
         TabIndex        =   33
         Top             =   1710
         Width           =   3060
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dia da semana"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   6735
         TabIndex        =   32
         Top             =   1290
         Width           =   2235
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inicial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   31
         Top             =   1290
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2130
         TabIndex        =   30
         Top             =   1290
         Width           =   660
      End
      Begin VB.Label lblResultado 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/01/2017"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   420
         Left            =   4935
         TabIndex        =   29
         Top             =   1710
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   4950
         TabIndex        =   28
         Top             =   1290
         Width           =   1530
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4695
         TabIndex        =   27
         Top             =   1740
         Width           =   180
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   30
      Left            =   135
      TabIndex        =   23
      Top             =   960
      Width           =   6585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CALCULADORA DE DIAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   510
      Left            =   1365
      TabIndex        =   22
      Top             =   210
      Width           =   5145
   End
   Begin VB.Image Image1 
      Height          =   810
      Left            =   390
      Picture         =   "frmPrincipal.frx":9C5E
      Top             =   60
      Width           =   810
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OldWndProc As Long

Private Sub btnCalc2_Click()
On Error Resume Next
If IsDate(txtHora1.Text) = False Or IsDate(txtHora2.Text) = False Then
    lblResultado2.Caption = "xx Horas"
Else
    lblResultado2.Caption = DiferencaHoras(txtHora1.Text, txtHora2.Text) & " Hora(s)"
End If
End Sub

Private Sub btnCalc3_Click()
On Error Resume Next
If txtValor1.Text = "" Or txtPercentual.Text = "" Then
    MsgBox "Informe corretamente todos os campos antes de continuar!", vbCritical, "Erro"
    txtValor1.SetFocus
    Exit Sub
End If

    Dim P As Double
    If Option4.Value = False Then
        If CDbl(txtPercentual.Text) > 100 Then
            MsgBox "Atenção, percentual é no máximo 100%", vbCritical, "Erro"
            txtPercentual.SetFocus
            Exit Sub
        End If
        If (CDbl(100 - CDbl(txtPercentual.Text)) - CInt(100 - CDbl(txtPercentual.Text))) = 0 Then
            P = CDbl("0," & CStr(100 - CDbl(txtPercentual.Text)))
        Else
            P = CDbl(CStr(100 - CDbl(txtPercentual.Text)))
        End If
    End If
    lblAjuda.Caption = ""
    
If Option1.Value = True Then
    lblAjuda.Caption = ""
    lblResultadoPercentual.Caption = "R$ " & Format((CDbl(txtValor1.Text) / 100) * CDbl(txtPercentual.Text), "#,##0.00")
ElseIf Option2.Value = True Then
    lblAjuda.Caption = "Diferença de R$" & Format((CDbl(txtValor1.Text) / P) - CDbl(txtValor1.Text), "#,##0.00")
    lblResultadoPercentual.Caption = "R$" & Format(CDbl(txtValor1.Text) / P, "#,##0.00")
ElseIf Option3.Value = True Then
    lblAjuda.Caption = "Diferença de R$" & Format(CDbl(txtValor1.Text) - (CDbl(txtValor1.Text) * P), "#,##0.00")
    lblResultadoPercentual.Caption = "R$" & Format((CDbl(txtValor1.Text) * P), "#,##0.00")
ElseIf Option4.Value = True Then
    lblAjuda.Caption = ""
    lblResultadoPercentual.Caption = Format(((CDbl(txtPercentual.Text) / CDbl(txtValor1.Text) * 100)), "#,##0.00") & " %"
ElseIf Option5.Value = True Then
    Dim p2 As Double
    p2 = (1 - P) + 1
    lblAjuda.Caption = "Diferença de R$ " & Format((CDbl(txtValor1.Text) * p2) - CDbl(txtValor1.Text), "#,##0.00")
    lblResultadoPercentual.Caption = Format(CDbl(txtValor1.Text) * p2, "#,##0.00")
End If

End Sub

Private Sub btnCalcular_Click()
On Error Resume Next
    If IsDate(Trim(txtData.Text)) = False Then
        MsgBox "Data inválida, verifique e tente novamente.", vbCritical, "Data inválida."
        txtData.SetFocus
        txtData.SelStart = 0
        txtData.SelLength = Len(txtData.Text)
        Exit Sub
    End If
    
    If IsNumeric(txtDias.Text) = False Then
        MsgBox "Informe o número de dias corretamente antes de continuar!", vbCritical, "Verifique e tente novamente."
        txtDias.SetFocus
        txtDias.SelLength = Len(txtDias.Text)
        Exit Sub
    End If
        
    Dim vData As Date
    Dim vDias As Integer
    
    vData = txtData.Text
    vDias = txtDias.Text
    
    lblResultado.Caption = Format(CDate(DateValue(CDate(txtData.Text)) + CDbl(txtDias.Text)), "dd/MM/yyyy")
    lblDiaSemana.Caption = DiaSemana(CDate(lblResultado.Caption))
    
End Sub

Private Sub btnDesligar_Click()
    Dim vData As String
    Dim vHora As String
    
    If btnDesligar.Caption = "Desligar" Then
        If IsDate(txtDesligar) Then
            btnDesligar.Caption = "Cancelar"
            If CDate(txtDesligar.Text) < CDate(Format(Now, "HH:mm")) Then
                vData = DateValue(Now) + 1
                vHora = txtDesligar.Text
            Else
                vData = DateValue(Now)
                vHora = txtDesligar.Text
            End If
            
            Shell "shutdown -s -f -t " & DateDiff("s", Now, CDate(Format(CDate(vData), "dd/MM/yyyy") & " " & Format(CDate(vHora), "hh:mm")))
        Else
            MsgBox "Hora inválida!", vbCritical, "Verifique e tente novamente."
        End If
    Else
        Shell "shutdown -a"
        btnDesligar.Caption = "Desligar"
     End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    If (IsDate(Right(MaskEdBox1.Text, Len(MaskEdBox1) - 1))) Then
        List1.AddItem MaskEdBox1.Text
        lblResultadoCalc.Caption = SomaHora() & " Horas"
        
        Dim s As String
        s = SomaHora()
        lblResultadoCalc1.Caption = Val(Left(s, Len(s) - 3)) & "," & Format((CDbl(Right(s, 2)) / 60) * 100, "00") & " Horas"
        
        MaskEdBox1.Text = "+__:__"
        MaskEdBox1.SetFocus
    Else
        MsgBox "Informe uma hora válida antes de continuar", vbCritical, "Hora inválida..."
        MaskEdBox1.SetFocus
        Exit Sub
    End If
End Sub

Function SomaHora() As String
On Error Resume Next
Dim x As Long
Dim total, horas As Double

For x = 0 To List1.ListCount - 1

If Left(List1.List(x), 1) = "+" Then
    total = total + TimeValue(Right(List1.List(x), Len(List1.List(x)) - 1))
Else
    total = total - TimeValue(Right(List1.List(x), Len(List1.List(x)) - 1))
End If
Next x

horas = Int(total)
horas = horas * 24


restohora = Format(total, "hh:mm")

horas = Mid(restohora, 1, InStr(1, restohora, ":") - 1) + horas

totalhoras = horas & Mid(restohora, InStr(1, restohora, ":"), Len(restohora))




SomaHora = totalhoras

End Function
Function SomaHora2() As String
On Error Resume Next
Dim x As Long
Dim total, horas As Double

For x = 0 To List2.ListCount - 1

If Left(List2.List(x), 1) = "+" Then
    total = total + CDbl(Right(List2.List(x), Len(List2.List(x)) - 1))
Else
    total = total - CDbl(Right(List2.List(x), Len(List2.List(x)) - 1))
End If
Next x

SomaHora2 = Format(total, "#,##0.0")

End Function
Private Sub Command2_Click()
On Error Resume Next
If Val(Right(M.Text, Len(M.Text) - 3)) >= 10 Then
    MsgBox "Os minutos são no máximo 9", vbCritical, "Campo incorreto."
    M.SetFocus
    Exit Sub
End If
List2.AddItem Left(M.Text, 3) & "," & Val(Right(M.Text, Len(M.Text) - 3))

lblCalcB10.Caption = SomaHora2() & " Horas"
Preenche1

M.Text = "+"
M.SelStart = 1
M.SetFocus
End Sub
Private Sub Preenche1()
On Error Resume Next
    If IsNumeric(SomaHora2()) Then
            Dim v() As String
            v = Split(SomaHora2, ",")
            lblCalcB101.Caption = MontaHora((CDbl(v(0)) * 3600) + (CDbl("0," & Right(SomaHora2(), 1)) * 60) * 60) & " Horas"
    Else
            lblCalcB101.Caption = "00:00 Horas"
    End If
End Sub
Private Sub Preenche2()
    
End Sub
Private Sub dtFinal_Change()
lblDias.Caption = DiferencaDias(dtInicial.Value, dtFinal.Value)
lblDias2.Caption = DiferencaemDias(dtInicial.Value, dtFinal.Value) & " Dias"
End Sub

Private Sub dtFinal_Click()
lblDias.Caption = DiferencaDias(dtInicial.Value, dtFinal.Value)
lblDias2.Caption = DiferencaemDias(dtInicial.Value, dtFinal.Value) & " Dias"
End Sub

Private Sub dtInicial_Change()
lblDias.Caption = DiferencaDias(dtInicial.Value, dtFinal.Value)
lblDias2.Caption = DiferencaemDias(dtInicial.Value, dtFinal.Value) & " Dias"
End Sub

Private Sub dtInicial_Click()
    lblDias.Caption = DiferencaDias(dtInicial.Value, dtFinal.Value)
    lblDias2.Caption = DiferencaemDias(dtInicial.Value, dtFinal.Value) & " Dias"
End Sub

Private Sub Form_Load()
Call AlwaysOnTop(Me, True)
Ret = RegisterHotKey(Me.hWnd, 1, MOD_ALT, _
VK_F2)
txtDesligar.Text = Format(Now, "HH:mm")
txtData.Text = Format(Date, "dd/MM/yyyy")
txtDias.Text = "060"
dtInicial.Value = Date
dtFinal.Value = Date
txtHora1.Text = "07:30"
txtHora2.Text = "12:00"
btnCalcular_Click
btnCalc2_Click

    With nid
    .cbSize = Len(nid)
    .hWnd = Me.hWnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon
    .szTip = "Cálculo Data - Hora" & vbNullChar
  End With
  
  Shell_NotifyIcon NIM_ADD, nid
  
  
  OldWndProc = SetWindowLong(Me.hWnd, _
GWL_WNDPROC, AddressOf _
WindowProc)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
  'this procedure receives the callbacks from the System Tray icon.
  Dim Result As Long
  Dim msg As Long
  
  'the value of X will vary depending upon the scalemode setting
  If Me.ScaleMode = vbPixels Then
    msg = x
  Else
    msg = x / Screen.TwipsPerPixelX
  End If
  
  Select Case msg
  Case WM_LBUTTONUP        '514 restore form window
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hWnd)
    Me.Show
    Call AlwaysOnTop(Me, True)
  Case WM_LBUTTONDBLCLK    '515 restore form window
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hWnd)
    Me.Show
    Call AlwaysOnTop(Me, True)
  Case WM_RBUTTONUP        '517 display popup menu
    Result = SetForegroundWindow(Me.hWnd)
    On Error Resume Next

  End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Me.WindowState = vbMinimized
Cancel = True
Call AlwaysOnTop(Me, False)
'Call SetWindowLong(Me.hwnd, GWL_WNDPROC, OldWndProc)

End Sub
Private Sub Form_Resize()
  'this is necessary to assure that the minimized window is hidden
  If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Public Sub BahMetodo()
On Error Resume Next
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hWnd)
    Me.Show
    Call AlwaysOnTop(Me, True)
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim men As Integer
If List1.Text <> "" And KeyCode = 46 Then
    men = MsgBox("Tem certeza que deseja excluir a hora selecionada?", vbQuestion + vbYesNo, "Hora selecionada: " & List1.Text)
        If men = vbYes Then
            List1.RemoveItem (List1.ListIndex)
            If CDbl(List1.ListCount) > 0 Then
                lblResultadoCalc.Caption = SomaHora() & " Horas"
                Dim s As String
                s = SomaHora()
                lblResultadoCalc1.Caption = Val(Left(s, Len(s) - 3)) & "," & Format((CDbl(Right(s, 2)) / 60) * 100, "00") & " Horas"
            Else
                lblResultadoCalc.Caption = "00 Horas"
                lblResultadoCalc1.Caption = "00:00 Horas"
            End If
        End If
End If
End Sub

Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim men As Integer
If List2.Text <> "" And KeyCode = 46 Then
    men = MsgBox("Tem certeza que deseja excluir a hora selecionada?", vbQuestion + vbYesNo, "Hora selecionada: " & List2.Text)
        If men = vbYes Then
            List2.RemoveItem (List2.ListIndex)
            If List2.ListCount > 0 Then
                lblCalcB10.Caption = SomaHora2() & " Horas"
                Preenche1
            Else
                lblCalcB10.Caption = "00:00 Horas"
                lblCalcB101.Caption = "00 Horas"
             End If
                
        End If
End If
End Sub

Private Sub M_Change()
On Error Resume Next
If IsNumeric(Right(M.Text, Len(M.Text) - 1)) Then
    
    If Len(Trim(M.Text)) = 4 Then
        lblHora2.Caption = "Hora - " & MontaHora((CDbl(Right(Left(M.Text, 3), 2)) * 3600) + (CDbl("0," & Right(M.Text, 1)) * 60) * 60)
    Else
        lblHora2.Caption = "Hora"
    End If
    
End If
End Sub

Private Sub M_GotFocus()
If M.Text = "" Then
    M.Text = "+"
    M.SelStart = 1
Else
M.SelStart = 1
End If
End Sub

Private Sub MaskEdBox1_Change()
On Error Resume Next
If IsNumeric(Right(MaskEdBox1.Text, 2)) Then
    lblHora1.Caption = "Hora - " & Val(Right(Left(MaskEdBox1.Text, 3), 2)) & "," & Format((CDbl(Right(MaskEdBox1.Text, 2)) / 60) * 100, "00")
Else
    lblHora1.Caption = "Hora"
End If
End Sub

Private Sub MaskEdBox1_GotFocus()
If MaskEdBox1.Text = "___:__" Then
    MaskEdBox1.Text = "+__:__"
    MaskEdBox1.SelStart = 1
Else
MaskEdBox1.SelStart = 1
End If
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 43 And KeyAscii <> 45 And KeyAscii <> 8 And MaskEdBox1.SelStart = 0 Then KeyAscii = 0
End Sub

Private Sub Option1_Click()
lbl1.Caption = "Valor"
lbl2.Caption = "Percentual"
txtValor1.Text = ""
txtPercentual.Text = ""
lblResultadoPercentual.Caption = ""
txtValor1.SetFocus
End Sub

Private Sub Option2_Click()
lbl1.Caption = "Valor"
lbl2.Caption = "Percentual"
txtValor1.Text = ""
txtPercentual.Text = ""
lblResultadoPercentual.Caption = ""
txtValor1.SetFocus
End Sub

Private Sub Option3_Click()
lbl1.Caption = "Valor"
lbl2.Caption = "Percentual"
txtValor1.Text = ""
txtPercentual.Text = ""
lblResultadoPercentual.Caption = ""
txtValor1.SetFocus
End Sub

Private Sub Option4_Click()
lbl1.Caption = "Valor 01"
lbl2.Caption = "Valor 02"
txtValor1.Text = ""
txtPercentual.Text = ""
lblResultadoPercentual.Caption = ""
txtValor1.SetFocus
End Sub
Private Sub Option5_Click()
lbl1.Caption = "Valor"
lbl2.Caption = "Percentual"
txtValor1.Text = ""
txtPercentual.Text = ""
lblResultadoPercentual.Caption = ""
txtValor1.SetFocus
End Sub
Private Sub txtData_GotFocus()
txtData.SelStart = 0
txtData.SelLength = Len(txtData.Text)
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub txtDesligar_GotFocus()
txtDesligar.SelStart = 0
txtDesligar.SelLength = Len(txtDesligar.Text)
End Sub

Private Sub txtDias_GotFocus()
txtDias.SelStart = 0
txtDias.SelLength = Len(txtDias.Text)
End Sub

Private Sub txtDias_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub txtDias_LostFocus()
txtDias.Text = Format(Val(txtDias.Text), "000")
End Sub
Public Function DiaSemana(Data As Date) As String
On Error Resume Next
    Dim dia As String
    
    Select Case Weekday(Data)
        Case 1:
            dia = "Domingo"
        Case 2:
            dia = "Segunda-Feira"
        Case 3:
            dia = "Terça-Feira"
        Case 4:
            dia = "Quarta-Feira"
        Case 5:
            dia = "Quinta-Feira"
        Case 6:
            dia = "Sexta-Feira"
        Case 7:
            dia = "Sábado"
    End Select
    DiaSemana = dia
End Function
Public Function DiferencaDias(data1 As Date, data2 As Date) As String
On Error Resume Next
    If data2 < data1 Then
        DiferencaDias = "xx"
    Else
        DiferencaDias = fncIdadeCompleta(data2, data1)
    End If
End Function
Public Function DiferencaemDias(data1 As Date, data2 As Date) As String
On Error Resume Next
    If data2 < data1 Then
        DiferencaemDias = "xx"
    Else
        DiferencaemDias = DateDiff("d", data1, data2)
    End If
End Function

Public Function fncIdadeCompleta(D1 As Date, d2 As Date) As String
Dim Anos As Byte, Meses, Dias As Byte, DataRef As Date
Dim Resultado As Boolean
 
If d2 > D1 Or d2 = 0 Then
    fncIdadeCompleta = ""
    Exit Function
End If

If d2 = D1 Then
    fncIdadeCompleta = 0
    Exit Function
End If
 
'Ajusta ano bissexto
d2 = IIf(Format(d2, "mm/dd") = "02/29", d2 - 1, d2)
 
Anos = Int((Format(D1, "yyyymmdd") - Format(d2, "yyyymmdd")) / 10000)
 
Resultado = (Format(d2, "mmdd") > Format(D1, "mmdd"))

DataRef = DateSerial(Year(D1) + Resultado, Format(d2, "mm"), Format(d2, "dd"))

Meses = DateDiff("m", DataRef, D1) + (Format(d2, "dd") > Format(D1, "dd"))
 
Resultado = (Format(d2, "dd") > Format(D1, "dd"))

DataRef = DateSerial(Year(D1), Format(D1, "mm") + Resultado, Format(d2, "dd"))
DataRef = IIf(Format(d2, "dd") <> Format(DataRef, "dd"), DataRef - Format(DataRef, "dd"), DataRef)

Dias = CDbl(D1) - CDbl(DataRef)

fncIdadeCompleta = IIf(Anos <= 1, IIf(Anos = 0, "", Anos & " ano "), Anos & " anos ") & _
                              IIf(Meses <= 1, IIf(Meses = 0, "", Meses & " mes "), Meses & " meses ") & _
                              IIf(Dias <= 1, IIf(Dias = 0, "", Dias & " dia "), Dias & " dias ")
End Function
Public Function DiferencaHoras(Hora1 As Date, Hora2 As Date) As String
On Error Resume Next
    Dim vHoras As String
    Dim vMin As String
    
    If Hora2 < Hora1 Then
        DiferencaHoras = "xx"
    Else
        DiferencaHoras = Format(CDate(TimeValue(Hora2) - TimeValue(Hora1)), "hh:mm")
        
        'vHoras = Format(CDbl(DateDiff("h", Hora1, Hora2)), "00")
        'vMin = Format(CDbl(DateDiff("m", Hora1, Hora2)), "00")
    End If
    
End Function
Private Sub txtHora2_GotFocus()
txtHora2.SelStart = 0
txtHora2.SelLength = Len(txtHora2.Text)
End Sub
Private Sub txtHora1_GotFocus()
txtHora1.SelStart = 0
txtHora1.SelLength = Len(txtHora1.Text)
End Sub

Private Sub txtValor_Change()

End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{Tab}")
If KeyAscii < 44 And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 27 And KeyAscii <> 13 Or KeyAscii > 57 Then
    KeyAscii = 0
    Exit Sub
End If

If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub txtPercentual_GotFocus()
txtPercentual.SelStart = 0
txtPercentual.SelLength = Len(txtPercentual.Text)
End Sub

Private Sub txtPercentual_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{Tab}")
If KeyAscii < 44 And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 27 And KeyAscii <> 13 Or KeyAscii > 57 Then
    KeyAscii = 0
    Exit Sub
End If

If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub txtPercentual_LostFocus()
txtPercentual.Text = Format(txtPercentual.Text, "#,##0.00")
End Sub

Private Sub txtValor1_GotFocus()
txtValor1.SelStart = 0
txtValor1.SelLength = Len(txtValor1.Text)
End Sub

Private Sub txtValor1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{Tab}")
If KeyAscii < 44 And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 27 And KeyAscii <> 13 Or KeyAscii > 57 Then
    KeyAscii = 0
    Exit Sub
End If

If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub txtValor1_LostFocus()
txtValor1.Text = Format(txtValor1.Text, "#,##0.00")
End Sub
