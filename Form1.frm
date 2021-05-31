VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "快讯排版生成器v0.2"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8850
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   8850
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "Form1.frx":048A
      Left            =   3120
      List            =   "Form1.frx":0494
      TabIndex        =   25
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   240
      TabIndex        =   24
      Top             =   6720
      Width           =   2775
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   240
      TabIndex        =   23
      Top             =   6120
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "选择颜色模式"
      Height          =   615
      Left            =   4560
      TabIndex        =   20
      Top             =   6360
      Width           =   4095
      Begin VB.OptionButton Option2 
         Caption         =   "黑夜模式"
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "白天模式"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   2400
      TabIndex        =   17
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   240
      TabIndex        =   14
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空输出框"
      Height          =   615
      Left            =   7200
      TabIndex        =   11
      Top             =   7200
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "生成卡狗代码并复制"
      Height          =   615
      Left            =   5040
      TabIndex        =   7
      Top             =   7200
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   1815
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3960
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Height          =   1575
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   5415
      Left            =   4560
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form1.frx":04A4
      Left            =   240
      List            =   "Form1.frx":051A
      TabIndex        =   0
      Text            =   "选择快讯来源"
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label13 
      Caption         =   "made by xiang_xge"
      Height          =   375
      Left            =   13920
      TabIndex        =   28
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "插入视频/图片"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "插入蓝色链接"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "v0.2"
      Height          =   255
      Left            =   7920
      TabIndex        =   19
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "卡狗式快讯排版BBCode生成器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label8 
      Caption         =   "脚注"
      Height          =   255
      Left            =   2400
      TabIndex        =   16
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "原文地址"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "翻译语种（不填默认为英语）"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   7200
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "译者名输入"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "快讯来源"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "译文输入"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "原文输入"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "BBCode代码输出"
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim image As String 'image为头像地址
Dim name As String 'name为用户昵称
Dim translator As String 'translator为译者名
Dim lang As String 'lang为语言名
Dim id As String 'id为推特ID

If Text5.Text = "" Then
lang = "英语"
Else
lang = Text5.Text
End If

translator = Text4.Text

Select Case Combo1.Text
Case "@Minecraft"
image = "https://z3.ax1x.com/2021/05/23/gO2zuR.jpg"
name = "Minecraft"
id = "@Minecraft"

Case "@Minecraft Earth"
image = "https://z3.ax1x.com/2021/05/23/gO2vv9.png"
name = "Minecraft Earth"
id = "@minecraftearth"

Case "@Minecraft Dungeons"
image = "https://z3.ax1x.com/2021/05/23/gO2jgJ.png"
name = "Minecraft Dungeons"
id = "@dungeonsgame"

Case "@Mojang"
image = "https://z3.ax1x.com/2021/05/23/gO2X34.jpg"
name = "Mojang"
id = "@Mojang"

Case "@Mojang Support"
image = "https://z3.ax1x.com/2021/05/23/gO2X34.jpg"
name = "Mojang Support"
id = "@MojangSupport"

Case "@Mojang Status"
image = "https://z3.ax1x.com/2021/05/23/gO2X34.jpg"
name = "Mojang Status"
id = "@MojangStatus"

Case "@Minecraft Education"
image = "https://z3.ax1x.com/2021/05/23/gO2bNT.png"
name = "Minecraft: Education Edition"
id = "@PlayCraftLearn"

Case "@Jeb"
image = "https://z3.ax1x.com/2021/05/23/gO2HEV.jpg"
name = "Jens Bergensten"
id = "@jeb_"

Case "@Dinnerbone"
image = "https://i.loli.net/2021/05/24/LAaHdNIGpcKskvP.jpg"
name = "Nathan Adams"
id = "@Dinnerbone"

Case "@JAPPA"
image = "https://z3.ax1x.com/2021/05/23/gO2oBq.jpg"
id = "@JasperBoerstra"
name = "JAPPA"

Case "@slicedlime"
id = "@slicedlime"
image = "https://z3.ax1x.com/2021/05/23/gO2Iun.jpg"
name = "slicedlime"

Case "@Adrian Ostergard"
image = "https://z3.ax1x.com/2021/05/23/gO24js.jpg"
name = "Adrian Ostergard"
id = "@adrian_ivl"

Case "@Maria Lemon"
image = "https://z3.ax1x.com/2021/05/23/gO2f3Q.jpg"
name = "Maria Lemón"
id = "@MiaLem_n"

Case "@LadyAgnes"
image = "https://z3.ax1x.com/2021/05/24/gjG7t0.png"
name = "LadyAgnes"
id = "@_LadyAgnes"

Case "@Cory Scheviak"
image = "https://z3.ax1x.com/2021/05/24/gjGTkq.png"
name = "Cory Scheviak"
id = "@Cojomax99"

Case "@tomcc"
image = "https://z3.ax1x.com/2021/05/24/gjGfXQ.jpg"
name = "Tommaso Checchi"
id = "@_tomcc"

Case "@Jeison"
image = "https://z3.ax1x.com/2021/05/24/gjG50s.jpg"
name = "Jeison S"
id = "@TamerJeison"

Case "@Jason Major"
image = "https://z3.ax1x.com/2021/05/24/gjG4mj.jpg"
name = "Jason Major"
id = "@argo_major"

Case "@Joshua D Bullard"
image = "https://z3.ax1x.com/2021/05/24/gjGI7n.png"
name = "Joshua D Bullard"
id = "@Jdbullard"

Case "@David Fries"
image = "https://z3.ax1x.com/2021/05/24/gjGW6g.jpg"
name = "David Fries"
id = "@JDavidFries"

Case "@Tanner Pearson"
image = "https://z3.ax1x.com/2021/05/24/gjGR1S.jpg"
name = "Tanner Pearson"
id = "@The_T_Pearson"

Case "@Helen Chiang"
image = "https://z3.ax1x.com/2021/05/24/gjG2p8.jpg"
name = "Helen Chiang"
id = "@Pr1ncessP1ggy"

Case "@Lydia Winters"
image = "https://z3.ax1x.com/2021/05/24/gjGcff.jpg"
name = "Lydia Winters"
id = "@LydiaWinters"

Case "@Jay"
image = "https://z3.ax1x.com/2021/05/24/gjG6tP.jpg"
name = "Jay Wells"
id = "@Mega_Spud"

Case "@Matt"
image = "https://z3.ax1x.com/2021/05/24/gjGykt.jpg"
name = "Matt Gartzke"
id = "@MattGartzke"

Case "@Marc Watson"
image = "https://z3.ax1x.com/2021/05/24/gjGrTI.jpg"
name = "Marc Watson"
id = "@Marc_IRL"

Case "@Saxs"
image = "https://z3.ax1x.com/2021/05/24/gjGwOH.jpg"
name = "Saxs"
id = "@Saxs"

Case "@Vu Bui"
image = "https://z3.ax1x.com/2021/05/24/gjGBmd.jpg"
name = "Vu Bui"
id = "@vubui"

Case "@John"
image = "https://z3.ax1x.com/2021/05/24/gjGD0A.jpg"
name = "John Hendricks"
id = "@JLtZD"

Case "@David"
image = "https://z3.ax1x.com/2021/05/24/gjGd6e.jpg"
name = "David"
id = "@CornerHardMC"

Case "@Josh Mulanax"
image = "https://z3.ax1x.com/2021/05/24/gjGJFx.jpg"
name = "Josh Mulanax"
id = "@JORAX79"

Case "@Pradnesh Patil"
image = "https://z3.ax1x.com/2021/05/24/gjGUSO.jpg"
name = "Pradnesh Patil"
id = "@pradneshpatil"

Case "@David Nisshagen"
image = "https://z3.ax1x.com/2021/05/24/gjGYY6.jpg"
name = "David Nisshagen"
id = "@DavidNisshagen"

Case "@Daniel Bjorkefors"
image = "https://z3.ax1x.com/2021/05/24/gjGtfK.jpg"
name = "Daniel Bjorkefors"
id = "@bjorkefors"

Case "@Keso"
image = "https://z3.ax1x.com/2021/05/24/gjGalD.jpg"
name = "Keso"
id = "@MaxHerngren"

Case "@Henrik"
image = "https://attachment.mcbbs.net/data/myattachment/forum/202105/17/145006gcwbcuzwmue0wgwk.png"
name = "Henrik Kniberg"
id = "@henrikkniberg"

End Select


Dim media As String
Dim link As String

If Text8.Text = "" Then
link = ""
Else
link = "[color=#1B95E0]" + Text8.Text + "[/color]"
End If

If Text9.Text = "" Then
media = ""
Else
Select Case Combo2.Text
Case "图片"
media = "[img]" + Text9.Text + "[/img]"
Case "视频"
media = "media=x,500,375]" + Text9.Text + "[/media]"
End Select
End If

If Option1.Value = True Then

Text1.Text = "[align=center][table=560,#FFFFFF]" + vbCrLf + _
"[tr][td][font=-apple-system, BlinkMacSystemFont,Segoe UI, Roboto, Helvetica, Arial, sans-serif][indent]" + vbCrLf + _
"[float=left][img=44,44]" + image + "[/img][/float][size=15px][b][color=#0F1419]" + name + "[/color][/b]" + _
vbCrLf + "[color=#5B7083]" + id + "[/color][/size]" + vbCrLf + vbCrLf _
+ "[color=#0F1419][size=23px]" + Text2.Text + "[/size]" + vbCrLf _
+ "[size=15px]由 " + translator + " 翻译自" + lang + "[/size]" + vbCrLf _
+ "[size=23px]" + Text3.Text + vbCrLf + link + vbCrLf + media + "[/size][/color][/indent][indent][size=15px][url=" + Text6.Text + "][color=#5B7083]" + Text7.Text + "[/color][/url][/size][/indent][/font]" + vbCrLf _
+ "[/td][/tr]" + vbCrLf _
+ "[/table][/align]"

Else

If Option2.Value = True Then

Text1.Text = "[align=center][table=560,#000000]" + vbCrLf + _
"[tr][td][font=-apple-system, BlinkMacSystemFont,Segoe UI, Roboto, Helvetica, Arial, sans-serif][indent]" + vbCrLf + _
"[float=left][img=44,44]" + image + "[/img][/float][size=15px][b][color=#D9D9D9]" + name + "[/color][/b]" + _
vbCrLf + "[color=#5B7083]" + id + "[/color][/size]" + vbCrLf + vbCrLf _
+ "[color=#D9D9D9][size=23px]" + Text2.Text + "[/size]" + vbCrLf _
+ "[color=#5B7083][size=15px]由 " + translator + " 翻译自" + lang + "[/size][/color]" + vbCrLf _
+ "[size=23px]" + Text3.Text + vbCrLf + link + vbCrLf + media + "[/size][/color][/indent][indent][size=15px][url=" + Text6.Text + "][color=#5B7083]" + Text7.Text + "[/color][/url][/size][/indent][/font]" + vbCrLf _
+ "[/td][/tr]" + vbCrLf _
+ "[/table][/align]"

Else

Text1.Text = "嗯......这是个彩蛋，但你不该发现它的，你这个肮脏的黑客"

End If
End If


Clipboard.Clear
Clipboard.SetText Text1.Text


End Sub

Private Sub Command2_Click()
Text1.Text = ""

End Sub

