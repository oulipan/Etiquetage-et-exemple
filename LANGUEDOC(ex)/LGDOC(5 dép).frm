VERSION 5.00
Object = "{05589FA0-C356-11CE-BF01-00AA0055595A}#2.0#0"; "AMOVIE.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15285
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   6.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   15285
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1815
      Left            =   13320
      TabIndex        =   9
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   3201
      _Version        =   393217
      BackColor       =   8454143
      ScrollBars      =   3
      TextRTF         =   $"LGDOC(5 dép).frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   12960
      TabIndex        =   8
      Top             =   5640
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin AMovieCtl.ActiveMovie ActiveMovie1 
      Height          =   420
      Left            =   12960
      TabIndex        =   7
      Top             =   4080
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   741
      ShowDisplay     =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13200
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   5760
      TabIndex        =   6
      Top             =   960
      Width           =   3855
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "xx"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   870
         TabIndex        =   5
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.CommandButton Aide 
      Caption         =   "Aide"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   10320
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   10800
      Top             =   1440
   End
   Begin VB.TextBox aremplir 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   0
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Zone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   7680
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   6720
      Visible         =   0   'False
      Width           =   11895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************** ex LANGUEDOC-ROUSSILLON ******************
'FICHIERS à joindre: "lgdoc.jpg" et "lgdoc.txt"
' ########## de multiples cartes/dessins/exos à trous peuvent être créés sur le même modèle:
' ------------- voir " " pour pointage des coordonnées des "étiquettes-trous" qui masqueront les mots à demander ------------
Dim ju As Integer
Dim j As Integer
Dim nchoix As Integer, etqs As Integer, erreur As Integer
Dim score As Integer, tfin As Integer, cfx As Integer, nb As Integer, osi
Dim Aid0 As String ' voir dans f.txt joint
Dim fic1 As String, fic2 As String, fic As String ' f.image ; f.texte ; "curdir"
Dim nt, ctr(30), npays(30)
Dim repe As String
Dim pays(30) As String
Dim cor(30) As String
Dim choix(20)
Dim af As Integer
Dim asie As Integer '(en cas d'une suite de multiples f.jpg et f.txt joints)
Dim total As Integer
Dim tps As Integer 'pour le timer
Dim g(30) As Integer, ht(30) As Integer
Dim fig3 As Integer
' ================================================================================================================================
'     ******************** pour mémoire et autres progr (un bouton "commande" et "boîte de dialogue pour choix d'image,f.rtf/.txt)**************************************
'#& Private Sub Cmdchxtxt_Click() ' =========================================================
'#& Frame3.Visible = False
'#& Cmdchxtxt.Visible = False
 '#&  Rtb1.Visible = True: 'Tsaisi1(0).Visible = True
  '#& Cmdchxtxt.Visible = False
  '#& On Error GoTo fin
  '#& ChDir CurDir
   '#& CommonDialog1.Filter = "Fichiers au format RTF(*.rtp)|*.rtp|"
 '#&  CommonDialog1.FilterIndex = 1
  '#& CommonDialog1.ShowOpen
   '#& Rtb1.LoadFile CommonDialog1.FileName, rtfRTF
   '#& fic = Left(CommonDialog1.FileName, Len(CommonDialog1.FileName) - 3) & "txt"
 '#& charge
'#& Exit Sub
'#& 'fichier chargé par défaut
'#& fin: Rtb1.LoadFile CurDir & "\valeurpresent.rtp", rtfRTF:
'#& Close #1: fic = CurDir & "\valeurpresent.txt": charge
'#& End Sub '===============================================================================
' ================================================================================================================================
Private Sub Aide_Click()
If af = 1 Then a = MsgBox("Clique sur le nom du département" + Chr(10) + "où clignote le curseur" + Chr(10) + Chr(10) + "Attention, certains de ces départements" + Chr(10) + "sont en Midi-Pyrénées" + Chr(10) + "ou en Provence-Côte d'Azur", , "5 départements à choisir:") Else a = MsgBox("Clique sur le nom de la ville (point rouge) " + Chr(10) + "dont l'étiquette porte un curseur clignotant.", , "Aid0:")
aremplir(nb).SetFocus
End Sub
Private Sub Label4_Click(Index As Integer)
If Index > nchoix Then Exit Sub
aremplir(nb).FontStrikethru = False: aremplir(nb).Text = Label4(Index).Caption: repe = Label4(Index).Caption
verif (nb)
End Sub
Private Sub Form_Load()
For xx = 0 To (Minute(Time) + Second(Time)): hhzz = Int(Rnd * 3): Next xx
fic = CurDir + "\": tfin = 0
Label2(0).Visible = False
RichTextBox1.Visible = False
ActiveMovie1.Visible = False: MMControl1.Visible = False '(en prévision)
Form1.WindowState = 0: Form1.Top = 100: Form1.Left = 1000: Form1.Height = 10000: Form1.Width = 12000
Form1.Caption = "   **** (ex)LANGUEDOC-ROUSSILLON : 5 départements du sud du LANGUEDOC/ 66 - 11 - 34 - 30 et (plus au Nord:) Lozère(48)..."
'If hhzz = 1 Then suite2: Exit Sub
'If hhzz = 2 Then suite: Exit Sub
fic1 = fic + "lgdoc.jpg"
fic2 = fic + "lgdoc.txt"
charge
Frame1.Caption = "Grandes villes de (l'ex) Languedoc-Roussillon"
Label1.Visible = False
Label1.Left = 4000: Label1.Top = 4800
etqs = nchoix + 1: total = nt + 1: j = 0: Load Label4(etqs): Label4(etqs) = "Attention, deux de ces villes" + Chr(10) + "ne sont pas dans le " _
+ Chr(10) + "LANGUEDOC-ROUSSILLON": Label4(etqs).Top = Label4(nchoix).Top + 500: Label4(etqs).Left = Label4(0).Left: Label4(etqs).FontSize _
= 8: Label4(etqs).FontItalic = True: Label4(etqs).Visible = True: Label4(etqs).BackColor = Frame1.BackColor
End Sub


Private Sub aremplir_GotFocus(Index As Integer)
nb = Index
End Sub

Private Sub aremplir_KeyPress(Index As Integer, KeyAscii As Integer)
If ctr(Index) = 1 Then KeyAscii = 0: aremplir(Index).Text = pays(Index): Exit Sub
aremplir(Index).FontStrikethru = False
'If fig3 = 0 Then Label1.Visible = False '# (pour autre progr)
'If cfx = 0 Then Label1.Visible = False:
nb = Index
car = Chr$(KeyAscii)
If InStr(" -éèàîôâêûï'" + Chr$(8) + Chr$(13), car) <> 0 Or (car >= "A" And car <= "z") Or (car >= "0" And car <= "9") Then
If car = Chr$(13) Then KeyAscii = 0: repe = Trim(aremplir(Index).Text): verif (Index): Exit Sub
KeyAscii = Asc(car)
Else
KeyAscii = 0
End If
End Sub
Private Sub charge()
Label1.Visible = False
Form1.Picture = LoadPicture(fic1)
Open fic2 For Input As #1
Input #1, tfre, lfre, nchoix: nchoix = nchoix - 1: If nchoix > -1 Then For xx = 0 To nchoix: Input #1, choix(xx): Next xx
Input #1, Aid0, nt
Do Until n = nt
Input #1, g(n), ht(n), pays(n): n = n + 1
Loop
Close #1: n = 0: nt = nt - 1
 If nchoix > -1 Then
For xx = 0 To nchoix: If xx > 0 Then Load Label4(xx):  Label4(xx).Top = Label4(xx - 1).Top + Label4(xx - 1).Height + 100
' Label4(xx).Left = Label4(xx - 1).Left + Label4(xx - 1).Width + 800 Else If xx > 2 Then Label4(xx).Left = Label4(xx - 2).Left + Label4(xx - 2).Width _
+ 200: Label4(xx).Top = Label4(xx - 2).Top Else If xx = 2 Then Label4(xx).Top = Label4(xx - 1).Top + Label4(xx - 1).Height + 100: Label4(xx).Left _
= Label4(xx - 1).Left: ' Label4(xx).Top = Label4(xx - 2).Top Else Label4(xx).Left = Label4(4).Left: Label4(xx).Top = Label4(7).Top + Label4(7).Height + 100
Label4(xx) = choix(xx): Label4(xx).Visible = True: Next xx: Frame1.Left = lfre: Frame1.Top = tfre: Frame1.Visible = True
Aide.Left = Frame1.Left: Aide.Top = Frame1.Top + Frame1.Height + 200
End If
 For xx = 0 To nt: If xx > 0 Then Load aremplir(xx)
aremplir(xx).Enabled = True: ctr(xx) = 0: aremplir(xx).Text = "": aremplir(xx).Visible = False: Next xx
For xx = 0 To nt: aremplir(xx).Left = g(xx): aremplir(xx).Top = ht(xx): aremplir(xx).Height = 150: aremplir(xx).BackColor = vbYellow: _
aremplir(xx).Width = 1300
'  aremplir(xx).Width = 150 * (Len(pays(xx)) + Int(Rnd * 3)): If xx = 1 Or xx = 2 Then aremplir(xx).BackColor = vbYellow Else If xx > 8 Then _
aremplir(xx).BackColor = vbCyan Else aremplir(xx).BackColor = &HC0FFC0: 'If xx = 2 Then aremplir(xx).BackColor = &HC0E0FF Else If xx = 3 Then _
aremplir(xx).BackColor = &HC0FFC0 Else aremplir(xx).BackColor = vbYellow
If Len(pays(xx)) > 12 Then aremplir(xx).Width = 2500: aremplir(xx).Height = 500 - 200 * (Len(pays(xx)) > 20)
If aremplir(xx).Width < 300 Then aremplir(xx).Width = 300
Next xx
nb = 0: cfx = 1
For xx = 0 To nt: aremplir(xx).Visible = True: Next xx:
For xx = 1 To Len(Aid0) - 5: If Mid(Aid0, xx, 4) = "    " Then ai2 = ai2 + Chr(10) + Chr(10): xx = xx + 3 Else ai2 = ai2 + Mid(Aid0, xx, 1)
Next xx: ai2 = ai2 + Mid(Aid0, xx, 4): Aid0 = ai2
For yy = 0 To nt: For xx = 1 To Len(cor(yy)) - 5: If Mid(cor(yy), xx, 4) = "    " Then ai2 = ai2 + Chr(10) + Chr(10): xx = xx + 3 Else ai2 = ai2 + Mid(cor(yy), xx, 1)
Next xx: ai2 = ai2 + Mid(cor(yy), xx, 4): cor(yy) = ai2 + Right(cor(yy), 1): ai2 = "": Next yy
'If af = 1 Then aremplir(0).SetFocus
End Sub
Private Sub clair0()
  j = ju - erreur: If j < 0.25 Then j = 0.25
score = j: enreg
MsgBox ("Tu es arrivé au bout!" + Chr(10) + "avec " + Str(score) + " points."):
End
End Sub
Private Sub claira()
 If nchoix > -1 Then
 For xx = 0 To etqs: If xx > 0 Then Unload Label4(xx)
Next xx
End If
ju = ju + j
For xx = 0 To nt: aremplir(xx).Enabled = True: If xx > 0 Then Unload aremplir(xx)
 ctr(xx) = 0: Next xx
End Sub
Private Sub verif(Index)
acc = 0: verep = pays(Index): For xx = 1 To Len(pays(Index)): If InStr("éèê", Mid$(pays(Index), xx, 1)) <> 0 Then Mid$(verep, xx, 1) = "e": acc = 1
If InStr("àâ", Mid$(pays(Index), xx, 1)) <> 0 Then Mid$(verep, xx, 1) = "a": acc = 1
If InStr("î", Mid$(pays(Index), xx, 1)) <> 0 Then Mid$(verep, xx, 1) = "i": acc = 1
If InStr("ï", Mid$(pays(Index), xx, 1)) <> 0 Then Mid$(verep, xx, 1) = "i": acc = 1
If InStr("ô", Mid$(pays(Index), xx, 1)) <> 0 Then Mid$(verep, xx, 1) = "o": acc = 1
If InStr("ùû", Mid$(pays(Index), xx, 1)) <> 0 Then Mid$(verep, xx, 1) = "u": acc = 1
Next xx
verep2 = repe: For xx = 1 To Len(repe): If InStr("éèê", Mid$(repe, xx, 1)) <> 0 Then Mid$(verep2, xx, 1) = "e": acc = acc + 2
If InStr("àâ", Mid$(repe, xx, 1)) <> 0 Then Mid$(verep2, xx, 1) = "a": acc = acc + 2
If InStr("î", Mid$(repe, xx, 1)) <> 0 Then Mid$(verep2, xx, 1) = "i": acc = acc + 2
If InStr("ï", Mid$(repe, xx, 1)) <> 0 Then Mid$(verep2, xx, 1) = "i": acc = acc + 2
If InStr("ô", Mid$(repe, xx, 1)) <> 0 Then Mid$(verep2, xx, 1) = "o": acc = acc + 2
If InStr("ùû", Mid$(repe, xx, 1)) <> 0 Then Mid$(verep2, xx, 1) = "u": acc = acc + 2
Next xx
'If Left$(repe, 1) = LCase$(Left$(pays(index), 1)) Then MsgBox ("Tu aurais pu mettre une majuscule!")
If repe <> pays(Index) And UCase$(verep2) = UCase$(verep) And acc = 1 Then MsgBox ("Pense aux accents!")
'If repe <> pays(index) And UCase$(verep2) = UCase$(verep) And acc > 1 Then MsgBox ("Tu as mis" + Str$(acc / 2) + " accent(s) en trop!")
If repe = pays(Index) Then MsgBox ("bravo")
If repe = pays(Index) Or UCase$(verep2) = UCase$(verep) Then
   aremplir(Index).Text = pays(Index): aremplir(Index).Enabled = False: Label1.Visible = False: cfx = 0: ctr(Index) = 1: j = j + 1:  If j >= total Then MsgBox ("Tu as tout trouvé!"): suite Else nwmot
  Else
  aremplir(Index).FontStrikethru = True
  erreur = erreur + 1
  cmt = cor(Index): cfx = 1: 'If Left$(UCase$(verep2), 3) = Left$(UCase$(verep), 3) Then cmt = cmt + "Le début est bien écrit."
  'If Left$(UCase$(verep2), 5) = Left$(UCase$(verep), 5) Then cmt = cmt + "Tu as commis une erreur d'orthographe au milieu ou à la fin du mot."
  tps = 0: a = MsgBox(cmt, , "Erreur"): Label1.Caption = cmt: ' Label1.Visible = True: 'Timer1.Enabled = True
  End If
End Sub
Private Sub nwmot()
nf = 0: For xx = 0 To total - 1: If ctr(xx) = 0 Then nf = xx: xx = total
Next xx: aremplir(nf).SetFocus
End Sub
Private Sub suite() ' ********** charger un second exo **************
claira
'If af = 2 Then suite3: Exit Sub ' à modif selon nb exos

If af = 1 Then clair0: Exit Sub ' à modif selon nb exos
fic1 = fic + "lgdoc.jpg"
fic2 = fic + "lgdocd.txt"
charge
etqs = nbchoix
Frame1.Height = 4800
Frame1.Caption = "DEPARTEMENTS du (ex)Languedoc-Roussillon" + Chr(10) + "(4/13 départements du sud du LANGUEDOC + Lozère"
af = 1: aremplir(0).SetFocus
etqs = nchoix + 1: total = nt + 1: j = 0: Load Label4(etqs): Label4(etqs) = "Attention, quatre de ces départements" + Chr(10) + " sont dans des " + Chr(10) + "régions limitrophes": Label4(etqs).Top = Label4(nchoix).Top + 500: Label4(etqs).Left = Label4(0).Left - 300: Label4(etqs).FontSize = 8: Label4(etqs).FontItalic = True: Label4(etqs).Visible = True: Label4(etqs).BackColor = Frame1.BackColor
End Sub
Private Sub suite3() ' ************ ex chargement 4e exo (à modif) ou arrêt
End
End Sub
Private Sub suite2() ' ************ ex chargement 3e exo *************
fic1 = CurDir + "\imp3.jpg"
fic2 = CurDir + "\imp2.txt"
charge
asie = 1: total = nt + 1: j = 0
End Sub

Private Sub Timer1_Timer()
If tps > 5 Then Timer1.Enabled = False: Timer2.Enabled = False: Label1.Caption = "": Label1.Visible = False: Exit Sub
Label1.ForeColor = QBColor(15): Timer1.Enabled = False: Timer2.Enabled = True: tps = tps + 1
End Sub

Private Sub Timer2_Timer()
Label1.ForeColor = QBColor(12): Timer2.Enabled = False: Timer1.Enabled = True
End Sub
'# pour autre progr ##############
'Private Sub sono()
'If dr = 2 Then aremplir(ns).SetFocus: MMControl1.Command = "Close": MMControl1.FileName = fic + son(ns) + ".wav": MMControl1.Command = "Open": MMControl1.Notify = False: MMControl1.Wait = False: MMControl1.Command = "Play"
'End Sub
Private Sub clair()
If Form1.WindowState = 2 Or dd1 = 0 Then Exit Sub
arrsto
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
arrsto
End Sub

Private Sub enreg()
'*** sauvegarde résultat ***
Open CurDir + "\score_LGDOC.txt" For Output As #1
Print #1, Date, " ==> ", Time
Print #1, Trim(Str$(score))
Close #1
Clipboard.Clear
Clipboard.SetText (Str(score) + "$")
End Sub

Private Sub arrsto()
If tfin = 1 Then Exit Sub Else j = ju - erreur: If j < 0 Then j = 0
score = j: enreg
MsgBox ("Au revoir" + Chr(10) + Str(score) + " point(s)")
End Sub


