VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15600
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   15600
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10800
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "chx IMAGE à charg"
      Height          =   495
      Left            =   9360
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   10920
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "pointeur2024.frx":0000
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1455
      Left            =   7080
      TabIndex        =   1
      Top             =   6480
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a1(100), b1(100), a(100), b(100), c(100), n As Integer, n1 As Integer, nl As Integer, cl(100) As String, fic As String, lfic As String, nf As Integer, tit(100) As String, fi(100) As String, pol As Integer
'pointeur plus succint repris pr moi- mars 2010
Private Sub Command1_Click()
    Command1.Visible = False
     CommonDialog1.Filter = "Fichiers au format JPG(*.jpg)|*.jpg|"
  ' Définit le filtre par défaut
  CommonDialog1.FilterIndex = 1
  CommonDialog1.ShowOpen
   Form1.Picture = LoadPicture(CommonDialog1.FileName, jpgJPG)
      fic = Left(CommonDialog1.FileName, Len(CommonDialog1.FileName) - 3) + "txt"
 On Error Resume Next
   pol = -1: Open fic For Input As #1: For xx = 1 To 5: Input #1, zz: Next xx: Input #1, nl: For xx = 0 To nl - 1: Input #1, cl(xx): Next xx: Input #1, n: For xx = 0 To n - 1: Input #1, a1(xx), b1(xx), c(xx): a(xx) = Trim(Str(a1(xx))): b(xx) = Trim(Str(b1(xx))): Next xx
   Close #1
     titre = InputBox("Nom qui sera affiché dans la liste du choix des exercices" + Chr(10) + Chr(13) + Chr(10) + Chr(13) + "PAR EXEMPLE:" + Chr(10) + Chr(13) + "son [en] (en,an,em,am ...)", "titre/sujet de l'exercice qui sera crée?")
   lfic = InputBox("Nom complet du chemin/fichier où sera listé le nouveau fichier créé?" + Chr(10) + Chr(13) + Chr(10) + Chr(13) + "PAR DEFAUT:" + Chr(10) + Chr(13) + "[C:\vbtxt\français\lecture CP\lficlec.txt] pour l'exercice ''lecture à trous'' de CP", "fichier liste?", Left(CurDir, 1) + ":\vbtxt\français\lecture CP\lficlec.txt")
  
   pol = 1: Open lfic For Input As #1: Input #1, nf: For xx = 1 To nf: Input #1, tit(xx), fi(xx): tit(xx) = Chr(34) + tit(xx) + Chr(34): Next xx
   Close #1
   nf = nf + 1: tit(nf) = Chr(34) + titre + Chr(34): fi(nf) = fic
   If InStr(fic, ".txt") > 0 Then fi(nf) = Left(fic, Len(fic) - 4)
      Open lfic For Output As #1
   Print #1, Trim(Str(nf))
   For xx = 1 To nf: Print #1, tit(xx) + "," + fi(xx): Next xx
   Print #1, ""
   Print #1, ""
      Print #1, "================================================================================"
    Print #1, "nb fic choix /,/titre1,chemin1/..."
   Close #1
   Text1.Visible = True: Label1.Visible = True: pol = 0
   Text1.SetFocus

'erreur:
'If pol = -1 Then n = 0: Resume Next
'If pol = 1 Then Resume Next

End Sub

'penser à taper d'abord le mot avant de cliquer sur l'endroit où le placer
Private Sub Form_Load()
Form1.Caption = "Saisie de mots à placer sur étiquettes.              Une fois tous les mots saisis, pour terminer, cliquez sur la croix (blanche sur fond rouge) ---->"
Form1.Show
Form1.CurrentX = 9000: Form1.CurrentY = 1000: Form1.ForeColor = vbRed: Form1.Print "Cliquez sur le bouton de commande ci-dessus"; Tab(140); " pour charger une image JPG"
Text1.Text = "": Label1 = "Taper le mot avant" + Chr(10) + "de cliquer sur le haut gauche de l'étiquette" + Chr(10) + "qui cachera l'élément que l'élève devra écrire ." + Chr(10) + Chr(10) + "SIGNES INTERDITS:[,][;][?][:]"
'\math\Left(CurDir, 1) + ":\vbgb\action1.jpg"  c:\vbtxt\lecag30a.jpg
'c:\vbgb\a6u3l3\
nf = 0: nl = 0: n = 0: n1 = 0: pol = -2
'c:\vbtxt\geo\vbtxt\geo\axbdi0
'c:\vbgb\there.jpg
'Form1.Picture = LoadPicture(Left(CurDir, 1) + ":\vbtxt\français\lecture CP\p38a.jpg ")
Text1.Visible = False: Label1.Visible = False
'aa = MsgBox("Si vous voulez compléter un fichier de données déjà commencé," + Chr(13) + Chr(10) + "fermez provisoirement ce programme " + Chr(13) + Chr(10) + "pour sauvegarder le fichier-texte déjà réalisé sous un autre nom provisoire" + Chr(13) + Chr(10) + "(sinon, il sera écrasé par celui que vous allez créer)" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "Si vous voulez d'abord sauvegarder sous un autre nom le fichier déjà commencé" + Chr(13) + Chr(10) + "cliquez [NON]" + Chr(13) + Chr(10), 36, "ATTENTION, voulez-vous continuer?")
'If aa = 7 Then aa = MsgBox("Quand vous aurez fini de créer le fichier complémentaire"): lfic = "non": End
End Sub
Private Sub form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Label1.Visible = True: Label1.Caption = "x=" + Str$(X) + " , " + "y=" + Str$(Y)
If Text1.Text = "" Then aa = MsgBox("Tapez le mot,SVP" + Chr(10) + Chr(13) + "avant de cliquer sur le haut gauche" + Chr(10) + Chr(13) + "de l'étiquette qui lui correspondra.", , "Le mot d'abord, merci."): Exit Sub
'Open Left(CurDir, 1) + ":\vbtxt\aaarep.txt" For Append As #1
a(n) = Trim(Str$(X)): b(n) = Trim(Str$(Y)): c(n) = Text1.Text
Open fic For Append As #1
Print #1, Trim(Str$(X)) + "," + Trim(Str$(Y)) + "," + Chr(34) + Text1.Text + Chr(34)
Close #1
' **************** par sécurité ***************************
Open CurDir + "\aaarep.txt" For Append As #1
Print #1, Trim(Str$(X)) + "," + Trim(Str$(Y)) + "," + Chr(34) + Text1.Text + Chr(34)
Close #1
Text1 = ""
n = n + 1: n1 = n1 + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
If pol = -2 Then End: Exit Sub
'tp = InputBox("Taille de police: 10....14?", "Taille police?", "14")
'he = InputBox("Hauteur d'étiquette:400...600 ( 500 conseillé pour police 14)?", "Hauteur étiquette?", "500")
'lc = InputBox("Largeur caractère:200...300 (230 à 250 pr 14)?", "Largeur caractère?", "250")
liste = Trim(Str$(nl + n1))
If nl > 0 Then For xx = 0 To nl - 1: liste = liste + "," + cl(xx): Next xx
For xx = nl To nl + n1 - 1: liste = liste + "," + c(xx): Next xx
'ht = InputBox("coord ht coin ht G d'affichage de l'aide (100 conseillé si en ht / environ 4000 pour mi-hteur) ?", "Coin haut de l'aide?", "100")
'g = InputBox("coord  G d'affichage de l'aide (7000~= milieu pour affichage 1152x864) ?", "coin gauche de l'aide?", "7000")
'Open Left(CurDir, 1) + ":\vbtxt\aaarep.txt" For Append As #1
Open fic For Output As #1 ' le fichier créé précédemment sera effacé par  [output]
'Print #1, "Les pointages ci-dessus doivent être effacés: ils n'ont été enregistrés par précaution en cas de coupure intempestive en cours de saisie."
Print #1, "14,500,250,100,7000"
Print #1, liste
Print #1, Trim(Str$(n))
For xx = 0 To n - 1: Print #1, a(xx) + "," + b(xx) + "," + Chr(34) + c(xx) + Chr(34): Next xx
Print #1, ""
Print #1, "==================================================================================="
Print #1, "police,htr étiquette,largeur car,coord ht d'affich d'aide(X),coord G d'affich d'aide(Y)"
Print #1, "nb étiqu aide,[mot1,mot2,...]"
Print #1, "nb d'étiqu"
Print #1, "coor GCHE de mot1,coord HT de mot1,mot1/etc.."
Print #1, fic
Print #1, "==================================================================================="
Close #1
aa = MsgBox("Le fichier texte que vous venez de réaliser" + Chr(10) + Chr(13) + "se trouve dans le même dossier que l'image, avec le même nom" + Chr(10) + Chr(13) + Chr(10) + Chr(13) + "Ouvrez-le pour vérifier ou modifier:" + Chr(10) + Chr(13) + fic, , fic)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
car = Chr$(KeyAscii)
If InStr(" -éèàîôâêûï'+/()€.,<>?:!ë" + Chr(34) + Chr$(8) + Chr(3) + Chr$(22) + Chr$(13), car) <> 0 Or (car >= "A" And car <= "z") Or (car >= "0" And car <= "9") Then
If car = Chr$(13) Then KeyAscii = 0: repe = Trim(Text1.Text): Exit Sub
KeyAscii = Asc(car)
Else
KeyAscii = 0
End If

End Sub
