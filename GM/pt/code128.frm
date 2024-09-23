VERSION 5.00
Begin VB.Form code128 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Codes barre 128 / 128 bar codes"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "code128.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   96
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   15
      Top             =   5010
      Width           =   8295
   End
   Begin VB.TextBox Text2 
      Height          =   465
      Left            =   2220
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   4500
      Width           =   1215
   End
   Begin VB.TextBox label5 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox label1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   96
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   2
      Top             =   1440
      Width           =   8295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Copier / Copy"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fermer / Close"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Grandzebu (Français)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6600
      MouseIcon       =   "code128.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Grandzebu (English)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6600
      MouseIcon       =   "code128.frx":0884
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   4740
      Width           =   1935
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Supplied under GNU GPL license by :"
      Height          =   195
      Left            =   3600
      TabIndex        =   11
      Top             =   4740
      Width           =   2895
   End
   Begin VB.Label Label10 
      Caption         =   "Here is the code string :"
      Height          =   195
      Left            =   4200
      TabIndex        =   10
      Top             =   320
      Width           =   2295
   End
   Begin VB.Label Label9 
      Caption         =   "Type your code here :"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   320
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Fourni sous license GNU GPL par :"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Voici la chaine de code :"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Voici le résultat / Here is the result :"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Tapez votre code ici :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "code128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2003 (Grandzebu)
'Ce programme ainsi que la police de caractères qui l'accompagne est libre, vous pouvez le redistribuer et/ou le modifier
'selon les termes de la Licence Publique Générale GNU publiée par la Free Software Foundation (version 2 ou bien toute
'autre version ultérieure choisie par vous).
'Les fonctions d'encodage des codes barres sont régies par la Licence Générale Publique Amoindrie GNU (GNU LGPL)
'Ce programme est distribué car potentiellement utile, mais SANS AUCUNE GARANTIE, ni explicite ni implicite,
'y compris les garanties de commercialisation ou d'adaptation dans un but spécifique. Reportez-vous à la Licence
'Publique Générale GNU pour plus de détails.
'Veuillez charger une copie de la license à l'adresse : http://www.gnu.org/licenses/
'Une traduction non officielle se trouve à l'adresse : http://gnu.mirror.fr/licenses/translations.fr.html

'Copyright (C) 2003 (Grandzebu)
'This program and the font which is supplied with it is free, you can redistribute it and/or
'modify it under the terms of the GNU General Public License as published by the Free Software Foundation
'either version 2 of the License, or (at your option) any later version.
'The barcode encoding functions are governed by the GNU Lesser General Public License (GNU LGPL)
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
'even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General
'Public License for more details.
'Please download a license copy at : http://www.gnu.org/licenses/

'V. 3.0.0

Option Explicit
Private CodeClair$, CodeBarre$
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Command2_Click()
  Clipboard.Clear
  Clipboard.SetText label5.Text
End Sub

Private Sub label1_Click()
  Text1.SetFocus
End Sub

Private Sub label5_Change()

End Sub

Private Sub Label6_Click()
  ShellExecute Me.hWnd, "open", "http://grandzebu.net", vbNullString, vbNullString, 3
End Sub

Private Sub Label8_Click()
  ShellExecute Me.hWnd, "open", "http://grandzebu.net/informatique/codbar-en/codbar.htm", vbNullString, vbNullString, 3
End Sub

Private Sub Text1_Change()
  Dim CodeBarre$
  CodeBarre$ = code128$(Text1)
  label5.Text = CodeBarre$
  label1.Text = CodeBarre$
End Sub

Public Function code128$(chaine$)
  'Cette fonction est régie par la Licence Générale Publique Amoindrie GNU (GNU LGPL)
  'This function is governed by the GNU Lesser General Public License (GNU LGPL)
  'V 2.0.0
  'Paramètres : une chaine
  'Parameters : a string
  'Retour : * une chaine qui, affichée avec la police CODE128.TTF, donne le code barre
  '         * une chaine vide si paramètre fourni incorrect
  'Return : * a string which give the bar code when it is dispayed with CODE128.TTF font
  '         * an empty string if the supplied parameter is no good
  Dim i%, checksum&, mini%, dummy%, tableB As Boolean
  code128$ = ""
  If Len(chaine$) > 0 Then
  'Vérifier si caractères valides
  'Check for valid characters
    For i% = 1 To Len(chaine$)
      Select Case Asc(Mid$(chaine$, i%, 1))
      Case 32 To 126, 203
      Case Else
        i% = 0
        Exit For
      End Select
    Next
    'Calculer la chaine de code en optimisant l'usage des tables B et C
    'Calculation of the code string with optimized use of tables B and C
    code128$ = ""
    tableB = True
    If i% > 0 Then
      i% = 1 'i% devient l'index sur la chaine / i% become the string index
      Do While i% <= Len(chaine$)
        If tableB Then
          'Voir si intéressant de passer en table C / See if interesting to switch to table C
          'Oui pour 4 chiffres au début ou à la fin, sinon pour 6 chiffres / yes for 4 digits at start or end, else if 6 digits
          mini% = IIf(i% = 1 Or i% + 3 = Len(chaine$), 4, 6)
          GoSub testnum
          If mini% < 0 Then 'Choix table C / Choice of table C
            If i% = 1 Then 'Débuter sur table C / Starting with table C
              code128$ = Chr$(210)
            Else 'Commuter sur table C / Switch to table C
              code128$ = code128$ & Chr$(204)
            End If
            tableB = False
          Else
            If i% = 1 Then code128$ = Chr$(209) 'Débuter sur table B / Starting with table B
          End If
        End If
        If Not tableB Then
          'On est sur la table C, essayer de traiter 2 chiffres / We are on table C, try to process 2 digits
          mini% = 2
          GoSub testnum
          If mini% < 0 Then 'OK pour 2 chiffres, les traiter / OK for 2 digits, process it
            dummy% = Val(Mid$(chaine$, i%, 2))
            dummy% = IIf(dummy% < 95, dummy% + 32, dummy% + 105)
            code128$ = code128$ & Chr$(dummy%)
            i% = i% + 2
          Else 'On n'a pas 2 chiffres, repasser en table B / We haven't 2 digits, switch to table B
            code128$ = code128$ & Chr$(205)
            tableB = True
          End If
        End If
        If tableB Then
          'Traiter 1 caractère en table B / Process 1 digit with table B
          code128$ = code128$ & Mid$(chaine$, i%, 1)
          i% = i% + 1
        End If
      Loop
      'Calcul de la clé de contrôle / Calculation of the checksum
      For i% = 1 To Len(code128$)
        dummy% = Asc(Mid$(code128$, i%, 1))
        dummy% = IIf(dummy% < 127, dummy% - 32, dummy% - 105)
        If i% = 1 Then checksum& = dummy%
        checksum& = (checksum& + (i% - 1) * dummy%) Mod 103
      Next
      'Calcul du code ASCII de la clé / Calculation of the checksum ASCII code
      checksum& = IIf(checksum& < 95, checksum& + 32, checksum& + 105)
      'Ajout de la clé et du STOP / Add the checksum and the STOP
      code128$ = code128$ & Chr$(checksum&) & Chr$(211)
    End If
  End If
  Exit Function
testnum:
  'si les mini% caractères à partir de i% sont numériques, alors mini%=0
  'if the mini% characters from i% are numeric, then mini%=0
  mini% = mini% - 1
  If i% + mini% <= Len(chaine$) Then
    Do While mini% >= 0
      If Asc(Mid$(chaine$, i% + mini%, 1)) < 48 Or Asc(Mid$(chaine$, i% + mini%, 1)) > 57 Then Exit Do
      mini% = mini% - 1
    Loop
  End If
Return
End Function
