VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Encodage d'une période de maladie"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6150
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public connu As Boolean
Public recuperationLigne As Integer ' pour récupérer les ligne d'un travailleur connu
Public dateDernierCertif As Date
Public dateNouveauCertif As Date
Public delaiEntreCertificats As Integer

Private Sub TextBox3_Change()

End Sub

Private Sub TextBox4_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Public Sub Valider_Click()

    Do While TextBox1.Value = "" Or TextBox2.Value = "" Or TextBox3.Value = "" Or TextBox4.Value = ""
    MsgBox ("Renseignements incomplets.")
    Exit Sub
    Loop

    Dim num_trav_form As Integer
    Dim nom As String
    Dim date_debut As Date
    Dim date_fin As Date
    Dim rechute As Boolean
    Dim nbre_lignes As Integer
    
    Sheets("MALADIE").Select
    
    num_trav_form = CInt(TextBox1.Value)
    nom = TextBox2.Value
    
    If Not IsDate(TextBox3.Value) Then
        MsgBox "Entrez une date de début valide.", vbInformation, "Date"
        Exit Sub
        Else
        date_debut = TextBox3.Value
    End If
    
    If Not IsDate(TextBox4.Value) Then
        MsgBox "Entrez une date de fin valide.", vbInformation, "Date"
        Exit Sub
        Else
        date_fin = TextBox4.Value
    End If
    
    rechute = CheckBox1.Value
    
    If date_debut > date_fin Then
    MsgBox "La date de début est antérieure à la date de fin.", vbInformation, "Vérification des dates"
    Exit Sub
    End If
    
    ' Call Module1.calculNombreLignes --> ne fonctionne pas dans cette feuille ???
    nbre_lignes = Cells.Find(What:="*", searchdirection:=xlPrevious).Row
    
    If connu Then ' s'il est connu il faut comparer la date de fin du dernier certif avec la date début du nouveau
        ' récupérer la date de fin du dernier certif
        dateDernierCertif = Range("D" & recuperationLigne).Value
        ' récupérer la date de début du nouveau certif
        dateNouveauCertif = date_debut
        delaiEntreCertificats = dateNouveauCertif - dateDernierCertif
            If delaiEntreCertificats <= 1 Then
                MsgBox ("INFO: la date de début du nouveau certificat est antérieure à la date de fin du dernier certificat.")
            End If
            ' si date début nouveau certif - date fin ancien certif < 15
        If delaiEntreCertificats < 15 Then
            ' - alors on garde le ligne et on ne change que la date de fin
            ' récupérer l'adresse de la dernière ligne de fin de certif
            ' on assigne la nouvelle date de fin
            Range("D" & recuperationLigne).Value = date_fin
        Else
            ' sinon >= 15
            ' on fait une nouvelle ligne
            ' code pour insérer les données formulaire à la bonne place
            For i = 4 To nbre_lignes
                If Range("a" & i).Value > num_trav_form Then
                Range("A" & i).EntireRow.Insert
                Range("a" & i) = num_trav_form
                Range("b" & i) = UCase(nom)
                Range("c" & i) = date_debut
                Range("d" & i) = date_fin
                    If rechute Then
                        Range("e" & i) = "O"
                    End If
                    Exit For
                End If
            Next
        End If
        GoTo continue
    Else
        ' on fait une nouvelle ligne
        ' code pour insérer les données formulaire à la bonne place
        For i = 4 To nbre_lignes
            If Range("a" & i).Value > num_trav_form Then
            Range("A" & i).EntireRow.Insert
            Range("a" & i) = num_trav_form
            Range("b" & i) = UCase(nom)
            Range("c" & i) = date_debut
            Range("d" & i) = date_fin
                If rechute Then
                    Range("e" & i) = "O"
                End If
                Exit For
            End If
        Next
    End If

continue:
    Columns("b").AutoFit
    Worksheets("304").Columns("B").AutoFit
    Worksheets("MALADIE").Columns("M").AutoFit
'    Worksheets("MALADIE").Columns("G:L").Hide

    Unload Me
    UserForm1.Show vbModeless
End Sub

Private Sub Quitter_Click()
    Unload Me
    UserForm2.Show vbModeless
End Sub

Private Sub TextBox1_AfterUpdate()
    ' ***************************************************************************************
    ' * cette sub sert à tester si le n° de travailleur encodé existe déjà dans la liste
    ' * active la cellule de la ligne si le n° est connu
    ' * pré-remplit le formulaire avec le nom si connu
    ' * récupère le numéro de ligne du dernier certificat de ce travailleur
    ' ***************************************************************************************
    
    ' ce code active la ligne correspondante au n° de dossier qu'on a entré dans le textbox1 du formulaire
    Dim nbre_lignes As Integer
    nbre_lignes = Cells.Find(What:="*", searchdirection:=xlPrevious).Row
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim FoundCell As Range
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    Dim numDossier As String
    numDossier = TextBox1.Value
    Set FoundCell = ws.Range("A:A").Find(What:=numDossier)
        If Not FoundCell Is Nothing Then ' si le travailleur est connu
            Range("A" & FoundCell.Row).Activate
            TextBox2.Value = Range("B" & FoundCell.Row).Value ' pré-remplit le formulaire avec le nom
            connu = True
            For i = FoundCell.Row To nbre_lignes ' depuis la ligne de la 1° occurrence du n° travailleur jusqu'au bout de la feuille
                If numDossier <> Range("A" & i + 1).Value Then ' si le n° de dossier de la ligne suivante est différent de l'actuel
                    recuperationLigne = i ' on récupère le n° de la ligne: c'est dans cette ligne qu'on éditera la date de fin de certif
                    GoTo continue
                End If
            Next i
        Else ' si travailleur inconnu
            MsgBox (numDossier & " ne figure pas dans la liste, il sera créé.")
            connu = False
        End If
continue:
End Sub

