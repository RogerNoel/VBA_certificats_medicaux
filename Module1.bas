Attribute VB_Name = "Module1"
Sub recopierFormules()
    ' recopie des formules de la ligne 3 sur les autres lignes
    Range("h3:n3").Select
    Selection.AutoFill Destination:=Range("h3:n" & nbre_lignes), Type:=xlFillDefault
    Range("P3").Select
    Selection.AutoFill Destination:=Range("P3:P" & nbre_lignes)
    Range("a3").Select
End Sub

Sub inverserCouleur()
    ' assignation des couleurs: 2 = blanc et 35 = vert clair
    couleur = 2
    inverse = 35
    tampon = 35
 ' la premi�re ligne sera blanche
    Range("a3:o3").Interior.ColorIndex = blanc
    
    Call calculNombreLignes
    
    ' code pour inverser les couleurs � chaque changement de travailleur
    For i = 4 To nbre_lignes
        If Cells(i, 1) = Cells(i - 1, 1) Then
        ' on garde la couleur du dernier enregistrement
        Else
            consecutif = False
            ' inverser la couleur avec un syst�me de tampon
            tampon = couleur
            couleur = inverse
            inverse = tampon
        End If
        ' on colorie la ligne selon la couleur choisie
        Range("a" & i & ":o" & i).Interior.ColorIndex = couleur
        ' quadrillage --->
        Range("a" & i & ":o" & i).Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
' --> fin quadrillage
    Next
' couleur = blanc
' inverse = vert
' tampon = vert
' ------------
' couleur = tampon ----> vert
' inverse = couleur -----> blanc
' tampon = inverse ------> blanc
End Sub

Sub calculNombreLignes()
    nbre_lignes = Cells.Find(What:="*", searchdirection:=xlPrevious).Row
End Sub

Sub maladieVers304()
' recherche des cellules "d�but 304" non vides pour les copier
    ' et les coller dans la feuille "304" par ordre roissant
    Dim num_trav As Integer
    Dim j As Integer
    j = 4 'j est le compteur pour rechercher l'endroit o� coller la ligne
    For i = 4 To nbre_lignes
        If Cells(i, 14).Value <> "" Then ' colonne "d�but 304"
            If Cells(i, 15).Value <> "OK" Then ' colonne "trait�"
                num_trav = Cells(i, 1).Value ' copie le n� du travailleur
                Rows(i).Select
                Selection.Copy
                Range("O" & i).Value = "OK" ' pour �viter de recopier une seconde fois vers 304
                Sheets("304").Select
                ' comparaison entre le num�ro de travailleur copi� et le num�ro de travailleur de la feuille 304
                Do
                    j = j + 1
                Loop Until Cells(j, 1) > num_trav
                Rows(j).Select
                Selection.Insert Shift:=xlDown
                Range("A" & j).PasteSpecial Paste:=xlPasteValues
                Range("C" & j).Value = Range("M" & j).Value
                ' Test suivant: si une p�riode B en suit une autre A qui a d�j� d�pass� les 30 jours, la ligne de
                ' la p�riode B va s'ajouter dans 304. Pour �viter cette r�p�tition je compare la nouvelle ligne que je colle
                ' avec la pr�c�dente: si la date de d�but maladie ET le n� de travailleur sont les m�mes, je supprime
                ' cette ligne qui vient d'�tre coll�e.
                If Range("C" & j).Value = Range("C" & j - 1).Value And Range("A" & j).Value = Range("A" & j - 1).Value Then
                    Rows(j - 1).EntireRow.Delete
                    j = j - 1 ' je d�cr�ment J car j'ai supprim� une ligne
                End If
                Rows(j).Interior.ColorIndex = 2
                Columns("F").AutoFit
                Sheets("MALADIE").Select
            End If
        End If
    Next
End Sub
