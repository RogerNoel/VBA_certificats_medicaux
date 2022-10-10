Attribute VB_Name = "Module11"
Public nbre_lignes As Integer
Public blanc As Integer
Public vert As Integer
Public couleur As String
Public tampon As String


Sub calcul()
Attribute calcul.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.ScreenUpdating = False
    Application.Calculation = xlAutomatic
    Sheets("MALADIE").Select
    Columns("F").AutoFit
    
    Call Module1.calculNombreLignes
    Call Module1.recopierFormules
    Call Module1.inverserCouleur
    Call Module1.maladieVers304
   
    Sheets("304").Select
    Range("H:L").EntireColumn.Hidden = True
    Call Module1.inverserCouleur
    
    Columns("M:N").AutoFit
    Sheets("MALADIE").Select
    Range("a4").Select
    MsgBox "Opération terminée"
End Sub
