Attribute VB_Name = "fn_status"
Option Explicit

Function test_status(ByVal status As String) As Boolean
'Fonction permettant de tester l'erreur http
    If status = 200 Then
        test_status = True
    Else
        MsgBox ("La requête n'a pas pu s'exécuter. Erreur : " & status)
        test_status = False
    End If
End Function


