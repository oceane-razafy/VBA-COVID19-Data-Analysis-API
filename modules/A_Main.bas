Attribute VB_Name = "A_Main"
Option Explicit
Option Base 1

Public Sub Main()

    Dim wb As Workbook, ws_current As Worksheet, ws_extraction As Worksheet, rg As Range 'fichier, feuille de donn�es, feuille 'Extraction des donn�es', plage de donn�es
    Dim nb_dates As Integer, row_dates As Integer 'nombre de dates, rang�e d'une date
    Dim typedata(3) As String, nb_pays As Integer 'tableau avec types de donn�es 'Infections', 'D�c�s', 'Taux de d�c�s', nombre de pays choisis
    Dim json_Text As String, name_ranges(3) As String 'requ�te Json, tableau avec le nom des plages de donn�es
    Dim i As Integer, j As Integer, status As Integer 'i boucle sur les feuilles, sur les plages de donn�es et j boucle sur les pays, statut de la requ�te
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'I.INITIALISATION DE CERTAINES VARIABLES
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Cr�ation des tableaux "typedata" et "name_ranges", avec les type de donn�es recherch�s et le nom des plages
    Call affectation_tables(typedata, name_ranges)
    
    'Affectation du fichier et de la feuille "Extraction" dans des variables
    Set wb = ThisWorkbook
    Set ws_extraction = wb.Worksheets("Extraction des donn�es")
    
    'Affectation du nombre de pays dans une variable
    nb_pays = ws_extraction.Range("PAYS").Count
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'II.EXTRACTION DES DONNEES
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Extraire la requ�te, et prendre la partie dont on a besoin
    json_Text = MakeQuery(Range("URL").Value, status)
    
    'Tester la requ�te
    If test_status(status) = False Then Exit Sub

    'Extraire la partie de la requ�te que l'on veut
    json_Text = cut_json(json_Text)

    'Boucle sur les feuilles "Infection", "D�c�s" et "TauxDeces"
    For i = LBound(typedata) To UBound(typedata)
    
         'Prendre la bonne plage de donn�es
         Set ws_current = wb.Worksheets(typedata(i))
         Set rg = ws_current.Range(name_ranges(i))
    
         'Nombre de lignes d'un tableau de donn�es
         nb_dates = rg.Rows.Count
    
        'Boucle sur les pays
        For j = 1 To nb_pays
        
            'Supprimer l'extraction des anciennes donn�es
            rg.Columns(j + 1).ClearContents
                
            'Boucle sur les dates
            'Extraire le nombre d'infections, celui de d�c�s ou le taux de d�c�s pour une certaine date, et pour un certain pays
            For row_dates = 1 To nb_dates
            
                putValue json_Text, j + 1, typedata(i), row_dates, rg
                
                'Si le pays n'a pas �t� trouv�
                If rg.Cells(row_dates, j + 1).Value = "PAYS NON TROUVE" Then
                    'Message d'erreur
                    MsgBox ("Le pays " & Chr(34) & rg.Cells(0, j + 1).Value & Chr(34) & " n'a pas �t� trouv�," & _
                    " veuillez en saisir un autre." & Chr(13) & Chr(10) & "Extraction interrompue")
                    'Arr�t de la macro
                    Exit Sub
                End If
            Next row_dates
            
            'Gestion d'erreur en cas de nombre manquant
            GestionErreur nb_dates, rg, j + 1
        Next j
        
        'Ajuster les axes des graphiques
        Graphique ws_current, rg, nb_dates, nb_pays
    Next i
End Sub

Private Sub affectation_tables(ByRef typedata() As String, ByRef name_ranges() As String)
'Proc�dure remplissant les tableaux "typedata", et "name_ranges"
    typedata(1) = "Infection"
    typedata(2) = "Deces"
    typedata(3) = "TauxDeces"
    
    name_ranges(1) = "rg_infection"
    name_ranges(2) = "rg_deces"
    name_ranges(3) = "rg_tauxdeces"
End Sub
   
Private Sub GestionErreur(ByVal nb_dates As Integer, ByVal rg As Range, ByVal pays As Integer)
'Proc�dure permettant de s'occuper du cas o� une donn�e est manquante pour une certaine date
    'si nombre est manquant, remplacer par la valeur du jour pr�c�dent
    'si le nombre manquant est celui de la derni�re date de la liste, remplacer par l'avant dernier nombre

    Const vide As String = "VIDE"
    
    Dim row_dates As Integer

    For row_dates = 1 To nb_dates
        If rg.Cells(row_dates, pays).Value = vide Then
            If row_dates <> nb_dates Then
                rg.Cells(row_dates, pays).Value = rg.Cells(row_dates + 1, pays).Value
            Else
                rg.Cells(row_dates, pays).Value = rg.Cells(row_dates - 1, pays).Value
            End If
        End If
    Next row_dates

End Sub

Private Sub Graphique(ByVal ws_current As Worksheet, ByVal rg As Range, ByVal nb_dates As Integer, ByVal nb_pays As Integer)
'Proc�dure permettant d'ajuster les axes sur les graphiques
    Dim chObj As ChartObject
    
    Set chObj = ws_current.ChartObjects(1)
    With chObj.Chart
        .Axes(xlValue).MinimumScale = WorksheetFunction.Min(rg.Cells(1, 2).Resize(nb_dates, nb_pays))
        .Axes(xlCategory).MinimumScale = rg.Cells(nb_dates, 1)
        .Axes(xlCategory).MaximumScale = rg.Cells(1, 1)
    End With
End Sub

Private Sub putValue(ByVal json_Text As String, ByVal pays As Integer, _
    ByVal type_data As String, ByVal row_dates As Integer, ByVal rg As Range)
'Proc�dure permettant de renseigner la cellule avec la donn�e r�cup�r�e
    Dim day As String
    day = Application.WorksheetFunction.Text _
    (rg.Cells(row_dates, 1).Value, "yyyy-mm-dd") & "T00:00:00"
    rg.Cells(row_dates, pays).Value = ExtractSubstringwithdate _
    (json_Text, day, rg.Cells(0, pays), type_data, ",")
End Sub
