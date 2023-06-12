Attribute VB_Name = "fn_Extraction"
Option Explicit

Public Function MakeQuery(ByVal url As String, ByRef status As Integer) As String
'Fonction qui � partir de l'url permet de r�cup�rer la requ�te sous format json
    Dim webS As MSXML2.XMLHTTP60
    Set webS = New MSXML2.XMLHTTP60
    
    Call webS.Open("GET", url, False)
    Call webS.send
    
    status = webS.status


    Let MakeQuery = webS.responseText
    

End Function

Function ExtractSubstringwithdate(ByVal jsonFile As String, ByVal markup_date As String, ByVal markup_pays As String _
    , ByVal markup_typedata As String, ByVal markup_virgule As String) As String
'Fonction permettant d'extraire le nombre d'infections, celui de d�c�s ou le taux de d�c�s � la date voulue
    
    Const taille_date As Integer = 19 'Taille d'une date dans la requ�te
    Const separation_date_pays As Integer = 29 'Taille s�parant un pays et sa date correspondante
    
    Dim index_date As Long, index_pays As Long, index_typedata As Long, indexEnd As Long
    
    'Trouver la date
    Let index_date = InStr(jsonFile, markup_date)
    
    'Date non trouv�e
    If index_date = 0 Then
        'Extraire "VIDE", qui sera remplac� plus tard dans le module "MAIN" par le nombre du jour pr�c�dent
        ExtractSubstringwithdate = "VIDE"
    Else
        'Trouver le pays
        Let index_pays = InStr(index_date, jsonFile, markup_pays)
        'Si le pays n'est pas trouv�
        If index_pays = 0 Then
           'La macro s'arr�tera dans "MAIN"
            ExtractSubstringwithdate = "PAYS NON TROUVE"
        Else
            'Si un nombre est manquant pour un pays, pour une certaine date, extraire "VIDE" remplac� plus tard par le nombre du jour pr�c�dent
            If Mid(jsonFile, index_date, taille_date) <> Mid(jsonFile, index_pays - separation_date_pays, taille_date) Then
                ExtractSubstringwithdate = "VIDE"
            Else
            'Sinon extraire le nombre
                Let index_typedata = InStr(index_pays, jsonFile, markup_typedata) + Len(markup_typedata)
                Let indexEnd = InStr(index_typedata, jsonFile, markup_virgule)
                Let ExtractSubstringwithdate = Mid(jsonFile, (index_typedata + 2), (indexEnd - index_typedata - 2))
            End If
        End If
    End If
End Function

Function cut_json(json_Text As String) As String
'Fonction permettant d'extraire que la partie du json dont on a besoin --> Partie "PaysData"
    Dim index_paysdata As String
    
    index_paysdata = InStr(json_Text, "PaysData")
    
    cut_json = Mid(json_Text, index_paysdata, Len(json_Text) - index_paysdata)

End Function
