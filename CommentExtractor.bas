'===============================================================================
' FDL Extractor - Macro VBA Word (v1.1)
'
' Exporte les commentaires du document Word actif vers un classeur Excel au
' format Fiche De Lecture (FDL), avec le NUMERO DE PAGE REEL calcule par le
' moteur de mise en page de Word (Range.Information(wdActiveEndPageNumber)) --
' le meme calcul que celui utilise a l'impression. Aucune estimation ici :
' c'est la valeur exacte, quelle que soit la complexite de la mise en forme
' (tableaux, images, styles, en-tetes/pieds de page...).
'
' Installation : voir INSTALL_VBA.md fourni avec ce fichier.
' Usage : ouvrir le document Word a traiter, puis lancer la macro
'         "ExporterCommentairesFDL" (Developpeur > Macros).
'===============================================================================

Option Explicit

Private Type CommentRowData
    PageNum As Long
    StartPos As Long
    Auteur As String
    DateStr As String
    priorite As String
    Remarque As String
End Type

Private Function CleanCommentText(ByVal rawText As String) As String
    Dim t As String
    t = rawText
    t = Replace(t, Chr$(13), " ")  ' marques de paragraphe
    t = Replace(t, Chr$(7), " ")   ' marques de fin de cellule (commentaires multi-paragraphes)
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    CleanCommentText = Trim$(t)
End Function

Sub ExporterCommentairesFDL()

    Dim wdDoc As Document
    Dim wdComment As Comment
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim fileName As String
    Dim boFo As String
    Dim version As String
    Dim rxVersion As Object
    Dim rxPriorite As Object
    Dim matches As Object

    Dim n As Long
    Dim rows() As CommentRowData
    Dim i As Long, j As Long
    Dim commentText As String
    Dim priorite As String

    If Application.Documents.Count = 0 Then
        MsgBox "Aucun document Word ouvert.", vbExclamation
        Exit Sub
    End If

    Set wdDoc = ActiveDocument

    If wdDoc.Comments.Count = 0 Then
        MsgBox "Ce document ne contient aucun commentaire.", vbInformation
        Exit Sub
    End If

    fileName = wdDoc.Name

    ' --- Deduction BO / FO depuis le nom de fichier -------------------------------
    boFo = "NA"
    If InStr(1, fileName, "BO", vbTextCompare) > 0 Then
        boFo = "BO"
    ElseIf InStr(1, fileName, "FO", vbTextCompare) > 0 Then
        boFo = "FO"
    End If

    ' --- Deduction du numero de version depuis le nom de fichier (motif VXX) -----
    Set rxVersion = CreateObject("VBScript.RegExp")
    rxVersion.IgnoreCase = True
    rxVersion.Global = False
    rxVersion.Pattern = "(?:^|[^A-Za-z0-9])V(\d{1,2})(?![0-9])"
    version = ""
    If rxVersion.Test(fileName) Then
        Set matches = rxVersion.Execute(fileName)
        version = matches(0).SubMatches(0)
    End If

    ' --- Regex de detection du tag de priorite "#0".."#9" dans le texte ----------
    Set rxPriorite = CreateObject("VBScript.RegExp")
    rxPriorite.IgnoreCase = True
    rxPriorite.Global = False
    rxPriorite.Pattern = "#\s*([0-9])\b"

    ' --- Collecte des commentaires (page reelle + tri par position) --------------
    n = wdDoc.Comments.Count
    ReDim rows(1 To n)

    For i = 1 To n
        Set wdComment = wdDoc.Comments(i)

        commentText = CleanCommentText(wdComment.Range.Text)

        priorite = ""
        If rxPriorite.Test(commentText) Then
            Set matches = rxPriorite.Execute(commentText)
            priorite = "P" & matches(0).SubMatches(0)
            commentText = Trim$(Replace(commentText, matches(0).Value, "", 1, 1))
        End If

        With rows(i)
            .PageNum = wdComment.Scope.Information(wdActiveEndPageNumber)
            .StartPos = wdComment.Scope.Start
            .Auteur = wdComment.Author
            .DateStr = Format$(wdComment.Date, "yyyy-mm-dd")
            .priorite = priorite
            .Remarque = commentText
        End With
    Next i

    ' --- Tri par (page, position dans le document) : tri a bulles, n reste petit -
    Dim tmp As CommentRowData
    For i = 1 To n - 1
        For j = 1 To n - i
            Dim needsSwap As Boolean
            needsSwap = False
            If rows(j).PageNum > rows(j + 1).PageNum Then
                needsSwap = True
            ElseIf rows(j).PageNum = rows(j + 1).PageNum And rows(j).StartPos > rows(j + 1).StartPos Then
                needsSwap = True
            End If
            If needsSwap Then
                tmp = rows(j)
                rows(j) = rows(j + 1)
                rows(j + 1) = tmp
            End If
        Next j
    Next i

    ' --- Creation du classeur Excel et ecriture des lignes ------------------------
    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0
    If xlApp Is Nothing Then
        MsgBox "Excel ne semble pas installe sur ce poste : impossible de creer le classeur.", vbCritical
        Exit Sub
    End If

    xlApp.Visible = True
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)
    xlSheet.Name = "Detail"

    Dim headers As Variant
    headers = Array("RefFDL", "Relecteur", "Date ouverture remarque", "Priorite", "JIRA", "BO / FO", "Version", "Page, Paragraphe", "Contenu - Si applicable", "Remarque")

    Dim c As Long
    For c = LBound(headers) To UBound(headers)
        xlSheet.Cells(1, c + 1).Value = headers(c)
    Next c
    xlSheet.rows(1).Font.Bold = True

    For i = 1 To n
        xlSheet.Cells(i + 1, 1).Value = i
        xlSheet.Cells(i + 1, 2).Value = rows(i).Auteur
        xlSheet.Cells(i + 1, 3).Value = rows(i).DateStr
        xlSheet.Cells(i + 1, 4).Value = rows(i).priorite
        xlSheet.Cells(i + 1, 5).Value = ""
        xlSheet.Cells(i + 1, 6).Value = boFo
        xlSheet.Cells(i + 1, 7).Value = version
        xlSheet.Cells(i + 1, 8).Value = rows(i).PageNum        ' <-- numero de page EXACT (Word)
        xlSheet.Cells(i + 1, 9).Value = ""
        xlSheet.Cells(i + 1, 10).Value = rows(i).Remarque
    Next i

    xlSheet.Columns.AutoFit
    xlSheet.Columns(10).ColumnWidth = 60

    MsgBox n & " commentaire(s) exporte(s) vers Excel, avec numero de page exact.", vbInformation

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

End Sub
