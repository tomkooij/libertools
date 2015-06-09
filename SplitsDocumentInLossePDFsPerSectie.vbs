Sub SplitDocumentInLossePDFsPerSectie()
    ' Splits Worddocument in losse PDFs per sectie
    ' Een word macro om het leven met Liber te verbeteren
    ' Tom Kooij juni 2015

    MsgBox "Deze Marco splits een document in PDFs per SECTIE EINDE. De laatste sectie wordt meestal overgeslagen"

    ' Sla huidig bestand op!
    ActiveDocument.Save

    If InStrRev(ActiveDocument.Name, ".") <> 0 Then
        BestandsNaamStr = Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1)
    Else
        BestandsNaamStr = ActiveDocument.Name
    End If

    ' kies uitvoer map
    Dim folder As FileDialog
    Set folder = Application.FileDialog(msoFileDialogFolderPicker)
    folder.Title = "Kies uitvoermap:"

    ' Als folder gekozen, dan doe je ding
    If folder.Show = -1 Then
        strFolder = folder.SelectedItems(1)

        ChangeFileOpenDirectory strFolder

        'Selecteer de eerste sectie
        Selection.GoTo What:=wdGoToSection, Which:=wdGoToFirst

        DocNum = 0

        ' sla laatste sectie over: geeft vaak foutmelding
        For SectionIndex = 1 To (ActiveDocument.Sections.Count - 1)

            ' kopieer de hele sectie naar nieuw document
            ActiveDocument.Bookmarks("\Section").Range.Copy
            Documents.Add
            Selection.Paste

            ' maak bestandsnaam als string
            ' eerste sectie heet _par_0 de eerste paragraaf (tweede sectie) heet _par_1 enzovoort
            strFileName = BestandsNaamStr & "_par_" & DocNum
            DocNum = DocNum + 1

            ' sla op als .docx
            ' ActiveDocument.SaveAs FileName:=strFileName & ".docx"

            ' exporteer het bestand als PDF
            ActiveDocument.ExportAsFixedFormat OutputFileName:=strFileName & ".pdf", _
                                   ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
                                   wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=99, _
                                   Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
                                   CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
                                   BitmapMissingFonts:=True, UseISO19005_1:=True

            ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges

            ' Move the selection to the next section in the document.
            Application.Browser.Next
       Next SectionIndex
    Else
        MsgBox "Geen map gekozen. Einde."
    End If
End Sub
