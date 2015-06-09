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

    ' Als map gekozen, dan doe je ding
    If folder.Show = -1 Then
        strFolder = folder.SelectedItems(1)

        ChangeFileOpenDirectory strFolder

        DocNum = 0

        ' Ga naar de eerste sectie
        Selection.GoTo What:=wdGoToSection, Which:=wdGoToFirst

        ' zorg dat Browser.Next werkt
        Application.Browser.Target = wdBrowseSection
        ' sla laatste sectie over: geeft vaak foutmelding
        For SectionIndex = 1 To (ActiveDocument.Sections.Count - 1)

            ' selecteer sectie
             Selection.Range.Sections.First.Range.Select

            ' maak bestandsnaam als string
            ' eerste sectie heet _par_0 de eerste paragraaf (tweede sectie) heet _par_1 enzovoort
            strFileName = BestandsNaamStr & "_par_" & DocNum
            DocNum = DocNum + 1

            ' exporteer DE SELECTIE  als PDF
            ActiveDocument.ExportAsFixedFormat OutputFileName:=strFileName & ".pdf", _
                                   ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
                                   wdExportOptimizeForPrint, Range:=wdExportSelection, _
                                   Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
                                   CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
                                   BitmapMissingFonts:=True, UseISO19005_1:=True

            'ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges

            ' ga naar volgende sectie
            Application.Browser.Next
       Next SectionIndex
    Else
        MsgBox "Geen map gekozen. Einde."
    End If
End Sub
