Attribute VB_Name = "SaveAndExport"
Sub SaveAndExport()

    ActiveDocument.Save
    
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
    Replace(ActiveDocument.FullName, ".docx", ".pdf"), _
    ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
    wdExportOptimizeForPrint, Range:=wdExportAllDocument, Item:= _
    wdExportDocumentContent, IncludeDocProps:=False, KeepIRM:=True, _
    CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
    BitmapMissingFonts:=True, UseISO19005_1:=False

End Sub
