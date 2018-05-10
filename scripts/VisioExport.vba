Sub Export2SVG()

    Dim vsoStencil As Visio.Document
    Dim vsoDocument As Visio.Document
    Dim vsoMaster As Visio.Master
    Dim vsoShape As Visio.Shape
    
    Set vsoStencil = Documents.Add("C:\Users\rbl\Downloads\OfficeSymbols_2012and2014\s_symbols_Concepts_2014.vss")
    Set vsoDocument = Documents.Add("")
    
    For Each vsoMaster In vsoStencil.Masters
        Debug.Print "Exporting; " & vsoMaster.Name & "…"
        Set vsoShape = vsoDocument.Pages(1).Drop(vsoMaster, 0#, 0#)
        
        vsoShape.Export ("C:\temp\Concepts\" & Replace(vsoMaster.Name, "/", "_") & ".svg")
        vsoShape.Delete
    Next
    
    vsoStencil.Close
    vsoDocument.Saved = True
    vsoDocument.Close
End Sub
