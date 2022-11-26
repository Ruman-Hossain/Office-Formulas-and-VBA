Sub UpdateAll()
    Dim sld As Slide
    Dim shp As Shape       
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Type = msoLinkedOLEObject Then
                If shp.OLEFormat.ProgID Like "Excel.*" Then
                    shp.LinkFormat.Update
                    'shp.LinkFormat.BreakLink
                End If
            End If
        Next shp
    Next sld       
End Sub