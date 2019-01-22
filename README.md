# MS_Word_Macro_Resize_Image
This macro is used to resize the all the images inside your document 

Sub ResizeImage()

Dim i As Long
With ActiveDocument
    For i = 1 To .InlineShapes.Count
        With .InlineShapes(i)
           .Height = 160
            .Width = 160
        End With
    Next i
End With


End Sub

