Sub InsertImageATSComments()
    Dim sh As Shape
    Dim sPath As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        If .Show Then
            sPath = .SelectedItems(1)
            ActiveSheet.Shapes.AddPicture Filename:=sPath, LinkToFile:=0, SaveWithDocument:=-1, _
            Left:=ActiveCell.Left + 2, Top:=ActiveCell.Top + 2, Width:=ActiveCell.Width - 3, Height:=ActiveCell.Height - 3

            Else
                MsgBox ("Cancelled.")
        End If
    End With
End Sub
