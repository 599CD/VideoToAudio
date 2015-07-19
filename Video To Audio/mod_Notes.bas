Attribute VB_Name = "mod_Notes"
'http://stackoverflow.com/questions/1085436/how-to-use-open-file-dialog-in-vb-6

Sub FileDialog()

    CommonDialog.Filter = "Apps (*.txt)|*.txt|All files (*.*)|*.*"
    CommonDialog.DefaultExt = "txt"
    CommonDialog.DialogTitle = "Select File"
    CommonDialog.ShowOpen
    
    'The FileName property gives you the variable you need to use
    'MsgBox CommonDialog.FileName

End Sub
