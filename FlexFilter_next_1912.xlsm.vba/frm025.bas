
Public Sub OKButton_Click()
    Me.Hide
    
    Call SavePDF
    
    ' Close all
    'Dim UForm As Object
    'Dim i As Integer
    'i = 0
    'For Each UForm In VBA.UserForms
        'Debug.Print UForm.Name
    '    UForm.Hide
    '    Unload VBA.UserForms(i)
    '   i = i + 1
    'Next
    dFunc.msgError = "Tak - din besvarelse er nu gemt !"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Tak - din besvarelse er nu gemt !")
End Sub

Public Sub Tilbage_Click()
    Me.Hide
    SFunc.ShowFunc ("frm024")
    'frm024.Show
End Sub

Private Sub SavePDF()
    ' Save PDF
    Dim PathString
    PathString = Application.ActiveWorkbook.Path
    PathString = PathString & "\SpørgeskemaBesvarelse.pdf"
    
    Worksheets("PDF").Activate
    
    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
    End With
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=PathString, _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
            :=False, OpenAfterPublish:=False

End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeStretch

End Sub