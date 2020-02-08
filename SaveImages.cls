VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' �ۑ����Ɏ������s�̏ꍇ�A���̊֐����g�p����
'Private Sub Document_DocumentSaved(ByVal Document As IVDocument)

' �摜�̕ۑ�(HTML�̕t���Ƃ���)
Public Sub SaveImages()
    Dim outputPath As String
    Dim options As String

    outputPath = ThisDocument.path & ThisDocument.Name

    options = "/quiet=True" _
            & " /target=" & outputPath _
            & " /priformat=png" _
            & " /openbrowser=False"

    ' # output file name
    ' xxx.files/png_[N].png
    Application.Addons("SaveAsWeb").Run options

    Call RenameFiles
End Sub

' �ۑ���f�B���N�g���̍쐬
Private Function CreateImgOutputPath()
    Dim imgOutputPath, msg As String
    imgOutputPath = ThisDocument.path & "images" ' "img_" & ThisDocument.Name

    msg = "[False]"
    If Dir(imgOutputPath, vbDirectory) = "" Then
        ' �t�H���_�����݂��Ȃ��ꍇ�A�V�K�쐬
        MkDir imgOutputPath
        msg = "[True]"
    End If
    msg = msg & imgOutputPath
    ' return
    CreateImgOutputPath = imgOutputPath
End Function

Private Function GetPngName(num As Integer)
    ' return
    GetPngName = "png_" & num & ".png"
End Function

' �t�@�C���R�s�[
Private Sub CopyFile(src As String, dst As String)
    Debug.Print (src)
    Debug.Print (dst)
    Call FileCopy(src, dst)
End Sub

' �t�@�C���폜
Private Sub DeleteFile(path As String)
    Call Kill(path)
End Sub

' �f�B���N�g���폜
Private Sub DeleteDir(path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.DeleteFolder(path, True)
    Set fso = Nothing
End Sub

' �t�@�C�������l�[��
Private Sub RenameFiles()
    Dim i As Integer
    Dim imgOutputPath As String
    Dim outputPath As String
    Dim visioNoExtName As String
    Dim htmlPngPath As String
    Dim ImgPngPath As String
    Dim htmlFile As String

    ' �g���q�Ȃ�Visio��
    visioNoExtName = Replace(ThisDocument.Name, ".vsd", "")
    ' HTML��̏o�͐�
    outputPath = ThisDocument.path & ThisDocument.Name & ".files"
    ' �ۑ���̍쐬
    imgOutputPath = CreateImgOutputPath()

    ' ���݃y�[�W�������[�v
    i = 1
    For Each Page In ThisDocument.pages
        htmlPngPath = outputPath & "\" & GetPngName(i)
        ImgPngPath = imgOutputPath & "\" & visioNoExtName & "_" & Page.Name & ".png"

        Call CopyFile(htmlPngPath, ImgPngPath)

        Debug.Print (i & "]" & Page.Name & ", " & htmlPngPath)
        i = i + 1
    Next

    ' �s�v�Ȃ��̂��폜
    htmlFile = ThisDocument.path & ThisDocument.Name & ".htm"
    DeleteFile (htmlFile)
    Call DeleteDir(outputPath)
End Sub
