VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 保存時に自動実行の場合、この関数を使用する
'Private Sub Document_DocumentSaved(ByVal Document As IVDocument)

' 画像の保存(HTMLの付属として)
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

' 保存先ディレクトリの作成
Private Function CreateImgOutputPath()
    Dim imgOutputPath, msg As String
    imgOutputPath = ThisDocument.path & "images" ' "img_" & ThisDocument.Name

    msg = "[False]"
    If Dir(imgOutputPath, vbDirectory) = "" Then
        ' フォルダが存在しない場合、新規作成
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

' ファイルコピー
Private Sub CopyFile(src As String, dst As String)
    Debug.Print (src)
    Debug.Print (dst)
    Call FileCopy(src, dst)
End Sub

' ファイル削除
Private Sub DeleteFile(path As String)
    Call Kill(path)
End Sub

' ディレクトリ削除
Private Sub DeleteDir(path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.DeleteFolder(path, True)
    Set fso = Nothing
End Sub

' ファイル名リネーム
Private Sub RenameFiles()
    Dim i As Integer
    Dim imgOutputPath As String
    Dim outputPath As String
    Dim visioNoExtName As String
    Dim htmlPngPath As String
    Dim ImgPngPath As String
    Dim htmlFile As String

    ' 拡張子なしVisio名
    visioNoExtName = Replace(ThisDocument.Name, ".vsd", "")
    ' HTML上の出力先
    outputPath = ThisDocument.path & ThisDocument.Name & ".files"
    ' 保存先の作成
    imgOutputPath = CreateImgOutputPath()

    ' 現在ページ数分ループ
    i = 1
    For Each Page In ThisDocument.pages
        htmlPngPath = outputPath & "\" & GetPngName(i)
        ImgPngPath = imgOutputPath & "\" & visioNoExtName & "_" & Page.Name & ".png"

        Call CopyFile(htmlPngPath, ImgPngPath)

        Debug.Print (i & "]" & Page.Name & ", " & htmlPngPath)
        i = i + 1
    Next

    ' 不要なものを削除
    htmlFile = ThisDocument.path & ThisDocument.Name & ".htm"
    DeleteFile (htmlFile)
    Call DeleteDir(outputPath)
End Sub

