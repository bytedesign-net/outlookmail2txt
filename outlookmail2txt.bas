Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
     'SaveText EntryIDCollection
    Dim EntryIDArray() As String
    EntryIDArray = Split(EntryIDCollection, ",")
    Dim i As Integer
    For i = 0 To UBound(EntryIDArray)
    SaveText EntryIDArray(i)
    Next i
End Sub
'
Private Sub SaveText(ByVal EntryIDCollection As String)
     Const AUTO_SAVE_TITLE = "重要" ' 件名に含まれるキーワード
     Dim i As Integer
     Dim myMsg
     ' メッセージの取得
     'Set myMsg = Session.GetItemFromID(EntryIDCollection)
      Set myMsg = Session.GetItemFromID(EntryID)
      
     'strMsg = myMsg.Body
     
     ' 件名にキーワードが含まれていたら
     If InStr(myMsg.Subject, AUTO_SAVE_TITLE) > 0 Then
        Dim NOW_DATE As String
        NOW_DATE = Format(Date, "yyyymmdd")
        NOW_TIME = Format(Time, "hhmmss")
        Dim TXT_FILE As String
        TXT_FILE = "C:\users\%username%\Desktop\Outlook重要メール" & NOW_DATE & NOW_TIME & ".txt" ' 保存する Text ファイルの名前
        
        myMsg.SaveAs TXT_FILE, olTXT
        MsgBox ("「デスクトップ\Outlook重要メール」にメッセージを保存しました")
     End If
End Sub

