'-----------------------------------------------------------------------------------------
'- 作成日　    　2020/90/04
'- 最終更新日　2020/09/09
'- 作成者  　    r_suzuki
'-----------------------------------------------------------------------------------------
Option Explicit

Sub main()

    aplrun
    filereadtext
    check
    IPreadtext
    IPdelete
    
    MsgBox "終了しました。"
    Application.DisplayAlerts = False
    Workbooks("レポート.xlsx").Save
    Workbooks("レポート.xlsx").Close
    Application.DisplayAlerts = True
End Sub

Private Sub aplrun()

    Dim ws
    
    Set ws = CreateObject("WScript.Shell")
    ws.CurrentDirectory = "C:\Users\******\Desktop" 'パスを指定
    ws.Run """ファイル抽出.vbs""", 1, True                        'apl名を指定し実行
    ws.Run """IP抽出.vbs""", 1, True
    
End Sub

Private Sub filereadtext()

    Dim wbName As String
    wbName = "レポート.xlsx"    'ファイル名指定
    
    Dim wb As Workbook
    Set wb = Workbooks.Add()    '保存用book作成
    wb.SaveAs wbName              'ファイル名変更
    
    Dim ws As Worksheet
    Set ws = Worksheets.Add()               'ワークシートを作成（作成したシートがアクティブになる）
    ActiveSheet.Name = "ファイル一覧"       'シート名変更
    
    MsgBox "ファイル一覧のテキストを選択してください。"
    Dim txtName As String
    txtName = Application.GetOpenFilename("テキストファイル,*.txt")
    
    If txtName <> "False" Then
        Open txtName For Input As #1
    End If
    
    Dim r As Long
    r = 1 '1行目から書き出す
    
    Do Until EOF(1)
    
        Dim buf As String
        Line Input #1, buf
        
        Dim aryLine As Variant '文字列格納用配列変数
        aryLine = Split(buf, vbTab) '読み込んだ行をタブ区切りで配列変数に格納
        
        Dim i As Long
        For i = LBound(aryLine) To UBound(aryLine)
            'インデックスが0から始まるので列番号に合わせるため+1
            Cells(r, i + 1) = aryLine(i)
        Next
        
            r = r + 1
    
    Loop
    
    Close #1
 
End Sub

Private Sub check()
    
    Dim rowsData As Long '行数カウント用の変数
    'rowsData = wsData.Cells(Rows.count, 1).End(xlUp).Row '最後の行数を取得
    
    Dim s As String
    Dim i As Integer
    Dim a As Integer
    Dim str(28) As String '標準APL格納
    
    str(0) = "Access.lnk"
    str(1) = "Accessibility"
    str(2) = "Accessories"
    str(3) = "Administrative Tools"
    str(4) = "Adobe Acrobat 2017.lnk"
    str(5) = "Adobe Acrobat Distiller 2017.lnk"
    str(6) = "CyberLink Media Suite"
    str(7) = "Excel.lnk"
    str(8) = "Google Chrome.lnk"
    str(9) = "Internet Explorer.lnk"
    str(10) = "Java"
    str(11) = "LAPS"
    str(12) = "Maintenance"
    str(13) = "McAfee"
    str(14) = "Microsoft Edge.lnk"
    str(15) = "Microsoft Endpoint Manager"
    str(16) = "Microsoft Office ツール"
    str(17) = "Microsoft System Center"
    str(18) = "NetMotion"
    str(19) = "Outlook.lnk"
    str(20) = "PowerPoint.lnk"
    str(21) = "Publisher.lnk"
    str(22) = "Server Manager.lnk"
    str(23) = "Skype for Business.lnk"
    str(24) = "StartUp"
    str(25) = "System Tools"
    str(26) = "VAIO"
    str(27) = "Word.lnk"
    
    Range("A1").Select
    
    i = 0
    a = 0
    
     '// 空セルまでループ
    Do
        s = ActiveCell.Offset(i, 0).Value   '// セル値を取得
        
        If s = "" Then                            '// セル値が未設定(空)の場合
            Exit Do                                  '// ループを抜ける
        
        ElseIf s = str(a) Then                  '//標準APLか確認
            a = a + 1
            
        Else
            Cells(i + 1, 1).Interior.Color = RGB(200, 200, 200) '//標準APLでないため色を付ける
        
        End If
        
            i = i + 1                               '// ループカウンタを加算
        
    Loop
        
End Sub

Private Sub IPreadtext()

    Dim ws As Worksheet
    Set ws = Worksheets.Add()               'ワークシートを作成（作成したシートがアクティブになる）
    ActiveSheet.Name = "IP情報"              'シート名変更
    
    MsgBox "IP情報のテキストを選択してください。"
    Dim txtName As String
    txtName = Application.GetOpenFilename("テキストファイル,*.txt")
    
    If txtName <> "False" Then
        Open txtName For Input As #1
    End If
    
    Dim r As Long
    r = 1 '1行目から書き出す
    
    Do Until EOF(1)
    
        Dim buf As String
        Line Input #1, buf
        
        Dim aryLine As Variant '文字列格納用配列変数
        aryLine = Split(buf, vbTab) '読み込んだ行をタブ区切りで配列変数に格納
        
        Dim i As Long
        For i = LBound(aryLine) To UBound(aryLine)
            'インデックスが0から始まるので列番号に合わせるため+1
            Cells(r, i + 1) = aryLine(i)
        Next
        
            r = r + 1
    
    Loop
    
    Close #1
    
End Sub

Private Sub IPdelete()

    Range("4:47").Delete
    Range("24:42").Delete
    
End Sub
'<---------------------------------- end of source ---------------------------------->
