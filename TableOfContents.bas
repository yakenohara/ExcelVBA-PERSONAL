Attribute VB_Name = "TableOfContents"
'<License>------------------------------------------------------------
'
' Copyright (c) 2018 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

'
'目次を作る
'フォーカスがあたっているセルを書き込み開始セルとみなし、
'全シート名のリンク付きリストを作ります
'
Sub TableOfContents()
    
    '変数宣言
    Dim writePlace As Range
    Dim numOfWorkSheets As Long
    Dim cout As Long
    
    Dim cautionMessage As String: cautionMessage = "このSubプロシージャは、" & vbLf & _
                                                   "現在の選択範囲に対して値の書き込みを行います。" & vbLf & vbLf & _
                                                   "実行しますか?"
    
    '実行確認
    retVal = MsgBox(cautionMessage, vbOKCancel + vbExclamation)
    If retVal <> vbOK Then
        Exit Sub
        
    End If
    
    'シート選択状態チェック
    If ActiveWindow.SelectedSheets.Count > 1 Then
        MsgBox "複数シートが選択されています" & vbLf & _
               "不要なシート選択を解除してください"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Set writePlace = Cells(Selection.Row, Selection.Column)
    numOfWorkSheets = ActiveWorkbook.Worksheets.Count
    
    '上書き確認
    If WorksheetFunction.CountA(Range(writePlace, Cells(writePlace.Row + numOfWorkSheets - 1, writePlace.Column))) > 0 Then
        yn = MsgBox("作成先のセルに値が入っています" & vbLf & vbLf & _
                    "上書きしますか？", _
                    vbOKCancel)
        
        If yn = vbCancel Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
    End If
    
    cout = 0
    For Each sh In ActiveWorkbook.Worksheets
        '進捗表示
        Application.StatusBar = "Progress:" & cout & "/" & numOfWorkSheets
    
        '書式を文字列型に変更
        writePlace.Clear
        writePlace.NumberFormatLocal = "@"
        
        'ハイパーリンクの作成
        ActiveSheet.Hyperlinks.Add _
                                Anchor:=writePlace, _
                                Address:="", _
                                SubAddress:="'" & sh.Name & "'!A1", _
                                TextToDisplay:="'" & sh.Name
                                
        '書き込み先セル位置の移動
        Set writePlace = Cells(writePlace.Row + 1, writePlace.Column)
        
        cout = cout + 1
    Next sh
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox ("Done!")
    
End Sub

