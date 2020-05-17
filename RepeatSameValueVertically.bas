Attribute VB_Name = "RepeatSameValueVertically"
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
'選択した範囲の空白でないセル内容を、
'空白セルにコピーする
'
Sub RepeatSameValueVertically()
    
    Dim r As Range
    Dim val As Variant
    Dim retVal As Integer
    
    Dim numOfRows As Long
    Dim numOfCols As Long
    
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
    
    '初期化
    numOfRows = Selection.Rows.Count
    numOfCols = Selection.Columns.Count
    
    '実行ループ
    For colFocus = 1 To numOfCols
    
        val = Selection(1).Offset(0, colFocus - 1).Value
    
        For rowFocus = 2 To numOfRows
            
            If Selection(1).Offset(rowFocus - 1, colFocus - 1).Value <> "" Then
            
                val = Selection(1).Offset(rowFocus - 1, colFocus - 1).Value '書き込み値を更新
                
            Else
                
                Selection(1).Offset(rowFocus - 1, colFocus - 1).Value = val '書き込み値で書き込み
                
            End If
            
        Next
    
    Next
    
    MsgBox "Done!"
    
End Sub
