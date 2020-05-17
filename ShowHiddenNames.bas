Attribute VB_Name = "ShowHiddenNames"
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

Sub ShowHiddenNames()
    
    '変数宣言
    Dim nameDef As Object
    Dim numOfInvisibleNames As Long: numOfInvisibleNames = 0
    
    '処理
    For Each nameDef In Names
        If Not (nameDef.Visible) Then
            numOfInvisibleNames = numOfInvisibleNames + 1
            nameDef.Visible = True
        End If
    Next
    
    MsgBox "Done!" & vbLf & vbLf & _
           "表示済みオブジェクトを含め、" & Format(Names.Count, "#,###;-#,###;0") & " 件処理しました" & vbLf & _
           "内、" & Format(numOfInvisibleNames, "#,###;-#,###;0") & "件を非表示状態から表示状態に変更しました"
    
End Sub
