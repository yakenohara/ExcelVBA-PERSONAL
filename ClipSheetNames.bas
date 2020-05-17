Attribute VB_Name = "ClipSheetNames"
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
'�J���Ă���u�b�N�̃V�[�g�ꗗ���N���b�v�{�[�h�ɓ\��t���܂�
'�N���b�v�{�[�h�ւ̓\��t����setClipBoad�̃R�����g���Q��
Sub ClipSheetNames()
    '�V�[�g���̕������ێ����܂�
    Dim workSheetNames As String
      
    For Each targetWorkSheet In Sheets
        workSheetNames = workSheetNames & targetWorkSheet.Name & vbCrLf
    
    Next
    
    '�N���b�v�{�[�h�ɐݒ肵�܂�
    SetCB (workSheetNames)

End Sub

'<�N���b�v�{�[�h����>-------------------------------------------

'�N���b�v�{�[�h�ɕ�������i�[
Private Sub SetCB(ByVal str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .Text = str
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub

'�N���b�v�{�[�h���當������擾
Private Sub GetCB(ByRef str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    If .CanPaste = True Then .Paste
    str = .Text
  End With
End Sub

'------------------------------------------</�N���b�v�{�[�h����>
 
