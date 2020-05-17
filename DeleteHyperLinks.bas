Attribute VB_Name = "DeleteHyperLinks"
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
'�I��͈͂̃n�C�p�[�����N���폜����
'
'
Public Sub DeleteHyperLinks()

    Dim hyperlinksObj As Hyperlinks
    Dim tmpBk As Workbook
    Dim tmpR As Range
    Dim nowSht As Worksheet
    Dim nowAddress As String

    Dim cautionMessage As String: cautionMessage = "����Sub�v���V�[�W���́A" & vbLf & _
                                                   "���݂̑I��͈͂ɑ΂��ĕύX���s���܂��B" & vbLf & vbLf & _
                                                   "���s���܂���?"
    
    '���s�m�F
    retVal = MsgBox(cautionMessage, vbOKCancel + vbExclamation)
    If retVal <> vbOK Then
        Exit Sub
        
    End If
    
    '�V�[�g�I����ԃ`�F�b�N
    If ActiveWindow.SelectedSheets.Count > 1 Then
        MsgBox "�����V�[�g���I������Ă��܂�" & vbLf & _
               "�s�v�ȃV�[�g�I�����������Ă�������"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    '�I��͈͂̕ۑ�
    Set nowSht = ActiveSheet
    nowAddress = Selection.Address
    
    '�V�[�g�͈͑S�I������Ă����ꍇ�́AUsedRange���Ɏ��܂�悤�Ƀg���~���O
    Set range_selection = trimWithUsedRange(Selection)
    
    '�n�C�p�[�����N�̍폜
    For Each c In range_selection
    
        Set hyperlinksObj = c.Hyperlinks
        numOfHyperlink = hyperlinksObj.Count
        
        If (c.Address = c.MergeArea.Cells(1, 1).Address) Then '�ΏۃZ���������Z���̍���łȂ��ꍇ�́A�X�L�b�v
        
            Set c = c.MergeArea
            
            If numOfHyperlink > 0 Then '�n�C�p�[�����N�����݂���ꍇ
                
                'tmpBook���Ȃ���΍쐬����
                If tmpBk Is Nothing Then
                    Set tmpBk = Workbooks.Add
                    
                End If
                
                Set tmpR = tmpBk.Sheets(1).Range(c.Address)
                
                '������tmpBook�̃Z����backup����
                c.Copy
                tmpR.PasteSpecial _
                    Paste:=xlPasteFormats, _
                    Operation:=xlNone, _
                    SkipBlanks:=False, _
                    Transpose:=False
                
                For counter = 1 To numOfHyperlink
    
                    hyperlinksObj(counter).Delete
    
                Next counter
                
                'buckup����������\��t����
                tmpR.Copy
                c.PasteSpecial _
                    Paste:=xlPasteFormats, _
                    Operation:=xlNone, _
                    SkipBlanks:=False, _
                    Transpose:=False
                
            End If
            
        End If
        
    Next c
    
    'tmpBook������Εۑ������ɍ폜����
    If Not (tmpBk Is Nothing) Then
        tmpBk.Close SaveChanges:=False
        
    End If
    
    '�I��͈͂̕���
    nowSht.Range(nowAddress).Select
    
    Application.ScreenUpdating = True
    
    MsgBox "Done!"
    
End Sub

'
' �Z���Q�Ɣ͈͂� UsedRange �͈͂Ɏ��܂�悤�Ƀg���~���O����
'
Private Function trimWithUsedRange(ByVal rangeObj As Range) As Range

    'variables
    Dim ret As Range
    Dim long_bottom_right_row_idx_of_specified As Long
    Dim long_bottom_right_col_idx_of_specified As Long
    Dim long_bottom_right_row_idx_of_used As Long
    Dim long_bottom_right_col_idx_of_used As Long

    '�w��͈͂̉E���ʒu�̎擾
    long_bottom_right_row_idx_of_specified = rangeObj.Item(1).Row + rangeObj.Rows.Count - 1
    long_bottom_right_col_idx_of_specified = rangeObj.Item(1).Column + rangeObj.Columns.Count - 1
    
    'UsedRange�̉E���ʒu�̎擾
    With rangeObj.Parent.UsedRange
        long_bottom_right_row_idx_of_used = .Item(1).Row + .Rows.Count - 1
        long_bottom_right_col_idx_of_used = .Item(1).Column + .Columns.Count - 1
    End With
    
    '�g���~���O
    Set ret = rangeObj.Parent.Range( _
        rangeObj.Item(1), _
        rangeObj.Parent.Cells( _
            IIf(long_bottom_right_row_idx_of_specified > long_bottom_right_row_idx_of_used, long_bottom_right_row_idx_of_used, long_bottom_right_row_idx_of_specified), _
            IIf(long_bottom_right_col_idx_of_specified > long_bottom_right_col_idx_of_used, long_bottom_right_col_idx_of_used, long_bottom_right_col_idx_of_specified) _
        ) _
    )
    
    '�i�[���ďI��
    Set trimWithUsedRange = ret
    
End Function


