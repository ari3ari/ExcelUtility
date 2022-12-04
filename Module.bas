Attribute VB_Name = "Module"
Option Explicit

' todo:������ɑΉ����Ă��邩�ǂ����̊m�F�B���݂͐��l�̂݁��ꉞ�Ή����Ă���Hhttps://excel-ubara.com/excel5/EXCEL846.html�@�v�m�F
' todo:���݂͗�����i�c�����j�̂ݑΉ��B���Â�s�������Ή��������H
Public Function XLOOKUP_AH(match As Variant, matchRange As Range, returnRange As Range, Optional ifNotFound As Variant = xlErrNA, Optional matchMode As Long = 0) As Variant
'Public Function XLOOKUP_AH(match As Variant, matchRange As Range, returnRange As Range, ifNotFound As Variant, matchMode As Long, searchMode As Long) As Variant :todo ���Â�searchMode�ɑΉ��������i�ォ��T���A������T���j
    Dim matchValue As Variant
    
    On Error GoTo errorProcess
    
    ' �����l�̎擾
    If TypeName(match) = "Range" Then
        ' �����l��Range�ŒP��Z���ȊO�̏ꍇ�̓G���[
        If match.Cells.Count <> 1 Then
            XLOOKUP_AH = CVErr(xlErrValue)
            Exit Function
        End If
        
        matchValue = match.value
    Else
        ' Range�ȊO�͒l�^�Ƃ��ď�������i����ȊO�̏ꍇ�̓G���[�j
        matchValue = match
    End If
    
    ' �s�����قȂ�ꍇ�̓G���[
    If matchRange.Rows.Count <> returnRange.Rows.Count Then
        XLOOKUP_AH = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' �P���̂ݑΉ��B������̏ꍇ�̓G���[�@todo:���Â���C�H
    If matchRange.Columns.Count <> 1 Or returnRange.Columns.Count <> 1 Then
        XLOOKUP_AH = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' ��r�p�̃R�[���o�b�N�֐��ݒ�
    Dim cb As ICallBack
    If matchMode >= 0 Then
        Set cb = New GE
    Else
        Set cb = New LE
    End If
    
    ' ����
    Dim cell As Range
    Dim nearest As Variant
    Dim nearestCell As Range
    nearest = 2000000000    ' �K���ɑ傫���l�������l�Ƃ���
    
    For Each cell In matchRange
        If cb.Predicate(cell.value, matchValue) Then
            ' ���߂��l�̏ꍇ�͒l�����ւ���
            If Abs(matchValue - nearest) > Abs(matchValue - cell.value) Then
                nearest = cell.value
                Set nearestCell = cell
            End If
        End If
    Next
    
    ' �������ʂ̃Z���̑��Έʒu���擾
    Dim relativeRow As Long
    relativeRow = nearestCell.Row - matchRange.Item(1).Row + 1
    
    ' �Ԃ�l��Ԃ�l�p��Range����擾
    Dim returnValue As Variant
    If nearestCell Is Nothing Then
        returnValue = ifNotFound
    Else
        returnValue = returnRange.Item(relativeRow, 1).value   ' �P���̂ݑΉ�
    End If

    ' ���S��v�̏ꍇ�Ō������ʂ�����łȂ��ꍇ��ifNotFound��Ԃ�l�ɐݒ�
    If matchMode = 0 Then
        If nearest <> matchValue Then
            returnValue = ifNotFound
        End If
    End If
    
    XLOOKUP_AH = returnValue
    
    Exit Function
    
errorProcess:
    ' �G���[����
    XLOOKUP_AH = CVErr(xlErrValue)
    
End Function

''' ���Ԃ������v�Z����
'''
''' parent�F�e�̍��ԃZ���B ex:2-3-1�ƂȂ��ė~�����ꍇ��2-3�̃Z�����w�肷��
''' delimiter�F���Ԃ̋�؂蕶���B�e���Ȃ��ꍇ�͋󕶎����w�肷��B�f�t�H���g�̓n�C�t���B ex:"-"���w�肵���ꍇ�� 2-3�B"."���w�肵���ꍇ�� 2.3�B
''' return�F���ԁB
''' attention�F�e�̍��Ԃ�ύX�����ꍇ�͎q�̍��Ԃ������Ōv�Z����邪���̍��Ԃ͎����Ōv�Z����Ȃ��ȂǁA���Ԃ����f����Ȃ��ꍇ������B
'''            ���̂��߁ACtrl + Alt + F9 �ŃZ�������v�Z�����s���邱�ƂŔ��f�����邱�Ƃ��ł���B
Public Function ITEM_NUMBER(parent As Range, Optional delimiter As String = "-") As String
    Dim FUNCTION_NAME As String
    FUNCTION_NAME = "ITEM_NUMBER"
    
    ' �����e�v�f�����Z��v�f�̐��𐔂���
    ' ���߂̌Z��Z���̔ԍ����C���N�������g������@�̂ق����J��Ԃ��͏��Ȃ��Ȃ邪�A
    ' �����Čv�Z�̂Ƃ��ɉ��̃Z������v�Z����Ă��܂��A���m�Ȕԍ����擾�ł��Ȃ����߁A
    ' �Z��v�f�����ׂăJ�E���g����������g�p
    Dim prevInputCell As Range
    Dim targetCell As Range
    Dim foundCount As Long
    Dim thisFirstArgStr As String
    foundCount = 0
    Set targetCell = Application.ThisCell
    thisFirstArgStr = ExtractFirstArgFromFormula(targetCell.Formula, FUNCTION_NAME)
    Do While True
        ' ������Ȃ��ꍇ�͌��݂̃Z�����ŏ��̎q�v�f�Ȃ̂Ń��[�v�𔲂���
        If Not FindPrevInputCell(targetCell, prevInputCell) Then
            Exit Do
        End If
        
        ' �e�v�f���قȂ�ꍇ�͌��݂̃Z�����ŏ��̎q�v�f�Ȃ̂Ń��[�v�𔲂���
        Dim prevFirstArgStr As String
        prevFirstArgStr = ExtractFirstArgFromFormula(targetCell.Formula, FUNCTION_NAME)
        If prevFirstArgStr = "" Or prevFirstArgStr <> thisFirstArgStr Then
            Exit Do
        End If
        
        If prevInputCell.Row < parent.Row Then
            Exit Do
        End If

        foundCount = foundCount + 1
        Set targetCell = prevInputCell
    Loop
    
    ' �ŏ��̎q�v�f�Ƃ��Ēl��Ԃ�
    ' todo:�A���t�@�x�b�g�����Â�Ή��B���݂͐����̂�
    If foundCount <= 0 Then
        ITEM_NUMBER = parent.Text & delimiter & 1
        Exit Function
    End If
    
    ' �Z��Z�����猻�݂̍��Ԃ��擾
    ' todo:�A���t�@�x�b�g�����Â�Ή��B���݂͐����̂�
    Dim parentItemStr As String
    If delimiter = "" Then
        ' �ŏ��̊K�w�i���e�v�f�Ȃ��j�̏ꍇ�̏���
        parentItemStr = ""
    Else
        parentItemStr = parent.Text
    End If
    
    ' �w��̃t�H�[�}�b�g�̍��Ԕ��s
    ITEM_NUMBER = parentItemStr & delimiter & CStr(foundCount + 1)
End Function

''' ���͂���Ă����̃Z����T���ĕԂ��B
'''
''' from�F�T���N�_�Z���B
''' prevInputCell�F�o�͕ϐ��B���������Z�����ݒ肳���B
''' return�F���͂���Ă���Z�������������ꍇ��True�B����ȊO��False�B
''' todo�FRange.End(xl~�j�̂悤�ɉ������ɂ��Ή�������
Private Function FindPrevInputCell(from As Range, ByRef prevInputCell) As Boolean
    Dim prevCell As Range
    Dim i As Integer
    
    If from.Row <= 1 Then
        FindPrevInputCell = False
    End If
    
    Dim hasFound As Boolean
    hasFound = False
    ' Set prevCell = from.End(direction)    �A�����ē��͂���Ă���̂�2����擾���Ă��܂��ꍇ�����邽�߃R�����g�A�E�g
    For i = from.Row - 1 To 1 Step -1
        If Not IsEmptyAH(Cells(i, from.Column)) Then
            hasFound = True
            Set prevCell = Cells(i, from.Column)
            Exit For
        End If
    Next
    
    If Not hasFound Then
        FindPrevInputCell = False
        Exit Function
    End If
    
    Set prevInputCell = prevCell
    FindPrevInputCell = True
End Function

''' �Z���̎�����ŏ��̈��������o���B
'''
''' expression�F�Z����Formula�B
''' functionName�F�Z���ɓ��͂���Ă���֐����B
''' return�F�Z���̎�����ŏ��̈�����Ԃ��B
''' todo�F���Âꎮ����͂��Ċ֐����A������Ԃ��o�[�W�����ɂ�����
Private Function ExtractFirstArgFromFormula(expression As String, functionName As String) As String
        Dim functionNameLen As Long
        ' �֐��������Ɋ܂܂�Ȃ��ꍇ
        If Not (InStr(expression, functionName) > 0) Then
            ExtractFirstArgFromFormula = ""
            Exit Function
        End If
        
        functionNameLen = Len(functionName) + 2 ' +2 is to include "=" and "(" ex: =Row(A1)
        ExtractFirstArgFromFormula = Mid(expression, functionNameLen + 1, InStr(expression, ",") - functionNameLen - 1) ' -1 is to exclude ","
End Function

''' �z��̗v�f�������߂�B
'''
''' ary�F�ΏۂƂȂ�z��B
''' return�F�z��̗v�f���B�����Ƃ��ď���������Ă��Ȃ��z����w�肵������-1�A�z��ȊO���w�肵������-100��Ԃ��B
''' src�Fhttps://qiita.com/nkojima/items/7f8299b3299226a97abb
Private Function CalcArrayLength(ary As Variant) As Long
    If (IsArray(ary)) Then
        If (IsInitialized(ary)) Then
            CalcArrayLength = UBound(ary) - LBound(ary) + 1
        Else
            CalcArrayLength = -1
        End If
    Else
        CalcArrayLength = -100
    End If

End Function

''' �w�肵���Z�����󕶎����ǂ���
'''
''' cell�F�󕶎����ǂ������肷��Z��
''' return�F�e�L�X�g���󕶎��̏ꍇ��True�B����ȊO��False�B
''' attention�F�������͂���Ă��Ă��\����͋󕶎��̏ꍇ��False���Ԃ�B
'''            �܂��󔒕��������͂���Ă���ꍇ��False���Ԃ�B
Private Function IsEmptyAH(cell As Range) As Boolean
    IsEmptyAH = LenB(cell.Text) <= 0
End Function

' �z�񂪏���������Ă��邩���`�F�b�N����B
'
' ary�F�ΏۂƂȂ�z��B
' return�F�z�񂪏������ς݂Ȃ�True�A�����łȂ����False��Ԃ��B
' src�Fhttps://qiita.com/nkojima/items/7f8299b3299226a97abb
Private Function IsInitialized(ary As Variant) As Boolean
    On Error GoTo NOT_INITIALIZED_ERROR
    Dim length As Long: length = UBound(ary)    ' ���I�z�񂪏���������Ă��Ȃ���΁A�����ŃG���[����������B
    IsInitialized = True
    Exit Function

' �z�񂪏���������Ă��Ȃ��ꍇ�͂����ɔ�΂����B
NOT_INITIALIZED_ERROR:
    IsInitialized = False
End Function
