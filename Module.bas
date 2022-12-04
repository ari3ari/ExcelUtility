Attribute VB_Name = "Module"
Option Explicit

' todo:文字列に対応しているかどうかの確認。現在は数値のみ→一応対応している？https://excel-ubara.com/excel5/EXCEL846.html　要確認
' todo:現在は列方向（縦方向）のみ対応。いづれ行方向も対応したい？
Public Function XLOOKUP_AH(match As Variant, matchRange As Range, returnRange As Range, Optional ifNotFound As Variant = xlErrNA, Optional matchMode As Long = 0) As Variant
'Public Function XLOOKUP_AH(match As Variant, matchRange As Range, returnRange As Range, ifNotFound As Variant, matchMode As Long, searchMode As Long) As Variant :todo いづれsearchModeに対応したい（上から探す、下から探す）
    Dim matchValue As Variant
    
    On Error GoTo errorProcess
    
    ' 検索値の取得
    If TypeName(match) = "Range" Then
        ' 検索値がRangeで単一セル以外の場合はエラー
        If match.Cells.Count <> 1 Then
            XLOOKUP_AH = CVErr(xlErrValue)
            Exit Function
        End If
        
        matchValue = match.value
    Else
        ' Range以外は値型として処理する（それ以外の場合はエラー）
        matchValue = match
    End If
    
    ' 行数が異なる場合はエラー
    If matchRange.Rows.Count <> returnRange.Rows.Count Then
        XLOOKUP_AH = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' 単一列のみ対応。複数列の場合はエラー　todo:いづれ改修？
    If matchRange.Columns.Count <> 1 Or returnRange.Columns.Count <> 1 Then
        XLOOKUP_AH = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' 比較用のコールバック関数設定
    Dim cb As ICallBack
    If matchMode >= 0 Then
        Set cb = New GE
    Else
        Set cb = New LE
    End If
    
    ' 検索
    Dim cell As Range
    Dim nearest As Variant
    Dim nearestCell As Range
    nearest = 2000000000    ' 適当に大きい値を初期値とする
    
    For Each cell In matchRange
        If cb.Predicate(cell.value, matchValue) Then
            ' より近い値の場合は値を入れ替える
            If Abs(matchValue - nearest) > Abs(matchValue - cell.value) Then
                nearest = cell.value
                Set nearestCell = cell
            End If
        End If
    Next
    
    ' 検索結果のセルの相対位置を取得
    Dim relativeRow As Long
    relativeRow = nearestCell.Row - matchRange.Item(1).Row + 1
    
    ' 返り値を返り値用のRangeから取得
    Dim returnValue As Variant
    If nearestCell Is Nothing Then
        returnValue = ifNotFound
    Else
        returnValue = returnRange.Item(relativeRow, 1).value   ' 単一列のみ対応
    End If

    ' 完全一致の場合で検索結果が同一でない場合はifNotFoundを返り値に設定
    If matchMode = 0 Then
        If nearest <> matchValue Then
            returnValue = ifNotFound
        End If
    End If
    
    XLOOKUP_AH = returnValue
    
    Exit Function
    
errorProcess:
    ' エラー処理
    XLOOKUP_AH = CVErr(xlErrValue)
    
End Function

''' 項番を自動計算する
'''
''' parent：親の項番セル。 ex:2-3-1となって欲しい場合は2-3のセルを指定する
''' delimiter：項番の区切り文字。親がない場合は空文字を指定する。デフォルトはハイフン。 ex:"-"を指定した場合は 2-3。"."を指定した場合は 2.3。
''' return：項番。
''' attention：親の項番を変更した場合は子の項番も自動で計算されるが孫の項番は自動で計算されないなど、項番が反映されない場合がある。
'''            そのため、Ctrl + Alt + F9 でセル自動計算を実行することで反映させることができる。
Public Function ITEM_NUMBER(parent As Range, Optional delimiter As String = "-") As String
    Dim FUNCTION_NAME As String
    FUNCTION_NAME = "ITEM_NUMBER"
    
    ' 同じ親要素を持つ兄弟要素の数を数える
    ' 直近の兄弟セルの番号をインクリメントする方法のほうが繰り返しは少なくなるが、
    ' 自動再計算のときに下のセルから計算されてしまい、正確な番号を取得できないため、
    ' 兄弟要素をすべてカウントする方式を使用
    Dim prevInputCell As Range
    Dim targetCell As Range
    Dim foundCount As Long
    Dim thisFirstArgStr As String
    foundCount = 0
    Set targetCell = Application.ThisCell
    thisFirstArgStr = ExtractFirstArgFromFormula(targetCell.Formula, FUNCTION_NAME)
    Do While True
        ' 見つからない場合は現在のセルが最初の子要素なのでループを抜ける
        If Not FindPrevInputCell(targetCell, prevInputCell) Then
            Exit Do
        End If
        
        ' 親要素が異なる場合は現在のセルが最初の子要素なのでループを抜ける
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
    
    ' 最初の子要素として値を返す
    ' todo:アルファベットもいづれ対応。現在は数字のみ
    If foundCount <= 0 Then
        ITEM_NUMBER = parent.Text & delimiter & 1
        Exit Function
    End If
    
    ' 兄弟セルから現在の項番を取得
    ' todo:アルファベットもいづれ対応。現在は数字のみ
    Dim parentItemStr As String
    If delimiter = "" Then
        ' 最初の階層（≒親要素なし）の場合の処理
        parentItemStr = ""
    Else
        parentItemStr = parent.Text
    End If
    
    ' 指定のフォーマットの項番発行
    ITEM_NUMBER = parentItemStr & delimiter & CStr(foundCount + 1)
End Function

''' 入力されている上のセルを探して返す。
'''
''' from：探す起点セル。
''' prevInputCell：出力変数。見つかったセルが設定される。
''' return：入力されているセルが見つかった場合はTrue。それ以外はFalse。
''' todo：Range.End(xl~）のように横方向にも対応したい
Private Function FindPrevInputCell(from As Range, ByRef prevInputCell) As Boolean
    Dim prevCell As Range
    Dim i As Integer
    
    If from.Row <= 1 Then
        FindPrevInputCell = False
    End If
    
    Dim hasFound As Boolean
    hasFound = False
    ' Set prevCell = from.End(direction)    連続して入力されているのに2つ上を取得してしまう場合があるためコメントアウト
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

''' セルの式から最初の引数を取り出す。
'''
''' expression：セルのFormula。
''' functionName：セルに入力されている関数名。
''' return：セルの式から最初の引数を返す。
''' todo：いづれ式を解析して関数名、引数を返すバージョンにしたい
Private Function ExtractFirstArgFromFormula(expression As String, functionName As String) As String
        Dim functionNameLen As Long
        ' 関数名が式に含まれない場合
        If Not (InStr(expression, functionName) > 0) Then
            ExtractFirstArgFromFormula = ""
            Exit Function
        End If
        
        functionNameLen = Len(functionName) + 2 ' +2 is to include "=" and "(" ex: =Row(A1)
        ExtractFirstArgFromFormula = Mid(expression, functionNameLen + 1, InStr(expression, ",") - functionNameLen - 1) ' -1 is to exclude ","
End Function

''' 配列の要素数を求める。
'''
''' ary：対象となる配列。
''' return：配列の要素数。引数として初期化されていない配列を指定した時は-1、配列以外を指定した時は-100を返す。
''' src：https://qiita.com/nkojima/items/7f8299b3299226a97abb
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

''' 指定したセルが空文字かどうか
'''
''' cell：空文字かどうか判定するセル
''' return：テキストが空文字の場合はTrue。それ以外はFalse。
''' attention：式が入力されていても表示上は空文字の場合はFalseが返る。
'''            また空白文字が入力されている場合もFalseが返る。
Private Function IsEmptyAH(cell As Range) As Boolean
    IsEmptyAH = LenB(cell.Text) <= 0
End Function

' 配列が初期化されているかをチェックする。
'
' ary：対象となる配列。
' return：配列が初期化済みならTrue、そうでなければFalseを返す。
' src：https://qiita.com/nkojima/items/7f8299b3299226a97abb
Private Function IsInitialized(ary As Variant) As Boolean
    On Error GoTo NOT_INITIALIZED_ERROR
    Dim length As Long: length = UBound(ary)    ' 動的配列が初期化されていなければ、ここでエラーが発生する。
    IsInitialized = True
    Exit Function

' 配列が初期化されていない場合はここに飛ばされる。
NOT_INITIALIZED_ERROR:
    IsInitialized = False
End Function
