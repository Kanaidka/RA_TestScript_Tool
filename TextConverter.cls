VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPH As PRSPH
Private mWb As Workbook

'プロパティ　Workbook(Set)
'別表データ検索の為、読込Workbookの保持
Public Property Set Workbook(ByRef iWb As Workbook)
    Set mWb = iWb
End Property

'関数　Convert
'引数　iText　変換するテキスト
'引数　iPH　PHデータ
'戻り値　変換後の文字列
'改行入りテキストの変換処理
'おもに改行分割と条件式のチェック
Public Function Convert(ByVal iText As String, ByRef iPH As PRSPH) As String

    Dim items() As String
    Dim lines() As String
    Dim ret As String
    Dim i As Long
    Dim j As Long
    Dim condition As Boolean
    
    'PHデータがなければ変換なしで戻す
    If iPH Is Nothing Then
        Convert = iText
        Exit Function
    End If
    
    '各関数で使用するのでクラス内変数に格納
    Set mPH = iPH
    
    ret = ""
    condition = True  '行を描画するかを初期化
    '改行コードを統一し、行ごとに分割
    lines = Split(Replace(iText, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(lines)
        '条件式をチェックし、表示対象なら文字列を結合していく
        If CheckCondition(lines(i), condition) Then
            ret = ret & ConvertLine(lines(i)) & vbCrLf
        End If
    Next
    
    '最後についてしまう改行を削除
    If Right(ret, Len(vbCrLf)) = vbCrLf Then
        ret = Left(ret, Len(ret) - Len(vbCrLf))
    End If
    
    Convert = ret

End Function

'関数　ConvertLine
'引数　iText　変換するテキスト
'戻り値　変換後の文字列
'1行の内容を分析し変換式を見つける
Private Function ConvertLine(ByVal iText As String) As String

    Dim items() As String
    Dim i As Long
    Dim ret As String
    
    '%で分割し、偶数要素を変換式分析する
    items = Split(iText, "%")
    For i = 0 To UBound(items)
        If i Mod 2 = 0 Then
            '奇数要素は変換なしで追加
            ret = ret & items(i)
        Else
            '変換式を分析し、値を戻す
            ret = ret & GetConvertText(items(i))
        End If
    Next

    ConvertLine = ret

End Function

'関数　GetConvertText
'引数　iQuery　変換式
'戻り値　変換後の文字列
'変換式を特定文字に変換する
Private Function GetConvertText(ByVal iQuery As String) As String
    
    Dim ret As String
    Dim items() As String
    ret = ""
    'クエリ要素で分割
    items = Split(Trim(iQuery), ",")
    
    '第二要素までない場合は空白を返す
    If UBound(items) < 1 Then
        GetConvertText = ret
        Exit Function
    End If
    
    '第二要素名で処理分岐する
    '変換処理を増やしたい場合はここを増やす
    Select Case UCase(items(1))
        Case "SHEET"
            ret = GetSheet(items)
        Case "LINE"
            ret = GetLine(items)
        Case "KEY"
            ret = GetKeyValue(items)
        Case "PREFIX"
            ret = GetPrefix(items)
        Case "BETWEEN"
            ret = GetBetween(items)
        Case "ALL"
            ret = GetPHText(items(0))
    End Select
    
    GetConvertText = ret
    
End Function

'関数　GetPHText
'引数　iText　取得判別文字列
'戻り値　PHデータ
'文字列を基に対応するPHデータを返す
Private Function GetPHText(ByVal iText As String) As String
    
    Dim ret As String
    
    '取得したい要素が増えた場合はここを増やす
    Select Case UCase(iText)
        Case "PR"
            ret = mPH.PR.text
        Case "UP"
            ret = mPH.UP.text
        Case "OP"
            ret = mPH.OP.text
        Case "OP_COMMENT"
            ret = mPH.OP.Comment
        Case "OP_プロセスシーケンス"
            ret = mPH.OP.ProcessSequence
        Case "ID"
            ret = mPH.id
        Case "PH"
            ret = mPH.text
        Case "COMMENT"
            ret = mPH.Comment
        Case "レシピパラメータ"
            ret = mPH.RecipeParameter
        Case "マテリアル"
            ret = mPH.Material
        Case "機器"
            ret = mPH.Equipment
        Case "場所"
            ret = mPH.Place
        Case "GMP署名"
            ret = mPH.GMP
        Case "プロセスシーケンス"
            ret = mPH.ProcessSequence
    End Select
    
    GetPHText = Trim(ret)

End Function

'関数　CheckCondition
'引数　iText　文字列
'引数　iCondition　以降の行を表示するか
'戻り値　該当行を表示するか
'条件式の内容と現在の状態から該当行と以降の行の表示/非表示を返す
Private Function CheckCondition(ByVal iText As String, ByRef iCondition As Boolean) As Long
    
    Dim text As String
    text = Trim(iText)
    
    '空白行の場合
    If text = "" Then
        iCondition = True  '以降の行を追加するか
        CheckCondition = False 'その行を追加するか
        Exit Function
    End If
    '条件式行の場合
    If Left(text, 1) = "#" And Right(text, 1) = "#" Then
        iCondition = CheckConditionItem(text) '以降の行を追加するか
        CheckCondition = False 'その行を追加するか
        Exit Function
    End If
    '上記条件に当てはまらない行は上の行の引継ぎ
    CheckCondition = iCondition

End Function

'関数　CheckConditionItem
'引数　iText　条件判別文字列
'戻り値　条件によって表示するかしないか
'式の内容を確認し、条件に当てはまるかを返す
Private Function CheckConditionItem(ByVal iText As String) As Boolean

    Dim text As String
    Dim leftValue As String
    Dim rightValue As String
    Dim ret As Boolean
    Dim items() As String
    Dim i As Long
    
    ret = True
    
    text = Trim(iText)
    
    '前方の#を除去
    Do While Len(text) > 1 And Left(text, 1) = "#"
        text = mID(text, 2)
    Loop
    
    '後方の#を除去
    Do While Len(text) > 1 And Right(text, 1) = "#"
        text = Left(text, Len(text) - 1)
    Loop
    
    '変換式がある場合は変換後にカンマ区切りで分割
    items = Split(Trim(ConvertLine(text)), ",")
    
    '要素が3つない場合は無条件でOK
    If UBound(items) < 2 Then
        CheckConditionItem = True
        Exit Function
    End If
    
    i = 0
    
    '3つの要素ごとにチェックを行い、AND条件で繋いでいく
    Do While UBound(items) >= i + 2
        Select Case UCase(items(i + 1))
            Case "="
                ret = ret And (items(i) = items(i + 2))
            Case "<>"
                ret = ret And (items(i) <> items(i + 2))
            Case "EXIST"
                ret = ret And (InStr(items(i), items(i + 2)) > 0)
            Case "!EXIST"
                ret = ret And (Not InStr(items(i), items(i + 2)) > 0)
        End Select
        i = i + 3
    Loop
 
    'チェック結果を返す
    CheckConditionItem = ret

End Function

'関数　GetLine
'引数　iArgs　変換式内容
'戻り値　変換文字列
'テキストを行分割し、特定行の値を返す
Private Function GetLine(ByRef iArgs() As String) As String

    Dim i As Long
    Dim lines() As String
    Dim s As String
    Dim fromIndex As Long
    Dim toIndex As Long
        
    If UBound(iArgs) < 2 Then
        GetLine = ""
        Exit Function
    End If
    
    lines = Split(GetPHText(iArgs(0)), vbLf)
    
    s = Replace(iArgs(2), "～", "")
    
    i = 0
    If IsNumeric(s) Then
        i = CLng(s) - 1
    End If
    
    fromIndex = i
    toIndex = i
    
    If Left(iArgs(2), 1) = "～" Then
        fromIndex = 0
    End If
    
    If Right(iArgs(2), 1) = "～" Or i > UBound(lines) Then
        toIndex = UBound(lines)
    End If
    
    s = ""
    For i = fromIndex To toIndex
        s = s & lines(i) & vbCrLf
    Next
    If Right(s, Len(vbCrLf)) = vbCrLf Then
        s = Left(s, Len(s) - Len(vbCrLf))
    End If
    
    GetLine = s

End Function

'関数　GetKeyValue
'引数　iArgs　変換式内容
'戻り値　変換文字列
'テキストから「〇〇：」を探し、その右辺を返す
Private Function GetKeyValue(ByRef iArgs() As String) As String

    Dim lines() As String
    Dim items() As String
    Dim delimiter As String
    Dim i As Long
    Dim ret As String
    Dim key As String
    
    
    If UBound(iArgs) < 2 Then
        GetKeyValue = ""
        Exit Function
    End If

    key = iArgs(2)
    ret = ""
    
    If key = "" Then
        GetKeyValue = ""
        Exit Function
    End If
    
    lines = Split(GetPHText(iArgs(0)), vbLf)
    
    For i = 0 To UBound(lines)
        delimiter = ""
        If InStr(lines(i), ":") > 0 Then
            delimiter = ":"
        ElseIf InStr(lines(i), "：") > 0 Then
            delimiter = "："
        End If
        If delimiter <> "" Then
            items = Split(lines(i), delimiter)
            If Trim(items(0)) = key Then
                ret = Trim(items(1))
                Exit For
            End If
        End If
    Next
    
    GetKeyValue = ret

End Function

'関数　GetPrefix
'引数　iArgs　変換式内容
'戻り値　変換文字列
'テキストを行分割し、特定文字列から始まる行の、特定文字列以降の文字列を返す
Private Function GetPrefix(ByRef iArgs() As String) As String

    Dim i As Long
    Dim lines() As String
    Dim key As String
    
    If UBound(iArgs) < 2 Then
        GetPrefix = ""
        Exit Function
    End If
    key = iArgs(2)
    
    lines = Split(GetPHText(iArgs(0)), vbLf)
    For i = 0 To UBound(lines)
        If Left(lines(i), Len(key)) = key Then
            GetPrefix = mID(lines(i), Len(key) + 1)
            Exit For
        End If
    Next
    
End Function

'関数　GetBetween
'引数　iArgs　変換式内容
'戻り値　変換文字列
'テキストの特定文字列Aと特定文字列Bの間の文字列を返す
Private Function GetBetween(ByRef iArgs() As String) As String

    Dim i As Long
    Dim s As String
    Dim item1() As String
    Dim item2() As String
    Dim before As String
    Dim after As String
    
    If UBound(iArgs) < 2 Then
        GetBetween = ""
        Exit Function
    End If
    before = iArgs(2)
    
    If UBound(iArgs) >= 3 Then
        after = iArgs(3)
    End If
    
    '後方検索文字列が\nの場合は改行までを検索する
    If UCase(after) = "\N" Then
        after = vbLf
    End If
    
    item1 = Split(GetPHText(iArgs(0)), before)
    
    'aaa[before]bbb[after]ccc
    '↓
    'aaa
    'bbb[after]ccc
    If UBound(item1) < 1 Then
        GetBetween = ""
        Exit Function
    Else
        s = item1(1)
    End If
    
    'bbb[after]ccc
    '↓
    'bbb
    'ccc
    item2 = Split(s, after)
    
    GetBetween = item2(0)

End Function

'関数　GetSheet
'引数　iArgs　変換式内容
'戻り値　変換文字列
'特定シート（別表）の同じID行を探し、指定した列名の値を返す
'列名と値を追加で指定し、条件の絞り込みを可能とする
'%シート名,SHEET,取得列名,チェック列名1,チェック値1,チェック列名2,チェック値2,…%
Private Function GetSheet(ByRef iArgs() As String) As String

    
    If UBound(iArgs) < 2 Then
        GetSheet = ""
        Exit Function
    End If
    
    'ワークブックが指定されていない
    If mWb Is Nothing Then
        GetSheet = ""
        Exit Function
    End If
    
    '指定したシートがない場合
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = mWb.Sheets(iArgs(0))
    On Error GoTo 0
    If ws Is Nothing Then
        GetSheet = ""
        Exit Function
    End If
    
    Dim colValue As New Column
    Dim col As New Column
    Dim i As Long
    Dim j As Long
    Dim find As Boolean
    Dim ret As String
    
    '取得先の列番号がない場合
    Call colValue.Init(iArgs(2))
    Call colValue.SetIndexPrefix(ws, 2)
    If colValue.Index = 0 Then
        GetSheet = ""
        Exit Function
    End If
    'IDの列番号がない場合
    Call col.Init("ID")
    Call col.SetIndexPrefix(ws, 2)
    If col.Index = 0 Then
        GetSheet = ""
        Exit Function
    End If
        
    Dim colIndex() As Long
    Dim text() As String
    '条件式の分だけ配列を先に拡張
    i = Fix((UBound(iArgs) - 2) / 2)
    ReDim colIndex(i)
    
    '最初の条件はIDとする
    colIndex(0) = col.Index
    ReDim text(i)
    text(0) = mPH.id

    'Index3,4、5,6、7,8・・・のように列名と値のセットを取得し、条件に追加する
    For i = 1 To UBound(colIndex)
        Call col.Init(iArgs(i * 2 + 1))
        '列名から列番号を探す
        Call col.SetIndexPrefix(ws, 2)
        colIndex(i) = col.Index
        text(i) = iArgs(i * 2 + 2)
    Next
    
    '対象シートのID列がなくなる行までを検索
    For i = 1 To ws.Cells(ws.Rows.Count, colIndex(0)).End(xlUp).row
        find = True
        '一つでも条件に当てはまらない場合はフラグをOFFしてスキップ
        For j = 0 To UBound(colIndex)
            If Trim(ws.Cells(i, colIndex(j)).Value) <> Trim(text(j)) Then
                find = False
                Exit For
            End If
        Next
        'すべての条件に当てはまる行を見つけた場合、該当行を返す
        If find Then
            ret = Trim(ws.Cells(i, colValue.Index).Value)
            Exit For
        End If
    Next
    
    GetSheet = ret

End Function
