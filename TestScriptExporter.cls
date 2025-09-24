VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestScriptExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'クラス　TestScript
'各データを集めてテストスクリプトデータの出力を管理します

Private Const ROW_INDEX_COLUMN_HEADER As Long = 10
Private Const ROW_INDEX_DATA As Long = 11

Private Const COLUMN_HEADER_TARGET As String = "対象"
Private Const COLUMN_HEADER_FILEPATH As String = "ファイルパス"
Private Const COLUMN_HEADER_IMPORT_SHEET As String = "取込シート名"
Private Const COLUMN_HEADER_OUTPUT_SHEET As String = "出力シート名"
Private Const COLUMN_HEADER_RESULT As String = "結果"

Private Const SHEET_NAME_FORMAT As String = "Format"
Private Const SHEET_NAME_MATRIX As String = "Matrix"
Private Const SHEET_NAME_TEMPLATE As String = "Template"

Private Const COPY_OTHER_SHEET_SUFFIX As String = "別表"

Private mColumnTarget As New Column
Private mColumnFilePath As New Column
Private mColumnImportSheet As New Column
Private mColumnOutputSheet As New Column
Private mColumnResult As New Column

Private mTestScript As New TestScript

'イベント　クラスの初期化
'クラスの初期化を行う
Private Sub Class_Initialize()
    
    Call mColumnTarget.Init(COLUMN_HEADER_TARGET)
    Call mColumnFilePath.Init(COLUMN_HEADER_FILEPATH)
    Call mColumnImportSheet.Init(COLUMN_HEADER_IMPORT_SHEET)
    Call mColumnOutputSheet.Init(COLUMN_HEADER_OUTPUT_SHEET)
    Call mColumnResult.Init(COLUMN_HEADER_RESULT)
    
End Sub

'関数　Export
'引数　iWs　出力対象ワークシート
'Matrixデータの読み込み
Public Sub Export(ByRef iWs As Worksheet)

    Dim app As Application
    Dim wb As Workbook
    Dim path  As String
    Dim inSheet As String
    Dim outSheet As String
    Dim rowIndex As Long
    Dim rowIndexLast As Long
    Dim dlg As New frmOutput
    Dim i As Long
    Dim target As Boolean
    Dim errMsg As String
    Dim errFlg As Boolean
    Dim fileLoop As Boolean
    Dim displayAlerts As Boolean
    Dim screenUpdating As Boolean
    
    fileLoop = False
    
On Error GoTo ERR_LABEL
    
    '進捗ダイアログを表示
    Call dlg.Show(False)

    Call dlg.AddLog("事前処理開始 ------------------------------")
        
    Call dlg.AddLog("Excel 準備中")
    Set app = Application
    displayAlerts = app.displayAlerts  'ダイアログ表示状態保存
    screenUpdating = app.screenUpdating '再描画状態の保存
    app.displayAlerts = False 'ダイアログ表示停止
    app.screenUpdating = False '描画停止
    Call dlg.AddLog("Excel 準備完了")
    
    '出力内容の列番号の取得
    Call dlg.AddLog("出力対象 読込中")
    Call mColumnTarget.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call mColumnFilePath.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call mColumnImportSheet.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call mColumnOutputSheet.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call mColumnResult.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    
    
    '1つでも列番号取得に失敗した場合はエラーとする
    Select Case 0
        Case mColumnTarget.Index
            errMsg = mColumnTarget.Name & "列が存在しない為、出力できません。"
            GoTo ERR_LABEL
        Case mColumnFilePath.Index
            errMsg = mColumnFilePath.Name & "列が存在しない為、出力できません。"
            GoTo ERR_LABEL
        Case mColumnTarget.Index
            errMsg = mColumnTarget.Name & "列が存在しない為、出力できません。"
            GoTo ERR_LABEL
        Case mColumnImportSheet.Index
            errMsg = mColumnImportSheet.Name & "列が存在しない為、出力できません。"
            GoTo ERR_LABEL
        Case mColumnOutputSheet.Index
            errMsg = mColumnOutputSheet.Name & "列が存在しない為、出力できません。"
            GoTo ERR_LABEL
        Case mColumnResult.Index
            errMsg = mColumnResult.Name & "列が存在しない為、出力できません。"
            GoTo ERR_LABEL
    End Select
    
    '出力対象の最終行取得
    rowIndexLast = iWs.Cells(iWs.Rows.Count, mColumnFilePath.Index).End(xlUp).row
    'データなしの場合
    If rowIndexLast < 1 Then
        errMsg = mColumnResult.Name & "出力データが存在しない為、出力できません。"
        GoTo ERR_LABEL
    End If
        
    '進捗率の設定
    dlg.NowValue = 0 '現在値
    dlg.MaxValue = rowIndexLast - ROW_INDEX_DATA + 1  '出力ファイル数
        
    Call dlg.AddLog("出力対象 読込完了")

    'FORMATデータ読込
    errMsg = "FORMATデータ の読み込みに失敗しました。"
    Call dlg.AddLog("FORMATデータ 読込中")
    Call mTestScript.LoadFormat(ThisWorkbook.Sheets(SHEET_NAME_FORMAT))
    Call dlg.AddLog("FORMATデータ 読込完了")
    
    'MATRIXデータ読込
    errMsg = "MATRIXデータ の読み込みに失敗しました。"
    Call dlg.AddLog("MATRIXデータ 読込中")
    Call mTestScript.LoadMatrix(ThisWorkbook.Sheets(SHEET_NAME_MATRIX))
    Call dlg.AddLog("MATRIXデータ 読込完了")
    
    'TEMPLATEデータ読込
    errMsg = "TEMPLATEデータ の読み込みに失敗しました。"
    Call dlg.AddLog("TEMPLATEデータ 読込中")
    Call mTestScript.LoadTemplate(ThisWorkbook.Sheets(SHEET_NAME_TEMPLATE))
    Call dlg.AddLog("TEMPLATEデータ 読込完了")
        
    '事前処理のほとんどがエラーなく終わったことを確認してから前回の結果を削除
    Call dlg.AddLog("前回結果 削除中")
    For rowIndex = ROW_INDEX_DATA To rowIndexLast
        iWs.Cells(rowIndex, mColumnResult.Index).Value = ""
    Next
    Call dlg.AddLog("前回結果 削除完了")
    Call dlg.AddLog("事前処理完了 ------------------------------")

    'テストスクリプト出力ループ開始
    fileLoop = True
    For rowIndex = ROW_INDEX_DATA To rowIndexLast
        
        target = iWs.Cells(rowIndex, mColumnTarget.Index).Value <> ""
        path = iWs.Cells(rowIndex, mColumnFilePath.Index).Value
        inSheet = iWs.Cells(rowIndex, mColumnImportSheet.Index).Value
        outSheet = iWs.Cells(rowIndex, mColumnOutputSheet.Index).Value
        dlg.Header = path
        
        Call dlg.AddLog("")
        Call dlg.AddLog("作業開始 ------------------------------")
        Call dlg.AddLog("ファイル：" & path)
        
        If target Then
            errFlg = False
            
            'Excelファイルを展開する
            errMsg = "ファイルの読み込みに失敗しました。"
            Call dlg.AddLog("ファイル 展開中")
            Set wb = app.Workbooks.Open(path)
            
            Call dlg.AddLog("ファイル 展開完了")
            
            If Not errFlg Then
                Call dlg.AddLog("シート：" & inSheet)
                errMsg = "シートの読込に失敗しました。"
                Call dlg.AddLog("シート 読込中")
                'PRSファイルの読込
                Call mTestScript.LoadPRS(wb.Sheets(inSheet))
                Call dlg.AddLog("シート 読込完了")
            End If
            
            '別表と作シートをコピーする
            errMsg = "別表コピーに失敗しました。"
            Call dlg.AddLog("別表コピー中")
            For i = 1 To wb.Sheets.Count
                If Right(wb.Sheets(i).Name, Len(COPY_OTHER_SHEET_SUFFIX)) = COPY_OTHER_SHEET_SUFFIX Then
                    Call dlg.AddLog("コピーシート：" & wb.Sheets(i).Name)
                    Call wb.Sheets(i).Copy(ThisWorkbook.Sheets(1))
                End If
            Next
            Call dlg.AddLog("別表コピー完了")
            
            If Not errFlg Then
                Call dlg.AddLog("テストスクリプト：" & outSheet)
                errMsg = "シートの出力に失敗しました。"
                Call dlg.AddLog("テストスクリプト 出力中")
                'テストスクリプトの出力
                Call mTestScript.Create(outSheet)
                Call dlg.AddLog("テストスクリプト 出力完了")
            End If
        
        
            If Not errFlg And Not wb Is Nothing Then
                'エラーじゃなければExcelを閉じる
                Call wb.Close
                iWs.Cells(rowIndex, mColumnResult.Index).Value = "完了"
            End If
            
        Else
        
            Call dlg.AddLog("対象外のためスキップします。")
            iWs.Cells(rowIndex, mColumnResult.Index).Value = "スキップ"

        End If
        Call dlg.AddLog("作業完了 ------------------------------")
        
        dlg.NowValue = dlg.NowValue + 1
    Next
    
    Call dlg.AddLog("終了処理開始 ------------------------------")
    Call dlg.AddLog("終了処理 開始")
    
    app.screenUpdating = screenUpdating
    app.displayAlerts = displayAlerts
    Set app = Nothing
    
    Call dlg.AddLog("終了処理 終了")
    Call dlg.AddLog("終了処理終了 ------------------------------")
        
Exit Sub
ERR_LABEL:
    
    'エラーダイアログがあればエラー内容を表示
    If Not dlg Is Nothing And errMsg <> "" Then
        Call dlg.AddLog(errMsg)
        Call dlg.AddLog(Err.Description)
    End If
    
    'テストスクリプト出力ループ中の場合
    If fileLoop Then
        
        'エラーフラグを立てる
        errFlg = True
        'ワークブックが開かれていたら閉じる
        If Not wb Is Nothing Then
            Call wb.Close
        End If
        
        '描画を止める
        Application.screenUpdating = False
        '一覧にエラーを記載
        iWs.Cells(rowIndex, mColumnResult.Index).Value = "エラー"
        'エラーになった処理以降を続行する。
        Resume Next
    
    End If

    'Excelの各状態を実行前の状態に戻す
    If Not app Is Nothing Then
        app.screenUpdating = screenUpdating
        app.displayAlerts = displayAlerts
        Set app = Nothing
    End If
    
End Sub

