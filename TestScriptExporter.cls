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

Private mColumnTarget As New Column
Private mColumnFilePath As New Column
Private mColumnImportSheet As New Column
Private mColumnOutputSheet As New Column
Private mColumnResult As New Column

Private mTestScript As New TestScript

Private Sub Class_Initialize()
    
    Call mColumnTarget.Init(COLUMN_HEADER_TARGET)
    Call mColumnFilePath.Init(COLUMN_HEADER_FILEPATH)
    Call mColumnImportSheet.Init(COLUMN_HEADER_IMPORT_SHEET)
    Call mColumnOutputSheet.Init(COLUMN_HEADER_OUTPUT_SHEET)
    Call mColumnResult.Init(COLUMN_HEADER_RESULT)
    
End Sub

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
    
    fileLoop = False
    
On Error GoTo ERR_LABEL
    
    Call dlg.Show(False)
    Application.ScreenUpdating = False

    Call dlg.AddLog("事前処理開始 ------------------------------")
        
    Call dlg.AddLog("Excel 準備中")
    Set app = New Application
    app.Visible = False
    app.DisplayAlerts = False
    app.ScreenUpdating = False
    Call dlg.AddLog("Excel 準備完了")
    
    Call dlg.AddLog("出力対象 読込中")
    Call mColumnTarget.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call mColumnFilePath.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call mColumnImportSheet.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call mColumnOutputSheet.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    Call mColumnResult.SetIndexPrefix(iWs, ROW_INDEX_COLUMN_HEADER)
    
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
    
    
    rowIndexLast = iWs.Cells(iWs.Rows.Count, mColumnFilePath.Index).End(xlUp).row
    If rowIndexLast < 1 Then
        errMsg = mColumnResult.Name & "出力データが存在しない為、出力できません。"
        GoTo ERR_LABEL
    End If
        
    dlg.NowValue = 0
    dlg.MaxValue = rowIndexLast - ROW_INDEX_DATA + 1
        
    Call dlg.AddLog("出力対象 読込完了")
    

    errMsg = "FORMATデータ の読み込みに失敗しました。"
    Call dlg.AddLog("FORMATデータ 読込中")
    Call mTestScript.LoadFormat(ThisWorkbook.Sheets(SHEET_NAME_FORMAT))
    Call dlg.AddLog("FORMATデータ 読込完了")
    
    errMsg = "MATRIXデータ の読み込みに失敗しました。"
    Call dlg.AddLog("MATRIXデータ 読込中")
    Call mTestScript.LoadMatrix(ThisWorkbook.Sheets(SHEET_NAME_MATRIX))
    Call dlg.AddLog("MATRIXデータ 読込完了")
    
    errMsg = "TEMPLATEデータ の読み込みに失敗しました。"
    Call dlg.AddLog("TEMPLATEデータ 読込中")
    Call mTestScript.LoadTemplate(ThisWorkbook.Sheets(SHEET_NAME_TEMPLATE))
    Call dlg.AddLog("TEMPLATEデータ 読込完了")
        
    Call dlg.AddLog("事前処理完了 ------------------------------")
    
    fileLoop = True
    For rowIndex = ROW_INDEX_DATA To rowIndexLast
        
        target = iWs.Cells(rowIndex, mColumnTarget.Index).value <> ""
        path = iWs.Cells(rowIndex, mColumnFilePath.Index).value
        inSheet = iWs.Cells(rowIndex, mColumnImportSheet.Index).value
        outSheet = iWs.Cells(rowIndex, mColumnOutputSheet.Index).value
        dlg.Header = path
        
        Call dlg.AddLog("")
        Call dlg.AddLog("作業開始 ------------------------------")
        Call dlg.AddLog("ファイル：" & path)
        
        If target Then
            errFlg = False
            
            errMsg = "ファイルの読み込みに失敗しました。"
            Call dlg.AddLog("ファイル 展開中")
            Set wb = app.Workbooks.Open(path)
            Call dlg.AddLog("ファイル 展開完了")
            
            If Not errFlg Then
                Call dlg.AddLog("シート：" & inSheet)
                
                errMsg = "シートの読込に失敗しました。"
                Call dlg.AddLog("シート 読込中")
                Call mTestScript.LoadPRS(wb.Sheets(inSheet))
                Call dlg.AddLog("シート 読込完了")
            End If
            
            If Not errFlg Then
                Call dlg.AddLog("テストスクリプト：" & outSheet)
                
                errMsg = "シートの出力に失敗しました。"
                Call dlg.AddLog("テストスクリプト 出力中")
                Call mTestScript.Create(outSheet)
                Call dlg.AddLog("テストスクリプト 出力完了")
            End If
        
            If Not errFlg Then
                Call wb.Close
            End If
        Else
        
            Call dlg.AddLog("対象外のためスキップします。")
        
        End If
        Call dlg.AddLog("作業完了 ------------------------------")
        
        iWs.Cells(rowIndex, mColumnResult.Index).value = "完了"
        
        
        dlg.NowValue = dlg.NowValue + 1
    Next
    
    Call dlg.AddLog("終了処理開始 ------------------------------")
    Call dlg.AddLog("終了処理 開始")
    
    app.ScreenUpdating = True
    app.DisplayAlerts = True
    Call app.Quit
    Set app = Nothing
    
    Call dlg.AddLog("終了処理 終了")
    Call dlg.AddLog("終了処理終了 ------------------------------")
    
    Application.ScreenUpdating = True
    
Exit Sub
ERR_LABEL:
    
    If Not dlg Is Nothing And errMsg <> "" Then
        Call dlg.AddLog(errMsg)
        Call dlg.AddLog(Err.Description)
    End If
    
    If fileLoop Then
    
        errFlg = True
        If Not wb Is Nothing Then
            Call wb.Close
        End If
        
        Application.ScreenUpdating = False
        
        Resume Next
    
    End If

    If Not app Is Nothing Then
        app.ScreenUpdating = True
        app.DisplayAlerts = True
        Call app.Quit
        Set app = Nothing
    End If
            
    Application.ScreenUpdating = True
    
End Sub

