Attribute VB_Name = "Module_ZipCode"
Option Explicit
' モジュール概要： 郵便局のCSVデータ（KEN_ALL.csv）をメモリへ展開し、住所から郵便番号を検索および着色する。
'   1. CSV読込時に複数行分割レコードを結合し、正規化住所と検索用インデックス（Dictionary）を作成する。
'   2. 住所検索時は「都道府県＋市区町村（郡省略対応）」で候補を絞り込み、「町域名」の最長一致で特定する。
'   3. 「以下に掲載がない場合」や「丁目・番地の範囲指定」「（その他）」「大字・字の表記ゆれ」に対応する。
'   4. 複数候補の郵便番号出力とマッチング情報の付記を行う。
'   5. 結果出力時、検索元のセルとマスタ文字列を比較し、一致部分を赤色で着色する。
' ==============================================================================EvaluateNumericRange

' [モジュール定数・構造体・変数の宣言]
    ' CSV列インデックス定数： Split(",")で分割した際の0始まり配列インデックスを定義する。
    Private Const CSV_COL_ZIP As Long = 2   ' 郵便番号（C列）
    Private Const CSV_COL_PREF As Long = 6  ' 都道府県（G列）
    Private Const CSV_COL_CITY As Long = 7  ' 市区町村（H列）
    Private Const CSV_COL_TOWN As Long = 8  ' 町域名（I列）

    ' 郵便番号マスタ構造体： 表示用生データと検索用正規化データを保持する。
    Private Type ZipMasterData
        ZipCode As String                   ' 郵便番号（7桁）
        PrefectureRaw As String             ' 都道府県（生データ）
        CityRaw As String                   ' 市区町村（生データ）
        TownRaw As String                   ' 町域名（生データ）
        FullAddressWithParen As String      ' 住所全体1（都道府県＋市区町村＋括弧を含む町域名）
        FullAddressNoParen As String        ' 住所全体2（都道府県＋市区町村＋括弧を除いた町域名）
        NormalizedPref As String            ' 正規化都道府県
        normalizedCity As String            ' 正規化市区町村
        NormalizedTown As String            ' 正規化町域名（括弧を含まない）
        NormalizedTownCore As String        ' コア町域名（先頭の「大字」「小字」「字」を除去した町域名）
        NormalizedNoPref As String          ' 正規化市区町村＋町域名（括弧を含む）
    End Type

    ' ファイルおよびシート名定数： 読込元CSVファイル名と入出力シート名を定義する。
    Private Const CSV_FILE_郵便コード As String = "\KEN_ALL.csv"
    Private Const SHEET_NAME_検索 As String = "検索"
    Private Const SHEET_NAME_郵便コード As String = "削除可（郵便コード）"

    ' モジュールレベルキャッシュ変数： マスタ全件配列と検索用Dictionaryをメモリ上に保持する。
    Private pMasterDataCache() As ZipMasterData
    Private pMasterCount As Long
    Private pIsCacheLoaded As Boolean
    Private pSearchIndex As Object          ' 検索インデックス（Key：正規化県＋市区町村、Value：Collection）
    Private pCityToPrefMap As Object        ' 逆引き辞書（Key：正規化市区町村、Value：都道府県生データ）
    
' ==============================================================================
' --- キャッシュ状態取得プロパティ ---
' マスタデータがメモリに展開済みかどうかを返す。
Public Property Get IsCacheLoaded() As Boolean
    IsCacheLoaded = pIsCacheLoaded
End Property

' ==============================================================================
' --- マスタ読込とシート出力処理 ---
' CSVファイルを読み込み、複数行結合と正規化を行ってメモリへ格納する。
Sub LoadMasterAndExportSheet()
    Dim csvFilePath As String
    Dim fso As Object
    Dim adoStream As Object
    Dim fileContent() As String
    Dim rawData() As String
    Dim csvLineText As String
    Dim lineCount As Long
    Dim i As Long

    Dim tempZipRaw As String
    Dim tempPrefRaw As String
    Dim tempCityRaw As String
    Dim tempTownRaw As String
    Dim dictKey As String
    Dim subKeyNoGun As String

    Dim isMergingParen As Boolean
    Dim recIdx As Long

    Dim wsMasterOutput As Worksheet
    Dim outputData() As Variant
    Dim tStart As Double
    Dim response As VbMsgBoxResult

    ' [実行確認とCSVファイルの存在チェック]
    ' 読込開始の確認： ユーザーへCSV読込の実行可否を確認する。
    response = MsgBox("郵便番号データ（CSV）をメモリに格納します。" & vbCrLf & _
                     " （郵便局HP 郵便番号データダウンロード）" & vbCrLf & _
                     " KEN_ALL.csv を同じフォルダに解凍しておく", vbYesNo + vbInformation, "マスタデータ読み込み")
    If response <> vbYes Then
        Exit Sub
    End If

    ' ファイル存在確認： ブックと同一フォルダにKEN_ALL.csvが存在するか検証する。
    csvFilePath = ThisWorkbook.Path & CSV_FILE_郵便コード
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(csvFilePath) Then
        MsgBox "エラー: CSVファイルが見つかりません。" & vbCrLf & csvFilePath, vbCritical
        Exit Sub
    End If

    tStart = Timer
    Call ToggleOptimizations(False)
    Call frm_msg処理中OP

    ' [ADODB.StreamによるCSVファイルの読み込み]
    ' テキスト一括読込： Shift_JIS形式でCSVを開き、改行コードをLFに統一して行配列へ分割する。
    Set adoStream = CreateObject("ADODB.Stream")
    On Error GoTo ReadError
    With adoStream
        .Charset = "Shift_JIS"
        .Type = 2
        .Open
        .LoadFromFile csvFilePath
        fileContent = Split(Replace(.ReadText, vbCrLf, vbLf), vbLf)
        .Close
    End With
    On Error GoTo 0

    lineCount = UBound(fileContent) + 1
    If lineCount < 1 Then
        MsgBox "CSVデータが空です。", vbExclamation
        GoTo Cleanup
    End If

    ' 確認用シート出力の選択： メモリ展開後にデバッグ用シートへ書き出すか確認する。
    response = MsgBox("確認用（郵便コード）を「" & SHEET_NAME_郵便コード & "」シートに出力しますか？" & vbCrLf & vbCrLf & _
                     "（*** メモリ格納は必ず実行されます。***）", vbYesNo + vbQuestion, "シート出力確認")
    If response = vbYes Then
        Set wsMasterOutput = GetOrCreateSheet(SHEET_NAME_郵便コード)
    End If
 
    ' [メモリ展開と検索用インデックスの構築]
    ' キャッシュ配列と辞書の初期化： 行数分の配列と検索用Dictionaryを生成する。
    ReDim pMasterDataCache(1 To lineCount)
    Set pSearchIndex = CreateObject("Scripting.Dictionary")
    Set pCityToPrefMap = CreateObject("Scripting.Dictionary")

    recIdx = 0
    isMergingParen = False

    ' 全行解析ループ： 各行をカンマ分割し、複数行に跨がる町域名を結合しながら構造体へ格納する。
    For i = 0 To lineCount - 1
        csvLineText = Trim(Replace(fileContent(i), vbCr, ""))
        If csvLineText <> "" Then
            rawData = Split(csvLineText, ",")

            If UBound(rawData) >= CSV_COL_TOWN Then
                tempZipRaw = FormatZip(Replace(rawData(CSV_COL_ZIP), """", ""))
                tempPrefRaw = Replace(rawData(CSV_COL_PREF), """", "")
                tempCityRaw = Replace(rawData(CSV_COL_CITY), """", "")
                tempTownRaw = Replace(rawData(CSV_COL_TOWN), """", "")

                ' 分割レコードの連結： 直前行の開き括弧が閉じておらず同一郵便番号・市区町村であれば町域名を結合する。
                If isMergingParen And recIdx > 0 Then
                    If pMasterDataCache(recIdx).ZipCode = tempZipRaw And _
                       pMasterDataCache(recIdx).PrefectureRaw = tempPrefRaw And _
                       pMasterDataCache(recIdx).CityRaw = tempCityRaw Then
                        
                        pMasterDataCache(recIdx).TownRaw = pMasterDataCache(recIdx).TownRaw & tempTownRaw
                        If InStr(tempTownRaw, "）") > 0 Or InStr(tempTownRaw, ")") > 0 Then
                            isMergingParen = False
                        End If
                        Call FinalizeMasterRecord(recIdx)
                        GoTo ContinueLoop
                    Else
                        isMergingParen = False
                    End If
                End If

                ' 新規レコードの登録： 配列インデックスを進めて基本項目を格納する。
                recIdx = recIdx + 1
                With pMasterDataCache(recIdx)
                    .ZipCode = tempZipRaw
                    .PrefectureRaw = tempPrefRaw
                    .CityRaw = tempCityRaw
                    .TownRaw = tempTownRaw
                End With

                ' 括弧未閉鎖の判定： 開き括弧があり閉じ括弧がない場合は次行結合フラグを有効にする。
                If (InStr(tempTownRaw, "（") > 0 And InStr(tempTownRaw, "）") = 0) Then
                    isMergingParen = True
                Else
                    isMergingParen = False
                End If

                ' 正規化文字列の生成： 構造体の検索用フィールドおよび表示用フィールドを確定する。
                Call FinalizeMasterRecord(recIdx)

                ' インデックス登録： 「都道府県＋市区町村」をキーとしてレコード番号をCollectionへ追加する。
                With pMasterDataCache(recIdx)
                    dictKey = .NormalizedPref & .normalizedCity
                    If Not pSearchIndex.Exists(dictKey) Then
                        Set pSearchIndex(dictKey) = New Collection
                    End If
                    pSearchIndex(dictKey).Add recIdx

                    ' 郡省略サブキーの登録： 郡名を省いた入力住所でも検索できるようサブキーを追加する。
                    subKeyNoGun = RemoveGunFromCity(.normalizedCity)
                    If subKeyNoGun <> .normalizedCity And subKeyNoGun <> "" Then
                        Dim altKey As String
                        altKey = .NormalizedPref & subKeyNoGun
                        If Not pSearchIndex.Exists(altKey) Then
                            Set pSearchIndex(altKey) = New Collection
                        End If
                        pSearchIndex(altKey).Add recIdx
                    End If

                    ' 都道府県逆引きマップの登録： 都道府県省略時に市区町村名から都道府県を特定する辞書を作成する。
                    If .normalizedCity <> "" Then
                        pCityToPrefMap(.normalizedCity) = .PrefectureRaw
                        If subKeyNoGun <> .normalizedCity And subKeyNoGun <> "" Then
                            If Not pCityToPrefMap.Exists(subKeyNoGun) Then
                                pCityToPrefMap(subKeyNoGun) = .PrefectureRaw
                            End If
                        End If
                    End If
                End With
            End If
        End If
ContinueLoop:
        If i Mod 2000 = 0 Then DoEvents
    Next i

    ' キャッシュ配列のリサイズ： 結合処理後の有効レコード数に合わせて配列サイズを確定する。
    pMasterCount = recIdx
    If pMasterCount > 0 Then
        ReDim Preserve pMasterDataCache(1 To pMasterCount)
    End If

    ' [確認用シートへの一括書き出し]
    ' シート出力の実行： 出力が選択された場合のみ2次元配列を構築してシートへ一括転送する。
    If response = vbYes And Not wsMasterOutput Is Nothing And pMasterCount > 0 Then
        ReDim outputData(1 To pMasterCount, 1 To 8)
        For i = 1 To pMasterCount
            With pMasterDataCache(i)
                outputData(i, 1) = .ZipCode
                outputData(i, 2) = .FullAddressWithParen
                outputData(i, 3) = .FullAddressNoParen
                outputData(i, 4) = .PrefectureRaw
                outputData(i, 5) = .CityRaw
                outputData(i, 6) = .TownRaw
                outputData(i, 7) = .NormalizedNoPref
                outputData(i, 8) = .NormalizedTown
            End With
        Next i

        With wsMasterOutput
            .Cells.Clear
            .Cells.Font.Name = "BIZ UDゴシック"
            .Cells.Font.Size = 9

            .Range("A1:H1").Value = Array("ZipCode", _
                                         "FullAddressWithParen", _
                                         "FullAddressNoParen", _
                                         "PrefectureRaw", _
                                         "CityRaw", _
                                         "TownRaw", _
                                         "NormalizedCityTown", _
                                         "NormalizedTownName")
            .Range("A2").Resize(pMasterCount, 8).Value = outputData

            .Columns("A").ColumnWidth = 8
            .Columns("B").ColumnWidth = 40
            .Columns("C").ColumnWidth = 30
            .Columns("D").ColumnWidth = 10
            .Columns("E").ColumnWidth = 20
            .Columns("F").ColumnWidth = 30
            .Columns("G").ColumnWidth = 20
            .Columns("H").ColumnWidth = 20
        End With
        Erase outputData
    End If

    ' [完了通知とリソース解放]
    pIsCacheLoaded = True

    Call frm_msg処理中CL
    MsgBox "マスタデータ（" & Format(pMasterCount, "#,##0") & "件）の読み込み完了。" & vbCrLf & _
             "処理時間: " & Format(Timer - tStart, "0.00") & "秒", vbInformation
Cleanup:
    Call ToggleOptimizations(True)
    Set adoStream = Nothing
    Set fso = Nothing
    Exit Sub

ReadError:
    Call frm_msg処理中CL
    MsgBox "ファイル読み込みエラー: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ==============================================================================
' --- マスタレコード確定処理 ---
' 1レコード分の表示用文字列および正規化文字列を生成する。
Private Sub FinalizeMasterRecord(ByVal recIdx As Long)
    Dim townForSearchRaw As String
    With pMasterDataCache(recIdx)
        ' 特殊文言の除外： 「以下に掲載がない場合」やビル除外注記のみのレコードは検索町名を空欄とする。
        townForSearchRaw = .TownRaw
        If townForSearchRaw = "以下に掲載がない場合" Or _
           (InStr(townForSearchRaw, "次のビルを除く") > 0 And Left(townForSearchRaw, 1) = "（") Then
            townForSearchRaw = ""
        End If

        ' 表示用文字列の設定： 括弧ありと括弧なしの完全住所文字列をそれぞれ保持する。
        .FullAddressWithParen = .PrefectureRaw & .CityRaw & .TownRaw
        .FullAddressNoParen = .PrefectureRaw & .CityRaw & RemoveRangeSuffixTownName(townForSearchRaw)

        ' 正規化文字列の設定： 都道府県、市区町村、町域名、および先頭の大字・字を除いたコア町名を保持する。
        .NormalizedPref = NormalizeAddress(.PrefectureRaw)
        .normalizedCity = NormalizeAddress(.CityRaw)
        .NormalizedTown = NormalizeAddress(RemoveRangeSuffixTownName(townForSearchRaw))
        .NormalizedTownCore = StripLeadingOazaAza(.NormalizedTown)
        .NormalizedNoPref = NormalizeAddress(.CityRaw & .TownRaw)
    End With
End Sub

' ==============================================================================
' --- 住所検索メイン処理 ---
' 検索シートの入力住所を読み込み、郵便番号の特定と赤色着色を実行する。
Sub ExecuteAddressSearch()
    Dim wsSearchInput As Worksheet
    Dim lastRow As Long
    Dim rngInputData As Range
    Dim arrInputAddress As Variant
    Dim arrResult() As Variant
    Dim i As Long
    Dim tStart As Double

    ' [事前検証と入力データの取得]
    ' キャッシュ状態の確認： マスタ未読込の場合はメッセージを表示して終了する。
    If Not pIsCacheLoaded Then
        MsgBox "先に「マスタ読み込み」を実行してください。", vbExclamation
        Exit Sub
    End If

    ' 入力範囲の取得： B列の最終行を取得し、データが存在しない場合は終了する。
    Set wsSearchInput = ThisWorkbook.Sheets(SHEET_NAME_検索)
    lastRow = wsSearchInput.Cells(wsSearchInput.Rows.Count, "B").End(xlUp).Row
    If lastRow < 3 Then MsgBox "データがありません。", vbExclamation: Exit Sub

    ' 入力配列の構築： 1行のみの場合でも2次元配列として扱えるよう格納する。
    Set rngInputData = wsSearchInput.Range("B3:B" & lastRow)
    If rngInputData.Rows.Count = 1 Then
        ReDim arrInputAddress(1 To 1, 1 To 1)
        arrInputAddress(1, 1) = rngInputData.Value
    Else
        arrInputAddress = rngInputData.Value
    End If

    ' [出力領域の初期化と検索ループの実行]
    ' シート初期化： 既存の検索結果およびB列の文字色をリセットし、ヘッダーを設定する。
    ReDim arrResult(1 To UBound(arrInputAddress, 1), 1 To 4)
    With wsSearchInput
        ' A列は名前や検証分類等の任意入力エリアのためクリア対象外とする。
'xx     .Range("A3").Resize(UBound(arrInputAddress, 1), 1).ClearContents
        .Range("B3").Resize(UBound(arrInputAddress, 1), 1).Font.ColorIndex = xlAutomatic
        .Range("C3").Resize(UBound(arrInputAddress, 1), 4).ClearContents
        .Range("C2:F2").Value = Array("郵便番号", "FullAddressNoParen", "TownRaw", "マッチング情報")
    End With

    tStart = Timer
    Call ToggleOptimizations(False)

    ' 検索ループ： 各行の住所文字列に対して郵便番号特定処理を実行する。
    For i = 1 To UBound(arrInputAddress, 1)
        Call FindMatchesForAddress(CStr(arrInputAddress(i, 1)), arrResult, i)
    Next i

    ' [検索結果の一括出力と一致文字の着色]
    ' 結果配列の書き込み： C?F列へ検索結果を一括で出力する。
    wsSearchInput.Range("C3").Resize(UBound(arrResult, 1), 4).Value = arrResult

    ' 一致部分の赤色着色： 該当ありの行に対して入力セルとマスタ文字列の照合着色を行う。
    For i = 1 To UBound(arrInputAddress, 1)
        If CStr(arrResult(i, 1)) <> "該当なし" And CStr(arrInputAddress(i, 1)) <> "" Then
            Call ColorMatchInCell(wsSearchInput, i + 2, _
                                  CStr(arrInputAddress(i, 1)), _
                                  CStr(arrResult(i, 2)), _
                                  CStr(arrResult(i, 3)))
        End If
    Next i

    Call ToggleOptimizations(True)
    MsgBox "検索完了 (" & UBound(arrInputAddress, 1) & "件)" & vbCrLf & _
             "処理時間: " & Format(Timer - tStart, "0.00") & "秒", vbInformation
End Sub

' ==============================================================================
' --- 住所別マッチング制御 ---
' 都道府県補完と市区町村インデックス検索を行い、町名マッチングを呼び出す。
Private Sub FindMatchesForAddress(ByVal rawAddress As String, ByRef resultArray() As Variant, ByVal rowIndex As Long)
    Dim trimmedRaw As String
    Dim inputSearchKey As String
    Dim inputNormalized As String
    Dim foundPref As String

    ' 初期値の設定： 検索結果の初期状態として「該当なし」をセットする。
    resultArray(rowIndex, 1) = "該当なし"
    resultArray(rowIndex, 2) = "***"
    resultArray(rowIndex, 3) = ""
    resultArray(rowIndex, 4) = ""

    trimmedRaw = Trim(rawAddress)
    If trimmedRaw = "" Then Exit Sub

    ' 都道府県の補完と正規化： 先頭に都道府県がない場合は市区町村名から逆引きして補完する。
    If IsAddressPrefectureStart(trimmedRaw) Then
        inputSearchKey = trimmedRaw
    Else
        foundPref = FindPrefectureByCity(trimmedRaw)
        inputSearchKey = foundPref & trimmedRaw
    End If

    ' 入力文字列の正規化： 番地区切り文字を保持した状態で全角化・漢数字変換を実行する。
    inputNormalized = NormalizeAddress(inputSearchKey)

    ' 市区町村インデックスの最長一致検索： 先頭から文字数を伸ばしながらDictionaryを直接参照する。
    Dim targetIndices As Collection
    Dim bestKey As String
    Dim tryLen As Long
    Dim candidatePrefix As String
    Dim maxSearchLen As Long

    maxSearchLen = Len(inputNormalized)
    If maxSearchLen > 16 Then maxSearchLen = 16

    For tryLen = 3 To maxSearchLen
        candidatePrefix = Left(inputNormalized, tryLen)
        If pSearchIndex.Exists(candidatePrefix) Then
            bestKey = candidatePrefix
        End If
    Next tryLen

    If bestKey = "" Then
        resultArray(rowIndex, 4) = "市区町村不一致"
        Exit Sub
    End If
    Set targetIndices = pSearchIndex(bestKey)

    ' 町域名マッチングの実行： 市区町村以降の文字列を切り出し、候補コレクションと照合する。
    Dim inputSuffixRaw As String
    Dim inputSuffixCore As String
    inputSuffixRaw = Mid(inputNormalized, Len(bestKey) + 1)
    inputSuffixCore = StripLeadingOazaAza(inputSuffixRaw)

    Dim matchedIdx As Long
    matchedIdx = ExecuteTownMatching(targetIndices, inputSuffixRaw, inputSuffixCore, resultArray, rowIndex)

    ' 京都市の通り名フォールバック： 京都市で不一致の場合、通り名表記を除去して再試行する。
    If matchedIdx = 0 And InStr(1, bestKey, "京都府京都市") = 1 Then
        Dim kyotoCleanSuffix As String
        kyotoCleanSuffix = StripKyotoTorina(inputSuffixRaw)
        If kyotoCleanSuffix <> inputSuffixRaw And kyotoCleanSuffix <> "" Then
            Call ExecuteTownMatching(targetIndices, kyotoCleanSuffix, StripLeadingOazaAza(kyotoCleanSuffix), resultArray, rowIndex)
        End If
    End If
End Sub

' ==============================================================================
' --- 町域名最長一致・範囲判定処理---
' 候補レコード群から最長一致かつ最高優先度の郵便番号を特定する。
Private Function ExecuteTownMatching(ByVal targetIndices As Collection, _
                                     ByVal inputSuffixRaw As String, _
                                     ByVal inputSuffixCore As String, _
                                     ByRef resultArray() As Variant, _
                                     ByVal rowIndex As Long) As Long
    Dim idx As Variant
    Dim master As ZipMasterData

    ' 候補管理用Dictionaryの初期化： 郵便番号ごとのマスタIndex、一致文字数、優先度を保持する。
    Dim candIdxMap As Object: Set candIdxMap = CreateObject("Scripting.Dictionary")
    Dim candLenMap As Object: Set candLenMap = CreateObject("Scripting.Dictionary")
    Dim candPriMap As Object: Set candPriMap = CreateObject("Scripting.Dictionary")

    Dim bestMatchTownLen As Long: bestMatchTownLen = -1
    Dim bestPriorityAtMaxLen As Long: bestPriorityAtMaxLen = 0
    Dim genericFallbackIdx As Long: genericFallbackIdx = -1
    Dim hadRangeRejection As Boolean: hadRangeRejection = False

    ' 候補コレクションの走査： 同一市区町村内の全マスタレコードに対して町名と括弧条件を検証する。
    For Each idx In targetIndices
        master = pMasterDataCache(idx)

        ' 代表レコードの退避： 「以下に掲載がない場合」はフォールバック用インデックスとして保持する。
        If master.TownRaw = "以下に掲載がない場合" Then
            genericFallbackIdx = idx
            GoTo NextMasterRecord
        End If

        Dim mTown As String
        Dim mTownCore As String
        mTown = master.NormalizedTown
        mTownCore = master.NormalizedTownCore
        If mTown = "" Then GoTo NextMasterRecord

        ' 町名一致判定： 完全一致、大字・字を除いたコア一致、または丁目ハイフン表記の一致を判定する。
        Dim isTownMatched As Boolean: isTownMatched = False
        Dim effectiveMatchLen As Long: effectiveMatchLen = 0
        Dim inputRest As String: inputRest = ""

        If InStr(1, inputSuffixRaw, mTown) = 1 Then
            isTownMatched = True
            effectiveMatchLen = Len(mTown)
            inputRest = Mid(inputSuffixRaw, Len(mTown) + 1)
        ElseIf mTownCore <> "" And InStr(1, inputSuffixCore, mTownCore) = 1 Then
            isTownMatched = True
            effectiveMatchLen = Len(mTownCore)
            inputRest = Mid(inputSuffixCore, Len(mTownCore) + 1)
        ElseIf (Right(mTown, 2) = "丁目" Or Right(mTown, 1) = "丁") Then
            ' 丁目省略ハイフン表記の判定： ベース町名が数字のみでない場合かつ後続がハイフン等の時のみ一致とする。
            Dim townBase As String
            If Right(mTown, 2) = "丁目" Then
                townBase = Left(mTown, Len(mTown) - 2)
            Else
                townBase = Left(mTown, Len(mTown) - 1)
            End If
            If townBase <> "" And Not IsAllFullWidthDigits(townBase) Then
                If InStr(1, inputSuffixCore, townBase) = 1 Then
                    Dim afterTownBase As String
                    afterTownBase = Mid(inputSuffixCore, Len(townBase) + 1)
                    If Left(afterTownBase, 1) = "－" Or afterTownBase = "" Then
                        isTownMatched = True
                        effectiveMatchLen = Len(mTown)
                        inputRest = afterTownBase
                    End If
                End If
            End If
        End If

        If isTownMatched Then
            ' 括弧内条件と優先度の評価： 条件一致は優先度3、標準レコードは2、「その他」は1を設定する。
            Dim isCandidate As Boolean: isCandidate = True
            Dim recordPriority As Long: recordPriority = 2

            Dim hasParen As Boolean
            hasParen = (InStr(master.TownRaw, "（") > 0)

            If hasParen Then
                Dim parenContent As String
                parenContent = ExtractParenContent(master.NormalizedNoPref)

                ' NormalizeAddressにより「その他」は「ソノ他」、「一円」は「１円」に変換されている点を考慮する。
                If (InStr(parenContent, "ソノ他") > 0 Or InStr(parenContent, "その他") > 0) And _
                   Not (parenContent Like "*[０-９]*") Then
                    recordPriority = 1
                ElseIf InStr(parenContent, "１円") > 0 Or InStr(master.TownRaw, "一円") > 0 Then
                    recordPriority = 2
                ElseIf parenContent Like "*[０-９]*" Then
                    ' 数値範囲の評価： 入力の番地・丁目数値を抽出し、括弧内の範囲条件と照合する。
                    Dim inputNum As Long
                    Dim inputSubNum As Long
                    inputNum = ExtractNumberFromSuffix(inputRest)
                    inputSubNum = ExtractSecondNumberFromSuffix(inputRest)

                    If inputNum > 0 Then
                        Dim rangeResult As Integer
                        rangeResult = EvaluateNumericRange(parenContent, inputNum, inputSubNum)
                        If rangeResult = 1 Then
                            recordPriority = 3
                        ElseIf rangeResult = -1 Then
                            isCandidate = False
                            hadRangeRejection = True
                        End If
                    Else
                        recordPriority = 2
                    End If
                Else
                    ' 括弧内小字・枝地名の評価： 「（渡戸沢「筍沢温泉」）」等の地名が入力後続に含まれるか判定する。
                    If Not IsGenericAnnotationParen(parenContent) Then
                        If IsSubTownParenMatch(parenContent, inputRest) Then
                            recordPriority = 3
                        Else
                            isCandidate = False
                        End If
                    End If
                End If
            End If

            ' 候補辞書の更新： 同一郵便番号でもより長い一致長または高い優先度があれば上書きする。
            If isCandidate Then
                Dim zCode As String
                zCode = master.ZipCode

                If Not candIdxMap.Exists(zCode) Then
                    candIdxMap(zCode) = CLng(idx)
                    candLenMap(zCode) = effectiveMatchLen
                    candPriMap(zCode) = recordPriority
                Else
                    If effectiveMatchLen > candLenMap(zCode) Or _
                       (effectiveMatchLen = candLenMap(zCode) And recordPriority > candPriMap(zCode)) Then
                        candIdxMap(zCode) = CLng(idx)
                        candLenMap(zCode) = effectiveMatchLen
                        candPriMap(zCode) = recordPriority
                    End If
                End If

                If effectiveMatchLen > bestMatchTownLen Then
                    bestMatchTownLen = effectiveMatchLen
                    bestPriorityAtMaxLen = recordPriority
                ElseIf effectiveMatchLen = bestMatchTownLen Then
                    If recordPriority > bestPriorityAtMaxLen Then
                        bestPriorityAtMaxLen = recordPriority
                    End If
                End If
            End If
        End If
NextMasterRecord:
    Next idx

    ' 最終候補の絞り込み： 最長一致長かつ最高優先度を持つ候補のみを抽出する。
    Dim finalResultZips As Object
    Set finalResultZips = CreateObject("Scripting.Dictionary")
    Dim matchInfo As String: matchInfo = "不一致"

    Dim key As Variant
    If bestMatchTownLen > 0 Then
        For Each key In candIdxMap.Keys
            If candLenMap(key) = bestMatchTownLen Then
                If candPriMap(key) = bestPriorityAtMaxLen Or _
                   (bestPriorityAtMaxLen = 2 And candPriMap(key) >= 1) Then
                    finalResultZips(key) = candIdxMap(key)
                End If
            End If
        Next key
    End If

    ' マッチング情報の決定： 町名一致、市区町村代表採用、または不一致のステータスを設定する。
    If finalResultZips.Count > 0 Then
        If bestPriorityAtMaxLen = 3 Then
            matchInfo = "町名＋範囲一致 (" & bestMatchTownLen & "文字)"
        Else
            matchInfo = "町名最長一致 (" & bestMatchTownLen & "文字)"
        End If
        ExecuteTownMatching = 1
    ElseIf genericFallbackIdx <> -1 Then
        finalResultZips(pMasterDataCache(genericFallbackIdx).ZipCode) = genericFallbackIdx
        If hadRangeRejection Then
            matchInfo = "市区町村代表採用(範囲外)"
        Else
            matchInfo = "市区町村代表採用"
        End If
        ExecuteTownMatching = 0
    Else
        If hadRangeRejection Then
            resultArray(rowIndex, 4) = "範囲外で除外"
        Else
            resultArray(rowIndex, 4) = "不一致"
        End If
        ExecuteTownMatching = 0
        Exit Function
    End If

    ' 結果配列への格納： 複数候補がある場合はカンマ区切りで郵便番号を連結し、代表町名をセットする。
    Dim zipRes As String
    Dim fullAddressRes As String
    Dim townRes As String
    Dim zipKey As Variant

    For Each zipKey In finalResultZips.Keys
        master = pMasterDataCache(finalResultZips(zipKey))

        zipRes = zipRes & FormatZipForOutput(master.ZipCode) & ", "

        If fullAddressRes = "" Then
            fullAddressRes = master.FullAddressNoParen
        End If

        If townRes = "" Then
            Dim resultTownRaw As String
            resultTownRaw = master.TownRaw
            If resultTownRaw = "以下に掲載がない場合" Then resultTownRaw = ""
            If InStr(resultTownRaw, "（") > 0 Then
                resultTownRaw = Left(resultTownRaw, InStr(resultTownRaw, "（") - 1)
            End If
            townRes = resultTownRaw
        End If
    Next zipKey

    If Len(zipRes) > 2 Then zipRes = Left(zipRes, Len(zipRes) - 2)

    resultArray(rowIndex, 1) = zipRes
    resultArray(rowIndex, 2) = fullAddressRes
    resultArray(rowIndex, 3) = townRes
    resultArray(rowIndex, 4) = matchInfo
End Function

' ==============================================================================
' --- 住所正規化処理 ---
' 全角化、カナ統一、異体字統一、ハイフン統一、および複合漢数字変換を行う。
Private Function NormalizeAddress(ByVal inputAddress As String) As String
    Dim result As String
    result = inputAddress
    If result = "" Then Exit Function

    ' 全角化とカタカナ化： 半角文字を全角へ揃え、ひらがなをカタカナへ変換する。
    result = StrConv(result, vbWide)
    result = StrConv(result, vbKatakana)

    ' 表記ゆれ文字の統一： 小書きカナや助詞のゆらぎを同一文字へ置換する。
    result = Replace(result, "ヶ", "ケ")
    result = Replace(result, "ヵ", "カ")
    result = Replace(result, "が", "ケ")
    result = Replace(result, "ガ", "ケ")
    result = Replace(result, "之", "ノ")
    result = Replace(result, "の", "ノ")
    result = Replace(result, "ッツ", "ツ")
    result = Replace(result, "ッ", "ツ")

    ' 異体字の統一： 旧字体や異体字を標準的な新字体へ置換する。
    result = Replace(result, "條", "条")
    result = Replace(result, "櫻", "桜")
    result = Replace(result, "澤", "沢")

    ' 区切り記号の統一： 各種ハイフンやダッシュ類を全角ハイフン「－」へ統一する。
    result = Replace(result, "-", "－")
    result = Replace(result, "ー", "－")
    result = Replace(result, "―", "－")
    result = Replace(result, "‐", "－")
    result = Replace(result, "ｰ", "－")

    ' 漢数字の変換： 十・百を含む漢数字列を全角アラビア数字へ変換する。
    result = ConvertKanjiNumerals(result)

    ' 空白の除去： 全角および半角のスペースを削除する。
    result = Replace(result, " ", "")
    result = Replace(result, " ", "")

    NormalizeAddress = result
End Function

' ==============================================================================
' --- 漢数字トークン抽出処理 ---
' 文字列内の連続する漢数字部分を切り出して算用数字変換へ渡す。
Private Function ConvertKanjiNumerals(ByVal src As String) As String
    Dim i As Long, n As Long
    Dim ch As String
    Dim kanjiBuf As String
    Dim outStr As String

    n = Len(src)
    For i = 1 To n
        ch = Mid(src, i, 1)
        If InStr("〇一二三四五六七八九十百", ch) > 0 Then
            kanjiBuf = kanjiBuf & ch
        Else
            If kanjiBuf <> "" Then
                outStr = outStr & ParseSingleKanjiNumber(kanjiBuf)
                kanjiBuf = ""
            End If
            outStr = outStr & ch
        End If
    Next i

    If kanjiBuf <> "" Then
        outStr = outStr & ParseSingleKanjiNumber(kanjiBuf)
    End If

    ConvertKanjiNumerals = outStr
End Function

' --- 複合漢数字の評価処理 ---
' 「十」「百」を含む漢数字トークンを数値計算し、全角数字文字列へ変換する。
Private Function ParseSingleKanjiNumber(ByVal kStr As String) As String
    ' 単純置換の分岐： 「十」「百」を含まない場合は各桁を直接全角数字へ置換する。
    If InStr(kStr, "十") = 0 And InStr(kStr, "百") = 0 Then
        Dim res As String
        res = kStr
        res = Replace(res, "〇", "０")
        res = Replace(res, "一", "１")
        res = Replace(res, "二", "２")
        res = Replace(res, "三", "３")
        res = Replace(res, "四", "４")
        res = Replace(res, "五", "５")
        res = Replace(res, "六", "６")
        res = Replace(res, "七", "７")
        res = Replace(res, "八", "８")
        res = Replace(res, "九", "９")
        ParseSingleKanjiNumber = res
        Exit Function
    End If

    ' 位取り計算の実行： 「十二」を12、「二十四」を24、「百二十三」を123として算出する。
    Dim total As Long: total = 0
    Dim currentDigit As Long: currentDigit = 0
    Dim idx As Long
    Dim c As String

    For idx = 1 To Len(kStr)
        c = Mid(kStr, idx, 1)
        Select Case c
            Case "一": currentDigit = 1
            Case "二": currentDigit = 2
            Case "三": currentDigit = 3
            Case "四": currentDigit = 4
            Case "五": currentDigit = 5
            Case "六": currentDigit = 6
            Case "七": currentDigit = 7
            Case "八": currentDigit = 8
            Case "九": currentDigit = 9
            Case "〇": currentDigit = 0
            Case "百"
                If currentDigit = 0 Then currentDigit = 1
                total = total + currentDigit * 100
                currentDigit = 0
            Case "十"
                If currentDigit = 0 Then currentDigit = 1
                total = total + currentDigit * 10
                currentDigit = 0
        End Select
    Next idx
    total = total + currentDigit

    ParseSingleKanjiNumber = StrConv(CStr(total), vbWide)
End Function

' ==============================================================================
' --- 先頭大字・字の除去処理 ---
' 町名先頭にある「大字」「小字」「字」のみを取り除く。
Private Function StripLeadingOazaAza(ByVal townStr As String) As String
    Dim s As String
    s = townStr
    If Left(s, 2) = "大字" And Len(s) > 2 Then
        s = Mid(s, 3)
    ElseIf Left(s, 2) = "小字" And Len(s) > 2 Then
        s = Mid(s, 3)
    ElseIf Left(s, 1) = "字" And Len(s) > 1 Then
        s = Mid(s, 2)
    End If
    StripLeadingOazaAza = s
End Function

' ==============================================================================
' --- 郡名除去処理 ---
' 町村の市区町村名から「〇〇郡」を除いた文字列を返す（「郡山市」等の市名は維持する）。
Private Function RemoveGunFromCity(ByVal normalizedCity As String) As String
    Dim gunPos As Long
    gunPos = InStr(normalizedCity, "郡")
    If gunPos > 0 And Right(normalizedCity, 1) <> "市" And gunPos < Len(normalizedCity) Then
        RemoveGunFromCity = Mid(normalizedCity, gunPos + 1)
    Else
        RemoveGunFromCity = normalizedCity
    End If
End Function

' ==============================================================================
' --- 京都市通り名除去処理 ---
' 「上ル」「下ル」「東入」「西入」までの通り名記述を先頭から除去する。
Private Function StripKyotoTorina(ByVal suffixStr As String) As String
    Dim markers As Variant
    Dim m As Variant
    Dim pos As Long
    Dim maxEndPos As Long: maxEndPos = 0

    markers = Array("上ル", "下ル", "東入ル", "西入ル", "東入", "西入")
    For Each m In markers
        pos = InStr(suffixStr, CStr(m))
        If pos > 0 Then
            If pos + Len(CStr(m)) - 1 > maxEndPos Then
                maxEndPos = pos + Len(CStr(m)) - 1
            End If
        End If
    Next m

    If maxEndPos > 0 And maxEndPos < Len(suffixStr) Then
        StripKyotoTorina = Mid(suffixStr, maxEndPos + 1)
    Else
        StripKyotoTorina = suffixStr
    End If
End Function

' ==============================================================================
' --- 全角数字のみで構成されているかの判定---
' 町名ベース部が数字のみ（「四丁目」の「四」等）かを判定する。
Private Function IsAllFullWidthDigits(ByVal strVal As String) As Boolean
    Dim i As Long
    If strVal = "" Then
        IsAllFullWidthDigits = False
        Exit Function
    End If
    For i = 1 To Len(strVal)
        If Not (Mid(strVal, i, 1) Like "[０-９]") Then
            IsAllFullWidthDigits = False
            Exit Function
        End If
    Next i
    IsAllFullWidthDigits = True
End Function

' ==============================================================================
' --- 汎用注釈括弧の判定 ---
' 括弧内が固有の小字地名ではなく「一円」「全域」「除ク」等の一般注記かを判定する。
Private Function IsGenericAnnotationParen(ByVal parenContent As String) As Boolean
    If parenContent = "" Then
        IsGenericAnnotationParen = True
    ElseIf InStr(parenContent, "ソノ他") > 0 Or _
           InStr(parenContent, "その他") > 0 Or _
           InStr(parenContent, "１円") > 0 Or _
           InStr(parenContent, "一円") > 0 Or _
           InStr(parenContent, "全域") > 0 Or _
           InStr(parenContent, "除ク") > 0 Or _
           InStr(parenContent, "除く") > 0 Or _
           InStr(parenContent, "以外") > 0 Or _
           InStr(parenContent, "マタガル") > 0 Or _
           InStr(parenContent, "地階") > 0 Or _
           InStr(parenContent, "階層") > 0 Then
        IsGenericAnnotationParen = True
    Else
        IsGenericAnnotationParen = False
    End If
End Function

' ==============================================================================
' --- 括弧内小字・枝地名の一致判定 ---
' 鉤括弧「」を除去した全体一致または分割キーワードの部分一致を判定する。
Private Function IsSubTownParenMatch(ByVal parenContent As String, ByVal inputRest As String) As Boolean
    Dim cleanRest As String
    Dim cleanParen As String
    Dim tokenStr As String
    Dim tokens() As String
    Dim tk As Variant
    Dim cleanTk As String

    IsSubTownParenMatch = False
    cleanRest = inputRest
    cleanRest = Replace(cleanRest, "（", "")
    cleanRest = Replace(cleanRest, "）", "")
    cleanRest = Replace(cleanRest, "(", "")
    cleanRest = Replace(cleanRest, ")", "")
    cleanRest = Replace(cleanRest, "「", "")
    cleanRest = Replace(cleanRest, "」", "")
    cleanRest = StripLeadingOazaAza(cleanRest)
    If cleanRest = "" Then Exit Function

    ' 1. 鉤括弧「」を除いた連結文字列での一致判定（例：「渡戸沢「筍沢温泉」」→「渡戸沢筍沢温泉」）
    cleanParen = Replace(Replace(parenContent, "「", ""), "」", "")
    cleanParen = StripLeadingOazaAza(cleanParen)
    If cleanParen <> "" And InStr(1, cleanRest, cleanParen) > 0 Then
        IsSubTownParenMatch = True
        Exit Function
    End If

    ' 2. 「」や読点で分割した個別地名キーワードでの一致判定（例：「渡戸沢」または「筍沢温泉」）
    tokenStr = Replace(parenContent, "「", "、")
    tokenStr = Replace(tokenStr, "」", "、")
    tokenStr = Replace(tokenStr, "・", "、")
    tokens = Split(tokenStr, "、")

    For Each tk In tokens
        cleanTk = StripLeadingOazaAza(Trim(CStr(tk)))
        If Len(cleanTk) >= 2 Then
            If InStr(1, cleanRest, cleanTk) > 0 Then
                IsSubTownParenMatch = True
                Exit Function
            End If
        End If
    Next tk
End Function

' ==============================================================================
' --- 町域名の括弧除去処理 ---
' 町域名に開き括弧が含まれる場合、括弧の手前までの文字列を返す。
Private Function RemoveRangeSuffixTownName(ByVal targetTownName As String) As String
    Dim openPos As Long

    openPos = InStr(targetTownName, "（")
    If openPos = 0 Then openPos = InStr(targetTownName, "(")

    If openPos > 0 Then
        RemoveRangeSuffixTownName = Left(targetTownName, openPos - 1)
    Else
        RemoveRangeSuffixTownName = targetTownName
    End If
End Function

' ==============================================================================
' --- 都道府県開始判定 ---
' 入力住所が47都道府県のいずれかから始まっているかを判定する。
Private Function IsAddressPrefectureStart(ByVal addressString As String) As Boolean
    Static prefList As Variant
    Dim cleanAddr As String
    cleanAddr = Trim(addressString)

    If IsEmpty(prefList) Then
        prefList = Array("北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県", _
                         "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県", _
                         "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県", _
                         "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県", _
                         "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県", _
                         "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県", _
                         "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県")
    End If

    Dim prefecture As Variant
    For Each prefecture In prefList
        If cleanAddr Like CStr(prefecture) & "*" Then
            IsAddressPrefectureStart = True
            Exit Function
        End If
    Next prefecture
    IsAddressPrefectureStart = False
End Function

' ==============================================================================
' --- 都道府県逆引き推測処理 ---
' 入力先頭の市区町村名から最長一致で都道府県名を特定する。
Private Function FindPrefectureByCity(ByVal rawAddress As String) As String
    Dim normInput As String
    Dim tryLen As Long
    Dim maxSearchLen As Long
    Dim candidateCity As String
    Dim matchedPref As String: matchedPref = ""

    normInput = NormalizeAddress(rawAddress)
    maxSearchLen = Len(normInput)
    If maxSearchLen > 12 Then maxSearchLen = 12

    For tryLen = 2 To maxSearchLen
        candidateCity = Left(normInput, tryLen)
        If pCityToPrefMap.Exists(candidateCity) Then
            matchedPref = CStr(pCityToPrefMap(candidateCity))
        End If
    Next tryLen

    FindPrefectureByCity = matchedPref
End Function

' ==============================================================================
' --- 第1数値抽出処理 ---
' 町名一致後の残り文字列から、最初の連続する数値（丁目または番地）を抽出する。
Private Function ExtractNumberFromSuffix(ByVal suffixAddress As String) As Long
    Dim i As Long
    Dim narrowSuffix As String
    Dim extractedNumStr As String
    Dim isDigitFound As Boolean

    ExtractNumberFromSuffix = 0
    If suffixAddress = "" Then Exit Function

    narrowSuffix = StrConv(suffixAddress, vbNarrow)

    For i = 1 To Len(narrowSuffix)
        Dim char As String
        char = Mid(narrowSuffix, i, 1)

        If char >= "0" And char <= "9" Then
            extractedNumStr = extractedNumStr & char
            isDigitFound = True
        Else
            If isDigitFound Then Exit For
            If char <> "-" And char <> "第" And char <> "字" And char <> "大" Then
                Exit For
            End If
        End If
    Next i

    If IsNumeric(extractedNumStr) And Len(extractedNumStr) <= 9 Then
        ExtractNumberFromSuffix = CLng(extractedNumStr)
    End If
End Function

' ==============================================================================
' --- 第2数値抽出処理 ---
' 「1丁目2番」や「1-2」のような入力から、2番目の数値（番地）を抽出する。
Private Function ExtractSecondNumberFromSuffix(ByVal suffixAddress As String) As Long
    Dim i As Long
    Dim narrowSuffix As String
    Dim extractedNumStr As String
    Dim groupIdx As Long: groupIdx = 0
    Dim inDigits As Boolean: inDigits = False

    ExtractSecondNumberFromSuffix = 0
    If suffixAddress = "" Then Exit Function

    narrowSuffix = StrConv(suffixAddress, vbNarrow)

    For i = 1 To Len(narrowSuffix)
        Dim char As String
        char = Mid(narrowSuffix, i, 1)

        If char >= "0" And char <= "9" Then
            If Not inDigits Then
                groupIdx = groupIdx + 1
                inDigits = True
            End If
            If groupIdx = 2 Then
                extractedNumStr = extractedNumStr & char
            End If
        Else
            inDigits = False
            If groupIdx >= 2 Then Exit For
        End If
    Next i

    If IsNumeric(extractedNumStr) And Len(extractedNumStr) <= 9 Then
        ExtractSecondNumberFromSuffix = CLng(extractedNumStr)
    End If
End Function

' ==============================================================================
' --- 括弧内文字列の抽出処理 ---
' 最外の開き括弧から閉じ括弧までの中身を切り出す。
Private Function ExtractParenContent(ByVal targetStr As String) As String
    Dim startPos As Long, endPos As Long
    startPos = InStr(targetStr, "（")
    If startPos = 0 Then startPos = InStr(targetStr, "(")

    endPos = InStrRev(targetStr, "）")
    If endPos = 0 Then endPos = InStrRev(targetStr, ")")

    If startPos > 0 And endPos > startPos Then
        ExtractParenContent = Mid(targetStr, startPos + 1, endPos - startPos - 1)
    ElseIf startPos > 0 Then
        ExtractParenContent = Mid(targetStr, startPos + 1)
    Else
        ExtractParenContent = ""
    End If
End Function

' ==============================================================================
' --- 数値範囲条件の評価処理 ---
' 括弧内の範囲指定と入力数値を比較し、1（一致）、-1（範囲外）、0（対象外）を返す。
Private Function EvaluateNumericRange(ByVal parenContent As String, _
                                      ByVal inputVal1 As Long, _
                                      Optional ByVal inputVal2 As Long = 0) As Integer
    Dim cleanStr As String
    Dim segments() As String
    Dim seg As Variant
    Dim sItem As String
    Dim isExclude As Boolean
    Dim hasValidRule As Boolean: hasValidRule = False
    Dim matchedPositiveRule As Boolean: matchedPositiveRule = False

    ' 複合指定（丁目＋特定番地）の判定： マスタ側に「丁目」と「番」の両方があり入力に番地がない場合は不一致とする。
    If InStr(parenContent, "丁目") > 0 And InStr(parenContent, "番") > 0 Then
        If inputVal2 = 0 Then
            EvaluateNumericRange = -1
            Exit Function
        End If
    End If

    ' 区切り記号の統一： チルダやハイフンを「～」に、読点や中黒を「、」に統一する。
    cleanStr = parenContent
    cleanStr = Replace(cleanStr, ChrW(&HFF5E), "～")
    cleanStr = Replace(cleanStr, ChrW(&H301C), "～")
    cleanStr = Replace(cleanStr, "~", "～")
    cleanStr = Replace(cleanStr, "－", "～")
    cleanStr = Replace(cleanStr, "-", "～")
    cleanStr = Replace(cleanStr, ",", "、")
    cleanStr = Replace(cleanStr, "・", "、")
    cleanStr = Replace(cleanStr, "及び", "、")
    cleanStr = Replace(cleanStr, "および", "、")

    segments = Split(cleanStr, "、")

    ' セグメント別判定ループ： カンマで分割された各条件に対して範囲内か除外対象かを判定する。
    For Each seg In segments
        sItem = Trim(CStr(seg))
        If sItem <> "" Then
            isExclude = (InStr(sItem, "除く") > 0 Or InStr(sItem, "以外") > 0)
            sItem = StripRangeUnits(sItem)

            If InStr(sItem, "～") > 0 Then
                Dim parts() As String
                Dim rStart As Long, rEnd As Long
                parts = Split(sItem, "～")
                rStart = ExtractFirstLong(parts(0))
                rEnd = ExtractFirstLong(parts(1))

                If rStart > 0 And rEnd > 0 Then
                    hasValidRule = True
                    If inputVal1 >= rStart And inputVal1 <= rEnd Then
                        If isExclude Then
                            EvaluateNumericRange = -1
                            Exit Function
                        Else
                            matchedPositiveRule = True
                        End If
                    End If
                End If
            Else
                Dim singleVal As Long
                singleVal = ExtractFirstLong(sItem)
                If singleVal > 0 Then
                    hasValidRule = True
                    If inputVal1 = singleVal Then
                        If isExclude Then
                            EvaluateNumericRange = -1
                            Exit Function
                        Else
                            matchedPositiveRule = True
                        End If
                    End If
                End If
            End If
        End If
    Next seg

    If matchedPositiveRule Then
        EvaluateNumericRange = 1
    ElseIf hasValidRule Then
        EvaluateNumericRange = -1
    Else
        EvaluateNumericRange = 0
    End If
End Function

' ==============================================================================
' --- 範囲単位文字の除去処理 ---
' 範囲比較用に「丁目」「番地」「地割」「を除く」等の文字列を除去する。
Private Function StripRangeUnits(ByVal s As String) As String
    Dim res As String
    res = s
    res = Replace(res, "丁目", "")
    res = Replace(res, "丁", "")
    res = Replace(res, "番地", "")
    res = Replace(res, "番", "")
    res = Replace(res, "号", "")
    res = Replace(res, "地割", "")
    res = Replace(res, "第", "")
    res = Replace(res, "を除く", "")
    res = Replace(res, "以外", "")
    StripRangeUnits = res
End Function

' ==============================================================================
' --- 文字列先頭数値の取得処理 ---
' 補助関数として文字列から最初の整数を返す。
Private Function ExtractFirstLong(ByVal s As String) As Long
    ExtractFirstLong = ExtractNumberFromSuffix(s)
End Function

' ==============================================================================
' --- Excel高速化設定の切替処理 ---
' 画面描画、イベント、自動計算の有効・無効を一括で切り替える。
Private Sub ToggleOptimizations(ByVal isEnabled As Boolean)
    With Application
        .ScreenUpdating = isEnabled
        .EnableEvents = isEnabled
        .Calculation = IIf(isEnabled, xlCalculationAutomatic, xlCalculationManual)
    End With
End Sub

' ==============================================================================
' --- シート取得・生成処理 ---
' 指定名のシートを取得し、存在しない場合は末尾に新規作成する。
Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

' ==============================================================================
' --- 郵便番号7桁整形処理 ---
' ハイフンや空白を除去し、先頭ゼロ落ちの6桁数字には0を補完する。
Private Function FormatZip(ByVal rawZip As String) As String
    Dim cleanZip As String

    cleanZip = Replace(rawZip, "-", "")
    cleanZip = Replace(rawZip, " ", "")

    If Len(cleanZip) = 6 And IsNumeric(cleanZip) Then
        FormatZip = "0" & cleanZip
    Else
        FormatZip = cleanZip
    End If
End Function

' ==============================================================================
' --- 出力用郵便番号整形処理 ---
' 7桁の郵便番号文字列へハイフン（123-4567形式）を付与する。
Private Function FormatZipForOutput(ByVal cleanZip7 As String) As String
    If Len(cleanZip7) = 7 And IsNumeric(cleanZip7) Then
        FormatZipForOutput = Left(cleanZip7, 3) & "-" & Right(cleanZip7, 4)
    Else
        FormatZipForOutput = cleanZip7
    End If
End Function

' ==============================================================================
' --- セル内一致文字の赤色着色処理 ---
' 入力住所とマスタ住所をポインタ比較し、一致した先頭文字群を赤色にする。
Private Sub ColorMatchInCell(ByVal ws As Worksheet, ByVal rowIdx As Long, ByVal rawAddress As String, _
                              ByVal masterFullAddress As String, ByVal masterTownRaw As String)
    Dim targetCell As Range
    Set targetCell = ws.Cells(rowIdx, "B")

    Dim cleanMasterBaseStr As String
    Dim masterPrefAndCityNormalized As String
    Dim masterTownNormalized As String

    ' マスタ文字列の正規化と分割： 都道府県＋市区町村部分と町名部分の境界を特定する。
    cleanMasterBaseStr = NormalizeAddress(masterFullAddress)
    masterTownNormalized = NormalizeAddress(RemoveRangeSuffixTownName(masterTownRaw))
    masterPrefAndCityNormalized = Left(cleanMasterBaseStr, Len(cleanMasterBaseStr) - Len(masterTownNormalized))

    ' 着色開始ポインタの決定： 入力住所で都道府県が省略されている場合はマスタ側の県名文字数分をスキップする。
    Dim currentMatchPos As Long: currentMatchPos = 1
    Dim suffixPos As Long
    Dim masterPrefLen As Long

    suffixPos = InStr(masterPrefAndCityNormalized, "県")
    If suffixPos = 0 Then suffixPos = InStr(masterPrefAndCityNormalized, "府")
    If suffixPos = 0 Then suffixPos = InStr(masterPrefAndCityNormalized, "道")
    If suffixPos = 0 Then suffixPos = InStr(masterPrefAndCityNormalized, "都")

    If suffixPos > 0 Then
        masterPrefLen = suffixPos
    Else
        masterPrefLen = 3
    End If

    If Not IsAddressPrefectureStart(rawAddress) Then
        currentMatchPos = masterPrefLen + 1
    End If

    ' 文字単位のポインタ照合ループ： 空白、大字・字の表記差、郡省略、複合漢数字を吸収しながら一致長を計測する。
    Dim currentRawPos As Long: currentRawPos = 1
    Dim matchedCharCount As Long: matchedCharCount = 0
    Dim rawLen As Long: rawLen = Len(rawAddress)
    Dim masterLen As Long: masterLen = Len(cleanMasterBaseStr)

    Do While currentRawPos <= rawLen And currentMatchPos <= masterLen
        Dim rawChar As String
        rawChar = Mid(rawAddress, currentRawPos, 1)

        ' 空白のスキップ： 入力側のスペースはポインタのみ進める。
        If rawChar = " " Or rawChar = " " Then
            currentRawPos = currentRawPos + 1
            GoTo NextCharLoop
        End If

        ' 入力側の大字・字のスキップ： 入力側にのみ「大字」「小字」「字」がある場合は入力ポインタを進める。
        If Mid(rawAddress, currentRawPos, 2) = "大字" Or Mid(rawAddress, currentRawPos, 2) = "小字" Then
            If Mid(cleanMasterBaseStr, currentMatchPos, 2) <> Mid(rawAddress, currentRawPos, 2) Then
                currentRawPos = currentRawPos + 2
                matchedCharCount = currentRawPos - 1
                GoTo NextCharLoop
            End If
        ElseIf rawChar = "字" Then
            If Mid(cleanMasterBaseStr, currentMatchPos, 1) <> "字" Then
                currentRawPos = currentRawPos + 1
                matchedCharCount = currentRawPos - 1
                GoTo NextCharLoop
            End If
        End If

        ' マスタ側の大字・字および郡名のスキップ： 入力側で省略されている場合はマスタポインタを進める。
        If Mid(cleanMasterBaseStr, currentMatchPos, 2) = "大字" And Mid(rawAddress, currentRawPos, 2) <> "大字" Then
            currentMatchPos = currentMatchPos + 2
            GoTo NextCharLoop
        ElseIf Mid(cleanMasterBaseStr, currentMatchPos, 1) = "字" And rawChar <> "字" Then
            currentMatchPos = currentMatchPos + 1
            GoTo NextCharLoop
        End If

        Dim gunPosInMaster As Long
        gunPosInMaster = InStr(currentMatchPos, masterPrefAndCityNormalized, "郡")
        If gunPosInMaster > currentMatchPos Then
            If NormalizeAddress(rawChar) <> Mid(cleanMasterBaseStr, currentMatchPos, 1) Then
                If NormalizeAddress(rawChar) = Mid(cleanMasterBaseStr, gunPosInMaster + 1, 1) Then
                    currentMatchPos = gunPosInMaster + 1
                End If
            End If
        End If

        ' 漢数字トークンの照合： 「十二」と「１２」など文字数が変化する漢数字列を一括比較する。
        If InStr("〇一二三四五六七八九十百", rawChar) > 0 Then
            Dim kEnd As Long: kEnd = currentRawPos
            Do While kEnd < rawLen And InStr("〇一二三四五六七八九十百", Mid(rawAddress, kEnd + 1, 1)) > 0
                kEnd = kEnd + 1
            Loop
            Dim rawKanjiToken As String
            Dim normKanjiToken As String
            rawKanjiToken = Mid(rawAddress, currentRawPos, kEnd - currentRawPos + 1)
            normKanjiToken = NormalizeAddress(rawKanjiToken)

            If Mid(cleanMasterBaseStr, currentMatchPos, Len(normKanjiToken)) = normKanjiToken Then
                currentRawPos = kEnd + 1
                currentMatchPos = currentMatchPos + Len(normKanjiToken)
                matchedCharCount = currentRawPos - 1
                GoTo NextCharLoop
            End If
        End If

        ' 通常文字の1文字比較： 正規化後の1文字が一致すれば双方のポインタを進める。
        Dim rawCharNormalized As String
        Dim normalizedChar As String
        rawCharNormalized = NormalizeAddress(rawChar)
        normalizedChar = Mid(cleanMasterBaseStr, currentMatchPos, 1)

        If rawCharNormalized = normalizedChar Then
            currentRawPos = currentRawPos + 1
            currentMatchPos = currentMatchPos + 1
            matchedCharCount = currentRawPos - 1
        Else
            Exit Do
        End If
NextCharLoop:
    Loop

    ' 赤色着色の適用： 一致した文字数分だけセル先頭からのフォント色を赤に設定する。
    With targetCell
        .Font.ColorIndex = xlAutomatic
        If matchedCharCount > 0 Then
            .Characters(Start:=1, Length:=matchedCharCount).Font.Color = RGB(255, 0, 0)
        End If
    End With
End Sub

' ==============================================================================
' --- 処理中フォーム表示処理 ---
' モードレスで進捗表示用UserFormを開く。
Sub frm_msg処理中OP()
    On Error Resume Next
'    If frm_msg処理中 Is Nothing Then
'        Load frm_msg処理中
'    End If
'    frm_msg処理中.Show (vbModeless)
    DoEvents
    On Error GoTo 0
End Sub

' ==============================================================================
' --- 処理中フォーム終了処理 ---
' 表示中の進捗UserFormをアンロードする。
Sub frm_msg処理中CL()
    On Error Resume Next
'    If Not frm_msg処理中 Is Nothing Then
'        If frm_msg処理中.Visible = True Then
'            Unload frm_msg処理中
'        End If
'    End If
    DoEvents
    On Error GoTo 0
End Sub

