Option Explicit

'! VBScript 用 Unit Test
'!
'! @author  miyamiya <rai.caver@gmail.com>
'! @date    2013/08/14
'! @version 0.1
Class VBSUnit

    '! ログを記録する
    '!
    '! @var dic          object Dictionary Object
    Private dic

    '! 現在テスト中のメソッド名
    '!
    '! @var name         string メソッド名
    Private name

    '! 結果をコンソール上に表示する
    '!
    '! @param  void
    '! @return void
    Public Function getResult
        getResut = dic
    End Function

    '! メソッド名をセットする
    '!
    '! @param  theName     string テストを行うクラス名
    '! @return void
    Public Property Let setName(byval theName)
        name = theName
        dic.Add name, CreateObject("Scripting.Dictionary")
        dic(name).Add "cnt"    , 0
        dic(name).Add "success", CreateObject("Scripting.Dictionary")
        dic(name).Add "failure", CreateObject("Scripting.Dictionary")
        dic(name).Add "error"  , CreateObject("Scripting.Dictionary")
    End Property

    '! 2つの変数(数値、文字、配列)が等しくない時に失敗を報告する
    '!
    '! @see    isEqual
    '! @param  theExpected   mixed 検査する左辺
    '! @param  theActual     mixed 検査する右辺
    '! @return               text 状態(success, failure, error)
    Public Function assertEquals(theExpected, theActual)
        Dim rtn : rtn = isEqual(theExpected, theActual)
        Select Case rtn
            Case True
                assertEquals = addSuccess("Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual))
            Case False
                assertEquals = addFailure("Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual))
            Case Else
                assertEquals = addError(rtn)
        End Select
    End Function

    '! 2つの変数(数値、文字、配列)が等しい時に失敗を報告する
    '!
    '! @see    isEqual
    '! @param  theExpected     mixed 検査する左辺
    '! @param  theActual       mixed 検査する右辺
    '! @return void
    Public Function assertNotEquals(theExpected, theActual)
        Dim rtn : rtn = isEqual(theExpected, theActual)
        Select Case rtn
            Case False
                assertNotEquals = addSuccess("Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual))
            Case True
                assertNotEquals = addFailure("Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual))
            Case Else
                assertNotEquals = addError(rtn)
        End Select
    End Function

    '! 変数が FALSE の時に失敗を報告する
    '!
    '! @param  theA     boolean 検査する値
    '! @return void
    Public Function assertTrue(theA)
        If VarType(theA) <> vbBoolean Then
            assertTrue = addError("1st Args is not Boolean -> " & toStringValue(theA))
            Exit Function
        End If

        If theA = True Then
            assertTrue = addSuccess("Arg1:" & toStringValue(theA))
            Exit Function
        End If
        assertTrue = addFailure("Arg1:" & toStringValue(theA))
    End Function


    '! 変数が True の時に失敗を報告する
    '!
    '! @param  theA     boolean 検査する値
    '! @return void
    Public Function assertFalse(theA)
        If VarType(theA) <> vbBoolean Then
            assertFalse = addError("1st Args is not Boolean -> " & toStringValue(theA))
            Exit Function
        End If

        If theA = False Then
            assertFalse = addSuccess("Arg1:" & toStringValue(theA))
            Exit Function
        End If
        assertFalse = addFailure("Arg1:" & toStringValue(theA))
    End Function

    '! 変数が Null 以外の時に失敗を報告する
    '!
    '! @param  theA     mixed 検査する値
    '! @return void
    Public Function assertNull(theA)
        If IsNull(theA) = True Then
            assertNull = addSuccess("Arg1:" & toStringValue(theA))
            Exit Function
        End If
        assertNull = addFailure("Arg1:" & toStringValue(theA))
    End Function

    '! 変数が Null の時に失敗を報告する
    '!
    '! @param  theA     mixed 検査する値
    '! @return void
    Public Function assertNotNull(theA)
        If IsNull(theA) = False Then
            assertNotNull = addSuccess("Arg1:" & toStringValue(theA))
            Exit Function
        End If
        assertNotNull = addFailure("Arg1:" & toStringValue(theA))
    End Function

    '! 変数が 空でない時に失敗を報告する
    '!
    '! @param  theA     mixed 検査する値
    '! @return void
    Public Function assertEmpty(theA)
        If is_empty(theA) = True Then
            assertEmpty = addSuccess("Arg1:" & toStringValue(theA))
            Exit Function
        End If
        assertEmpty = addFailure("Arg1:" & toStringValue(theA))
    End Function

    '! 変数が 空の時に失敗を報告する
    '!
    '! @param  theA     mixed 検査する値
    '! @return void
    Public Function assertNotEmpty(theA)
        If is_empty(theA) = False Then
            assertNotEmpty = addSuccess("Arg1:" & toStringValue(theA))
            Exit Function
        End If
        assertNotEmpty = addFailure("Arg1:" & toStringValue(theA))
    End Function

    '! 第2引数の配列の中に、第1引数が含まれていない時に失敗を報告する
    '!
    '! @see    isContains
    '! @param  theNeedles     mixed 検査対象となる値
    '! @param  theHaystack    array 検査する配列
    '! @return void
    Public Function assertContains(theNeedles, theHaystack)
        If IsArray(theHaystack) = False or IsObject(theNeedles) = True Then
            assertContains = addError("Arg1:" & toStringValue(theNeedles) & ", Arg2:" & toStringValue(theHaystack))
            Exit Function
        End If

        If (isContains(theNeedles, theHaystack) = True) Then
            assertContains = addSuccess("Arg1:" & toStringValue(theNeedles) & ", Arg2:" & toStringValue(theHaystack))
            Exit Function
        End If
        assertContains = addFailure("Arg1:" & toStringValue(theNeedles) & ", Arg2:" & toStringValue(theHaystack))
    End Function

    '! 第2引数の配列の中に、第1引数が含まれている時に失敗を報告する
    '!
    '! @see    isContains
    '! @param  theNeedles     mixed 検査対象となる値
    '! @param  theHaystack    array 検査する配列
    '! @return void
    Public Function assertNotContains(theNeedles, theHaystack)
        If IsArray(theHaystack) = False or IsObject(theNeedles) = True Then
            assertNotContains = addError("Arg1:" & toStringValue(theNeedles) & ", Arg2:" & toStringValue(theHaystack))
            Exit Function
        End If

        If (isContains(theNeedles, theHaystack) = False) Then
            assertNotContains = addSuccess("Arg1:" & toStringValue(theNeedles) & ", Arg2:" & toStringValue(theHaystack))
            Exit Function
        End If
        assertNotContains = addFailure("Arg1:" & toStringValue(theNeedles) & ", Arg2:" & toStringValue(theHaystack))
    End Function


    '! 第2引数が第1引数で始まっていない時に失敗を報告する
    '!
    '! @param  thePrefix      string 検査対象となる値
    '! @param  theString      string 検査する文字列
    '! @return void
    Public Function assertStringStartsWith(thePrefix, theString)
        If (VarType(thePrefix) <> vbString Or VarType(theString) <> vbString) Then
            assertStringStartsWith = addError("Arg1:" & toStringValue(thePrefix) & ", Arg2:" & toStringValue(theString))
            Exit Function
        End If
        If (InStr(theString, thePrefix) = 1) Then
            assertStringStartsWith = addSuccess("Arg1:" & toStringValue(thePrefix) & ", Arg2:" & toStringValue(theString))
            Exit Function
        End If
        assertStringStartsWith = addFailure("Arg1:" & toStringValue(thePrefix) & ", Arg2:" & toStringValue(theString))
    End Function

    '! 第2引数が第1引数で始まっている時に失敗を報告する
    '!
    '! @param  thePrefix      string 検査対象となる値
    '! @param  theString      string 検査する文字列
    '! @return void
    Public Function assertStringStartsNotWith(thePrefix, theString)
        If (VarType(thePrefix) <> vbString Or VarType(theString) <> vbString) Then
            assertStringStartsNotWith = addError("Arg1:" & toStringValue(thePrefix) & ", Arg2:" & toStringValue(theString))
            Exit Function
        End If

        If (InStr(theString, thePrefix) = 1) Then
            assertStringStartsNotWith = addFailure("Arg1:" & toStringValue(thePrefix) & ", Arg2:" & toStringValue(theString))
            Exit Function
        End If
        assertStringStartsNotWith = addSuccess("Arg1:" & toStringValue(thePrefix) & ", Arg2:" & toStringValue(theString))
    End Function


    '! 第2引数が第1引数で終わっていない時に失敗を報告する
    '!
    '! @param  theSuffix      string 検査対象となる値
    '! @param  theString      string 検査する文字列
    '! @return void
    Public Function assertStringEndsWith(theSuffix, theString)
        If (VarType(theSuffix) <> vbString Or VarType(theString) <> vbString) Then
            assertStringEndsWith = addError("Arg1:" & toStringValue(theSuffix) & ", Arg2:" & toStringValue(theString))
            Exit Function
        End If

        If (InStr(StrReverse(theString), StrReverse(theSuffix)) = 1) Then
            assertStringEndsWith = addSuccess("Arg1:" & toStringValue(theSuffix) & ", Arg2:" & toStringValue(theString))
            Exit Function
        End If
        assertStringEndsWith = addFailure("Arg1:" & toStringValue(theSuffix) & ", Arg2:" & toStringValue(theString))
    End Function

    '! 第2引数が第1引数で終わっている時に失敗を報告する
    '!
    '! @param  theSuffix      string 検査対象となる値
    '! @param  theString      string 検査する文字列
    '! @return void
    Public Function assertStringEndsNotWith(theSuffix, theString)
        If (VarType(theSuffix) <> vbString Or VarType(theString) <> vbString) Then
            assertStringEndsNotWith = addError("Arg1:" & toStringValue(theSuffix) & ", Arg2:" & toStringValue(theString))
            Exit Function
        End If

        If (InStr(StrReverse(theString), StrReverse(theSuffix)) = 1) Then
            assertStringEndsNotWith = addFailure("Arg1:" & toStringValue(theSuffix) & ", Arg2:" & toStringValue(theString))
            Exit Function
        End If
        assertStringEndsNotWith = addSuccess("Arg1:" & toStringValue(theSuffix) & ", Arg2:" & toStringValue(theString))
    End Function


    '! 第2引数の型が第1引数でない時に失敗を報告する
    '!
    '! @see    equalInternalType
    '! @param  theExpected    string タイプ
    '!           Empty     ：未初期化
    '!           Null      ：Null 値
    '!           Integer   ：整数型
    '!           Long      ：長整数型
    '!           Single    ：単精度浮動小数点数型
    '!           Double    ：倍精度浮動小数点数型
    '!           Currency  ：通貨型
    '!           Date      ：日付型
    '!           String    ：文字列型
    '!           Object    ：オートメーションオブジェクト型
    '!           Error     ：エラー型
    '!           Boolean   ：ブール型
    '!           Variant   ：バリアント型
    '!           DataObject：非オートメーションオブジェクト型
    '!           Byte      ：バイト型
    '!           Array     ：配列
    '! @param  theActual      mixed  検査する値
    '! @return void
    Public Function assertInternalType(theExpected, theActual)
        If (VarType(theExpected) <> vbString) Then
            assertInternalType = addError("Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual))
            Exit Function
        End If
        If (equalInternalType(theExpected, theActual) = True) Then
            assertInternalType = addSuccess("Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual))
            Exit Function
        End If
        assertInternalType = addFailure("Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual))
    End Function

    '! 第2引数の型が第1引数の時に失敗を報告する
    '!
    '! @see    equalInternalType
    '! @param  theExpected    string タイプ
    '!           Empty     ：未初期化
    '!           Null      ：Null 値
    '!           Integer   ：整数型
    '!           Long      ：長整数型
    '!           Single    ：単精度浮動小数点数型
    '!           Double    ：倍精度浮動小数点数型
    '!           Currency  ：通貨型
    '!           Date      ：日付型
    '!           String    ：文字列型
    '!           Object    ：オートメーションオブジェクト型
    '!           Error     ：エラー型
    '!           Boolean   ：ブール型
    '!           Variant   ：バリアント型
    '!           DataObject：非オートメーションオブジェクト型
    '!           Byte      ：バイト型
    '!           Array     ：配列
    '! @param  theActual      mixed  検査する値
    '! @return void
    Public Function assertNotInternalType(theExpected, theActual)
        If (VarType(theExpected) <> vbString) Then
            assertNotInternalType = addError("Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual))
            Exit Function
        End If
        If (equalInternalType(theExpected, theActual) = False) Then
            assertNotInternalType = addSuccess("Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual))
            Exit Function
        End If
        assertNotInternalType = addFailure("Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual))
    End Function






    '! 第1引数の場所にあるテキストファイルの中から第2引数の文字が含まれていない時に失敗を報告する
    '!
    '| @see    doSearchText
    '! @param  theExpectedFile    string 検査対象となるファイルのパス
    '! @param  theActualString    string 検査する文字列
    '! @return void
    Public Function assertStringEqualsFile(theExpectedFile, theActualString)
        Dim txt : txt = "Arg1:" & toStringValue(theExpectedFile) & ", Arg2:" & toStringValue(theActualString)
        If (file_exists(theExpectedFile) = True and VarType(theActualString) = vbString) Then
            If (doSearchText(theExpectedFile, theActualString) = True) Then
                assertStringEqualsFile = addSuccess(txt)
                Exit Function
            End If
            assertStringEqualsFile = addFailure(txt)
            Exit Function
        End If
        assertStringEqualsFile = addError(txt)
    End Function

    '! 第1引数の場所にあるテキストファイルの中から第2引数の文字が含るまれていない時に失敗を報告する
    '!
    '| @see    doSearchText
    '! @param  thePrefix      string 検査対象となるファイルのパス
    '! @param  theString      string 検査する文字列
    '! @return void
    Public Function assertStringNotEqualsFile(theExpectedFile, theActualString)
        Dim txt : txt = "Arg1:" & toStringValue(theExpectedFile) & ", Arg2:" & toStringValue(theActualString)
        If (file_exists(theExpectedFile) = True and VarType(theActualString) = vbString) Then
            If (doSearchText(theExpectedFile, theActualString) = True) Then
                assertStringNotEqualsFile = addFailure(txt)
                Exit Function
            End If
            assertStringNotEqualsFile = addSuccess(txt)
            Exit Function
        End If
        assertStringNotEqualsFile = addError(txt)
    End Function

    '! 第1引数の場所にあるテキストファイルと第2引数の場所にあるテキストファイルが違う時に失敗を報告する
    '!
    '| @see    doFileDiff
    '! @param  theExpected    string 検査対象となるファイルのパス
    '! @param  theActual      string 比較対象となるファイルのパス
    '! @return void
    Public Function assertFileEquals(theExpected, theActual)
        Select Case doFileDiff(theExpected, theActual)
            Case True
                assertFileEquals = addSuccess("Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual))
            Case False
                assertFileEquals = addFailure("Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual))
            Case Else
                assertFileEquals = addError("Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual))
        End Select
    End Function


    '! 第1引数の場所にあるテキストファイルと第1引数の場所にあるテキストファイルが同じ時に失敗を報告する
    '!
    '| @see    doFileDiff
    '! @param  theExpected    string 検査対象となるファイルのパス
    '! @param  theActual      string 比較対象となるファイルのパス
    '! @return void
    Public Function assertFileNotEquals(theExpected, theActual)
        Dim txt : txt = "Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual)
        Select Case doFileDiff(theExpected, theActual)
            Case True
                assertFileNotEquals = addFailure(txt)
            Case False
                assertFileNotEquals = addSuccess(txt)
            Case Else
                assertFileNotEquals = addError(txt)
        End Select
    End Function

    '! 第1引数の場所にファイルが存在しない時に失敗を報告する
    '!
    '| @see    file_exists
    '! @param  theFileName      string 検査対象となるファイルのパス
    '! @return void
    Public Function assertFileExists(theFileName)
        Dim txt : txt = "Arg1:" & toStringValue(theFileName)
        Select Case file_exists(theFileName)
            Case True
                assertFileExists = addSuccess(txt)
            Case False
                assertFileExists = addFailure(txt)
            Case Else
                assertFileExists = addError(txt)
        End Select
    End Function

    '! 第1引数の場所にファイルが存在する時に失敗を報告する
    '!
    '| @see    file_exists
    '! @param  theFileName      string 検査対象となるファイルのパス
    '! @return void
    Public Function assertFileNotExists(theFileName)
        Dim txt : txt = "Arg1:" & toStringValue(theFileName)
        Select Case file_exists(theFileName)
            Case True
                assertFileNotExists = addFailure(txt)
            Case False
                assertFileNotExists = addSuccess(txt)
            Case Else
                assertFileNotExists = addError(txt)
        End Select
    End Function

    '! theExpected < theActual の時に成功、それ以外で失敗を報告する
    '!
    '! @param  theExpected    比較する数値の左辺
    '! @param  theActual      比較する数値の右辺
    '! @return void
    Public Function assertGreaterThan(theExpected, theActual)
        Dim txt : txt = "Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual)
        If (is_numeric(theExpected) = False or is_numeric(theActual) = False) Then
            assertGreaterThan = addError(txt)
            Exit Function
        End If
        If (theExpected < theActual) Then
            assertGreaterThan = addSuccess(txt)
            Exit Function
        End If
        assertGreaterThan = addFailure(txt)
    End Function

    '! theExpected <= theActual の時に成功、それ以外で失敗を報告する
    '!
    '! @param  theExpected    比較する数値の左辺
    '! @param  theActual      比較する数値の右辺
    '! @return void
    Public Function assertGreaterThanOrEqual(theExpected, theActual)
        Dim txt : txt = "Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual)
        If (is_numeric(theExpected) = False or is_numeric(theActual) = False) Then
            assertGreaterThanOrEqual = addError(txt)
            Exit Function
        End If
        If (theExpected <= theActual) Then
            assertGreaterThanOrEqual = addSuccess(txt)
            Exit Function
        End If
        assertGreaterThanOrEqual = addFailure(txt)
    End Function

    '! theExpected > theActual の時に成功、それ以外で失敗を報告する
    '!
    '! @param  theExpected    比較する数値の左辺
    '! @param  theActual      比較する数値の右辺
    '! @return void
    Public Function assertLessThan(theExpected, theActual)
        Dim txt : txt = "Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual)
        If (is_numeric(theExpected) = False or is_numeric(theActual) = False) Then
            assertLessThan = addError(txt)
            Exit Function
        End If
        If (theExpected > theActual) Then
            assertLessThan = addSuccess(txt)
            Exit Function
        End If
        assertLessThan = addFailure(txt)
    End Function

    '! theExpected >= theActual の時に成功、それ以外で失敗を報告する
    '!
    '! @param  theExpected    比較する数値の左辺
    '! @param  theActual      比較する数値の右辺
    '! @return void
    Public Function assertLessThanOrEqual(theExpected, theActual)
        Dim txt : txt = "Arg1:" & toStringValue(theExpected) & ", Arg2:" & toStringValue(theActual)
        If (is_numeric(theExpected) = False or is_numeric(theActual) = False) Then
            assertLessThanOrEqual = addError(txt)
            Exit Function
        End If
        If (theExpected >= theActual) Then
            assertLessThanOrEqual = addSuccess(txt)
            Exit Function
        End If
        assertLessThanOrEqual = addFailure(txt)
    End Function

    '! 第2引数の配列の要素数と第1引数が違う時に失敗を報告する
    '!
    '! @param  theExpectedCount  integer 比較する数字
    '! @param  theHaystack       array 対象の配列
    '! @return void
    Public Function assertCount(theExpectedCount, theHaystack)
        Dim txt : txt = "Arg1:" & toStringValue(theExpectedCount) & ", Arg2:" & toStringValue(theHaystack)
        If (is_numeric(theExpectedCount) = False Or usedArray(theHaystack) = False) Then
            assertCount = addError(txt)
            Exit Function
        End If
        If ((UBound(theHaystack) - LBound(theHaystack) + 1) = theExpectedCount) Then
            assertCount = addSuccess(txt)
            Exit Function
        End If
        assertCount = addFailure(txt)
    End Function


    '! 第2引数の配列の要素数と第1引数が同じ時に失敗を報告する
    '!
    '! @param  theExpectedCount  integer 比較する数字
    '! @param  theHaystack       array 対象の配列
    '! @return void
    Public Function assertNotCount(theExpectedCount, theHaystack)
        Dim txt : txt = "Arg1:" & toStringValue(theExpectedCount) & ", Arg2:" & toStringValue(theHaystack)
        If (is_numeric(theExpectedCount) = False Or usedArray(theHaystack) = False) Then
            assertNotCount = addError(txt)
            Exit Function
        End If
        If ((UBound(theHaystack) - LBound(theHaystack) + 1) = theExpectedCount) Then
            assertNotCount = addFailure(txt)
            Exit Function
        End If
        assertNotCount = addSuccess(txt)
    End Function

    '--- Private Method ---

    '! ログに吐き出せる形式に変換する
    '!
    '! @param  theArg        mixed 変換する引数
    '! @return               String 形式
    Private Function toStringValue(theArg)
        Dim i, sp, txt
        If(IsObject(theArg)) Then
            toStringValue = "is Object"
            Exit Function
        End If

        If(IsArray(theArg)) Then
            sp  = ""
            txt = "Array{"
            If (usedArray(theArg) = True) Then
                For i = LBound(theArg) to UBound(theArg) Step 1
                    txt = txt & sp & i & " => " & theArg(i)
                    sp = ", "
                Next
            End If
            txt = "}"
            toStringValue = txt
            Exit Function
        End If

        If VarType(theArg) = vbBoolean Then
            If (theArg = True) Then
                toStringValue = "True(boolean)"
                Exit Function
            End If
            toStringValue = "False(boolean)"
            Exit Function
        End If

        toStringValue = theArg
    End Function

    '! コンストラクタ
    '!
    '! @param  void
    '! @return void
    Private Sub Class_Initialize()
        Set dic = CreateObject("Scripting.Dictionary")
    End Sub

    '! デストラクタ
    '!
    '! @param  void
    '! @return void
    Private Sub Class_Terminate()
        Set dic = Nothing
    End Sub

    '! 成功のカウンターを進める
    '!
    '! @param  theMsg     string 成功メッセージ
    '! @return void
    Private Function addSuccess(theMsg)
        With dic.item(name)
            .item("cnt") = .item("cnt") + 1
            .item("success").Add .item("cnt"), theMsg
        End With
        addSuccess = "success"
    End Function

    '! 失敗のカウンターを進める
    '!
    '! @param  theMsg     string 失敗メッセージ
    '! @return void
    Private Function addFailure(theMsg)
        With dic.item(name)
            .item("cnt") = .item("cnt") + 1
            .item("failure").Add .item("cnt"), theMsg
        End With
        addFailure = "failure"
    End Function

   '! エラーのカウンターを進める
    '!
    '! @param  theMsg     string エラーメッセージ
    '! @return void
    Private Function addError(theMsg)
        With dic.item(name)
            .item("cnt") = .item("cnt") + 1
            .item("error").Add .item("cnt"), theMsg
        End With
        addError = "error"
    End Function


    '! 2つの変数(数値、文字、配列)が同じか比較する
    '!
    '! @see    isValueEqual, isArrayEqual
    '! @param  theExpected   mixed 検査する左辺
    '! @param  theActual     mixed 検査する右辺
    '! @return               string success:等しい、failure:等しくない, error: エラー
    Private Function isEqual(theExpected, theActual)
        If (IsObject(theExpected)) Then
            isEqual = "1st Args is Object"
            Exit Function
        End If

        If (IsObject(theActual)) Then
            isEqual = "2st Args is Object"
            Exit Function
        End If

        If IsArray(theExpected) = True And IsArray(theActual) = True Then
            isEqual = isArrayEqual(theExpected, theActual)
            Exit Function
        End If

        isEqual = isValueEqual(theExpected, theActual)
    End Function

    '! 2つの配列が同じか比較する
    '!
    '! @see    isValueEqual
    '! @param  theExpected   array 検査する左辺
    '! @param  theActual     array 検査する右辺
    '! @return               boolean True:等しい、False:等しくない
    Private Function isArrayEqual(theExpected, theActual)
        If IsArray(theExpected) = False Or IsArray(theActual) = False Then
            isArrayEqual = False
            Exit Function
        End If

        If (Ubound(theExpected) <> Ubound(theActual)) Then
            isArrayEqual = False
            Exit Function
        End If

        Dim i
        For i = 0 To Ubound(theExpected) Step 1
            If (isValueEqual(theExpected(i), theActual(i)) = False) Then
                isArrayEqual = False
                Exit Function
            End If
        Next
        isArrayEqual = True
    End Function

    '! 2つの変数(数値、文字)が同じか比較する
    '!
    '! @param  theExpected   mixed 検査する左辺
    '! @param  theActual     mixed 検査する右辺
    '! @return               boolean True:等しい、False:等しくない
    Private Function isValueEqual(theExpected, theActual)
        If IsArray(theExpected) = True Or IsArray(theActual) = True _
           Or IsObject(theExpected) = True Or IsObject(theActual) = True Then
            isValueEqual = False
            Exit Function
        End If

        If IsNull(theExpected) = True And IsNull(theActual) = True Then
            isValueEqual = True
            Exit Function
        End If

        If (VarType(theExpected) = VarType(theActual) And theExpected = theActual) Then
            isValueEqual = True
            Exit Function
        End If
        isValueEqual = False
    End Function


    '! 配列が使用可能かどうか判定する
    '!
    '! @param  theA     array 検査する値
    '! @return void
    Private Function usedArray(theA)
        If (IsArray(theA) = True) Then
            Dim i
            For Each i In theA
                usedArray = True
                Exit Function
            Next
        End If
        usedArray = False
    End Function

    '! 値が空と同等の値かどうか確認
    '!
    '! @param  theA     mixed 検査する値
    '! @return void
    Private Function is_empty(theA)
        If (IsObject(theA) = True) Then
            is_empty = False
            Exit Function
        End If

        If (VarType(theA) = vbEmpty Or VarType(theA) = vbNull) Then
            is_empty = True
            Exit Function
        End If

        If (IsArray(theA) = True) Then
            If (usedArray(theA) = False) Then
                is_empty = True
                Exit Function
            End If
            is_empty = False
            Exit Function
        End If

        If (VarType(theA) = vbInteger Or VarType(theA) = vbLong _
            Or VarType(theA) = vbSingle Or VarType(theA) = vbDouble) Then
            If (theA = 0) Then
                is_empty = True
                Exit Function
            End If
        End If

        If (theA = "" Or theA = False) Then
            is_empty = True
            Exit Function
        End If
        is_empty = False
    End Function

    '! 第2引数の配列の中に、第1引数が含まれている時に失敗を報告する
    '!
    '! @param  theNeedles     mixed 検査対象となる値
    '! @param  theHaystack    array 検査する配列
    '! @return void
    Private Function isContains(theNeedles, theHaystack)
        Dim c, t, j

        If usedArray(theHaystack) = False Then
            isContains = False
            Exit Function
        End If

        t = "'" & join(theHaystack, "','") & "'"
        If IsArray(theNeedles) = False Then
            theNeedles = Array(theNeedles)
        End If

        For Each c In theNeedles
            j = InStr(t, "'" & c & "'")
            If IsNumeric(j) = True Then
                If j > 0 Then
                    isContains = True
                    Exit Function
                End If
            End If
        Next
    End Function

    '!  第2引数の型が第1引数の時に True を返す
    '!
    '! @param  theExpected    string タイプ
    '! @param  theActual      mixed  検査する値
    '! @return boolean
    Private Function equalInternalType(theExpected, theActual)
        If (LCase(theExpected) = "array") Then
            If (IsArray(theActual) = True) Then
                equalInternalType = True
                Exit Function
            End If
            equalInternalType = False
            Exit Function
        End If

        If (LCase(theExpected) = "object") Then
            If (IsObject(theActual) = True) Then
                equalInternalType = True
                Exit Function
            End If
            equalInternalType = False
            Exit Function
        End If

        If (eval("vb"&theExpected) = VarType(theActual)) Then
            equalInternalType = True
            Exit Function
        End If
        equalInternalType = False
    End Function

    '! theExpectedFile の場所にあるテキストファイルの中から theActualString を検索する
    '!
    '! @param  thePrefix      string 検査対象となるファイルのパス
    '! @param  theString      string 検査する文字列
    '! @return boolean        文字列がヒットしたら True, しない時に False
    Private Function doSearchText(theExpectedFile, theActualString)
        If (CreateObject("Scripting.FileSystemObject").GetFile(theExpectedFile).size = 0) Then
            doSearchText = False
            Exit Function
        End If
        Dim txt
        txt = CreateObject("Scripting.FileSystemObject").OpenTextFile(theExpectedFile).ReadAll()
        If (InStr(txt, theActualString) = 0 OR IsNull(InStr(txt, theActualString)) = True) Then
            doSearchText = False
            Exit Function
        End If
        doSearchText = True
    End Function

    '! 第1引数の場所にあるテキストファイルと第2引数の場所にあるテキストファイルを比較する
    '!
    '! @param  theExpected    string 検査対象となるファイルのパス
    '! @param  theActual      string 比較対象となるファイルのパス
    '! @return void
    Private Function doFileDiff(theExpected, theActual)
        If (file_exists(theExpected) = True And file_exists(theActual) = True) Then
            Dim expected, actual
            expected = CreateObject("Scripting.FileSystemObject").OpenTextFile(theExpected).ReadAll()
            actual = CreateObject("Scripting.FileSystemObject").OpenTextFile(theActual).ReadAll()
            If (expected = actual) Then
                doFileDiff = True
                Exit Function
            End If
            doFileDiff = False
            Exit Function
        End If
        doFileDiff = Null
    End Function

    '! ファイルの存在確認
    '!
    '! @param  theFileName      string 検査対象となるファイルのパス
    '! @return                  True: 存在、False: なし、Null: 引数エラー
    Private Function file_exists(theFileName)
        If (VarType(theFileName) <> vbString) Then
            file_exists = Null
            Exit Function
        End If
        file_exists = CreateObject("Scripting.FileSystemObject").FileExists(theFileName)
    End Function

    '! 引数が数値型なら True を返す
    '!
    '! @param  theValue    型を確認する値
    '! @return             True: 数値、False: 数値以外
    Private Function is_numeric(theValue)
        Select Case VarType(theValue)
            Case vbInteger, vbLong, vbSingle, vbDouble
                is_numeric = True
            Case Else
                is_numeric = False
        End Select
    End Function
End Class
