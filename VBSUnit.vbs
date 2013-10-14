Option Explicit

'! VBScript 用 Unit Test
'!
'! @author  miyamiya <rai.caver@gmail.com>
'! @date    2013/08/14
'! @version 0.2
Class VBSUnit

    '! ログを記録する
    '!
    '! @var dic          object Dictionary Object
    Private dic

    '! 現在テスト中のメソッド名
    '!
    '! @var name         string メソッド名
    Private name

    '! 結果を返す
    '!
    '! @param  void
    '! @return           object Dictionary Object
    Public Function getResult()
        Set getResult = dic
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

    '! 2つの変数(数値、文字、配列)が等しいときに成功
    '!
    '! @see    assertEqualsLogic
    '! @param  theExpected   mixed 検査する左辺
    '! @param  theActual     mixed 検査する右辺
    '! @return               text 状態(success, failure, error)
    Public Function assertEquals(theExpected, theActual)
        assertEquals = assertEqualsLogic(theExpected, theActual, True)
    End Function

    '! 2つの変数(数値、文字、配列)が等しくないとき成功
    '!
    '! @see    assertEqualsLogic
    '! @param  theExpected     mixed 検査する左辺
    '! @param  theActual       mixed 検査する右辺
    '! @return void
    Public Function assertNotEquals(theExpected, theActual)
          assertNotEquals = assertEqualsLogic(theExpected, theActual, False)
    End Function

    '! 変数が True のときに成功
    '!
    '! @see    assertBooleanLogic
    '! @param  theCondition   boolean 検査する値
    '! @return void
    Public Function assertTrue(theCondition)
        assertTrue = assertBooleanLogic(theCondition, True)
    End Function

    '! 変数が False のときに成功
    '!
    '! @see    assertBooleanLogic
    '! @param  theCondition   boolean 検査する値
    '! @return void
    Public Function assertFalse(theCondition)
        assertFalse = assertBooleanLogic(theCondition, False)
    End Function

    '! 変数が Null のときに成功
    '!
    '! @param  theCondition   mixed 検査する値
    '! @return void
    Public Function assertNull(theCondition)
        If IsNull(theCondition) = True Then
            assertNull = addSuccess(gen1ArgToMsg(theCondition))
            Exit Function
        End If
        assertNull = addFailure(gen1ArgToMsg(theCondition))
    End Function

    '! 変数が Null 以外のときに成功
    '!
    '! @param  theA     mixed 検査する値
    '! @return void
    Public Function assertNotNull(theCondition)
        If IsNull(theCondition) = False Then
            assertNotNull = addSuccess(gen1ArgToMsg(theCondition))
            Exit Function
        End If
        assertNotNull = addFailure(gen1ArgToMsg(theCondition))
    End Function


    '! 変数の値が空のときに成功
    '!
    '! @param  theCondition   mixed 検査する値
    '! @return void
    Public Function assertEmpty(theCondition)
        If is_empty(theCondition) = True Then
            assertEmpty = addSuccess(gen1ArgToMsg(theCondition))
            Exit Function
        End If
        assertEmpty = addFailure(gen1ArgToMsg(theCondition))
    End Function

    '! 変数の値が空でないときに成功
    '!
    '! @param  theCondition   mixed 検査する値
    '! @return void
    Public Function assertNotEmpty(theCondition)
        If is_empty(theCondition) = False Then
            assertNotEmpty = addSuccess(gen1ArgToMsg(theCondition))
            Exit Function
        End If
        assertNotEmpty = addFailure(gen1ArgToMsg(theCondition))
    End Function

    '! 第2引数の配列の中に、第1引数が含まれているときに成功
    '!
    '! @see    assertContainsLogic
    '! @param  theNeedles     mixed 検査対象となる値
    '! @param  theHaystack    array 検査する配列
    '! @return void
    Public Function assertContains(theNeedles, theHaystack)
        assertContains = assertContainsLogic(theNeedles, theHaystack, True)
    End Function

    '! 第2引数の配列の中に、第1引数が含まれていないときに成功
    '!
    '! @see    assertContainsLogic
    '! @param  theNeedles     mixed 検査対象となる値
    '! @param  theHaystack    array 検査する配列
    '! @return void
    Public Function assertNotContains(theNeedles, theHaystack)
        assertNotContains = assertContainsLogic(theNeedles, theHaystack, False)
    End Function

    '! 第2引数が第1引数で始まっているときに成功
    '!
    '! @see    assertStringStartsLogic
    '! @param  thePrefix      string 検査対象となる値
    '! @param  theString      string 検査する文字列
    '! @return void
    Public Function assertStringStartsWith(thePrefix, theString)
        assertStringStartsWith = assertStringStartsLogic(thePrefix, theString, False, True)
    End Function

    '! 第2引数が第1引数で始まっていないときに成功
    '!
    '! @see    assertStringStartsLogic
    '! @param  thePrefix      string 検査対象となる値
    '! @param  theString      string 検査する文字列
    '! @return void
    Public Function assertStringStartsNotWith(thePrefix, theString)
        assertStringStartsNotWith = assertStringStartsLogic(thePrefix, theString, False, False)
    End Function

    '! 第2引数が第1引数で終わっているときに成功
    '!
    '! @param  theSuffix      string 検査対象となる値
    '! @param  theString      string 検査する文字列
    '! @return void
    Public Function assertStringEndsWith(theSuffix, theString)
        assertStringEndsWith = assertStringStartsLogic(theSuffix, theString, True, True)
    End Function

    '! 第2引数が第1引数で終わっていないときに成功
    '!
    '! @param  theSuffix      string 検査対象となる値
    '! @param  theString      string 検査する文字列
    '! @return void
    Public Function assertStringEndsNotWith(theSuffix, theString)
        assertStringEndsNotWith = assertStringStartsLogic(theSuffix, theString, True, False)
    End Function

    '! 第2引数の型が第1引数のときに成功
    '!
    '! @see    assertInternalTypeLogic
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
        assertInternalType = assertInternalTypeLogic(theExpected, theActual, True)
    End Function

    '! 第2引数の型が第1引数のでないときに成功
    '!
    '! @see    assertInternalTypeLogic
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
        assertNotInternalType = assertInternalTypeLogic(theExpected, theActual, False)
    End Function

    '! 第1引数の場所にあるテキストファイルの中から第2引数の文字が含まれているときに成功
    '!
    '| @see    assertStringEqualsFileLogic
    '! @param  theExpectedFile    string 検査対象となるファイルのパス
    '! @param  theActualString    string 検査する文字列
    '! @return void
    Public Function assertStringEqualsFile(theExpectedFile, theActualString)
        assertStringEqualsFile = assertStringEqualsFileLogic(theExpectedFile, theActualString, True)
    End Function

    '! 第1引数の場所にあるテキストファイルの中から第2引数の文字が含まれていないときに成功
    '!
    '| @see    assertStringEqualsFileLogic
    '! @param  thePrefix      string 検査対象となるファイルのパス
    '! @param  theString      string 検査する文字列
    '! @return void
    Public Function assertStringNotEqualsFile(theExpectedFile, theActualString)
        assertStringNotEqualsFile = assertStringEqualsFileLogic(theExpectedFile, theActualString, False)
    End Function

    '! 第1引数の場所にあるテキストファイルと第2引数の場所にあるテキストファイルが同じときに成功
    '!
    '| @see    assertFileEqualsLogic
    '! @param  theExpected    string 検査対象となるファイルのパス
    '! @param  theActual      string 比較対象となるファイルのパス
    '! @return void
    Public Function assertFileEquals(theExpected, theActual)
        assertFileEquals = assertFileEqualsLogic(theExpected, theActual, True)
    End Function


    '! 第1引数の場所にあるテキストファイルと第1引数の場所にあるテキストファイルが違うときに成功
    '!
    '| @see    assertFileEqualsLogic
    '! @param  theExpected    string 検査対象となるファイルのパス
    '! @param  theActual      string 比較対象となるファイルのパス
    '! @return void
    Public Function assertFileNotEquals(theExpected, theActual)
        assertFileNotEquals = assertFileEqualsLogic(theExpected, theActual, False)
    End Function

    '! 第1引数の場所にファイルが存在しない時に失敗を報告する
    '!
    '| @see    file_exists
    '! @param  theFileName      string 検査対象となるファイルのパス
    '! @return void
    Public Function assertFileExists(theFileName)
        Select Case file_exists(theFileName)
            Case True
                assertFileExists = addSuccess(gen1ArgToMsg(theFileName))
            Case False
                assertFileExists = addFailure(gen1ArgToMsg(theFileName))
            Case Else
                assertFileExists = addError("File Not Found."_
                & vbLf & gen1ArgToMsg(theFileName))
        End Select
    End Function

    '! 第1引数の場所にファイルが存在する時に失敗を報告する
    '!
    '| @see    file_exists
    '! @param  theFileName      string 検査対象となるファイルのパス
    '! @return void
    Public Function assertFileNotExists(theFileName)
        Select Case file_exists(theFileName)
            Case True
                assertFileNotExists = addFailure(gen1ArgToMsg(theFileName))
            Case False
                assertFileNotExists = addSuccess(gen1ArgToMsg(theFileName))
            Case Else
                assertFileNotExists = addError("File Not Found."_
                & vbLf & gen1ArgToMsg(theFileName))
        End Select
    End Function

    '! theExpected < theActual の時に成功、それ以外で失敗を報告する
    '!
    '! @see    assertRelationalOperatorLogic
    '! @param  theExpected    比較する数値の左辺
    '! @param  theActual      比較する数値の右辺
    '! @return void
    Public Function assertGreaterThan(theExpected, theActual)
        assertGreaterThan = assertRelationalOperatorLogic(theExpected, theActual, "<")
    End Function

    '! theExpected <= theActual の時に成功、それ以外で失敗を報告する
    '!
    '! @see    assertRelationalOperatorLogic
    '! @param  theExpected    比較する数値の左辺
    '! @param  theActual      比較する数値の右辺
    '! @return void
    Public Function assertGreaterThanOrEqual(theExpected, theActual)
        assertGreaterThanOrEqual = assertRelationalOperatorLogic(theExpected, theActual, "<=")
    End Function

    '! theExpected > theActual の時に成功、それ以外で失敗を報告する
    '!
    '! @see    assertRelationalOperatorLogic
    '! @param  theExpected    比較する数値の左辺
    '! @param  theActual      比較する数値の右辺
    '! @return void
    Public Function assertLessThan(theExpected, theActual)
        assertLessThan = assertRelationalOperatorLogic(theExpected, theActual, ">")
    End Function

    '! theExpected >= theActual の時に成功、それ以外で失敗を報告する
    '!
    '! @see    assertRelationalOperatorLogic
    '! @param  theExpected    比較する数値の左辺
    '! @param  theActual      比較する数値の右辺
    '! @return void
    Public Function assertLessThanOrEqual(theExpected, theActual)
        assertLessThanOrEqual = assertRelationalOperatorLogic(theExpected, theActual, ">=")
    End Function

    '! 第2引数の配列の要素数と第1引数が違う時に失敗を報告する
    '!
    '! @see    assertCountLogic
    '! @param  theExpectedCount  integer 比較する数字
    '! @param  theHaystack       array 対象の配列
    '! @return void
    Public Function assertCount(theExpectedCount, theHaystack)
        assertCount = assertCountLogic(theExpectedCount, theHaystack, True)
    End Function


    '! 第2引数の配列の要素数と第1引数が同じ時に失敗を報告する
    '!
    '! @see    assertCountLogic
    '! @param  theExpectedCount  integer 比較する数字
    '! @param  theHaystack       array 対象の配列
    '! @return void
    Public Function assertNotCount(theExpectedCount, theHaystack)
        assertNotCount = assertCountLogic(theExpectedCount, theHaystack, False)
    End Function

    '! assertEquals, assertNotEquals のロジック部分
    '!
    '! @param  theExpectedCount  integer 比較する数字
    '! @param  theHaystack       array 対象の配列
    '! @param  theFlg            boolean 成功の時の返り値
    '! @return                   text 状態(success, failure, error)
    Private Function assertCountLogic(theExpectedCount, theHaystack, theFlg)
        If (is_numeric(theExpectedCount) = False) Then
            assertCountLogic = addError("1st Arg Not Numeric."_
                & vbLf & gen2ArgToMsg(theExpectedCount, theHaystack))
            Exit Function
        End If
        If (usedArray(theHaystack) = False) Then
            assertCountLogic = addError("2st Arg cann't used array."_
                & vbLf & gen2ArgToMsg(theExpectedCount, theHaystack))
            Exit Function
        End If

        If (((UBound(theHaystack) - LBound(theHaystack) + 1) = theExpectedCount) = theFlg) Then
            assertCountLogic = addSuccess(gen2ArgToMsg(theExpectedCount, theHaystack))
            Exit Function
        End If
        assertCountLogic = addFailure(gen2ArgToMsg(theExpectedCount, theHaystack))
    End Function

    '--- Private Method (Logic) ---------------------------------------------

    '! assertEquals, assertNotEquals のロジック部分
    '!
    '! @see    isEqual
    '! @param  theExpected   mixed   検査する左辺
    '! @param  theActual     mixed   検査する右辺
    '! @param  theFlg        boolean 成功の時の返り値
    '! @return               text 状態(success, failure, error)
    Private Function assertEqualsLogic(theExpected, theActual, theFlg)
        On Error Resume Next
        Dim rtn : rtn  = isEqual(theExpected, theActual)
        Dim emsg: emsg = genErrorMsg()
        On Error Goto 0

        If (emsg <> "") Then
            assertEqualsLogic = addError(_
                emsg & vbLf & gen2ArgToMsg(theExpected, theActual))
            Exit Function
        End If

        Select Case rtn
            Case theFlg
                assertEqualsLogic = addSuccess(gen2ArgToMsg(theExpected, theActual))
            Case Else
                assertEqualsLogic = addFailure(gen2ArgToMsg(theExpected, theActual))
        End Select
    End Function

    '! assertEquals, assertNotEquals のロジック部分
    '!
    '! @param  theCondition  mixed   検査する値
    '! @param  theFlg        boolean 成功の時の返り値
    '! @return               text 状態(success, failure, error)
    Private Function assertBooleanLogic(theCondition, theFlg)
        If VarType(theCondition) <> vbBoolean Then
            assertBooleanLogic = addError("1st Args is not Boolean"_
                & vbLf & gen1ArgToMsg(theCondition))
            Exit Function
        End If

        Select Case theCondition
            Case theFlg
                assertBooleanLogic = addSuccess(gen1ArgToMsg(theCondition))
            Case Else
                assertBooleanLogic = addFailure(gen1ArgToMsg(theCondition))
        End Select
    End Function

    '! assertContains, assertNotContains のロジック部分
    '!
    '! @param  theNeedles     mixed 検査対象となる値
    '! @param  theHaystack    array 検査する配列
    '! @param  theFlg         boolean 成功の時の返り値
    '! @return                text 状態(success, failure, error)
    Private Function assertContainsLogic(theNeedles, theHaystack, theFlg)
        If IsArray(theHaystack) = False or IsObject(theNeedles) = True Then
            assertContainsLogic = addError("Args Type Error"_
                & vbLf & gen2ArgToMsg(theNeedles, theHaystack))
            Exit Function
        End If

        If (isContains(theNeedles, theHaystack) = theFlg) Then
            assertContainsLogic = addSuccess(gen2ArgToMsg(theNeedles, theHaystack))
            Exit Function
        End If
        assertContainsLogic = addFailure(gen2ArgToMsg(theNeedles, theHaystack))
    End Function

    '! assertStringStartsWith, assertStringStartsNotWith, assertStringEndsWith, assertStringEndsNotWith のロジック部分
    '!
    '! @param  theTarget      string 検査対象となる値
    '! @param  theString      string 検査する文字列
    '! @param  theReverse     boolean 文字列を反転するなら True
    '! @param  theFlg         boolean 成功の時の返り値
    '! @return                text 状態(success, failure, error)
    Private Function assertStringStartsLogic(byval theTarget, byval theString, theReverse, theFlg)
        If (VarType(theTarget) <> vbString Or VarType(theString) <> vbString) Then
            assertStringStartsLogic = addError("Args is Not String"_
                & vbLf & gen2ArgToMsg(theTarget, theString))
            Exit Function
        End If

        If (theReverse = True) Then
            theTarget = StrReverse(theTarget)
            theString = StrReverse(theString)
        End If

        If ((InStr(theString, theTarget) = 1) = theFlg) Then
            assertStringStartsLogic = addSuccess(gen2ArgToMsg(theTarget, theString))
            Exit Function
        End If
        assertStringStartsLogic = addFailure(gen2ArgToMsg(theTarget, theString))
    End Function

    '! assertInternalType, assertNotInternalType のロジック部分
    '!
    '! @see    equalInternalType
    '! @param  theExpected   string  タイプ
    '! @param  theActual     mixed   検査する右辺
    '! @param  theFlg        boolean 成功の時の返り値
    '! @return               text 状態(success, failure, error)
    Private Function assertInternalTypeLogic(theExpected, theActual, theFlg)

        If (VarType(theExpected) <> vbString) Then
            assertInternalTypeLogic = addError("1st Args is not String"_
                & vbLf & gen2ArgToMsg(theExpected, theActual))
            Exit Function
        End If

        If (equalInternalType(theExpected, theActual) = theFlg) Then
            assertInternalTypeLogic = addSuccess(gen2ArgToMsg(theExpected, theActual))
            Exit Function
        End If
        assertInternalTypeLogic = addFailure(gen2ArgToMsg(theExpected, theActual))
    End Function

    '! assertStringEqualsFile, assertStringNotEqualsFile のロジック部分
    '!
    '| @see    doSearchText
    '! @param  thePrefix      string 検査対象となるファイルのパス
    '! @param  theString      string 検査する文字列
    '! @param  theFlg         boolean 成功の時の返り値
    '! @return                text 状態(success, failure, error)
    Private Function assertStringEqualsFileLogic(theExpectedFile, theActualString, theFlg)
        If (file_exists(theExpectedFile) = False) Then
            assertStringEqualsFileLogic = addError("File Not Found"_
                & vbLf & gen2ArgToMsg(theExpectedFile, theActualString))
            Exit Function
        End If

        If (VarType(theActualString) <> vbString) Then
            assertStringEqualsFileLogic = addError("2st Args is not String"_
                & vbLf & gen2ArgToMsg(theExpectedFile, theActualString))
            Exit Function
        End If

        If ((doSearchText(theExpectedFile, theActualString) = True) = theFlg) Then
            assertStringEqualsFileLogic = addSuccess(gen2ArgToMsg(theExpectedFile, theActualString))
            Exit Function
        End If

        assertStringEqualsFileLogic = addFailure(gen2ArgToMsg(theExpectedFile, theActualString))
    End Function

    '! assertFileEquals, assertFileNotEquals のロジック部分
    '!
    '| @see    doFileDiff
    '! @param  theExpected    string 検査対象となるファイルのパス
    '! @param  theActual      string 比較対象となるファイルのパス
    '! @param  theFlg         boolean 成功の時の返り値
    '! @return                text 状態(success, failure, error)
    Private Function assertFileEqualsLogic(theExpected, theActual, theFlg)
        If (VarType(theExpected) <> vbString) Then
            assertFileEqualsLogic = addError("1st Args is not String"_
                & vbLf & gen2ArgToMsg(theExpected, theActual))
            Exit Function
        End If

        If (VarType(theActual) <> vbString) Then
            assertFileEqualsLogic = addError("1st Args is not String"_
                & vbLf & gen2ArgToMsg(theExpected, theActual))
            Exit Function
        End If

        If (file_exists(theExpected) <> True) Then
            assertFileEqualsLogic = addError("(1st Args) File Not Found."_
                & vbLf & gen2ArgToMsg(theExpected, theActual))
            Exit Function
        End If

        If (file_exists(theActual) <> True) Then
            assertFileEqualsLogic = addError("(2st Args) File Not Found."_
                & vbLf & gen2ArgToMsg(theExpected, theActual))
            Exit Function
        End If

        If (doFileDiff(theExpected, theActual) = theFlg) Then
            assertFileEqualsLogic = addSuccess(gen2ArgToMsg(theExpected, theActual))
            Exit Function
        End If
        assertFileEqualsLogic = addFailure(gen2ArgToMsg(theExpected, theActual))
    End Function

    '! assertGreaterThan, assertGreaterThanOrEqual, assertLessThan, assertLessThanOrEqual のロジック部分
    '!
    '! @see    is_numeric
    '! @param  theExpected    比較する数値の左辺
    '! @param  theActual      比較する数値の右辺
    '! @param  theOperator    string 比較演算子
    '! @return                text 状態(success, failure, error)
    Private Function assertRelationalOperatorLogic(theExpected, theActual, theOperator)
        If (is_numeric(theExpected) = False) Then
            assertRelationalOperatorLogic = addError("1st Arg Not Numeric."_
                & vbLf & gen2ArgToMsg(theExpected, theActual))
            Exit Function
        End If
        If (is_numeric(theActual) = False) Then
            assertRelationalOperatorLogic = addError("2st Arg Not Numeric."_
                & vbLf & gen2ArgToMsg(theExpected, theActual))
            Exit Function
        End If

        On Error Resume Next
        Dim rtn: rtn = eval(theExpected & theOperator & theActual)
        Dim emsg: emsg = genErrorMsg()
        On Error Goto 0

        If (emsg <> "") Then
            assertRelationalOperatorLogic = addError(_
                emsg & vbLf & gen2ArgToMsg(theExpected, theActual))
            Exit Function
        End If

        If (rtn = True) Then
            assertRelationalOperatorLogic = addSuccess(gen2ArgToMsg(theExpected, theActual))
            Exit Function
        End If
        assertRelationalOperatorLogic = addFailure(gen2ArgToMsg(theExpected, theActual))
    End Function




    '--- Private Method (Util) ---------------------------------------------

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

    '! 一つの引数をログ用メッセージに加工する
    '!
    '! @param   theArg      mixed:引数
    '! @return              String:メッセージ
    Private Function gen1ArgToMsg(theArg)
        gen1ArgToMsg = "Arg => " & toStringValue(theArg)
    End Function

    '! 二つの引数をログ用メッセージに加工する
    '!
    '! @param   theArg1     mixed:引数
    '! @param   theArg2     mixed:引数
    '! @return              String:メッセージ
    Private Function gen2ArgToMsg(theArg1, theArg2)
        gen2ArgToMsg = "Arg1 => " & toStringValue(theArg1) _
                     & ", Arg2 => " & toStringValue(theArg2)
    End Function

    '! エラーが発生している時にエラーメッセージを返す
    '!
    '! @param  void
    '! @return             エラーメッセージ
    Private Function genErrorMsg()
        If (Err.Number = 0) Then
            genErrorMsg = ""
            Exit Function
        End If
        genErrorMsg = Err.Description & "(" & Err.Number & ")"
    End Function

    '! 成功のカウンターを進める
    '!
    '! @param  theMsg     string 成功メッセージ
    '! @return void
    Private Function addSuccess(theMsg)
        With dic.item(name)
            .item("success").Add .item("cnt"), theMsg
            .item("cnt") = .item("cnt") + 1
        End With
        addSuccess = "success"
    End Function

    '! 失敗のカウンターを進める
    '!
    '! @param  theMsg     string 失敗メッセージ
    '! @return void
    Private Function addFailure(theMsg)
        With dic.item(name)
            .item("failure").Add .item("cnt"), theMsg
            .item("cnt") = .item("cnt") + 1
        End With
        addFailure = "failure"
    End Function

   '! エラーのカウンターを進める
    '!
    '! @param  theMsg     string エラーメッセージ
    '! @return void
    Private Function addError(theMsg)
        With dic.item(name)
            .item("error").Add .item("cnt"), theMsg
            .item("cnt") = .item("cnt") + 1
        End With
        addError = "error"
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

    '! 2つの変数(数値、文字、配列)が同じか比較する
    '!
    '! @see    isValueEqual, isArrayEqual
    '! @param  theExpected   mixed 検査する左辺
    '! @param  theActual     mixed 検査する右辺
    '! @return               string success:等しい、failure:等しくない, error: エラー
    Private Function isEqual(theExpected, theActual)
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


End Class
