Execute CreateObject("Scripting.FileSystemObject").OpenTextFile("../VBSUnit.vbs").ReadAll()

'! VBSUnit Test
'!
'! @author  miyamiya <rai.caver@gmail.com>
'! @date    2013/08/14
'! @version 0.1
Class TestVBSUnit

    Public  unit
    private sp

    Private Function judge(theA, theB)
        If (theA = theB) Then
            WScript.StdOut.Write "o"
            Exit Function
        End If
        WScript.Echo "x"
        WScript.Quit
    End Function

    Private Function genTitle(theMsg)
        WScript.Echo sp & theMsg
        sp = vbCrLf & vbCrLf
    End Function

    Private Function genSubTitle(theMsg, theFlg)
        If (theFlg = 1) Then
            WScript.Echo 
        End If
        WScript.StdOut.Write "  " & theMsg & "-> "
    End Function

    Public Function testAssertEquals
        Call genTitle("testAssertEquals")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertEquals("a", "a"), "success")
        Call judge(unit.assertEquals(0, 0), "success")
        Call judge(unit.assertEquals(1, 1), "success")
        Call judge(unit.assertEquals(Array("1","2"), Array("1","2")), "success")
        Call judge(unit.assertEquals(Null, Null), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertEquals("a", "b"), "failure")
        Call judge(unit.assertEquals(1, 2), "failure")
        Call judge(unit.assertEquals(1, "1"), "failure")
        Call judge(unit.assertEquals("", Null), "failure")
        Call judge(unit.assertEquals(Array("1",2), Array("1","2")), "failure")

        'Error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertEquals(CreateObject("Scripting.Dictionary"), "a"), "error")
        Call judge(unit.assertEquals("a", CreateObject("Scripting.Dictionary")), "error")
        Call judge(unit.assertEquals(CreateObject("Scripting.Dictionary"), CreateObject("Scripting.Dictionary")), "error")
    End Function


    Public Function testAssertNotEquals
        Call genTitle("testAssertNotEquals")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertNotEquals("a", "b"), "success")
        Call judge(unit.assertNotEquals(1, 2), "success")
        Call judge(unit.assertNotEquals(1, "1"), "success")
        Call judge(unit.assertNotEquals(Array("1",2), Array("1","2")), "success")
        Call judge(unit.assertNotEquals("", Null), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertNotEquals("a", "a"), "failure")
        Call judge(unit.assertNotEquals(0, 0), "failure")
        Call judge(unit.assertNotEquals(1, 1), "failure")
        Call judge(unit.assertNotEquals(Array("1","2"), Array("1","2")), "failure")
        Call judge(unit.assertNotEquals(Null, Null), "failure")

        'Error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertNotEquals(CreateObject("Scripting.Dictionary"), "a"), "error")
        Call judge(unit.assertNotEquals("a", CreateObject("Scripting.Dictionary")), "error")
        Call judge(unit.assertNotEquals(CreateObject("Scripting.Dictionary"), CreateObject("Scripting.Dictionary")), "error")
    End Function


    Public Function testAssertTrue
        Call genTitle("testAssertTrue")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertTrue(True), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertTrue(False), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertTrue("True"), "error")
        Call judge(unit.assertTrue("False"), "error")
        Call judge(unit.assertTrue(1), "error")
        Call judge(unit.assertTrue(Array(1)), "error")
        Call judge(unit.assertTrue(CreateObject("Scripting.Dictionary")), "error")
        Call judge(unit.assertTrue(Null), "error")
    End Function

    Public Function testAssertFalse
        Call genTitle("testAssertFalse")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertFalse(False), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertFalse(True), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertFalse("True"), "error")
        Call judge(unit.assertFalse("False"), "error")
        Call judge(unit.assertFalse(1), "error")
        Call judge(unit.assertFalse(Array(1)), "error")
        Call judge(unit.assertFalse(CreateObject("Scripting.Dictionary")), "error")
        Call judge(unit.assertTrue(Null), "error")
    End Function


    Public Function testAssertNull
        Call genTitle("testAssertNull")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertNull(Null), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertNull("Null"), "failure")
        Call judge(unit.assertNull(""), "failure")
        Call judge(unit.assertNull(True), "failure")
        Call judge(unit.assertNull(False), "failure")
        Call judge(unit.assertNull(0), "failure")
        Call judge(unit.assertNull(Array(0)), "failure")
        Call judge(unit.assertNull(CreateObject("Scripting.Dictionary")), "failure")
    End Function

    Public Function testAssertNotNull
        Call genTitle("testAssertNotNull")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertNotNull("Null"), "success")
        Call judge(unit.assertNotNull(""), "success")
        Call judge(unit.assertNotNull(True), "success")
        Call judge(unit.assertNotNull(False), "success")
        Call judge(unit.assertNotNull(0), "success")
        Call judge(unit.assertNotNull(Array(0)), "success")
        Call judge(unit.assertNotNull(CreateObject("Scripting.Dictionary")), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertNotNull(Null), "failure")
    End Function


    Public Function testAssertEmpty
        Call genTitle("testAssertEmpty")
        'success
        Call genSubTitle("success", 0)
        Dim hoge
        Call judge(unit.assertEmpty(hoge), "success")
        Call judge(unit.assertEmpty(""), "success")
        Call judge(unit.assertEmpty(0), "success")
        Dim fuga()
        Call judge(unit.assertEmpty(fuga), "success")
        Call judge(unit.assertEmpty(Array()), "success")
        Call judge(unit.assertEmpty(False), "success")
        Call judge(unit.assertEmpty(0.0), "success")
        Call judge(unit.assertEmpty(Null), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertEmpty("Null"), "failure")
        Call judge(unit.assertEmpty(Array(0)), "failure")
        Call judge(unit.assertEmpty("0"), "failure")
        Call judge(unit.assertEmpty(1), "failure")
        Call judge(unit.assertEmpty(CreateObject("Scripting.Dictionary")), "failure")
    End Function


    Public Function testAssertNotEmpty
        Call genTitle("testAssertNotEmpty")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertNotEmpty("Null"), "success")
        Call judge(unit.assertNotEmpty(Array(0)), "success")
        Call judge(unit.assertNotEmpty("0"), "success")
        Call judge(unit.assertNotEmpty(1), "success")
        Call judge(unit.assertNotEmpty(CreateObject("Scripting.Dictionary")), "success")

        'failure
        Call genSubTitle("failure", 1)
        Dim hoge
        Call judge(unit.assertNotEmpty(hoge), "failure")
        Call judge(unit.assertNotEmpty(""), "failure")
        Call judge(unit.assertNotEmpty(0), "failure")
        Dim fuga()
        Call judge(unit.assertNotEmpty(fuga), "failure")
        Call judge(unit.assertNotEmpty(Array()), "failure")
        Call judge(unit.assertNotEmpty(False), "failure")
        Call judge(unit.assertNotEmpty(0.0), "failure")
        Call judge(unit.assertNotEmpty(Null), "failure")
    End Function

    Public Function testAssertContains
        Call genTitle("testAssertContains")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertContains("d", Array("a", "b", "c", "d")), "success")
        Call judge(unit.assertContains(Array("35", "c"), Array("a", "b", "c", "d")), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertContains("e", Array("a", "b", "c", "d")), "failure")
        Call judge(unit.assertContains(Array("35", "e"), Array("a", "b", "c", "d")), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertNotContains(CreateObject("Scripting.Dictionary"), Array("a")), "error")
        Call judge(unit.assertNotContains("a", "a"), "error")
        Call judge(unit.assertContains("a", "a"), "error")
        Call judge(unit.assertContains("a", 1), "error")
        Call judge(unit.assertContains("a", CreateObject("Scripting.Dictionary")), "error")
    End Function


    Public Function testAssertNotContains
        Call genTitle("testAssertNotContains")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertNotContains("e", Array("a", "b", "c", "d")), "success")
        Call judge(unit.assertNotContains(Array("35", "e"), Array("a", "b", "c", "d")), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertNotContains("d", Array("a", "b", "c", "d")), "failure")
        Call judge(unit.assertNotContains(Array("35", "c"), Array("a", "b", "c", "d")), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertNotContains(CreateObject("Scripting.Dictionary"), Array("a")), "error")
        Call judge(unit.assertNotContains("a", "a"), "error")
        Call judge(unit.assertNotContains("a", 1), "error")
        Call judge(unit.assertNotContains("a", CreateObject("Scripting.Dictionary")), "error")
    End Function


    Public Function testAssertStringStartsWith
        Call genTitle("testAssertStringStartsWith")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertStringStartsWith("abc", "abcdefg"), "success")
        Call judge(unit.assertStringStartsWith("AbC", "AbCdefg"), "success")
        Call judge(unit.assertStringStartsWith("ñæì˙", "ñæì˙ÇÕê∞ÇÍ"), "success")
        Call judge(unit.assertStringStartsWith("123", "1234444"), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertStringStartsWith("bc", "abcdefg"), "failure")
        Call judge(unit.assertStringStartsWith("ñæå„ì˙", "ñæì˙ÇÕê∞ÇÍ"), "failure")
        Call judge(unit.assertStringStartsWith("134", "1234444"), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertStringStartsWith(123, "abcdefg"), "error")
        Call judge(unit.assertStringStartsWith("abcdefg", 123), "error")
        Call judge(unit.assertStringStartsWith("abcdefg", Array()), "error")
        Call judge(unit.assertStringStartsWith("abcdefg", False), "error")
        Call judge(unit.assertStringStartsWith("abcdefg", Null), "error")
        Call judge(unit.assertStringStartsWith("abcdefg", CreateObject("Scripting.Dictionary")), "error")
    End Function

    Public Function testAssertStringStartsNotWith
        Call genTitle("testAssertStringStartsNotWith")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertStringStartsNotWith("bc", "abcdefg"), "success")
        Call judge(unit.assertStringStartsNotWith("ñæå„ì˙", "ñæì˙ÇÕê∞ÇÍ"), "success")
        Call judge(unit.assertStringStartsNotWith("134", "1234444"), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertStringStartsNotWith("abc", "abcdefg"), "failure")
        Call judge(unit.assertStringStartsNotWith("AbC", "AbCdefg"), "failure")
        Call judge(unit.assertStringStartsNotWith("ñæì˙", "ñæì˙ÇÕê∞ÇÍ"), "failure")
        Call judge(unit.assertStringStartsNotWith("123", "1234444"), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertStringStartsNotWith(123, "abcdefg"), "error")
        Call judge(unit.assertStringStartsNotWith("abcdefg", 123), "error")
        Call judge(unit.assertStringStartsNotWith("abcdefg", Array()), "error")
        Call judge(unit.assertStringStartsNotWith("abcdefg", False), "error")
        Call judge(unit.assertStringStartsNotWith("abcdefg", Null), "error")
        Call judge(unit.assertStringStartsNotWith("abcdefg", CreateObject("Scripting.Dictionary")), "error")
    End Function

    Public Function testAssertStringEndsWith
        Call genTitle("testAssertStringEndsWith")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertStringEndsWith("efg", "abcdefg"), "success")
        Call judge(unit.assertStringEndsWith("efG", "AbcdefG"), "success")
        Call judge(unit.assertStringEndsWith("ê∞ÇÍ", "ñæì˙ÇÕê∞ÇÍ"), "success")
        Call judge(unit.assertStringEndsWith("44", "1234444"), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertStringEndsWith("eg", "abcdefg"), "failure")
        Call judge(unit.assertStringEndsWith("ê∞ÇÍÇΩ", "ñæì˙ÇÕê∞ÇÍ"), "failure")
        Call judge(unit.assertStringEndsWith("244", "1234444"), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertStringEndsWith(123, "abcdefg"), "error")
        Call judge(unit.assertStringEndsWith("abcdefg", 123), "error")
        Call judge(unit.assertStringEndsWith("abcdefg", Array()), "error")
        Call judge(unit.assertStringEndsWith("abcdefg", False), "error")
        Call judge(unit.assertStringEndsWith("abcdefg", Null), "error")
        Call judge(unit.assertStringEndsWith("abcdefg", CreateObject("Scripting.Dictionary")), "error")
    End Function


    Public Function testAssertStringEndsNotWith
        Call genTitle("testAssertStringEndsNotWith")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertStringEndsNotWith("eg", "abcdefg"), "success")
        Call judge(unit.assertStringEndsNotWith("ê∞ÇÍÇΩ", "ñæì˙ÇÕê∞ÇÍ"), "success")
        Call judge(unit.assertStringEndsNotWith("244", "1234444"), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertStringEndsNotWith("efg", "abcdefg"), "failure")
        Call judge(unit.assertStringEndsNotWith("efG", "AbcdefG"), "failure")
        Call judge(unit.assertStringEndsNotWith("ê∞ÇÍ", "ñæì˙ÇÕê∞ÇÍ"), "failure")
        Call judge(unit.assertStringEndsNotWith("44", "1234444"), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertStringEndsNotWith(123, "abcdefg"), "error")
        Call judge(unit.assertStringEndsNotWith("abcdefg", 123), "error")
        Call judge(unit.assertStringEndsNotWith("abcdefg", Array()), "error")
        Call judge(unit.assertStringEndsNotWith("abcdefg", False), "error")
        Call judge(unit.assertStringEndsNotWith("abcdefg", Null), "error")
        Call judge(unit.assertStringEndsNotWith("abcdefg", CreateObject("Scripting.Dictionary")), "error")
    End Function



    Public Function testAssertInternalType
        Call genTitle("testAssertInternalType")
        Dim hoge
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertInternalType("empty", hoge), "success")
        Call judge(unit.assertInternalType("null", Null), "success")
        Call judge(unit.assertInternalType("integer", 333), "success")
        Call judge(unit.assertInternalType("long", CLng(333)), "success")
        Call judge(unit.assertInternalType("single", CSng(1.1)), "success")
        Call judge(unit.assertInternalType("double", 1.1), "success")
        Call judge(unit.assertInternalType("currency", CCur(1.1)), "success")
        Call judge(unit.assertInternalType("date", Date), "success")
        Call judge(unit.assertInternalType("date", Now), "success")
        Call judge(unit.assertInternalType("date", #2013/08/13#), "success")
        Call judge(unit.assertInternalType("string", "hoge"), "success")
        Call judge(unit.assertInternalType("string", "2013"), "success")
        Call judge(unit.assertInternalType("object", CreateObject("Scripting.Dictionary")), "success")
        'type:error -> No Test
        Call judge(unit.assertInternalType("boolean", True), "success")
        Call judge(unit.assertInternalType("boolean", False), "success")
        'type:variant -> No Test
        'type:DataObject -> No Test
        'type:Byte -> No Test
        Call judge(unit.assertInternalType("array", Array()), "success")

        'failure
        Call genSubTitle("failure", 1)
        hoge = 2
        Call judge(unit.assertInternalType("empty", hoge), "failure")
        Call judge(unit.assertInternalType("null", True), "failure")
        Call judge(unit.assertInternalType("integer", "333"), "failure")
        Call judge(unit.assertInternalType("long", 333), "failure")
        Call judge(unit.assertInternalType("single", 333), "failure")
        Call judge(unit.assertInternalType("double", CSng(1.1)), "failure")
        Call judge(unit.assertInternalType("currency", 1.1), "failure")
        Call judge(unit.assertInternalType("date", "2013/08/13"), "failure")
        Call judge(unit.assertInternalType("date", "Now"), "failure")
        Call judge(unit.assertInternalType("date", "#2013/08/13#"), "failure")
        Call judge(unit.assertInternalType("string", 234), "failure")
        Call judge(unit.assertInternalType("object", Array()), "failure")
        'type:error -> No Test
        Call judge(unit.assertInternalType("boolean", "True"), "failure")
        Call judge(unit.assertInternalType("boolean", "False"), "failure")
        'type:variant -> No Test
        'type:DataObject -> No Test
        'type:Byte -> No Test
        Call judge(unit.assertInternalType("array", "Array()"), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertInternalType(Array(), "Array"), "error")
        Call judge(unit.assertInternalType(True, "Array"), "error")
        Call judge(unit.assertInternalType(124, "Array"), "error")
        Call judge(unit.assertInternalType(Null, "Array"), "error")
    End Function

    Public Function testAssertNotInternalType
        Call genTitle("testAssertNotInternalType")
        'success
        Call genSubTitle("success", 0)
        Dim fuga : fuga = 2
        Call judge(unit.assertNotInternalType("empty", fuga), "success")
        Call judge(unit.assertNotInternalType("null", True), "success")
        Call judge(unit.assertNotInternalType("integer", "333"), "success")
        Call judge(unit.assertNotInternalType("long", 333), "success")
        Call judge(unit.assertNotInternalType("single", 333), "success")
        Call judge(unit.assertNotInternalType("double", CSng(1.1)), "success")
        Call judge(unit.assertNotInternalType("currency", 1.1), "success")
        Call judge(unit.assertNotInternalType("date", "2013/08/13"), "success")
        Call judge(unit.assertNotInternalType("date", "Now"), "success")
        Call judge(unit.assertNotInternalType("date", "#2013/08/13#"), "success")
        Call judge(unit.assertNotInternalType("string", 234), "success")
        Call judge(unit.assertNotInternalType("object", Array()), "success")
        'type:error -> No Test
        Call judge(unit.assertNotInternalType("boolean", "True"), "success")
        Call judge(unit.assertNotInternalType("boolean", "False"), "success")
        'type:variant -> No Test
        'type:DataObject -> No Test
        'type:Byte -> No Test
        Call judge(unit.assertNotInternalType("array", "Array()"), "success")

        'failure
        Call genSubTitle("failure", 1)
        Dim hoge
        Call judge(unit.assertNotInternalType("empty", hoge), "failure")
        Call judge(unit.assertNotInternalType("null", Null), "failure")
        Call judge(unit.assertNotInternalType("integer", 333), "failure")
        Call judge(unit.assertNotInternalType("long", CLng(333)), "failure")
        Call judge(unit.assertNotInternalType("single", CSng(1.1)), "failure")
        Call judge(unit.assertNotInternalType("double", 1.1), "failure")
        Call judge(unit.assertNotInternalType("currency", CCur(1.1)), "failure")
        Call judge(unit.assertNotInternalType("date", Date), "failure")
        Call judge(unit.assertNotInternalType("date", Now), "failure")
        Call judge(unit.assertNotInternalType("date", #2013/08/13#), "failure")
        Call judge(unit.assertNotInternalType("string", "hoge"), "failure")
        Call judge(unit.assertNotInternalType("string", "2013"), "failure")
        Call judge(unit.assertNotInternalType("object", CreateObject("Scripting.Dictionary")), "failure")
        'type:error -> No Test
        Call judge(unit.assertNotInternalType("boolean", True), "failure")
        Call judge(unit.assertNotInternalType("boolean", False), "failure")
        'type:variant -> No Test
        'type:DataObject -> No Test
        'type:Byte -> No Test
        Call judge(unit.assertNotInternalType("array", Array()), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertNotInternalType(Array(), "Array"), "error")
        Call judge(unit.assertNotInternalType(True, "Array"), "error")
        Call judge(unit.assertNotInternalType(124, "Array"), "error")
        Call judge(unit.assertNotInternalType(Null, "Array"), "error")
    End Function



    Public Function testAssertStringEqualsFile
        Call genTitle("testAssertStringEqualsFile")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertStringEqualsFile("search.txt", "åüçı"), "success")
        Call judge(unit.assertStringEqualsFile("search.txt", "---"), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertStringEqualsFile("search.txt", "åüçıÇ†"), "failure")
        Call judge(unit.assertStringEqualsFile("search.txt", "aaaa"), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertStringEqualsFile("search.txt", Array()), "error")
        Call judge(unit.assertStringEqualsFile("search.txt", True), "error")
        Call judge(unit.assertStringEqualsFile("search.txt", Null), "error")
        Call judge(unit.assertStringEqualsFile("search.txt", CreateObject("Scripting.Dictionary")), "error")
        Call judge(unit.assertStringEqualsFile("filenotfound.txt", "åüçıÇ†"), "error")
    End Function

    Public Function testAssertStringNotEqualsFile
        Call genTitle("testAssertStringNotEqualsFile")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertStringNotEqualsFile("search.txt", "åüçıÇ†"), "success")
        Call judge(unit.assertStringNotEqualsFile("search.txt", "aaaa"), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertStringNotEqualsFile("search.txt", "åüçı"), "failure")
        Call judge(unit.assertStringNotEqualsFile("search.txt", "---"), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertStringNotEqualsFile("search.txt", Array()), "error")
        Call judge(unit.assertStringNotEqualsFile("search.txt", True), "error")
        Call judge(unit.assertStringNotEqualsFile("search.txt", Null), "error")
        Call judge(unit.assertStringNotEqualsFile("search.txt", CreateObject("Scripting.Dictionary")), "error")
        Call judge(unit.assertStringNotEqualsFile("filenotfound.txt", "åüçıÇ†"), "error")
    End Function


    Public Function testAssertFileEquals
        Call genTitle("testAssertFileEquals")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertFileEquals("equalA.txt", "equalB.txt"), "success")
        Call judge(unit.assertFileEquals("equalA.png", "equalB.png"), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertFileEquals("equalA.txt", "notEqualB.txt"), "failure")
        Call judge(unit.assertFileEquals("equalA.png", "notEqualB.png"), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertFileEquals("search.txt", Array()), "error")
        Call judge(unit.assertFileEquals("search.txt", True), "error")
        Call judge(unit.assertFileEquals("search.txt", Null), "error")
        Call judge(unit.assertFileEquals("search.txt", CreateObject("Scripting.Dictionary")), "error")
        Call judge(unit.assertFileEquals("filenotfound.txt", "åüçıÇ†"), "error")
    End Function


    Public Function testAssertFileNotEquals
        Call genTitle("testAssertFileNotEquals")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertFileNotEquals("equalA.txt", "notEqualB.txt"), "success")
        Call judge(unit.assertFileNotEquals("equalA.png", "notEqualB.png"), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertFileNotEquals("equalA.txt", "equalB.txt"), "failure")
        Call judge(unit.assertFileNotEquals("equalA.png", "equalB.png"), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertFileNotEquals("search.txt", Array()), "error")
        Call judge(unit.assertFileNotEquals("search.txt", True), "error")
        Call judge(unit.assertFileNotEquals("search.txt", Null), "error")
        Call judge(unit.assertFileNotEquals("search.txt", CreateObject("Scripting.Dictionary")), "error")
        Call judge(unit.assertFileNotEquals("filenotfound.txt", "åüçıÇ†"), "error")
    End Function

    Public Function testAssertFileExists
        Call genTitle("testAssertFileExists")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertFileExists("equalA.txt"), "success")
        Call judge(unit.assertFileExists("equalA.png"), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertFileExists("equalAA.txt"), "failure")
        Call judge(unit.assertFileExists("equalAA.png"), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertFileExists(Array()), "error")
        Call judge(unit.assertFileExists(True), "error")
        Call judge(unit.assertFileExists(Null), "error")
        Call judge(unit.assertFileExists(CreateObject("Scripting.Dictionary")), "error")
    End Function

    Public Function testAssertFileNotExists
        Call genTitle("testAssertFileNotExists")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertFileNotExists("equalAA.txt"), "success")
        Call judge(unit.assertFileNotExists("equalAA.png"), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertFileNotExists("equalA.txt"), "failure")
        Call judge(unit.assertFileNotExists("equalA.png"), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertFileNotExists(Array()), "error")
        Call judge(unit.assertFileNotExists(True), "error")
        Call judge(unit.assertFileNotExists(Null), "error")
        Call judge(unit.assertFileNotExists(CreateObject("Scripting.Dictionary")), "error")
    End Function

    Public Function testAssertGreaterThan
        Call genTitle("testAssertGreaterThan")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertGreaterThan(20, 30), "success")
        Call judge(unit.assertGreaterThan(-20, 0), "success")
        Call judge(unit.assertGreaterThan(-20.1, -20), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertGreaterThan(30, 20), "failure")
        Call judge(unit.assertGreaterThan(20, 20), "failure")
        Call judge(unit.assertGreaterThan(0, -20), "failure")
        Call judge(unit.assertGreaterThan(-20, -20), "failure")
        Call judge(unit.assertGreaterThan(-20, -20.1), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertGreaterThan("30", 20), "error")
        Call judge(unit.assertGreaterThan(30, "20"), "error")
        Call judge(unit.assertGreaterThan(Array(), 20), "error")
        Call judge(unit.assertGreaterThan(30, Array()), "error")
        Call judge(unit.assertGreaterThan(Null, 20), "error")
        Call judge(unit.assertGreaterThan(30, Null), "error")
        Call judge(unit.assertGreaterThan(True, 20), "error")
        Call judge(unit.assertGreaterThan(30, True), "error")
        Call judge(unit.assertGreaterThan(CreateObject("Scripting.Dictionary"), 20), "error")
        Call judge(unit.assertGreaterThan(30, CreateObject("Scripting.Dictionary")), "error")
    End Function

    Public Function testAssertGreaterThanOrEqual
        Call genTitle("testAssertGreaterThanOrEqual")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertGreaterThanOrEqual(20, 30), "success")
        Call judge(unit.assertGreaterThanOrEqual(20, 20), "success")
        Call judge(unit.assertGreaterThanOrEqual(-20, -20), "success")
        Call judge(unit.assertGreaterThanOrEqual(-20, 0), "success")
        Call judge(unit.assertGreaterThanOrEqual(-20.1, -20), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertGreaterThanOrEqual(30, 20), "failure")
        Call judge(unit.assertGreaterThanOrEqual(0, -20), "failure")
        Call judge(unit.assertGreaterThanOrEqual(-20, -20.1), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertGreaterThanOrEqual("30", 20), "error")
        Call judge(unit.assertGreaterThanOrEqual(30, "20"), "error")
        Call judge(unit.assertGreaterThanOrEqual(Array(), 20), "error")
        Call judge(unit.assertGreaterThanOrEqual(30, Array()), "error")
        Call judge(unit.assertGreaterThanOrEqual(Null, 20), "error")
        Call judge(unit.assertGreaterThanOrEqual(30, Null), "error")
        Call judge(unit.assertGreaterThanOrEqual(True, 20), "error")
        Call judge(unit.assertGreaterThanOrEqual(30, True), "error")
        Call judge(unit.assertGreaterThanOrEqual(CreateObject("Scripting.Dictionary"), 20), "error")
        Call judge(unit.assertGreaterThanOrEqual(30, CreateObject("Scripting.Dictionary")), "error")
    End Function


    Public Function testAssertLessThan
        Call genTitle("testAssertLessThan")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertLessThan(30, 20), "success")
        Call judge(unit.assertLessThan(0, -20), "success")
        Call judge(unit.assertLessThan(-20, -20.1), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertLessThan(20, 30), "failure")
        Call judge(unit.assertLessThan(20, 20), "failure")
        Call judge(unit.assertLessThan(-20, 0), "failure")
        Call judge(unit.assertLessThan(-20, -20), "failure")
        Call judge(unit.assertLessThan(-20.1, -20), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertLessThan("30", 20), "error")
        Call judge(unit.assertLessThan(30, "20"), "error")
        Call judge(unit.assertLessThan(Array(), 20), "error")
        Call judge(unit.assertLessThan(30, Array()), "error")
        Call judge(unit.assertLessThan(Null, 20), "error")
        Call judge(unit.assertLessThan(30, Null), "error")
        Call judge(unit.assertLessThan(True, 20), "error")
        Call judge(unit.assertLessThan(30, True), "error")
        Call judge(unit.assertLessThan(CreateObject("Scripting.Dictionary"), 20), "error")
        Call judge(unit.assertLessThan(30, CreateObject("Scripting.Dictionary")), "error")
    End Function

    Public Function testAssertLessThanOrEqual
        Call genTitle("testAssertLessThanOrEqual")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertLessThanOrEqual(30, 20), "success")
        Call judge(unit.assertLessThanOrEqual(20, 20), "success")
        Call judge(unit.assertLessThanOrEqual(0, -20), "success")
        Call judge(unit.assertLessThanOrEqual(-20, -20), "success")
        Call judge(unit.assertLessThanOrEqual(-20, -20.1), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertLessThanOrEqual(20, 30), "failure")
        Call judge(unit.assertLessThanOrEqual(-20, 0), "failure")
        Call judge(unit.assertLessThanOrEqual(-20.1, -20), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertLessThanOrEqual("30", 20), "error")
        Call judge(unit.assertLessThanOrEqual(30, "20"), "error")
        Call judge(unit.assertLessThanOrEqual(Array(), 20), "error")
        Call judge(unit.assertLessThanOrEqual(30, Array()), "error")
        Call judge(unit.assertLessThanOrEqual(Null, 20), "error")
        Call judge(unit.assertLessThanOrEqual(30, Null), "error")
        Call judge(unit.assertLessThanOrEqual(True, 20), "error")
        Call judge(unit.assertLessThanOrEqual(30, True), "error")
        Call judge(unit.assertLessThanOrEqual(CreateObject("Scripting.Dictionary"), 20), "error")
        Call judge(unit.assertLessThanOrEqual(30, CreateObject("Scripting.Dictionary")), "error")
    End Function

    Public Function testAssertCount
        Call genTitle("testAssertCount")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertCount(2, Array(1, 2)), "success")
        Call judge(unit.assertCount(5, Array(1, 2, 3, 4, 5)), "success")

        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertCount(3, Array(1, 2)), "failure")
        Call judge(unit.assertCount(-5, Array(1, 2, 3, 4, 5)), "failure")

        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertCount("2", Array(1, 2)), "error")
        Call judge(unit.assertCount(Array(1, 2), "2"), "error")
        Call judge(unit.assertCount(30, Array()), "error")
        Call judge(unit.assertCount(Null, Array(1, 2)), "error")
        Call judge(unit.assertCount(True, Array(1, 2)), "error")
        Call judge(unit.assertCount(CreateObject("Scripting.Dictionary"), Array(1, 2)), "error")
    End Function

    Public Function testAssertNotCount
        Call genTitle("testAssertNotCount")
        'success
        Call genSubTitle("success", 0)
        Call judge(unit.assertNotCount(3, Array(1, 2)), "success")
        Call judge(unit.assertNotCount(-5, Array(1, 2, 3, 4, 5)), "success")
        'failure
        Call genSubTitle("failure", 1)
        Call judge(unit.assertNotCount(2, Array(1, 2)), "failure")
        Call judge(unit.assertNotCount(5, Array(1, 2, 3, 4, 5)), "failure")


        'error
        Call genSubTitle("  error", 1)
        Call judge(unit.assertNotCount("2", Array(1, 2)), "error")
        Call judge(unit.assertNotCount(Array(1, 2), "2"), "error")
        Call judge(unit.assertNotCount(30, Array()), "error")
        Call judge(unit.assertNotCount(Null, Array(1, 2)), "error")
        Call judge(unit.assertNotCount(True, Array(1, 2)), "error")
        Call judge(unit.assertNotCount(CreateObject("Scripting.Dictionary"), Array(1, 2)), "error")
    End Function


    Private Sub Class_Initialize()
        Set unit = new VBSUnit
        sp   = ""
    End Sub
    Private Sub Class_Terminate()
        Set unit = Nothing
    End Sub

End Class


Dim obj : Set obj = new TestVBSUnit


Dim testCase : testCase = Array(_
      "testAssertEquals" _
    , "testAssertNotEquals" _
    , "testAssertTrue" _
    , "testAssertFalse" _
    , "testAssertNull" _
    , "testAssertNotNull" _
    , "testAssertEmpty" _
    , "testAssertNotEmpty" _
    , "testAssertContains" _
    , "testAssertNotContains" _
    , "testAssertStringStartsWith" _
    , "testAssertStringStartsNotWith" _
    , "testAssertStringEndsWith" _
    , "testAssertStringEndsNotWith" _
    , "testAssertInternalType" _
    , "testAssertNotInternalType" _
    , "testAssertStringEqualsFile" _
    , "testAssertStringNotEqualsFile" _
    , "testAssertFileEquals" _
    , "testAssertFileNotEquals" _
    , "testAssertFileExists" _
    , "testAssertFileNotExists" _
    , "testAssertGreaterThan" _
    , "testAssertGreaterThanOrEqual" _
    , "testAssertLessThan" _
    , "testAssertLessThanOrEqual" _
    , "testAssertCount" _
    , "testAssertNotCount" _
)
For Each testName In testCase
    obj.unit.setName =  testName
    Execute("obj." & testName)
Next
