VBSUnit
=====

Description
-----------
Unit Test for VBScript.


Usage
-----

1. create class file

    Hoge.class
    ```
    Class Hoge
      Public Function Plus(a, b)
        Plus = a + b
      End Function

      Public Function Minus(a, b)
        Minus = a - b
      End Function

      Public Function chainCharacter(a, b)
        chainCharacter = a & b
      End Function
    End Class
    ```

2. create test class file

    test.Hoge.class
    ```
    Class testHoge
      Public unit
      Public obj

      Public Function testPlus
        unit.setName = "testPlus"
        Call unit.assertEquals(obj.Plus(2,3), 5)
        Call unit.assertEquals(obj.Plus(-2,5), 3)
        Call unit.assertEquals(obj.Plus(0,0), 0)
      End Function

      Public Function testMinus
        unit.setName = "testMinus"
        Call unit.assertEquals(obj.Minus(3,3), 0)
        Call unit.assertEquals(obj.Minus(100,5), 95)
      End Function

      Public Function testChainCharacter
        unit.setName = "testChainCharacter"
        Call unit.assertEquals(obj.chainCharacter("hoge", "fuga"), "hogefuga")
        Call unit.assertNotEquals(obj.chainCharacter(1, 1), 2)
      End Function

      Private Sub Class_Initialize
        Set unit = new VBSUnit
        Set obj  = new Hoge
      End Sub
      Private Sub Class_Terminate
        Set obj  = Nothing
        Set unit = Nothing
      End Sub
    End Class
    ```

3. TestRunner


Assertions
-------------------------------

- assertEquals
- assertNotEquals
- assertTrue
- assertFalse
- assertNull
- assertNotNull
- assertEmpty
- assertNotEmpty
- assertContains
- assertNotContains
- assertStringStartsWith
- assertStringStartsNotWith
- assertStringEndsWith
- assertStringEndsNotWith
- assertInternalType
- assertNotInternalType
- assertStringEqualsFile
- assertStringNotEqualsFile
- assertFileEquals
- assertFileNotEquals
- assertFileExists
- assertFileNotExists
- assertGreaterThan
- assertGreaterThanOrEqual
- assertLessThan
- assertLessThanOrEqual
- assertCount
- assertNotCount


TestRunner
======

Usage
-----
```
TestRunner.wsf /d:DIRECTORY_PATH [/o:OUTPUT_LOGFILE]
TestRunner.wsf /?

  /?                  Print this help.
  /d:DIRECTORY_PATH   Target Directory Path
  /o:OUTPUT_LOGFILE   Output directory. Default Current directory.
```


Example
-----

```
cscript //nologo TestRunner.wsf /d:C:\VBS
cscript //nologo TestRunner.wsf /d:C:\VBS
cscript //nologo TestRunner.wsf /d:C:\VBS /o:C:\VBS\log
```

