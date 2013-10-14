Option Explicit

'! VBScript �p Unit Test
'!
'! @author  miyamiya <rai.caver@gmail.com>
'! @date    2013/08/14
'! @version 0.2
Class VBSUnit

    '! ���O���L�^����
    '!
    '! @var dic          object Dictionary Object
    Private dic

    '! ���݃e�X�g���̃��\�b�h��
    '!
    '! @var name         string ���\�b�h��
    Private name

    '! ���ʂ�Ԃ�
    '!
    '! @param  void
    '! @return           object Dictionary Object
    Public Function getResult()
        Set getResult = dic
    End Function

    '! ���\�b�h�����Z�b�g����
    '!
    '! @param  theName     string �e�X�g���s���N���X��
    '! @return void
    Public Property Let setName(byval theName)
        name = theName
        dic.Add name, CreateObject("Scripting.Dictionary")
        dic(name).Add "cnt"    , 0
        dic(name).Add "success", CreateObject("Scripting.Dictionary")
        dic(name).Add "failure", CreateObject("Scripting.Dictionary")
        dic(name).Add "error"  , CreateObject("Scripting.Dictionary")
    End Property

    '! 2�̕ϐ�(���l�A�����A�z��)���������Ƃ��ɐ���
    '!
    '! @see    assertEqualsLogic
    '! @param  theExpected   mixed �������鍶��
    '! @param  theActual     mixed ��������E��
    '! @return               text ���(success, failure, error)
    Public Function assertEquals(theExpected, theActual)
        assertEquals = assertEqualsLogic(theExpected, theActual, True)
    End Function

    '! 2�̕ϐ�(���l�A�����A�z��)���������Ȃ��Ƃ�����
    '!
    '! @see    assertEqualsLogic
    '! @param  theExpected     mixed �������鍶��
    '! @param  theActual       mixed ��������E��
    '! @return void
    Public Function assertNotEquals(theExpected, theActual)
          assertNotEquals = assertEqualsLogic(theExpected, theActual, False)
    End Function

    '! �ϐ��� True �̂Ƃ��ɐ���
    '!
    '! @see    assertBooleanLogic
    '! @param  theCondition   boolean ��������l
    '! @return void
    Public Function assertTrue(theCondition)
        assertTrue = assertBooleanLogic(theCondition, True)
    End Function

    '! �ϐ��� False �̂Ƃ��ɐ���
    '!
    '! @see    assertBooleanLogic
    '! @param  theCondition   boolean ��������l
    '! @return void
    Public Function assertFalse(theCondition)
        assertFalse = assertBooleanLogic(theCondition, False)
    End Function

    '! �ϐ��� Null �̂Ƃ��ɐ���
    '!
    '! @param  theCondition   mixed ��������l
    '! @return void
    Public Function assertNull(theCondition)
        If IsNull(theCondition) = True Then
            assertNull = addSuccess(gen1ArgToMsg(theCondition))
            Exit Function
        End If
        assertNull = addFailure(gen1ArgToMsg(theCondition))
    End Function

    '! �ϐ��� Null �ȊO�̂Ƃ��ɐ���
    '!
    '! @param  theA     mixed ��������l
    '! @return void
    Public Function assertNotNull(theCondition)
        If IsNull(theCondition) = False Then
            assertNotNull = addSuccess(gen1ArgToMsg(theCondition))
            Exit Function
        End If
        assertNotNull = addFailure(gen1ArgToMsg(theCondition))
    End Function


    '! �ϐ��̒l����̂Ƃ��ɐ���
    '!
    '! @param  theCondition   mixed ��������l
    '! @return void
    Public Function assertEmpty(theCondition)
        If is_empty(theCondition) = True Then
            assertEmpty = addSuccess(gen1ArgToMsg(theCondition))
            Exit Function
        End If
        assertEmpty = addFailure(gen1ArgToMsg(theCondition))
    End Function

    '! �ϐ��̒l����łȂ��Ƃ��ɐ���
    '!
    '! @param  theCondition   mixed ��������l
    '! @return void
    Public Function assertNotEmpty(theCondition)
        If is_empty(theCondition) = False Then
            assertNotEmpty = addSuccess(gen1ArgToMsg(theCondition))
            Exit Function
        End If
        assertNotEmpty = addFailure(gen1ArgToMsg(theCondition))
    End Function

    '! ��2�����̔z��̒��ɁA��1�������܂܂�Ă���Ƃ��ɐ���
    '!
    '! @see    assertContainsLogic
    '! @param  theNeedles     mixed �����ΏۂƂȂ�l
    '! @param  theHaystack    array ��������z��
    '! @return void
    Public Function assertContains(theNeedles, theHaystack)
        assertContains = assertContainsLogic(theNeedles, theHaystack, True)
    End Function

    '! ��2�����̔z��̒��ɁA��1�������܂܂�Ă��Ȃ��Ƃ��ɐ���
    '!
    '! @see    assertContainsLogic
    '! @param  theNeedles     mixed �����ΏۂƂȂ�l
    '! @param  theHaystack    array ��������z��
    '! @return void
    Public Function assertNotContains(theNeedles, theHaystack)
        assertNotContains = assertContainsLogic(theNeedles, theHaystack, False)
    End Function

    '! ��2��������1�����Ŏn�܂��Ă���Ƃ��ɐ���
    '!
    '! @see    assertStringStartsLogic
    '! @param  thePrefix      string �����ΏۂƂȂ�l
    '! @param  theString      string �������镶����
    '! @return void
    Public Function assertStringStartsWith(thePrefix, theString)
        assertStringStartsWith = assertStringStartsLogic(thePrefix, theString, False, True)
    End Function

    '! ��2��������1�����Ŏn�܂��Ă��Ȃ��Ƃ��ɐ���
    '!
    '! @see    assertStringStartsLogic
    '! @param  thePrefix      string �����ΏۂƂȂ�l
    '! @param  theString      string �������镶����
    '! @return void
    Public Function assertStringStartsNotWith(thePrefix, theString)
        assertStringStartsNotWith = assertStringStartsLogic(thePrefix, theString, False, False)
    End Function

    '! ��2��������1�����ŏI����Ă���Ƃ��ɐ���
    '!
    '! @param  theSuffix      string �����ΏۂƂȂ�l
    '! @param  theString      string �������镶����
    '! @return void
    Public Function assertStringEndsWith(theSuffix, theString)
        assertStringEndsWith = assertStringStartsLogic(theSuffix, theString, True, True)
    End Function

    '! ��2��������1�����ŏI����Ă��Ȃ��Ƃ��ɐ���
    '!
    '! @param  theSuffix      string �����ΏۂƂȂ�l
    '! @param  theString      string �������镶����
    '! @return void
    Public Function assertStringEndsNotWith(theSuffix, theString)
        assertStringEndsNotWith = assertStringStartsLogic(theSuffix, theString, True, False)
    End Function

    '! ��2�����̌^����1�����̂Ƃ��ɐ���
    '!
    '! @see    assertInternalTypeLogic
    '! @param  theExpected    string �^�C�v
    '!           Empty     �F��������
    '!           Null      �FNull �l
    '!           Integer   �F�����^
    '!           Long      �F�������^
    '!           Single    �F�P���x���������_���^
    '!           Double    �F�{���x���������_���^
    '!           Currency  �F�ʉ݌^
    '!           Date      �F���t�^
    '!           String    �F������^
    '!           Object    �F�I�[�g���[�V�����I�u�W�F�N�g�^
    '!           Error     �F�G���[�^
    '!           Boolean   �F�u�[���^
    '!           Variant   �F�o���A���g�^
    '!           DataObject�F��I�[�g���[�V�����I�u�W�F�N�g�^
    '!           Byte      �F�o�C�g�^
    '!           Array     �F�z��
    '! @param  theActual      mixed  ��������l
    '! @return void
    Public Function assertInternalType(theExpected, theActual)
        assertInternalType = assertInternalTypeLogic(theExpected, theActual, True)
    End Function

    '! ��2�����̌^����1�����̂łȂ��Ƃ��ɐ���
    '!
    '! @see    assertInternalTypeLogic
    '! @param  theExpected    string �^�C�v
    '!           Empty     �F��������
    '!           Null      �FNull �l
    '!           Integer   �F�����^
    '!           Long      �F�������^
    '!           Single    �F�P���x���������_���^
    '!           Double    �F�{���x���������_���^
    '!           Currency  �F�ʉ݌^
    '!           Date      �F���t�^
    '!           String    �F������^
    '!           Object    �F�I�[�g���[�V�����I�u�W�F�N�g�^
    '!           Error     �F�G���[�^
    '!           Boolean   �F�u�[���^
    '!           Variant   �F�o���A���g�^
    '!           DataObject�F��I�[�g���[�V�����I�u�W�F�N�g�^
    '!           Byte      �F�o�C�g�^
    '!           Array     �F�z��
    '! @param  theActual      mixed  ��������l
    '! @return void
    Public Function assertNotInternalType(theExpected, theActual)
        assertNotInternalType = assertInternalTypeLogic(theExpected, theActual, False)
    End Function

    '! ��1�����̏ꏊ�ɂ���e�L�X�g�t�@�C���̒������2�����̕������܂܂�Ă���Ƃ��ɐ���
    '!
    '| @see    assertStringEqualsFileLogic
    '! @param  theExpectedFile    string �����ΏۂƂȂ�t�@�C���̃p�X
    '! @param  theActualString    string �������镶����
    '! @return void
    Public Function assertStringEqualsFile(theExpectedFile, theActualString)
        assertStringEqualsFile = assertStringEqualsFileLogic(theExpectedFile, theActualString, True)
    End Function

    '! ��1�����̏ꏊ�ɂ���e�L�X�g�t�@�C���̒������2�����̕������܂܂�Ă��Ȃ��Ƃ��ɐ���
    '!
    '| @see    assertStringEqualsFileLogic
    '! @param  thePrefix      string �����ΏۂƂȂ�t�@�C���̃p�X
    '! @param  theString      string �������镶����
    '! @return void
    Public Function assertStringNotEqualsFile(theExpectedFile, theActualString)
        assertStringNotEqualsFile = assertStringEqualsFileLogic(theExpectedFile, theActualString, False)
    End Function

    '! ��1�����̏ꏊ�ɂ���e�L�X�g�t�@�C���Ƒ�2�����̏ꏊ�ɂ���e�L�X�g�t�@�C���������Ƃ��ɐ���
    '!
    '| @see    assertFileEqualsLogic
    '! @param  theExpected    string �����ΏۂƂȂ�t�@�C���̃p�X
    '! @param  theActual      string ��r�ΏۂƂȂ�t�@�C���̃p�X
    '! @return void
    Public Function assertFileEquals(theExpected, theActual)
        assertFileEquals = assertFileEqualsLogic(theExpected, theActual, True)
    End Function


    '! ��1�����̏ꏊ�ɂ���e�L�X�g�t�@�C���Ƒ�1�����̏ꏊ�ɂ���e�L�X�g�t�@�C�����Ⴄ�Ƃ��ɐ���
    '!
    '| @see    assertFileEqualsLogic
    '! @param  theExpected    string �����ΏۂƂȂ�t�@�C���̃p�X
    '! @param  theActual      string ��r�ΏۂƂȂ�t�@�C���̃p�X
    '! @return void
    Public Function assertFileNotEquals(theExpected, theActual)
        assertFileNotEquals = assertFileEqualsLogic(theExpected, theActual, False)
    End Function

    '! ��1�����̏ꏊ�Ƀt�@�C�������݂��Ȃ����Ɏ��s��񍐂���
    '!
    '| @see    file_exists
    '! @param  theFileName      string �����ΏۂƂȂ�t�@�C���̃p�X
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

    '! ��1�����̏ꏊ�Ƀt�@�C�������݂��鎞�Ɏ��s��񍐂���
    '!
    '| @see    file_exists
    '! @param  theFileName      string �����ΏۂƂȂ�t�@�C���̃p�X
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

    '! theExpected < theActual �̎��ɐ����A����ȊO�Ŏ��s��񍐂���
    '!
    '! @see    assertRelationalOperatorLogic
    '! @param  theExpected    ��r���鐔�l�̍���
    '! @param  theActual      ��r���鐔�l�̉E��
    '! @return void
    Public Function assertGreaterThan(theExpected, theActual)
        assertGreaterThan = assertRelationalOperatorLogic(theExpected, theActual, "<")
    End Function

    '! theExpected <= theActual �̎��ɐ����A����ȊO�Ŏ��s��񍐂���
    '!
    '! @see    assertRelationalOperatorLogic
    '! @param  theExpected    ��r���鐔�l�̍���
    '! @param  theActual      ��r���鐔�l�̉E��
    '! @return void
    Public Function assertGreaterThanOrEqual(theExpected, theActual)
        assertGreaterThanOrEqual = assertRelationalOperatorLogic(theExpected, theActual, "<=")
    End Function

    '! theExpected > theActual �̎��ɐ����A����ȊO�Ŏ��s��񍐂���
    '!
    '! @see    assertRelationalOperatorLogic
    '! @param  theExpected    ��r���鐔�l�̍���
    '! @param  theActual      ��r���鐔�l�̉E��
    '! @return void
    Public Function assertLessThan(theExpected, theActual)
        assertLessThan = assertRelationalOperatorLogic(theExpected, theActual, ">")
    End Function

    '! theExpected >= theActual �̎��ɐ����A����ȊO�Ŏ��s��񍐂���
    '!
    '! @see    assertRelationalOperatorLogic
    '! @param  theExpected    ��r���鐔�l�̍���
    '! @param  theActual      ��r���鐔�l�̉E��
    '! @return void
    Public Function assertLessThanOrEqual(theExpected, theActual)
        assertLessThanOrEqual = assertRelationalOperatorLogic(theExpected, theActual, ">=")
    End Function

    '! ��2�����̔z��̗v�f���Ƒ�1�������Ⴄ���Ɏ��s��񍐂���
    '!
    '! @see    assertCountLogic
    '! @param  theExpectedCount  integer ��r���鐔��
    '! @param  theHaystack       array �Ώۂ̔z��
    '! @return void
    Public Function assertCount(theExpectedCount, theHaystack)
        assertCount = assertCountLogic(theExpectedCount, theHaystack, True)
    End Function


    '! ��2�����̔z��̗v�f���Ƒ�1�������������Ɏ��s��񍐂���
    '!
    '! @see    assertCountLogic
    '! @param  theExpectedCount  integer ��r���鐔��
    '! @param  theHaystack       array �Ώۂ̔z��
    '! @return void
    Public Function assertNotCount(theExpectedCount, theHaystack)
        assertNotCount = assertCountLogic(theExpectedCount, theHaystack, False)
    End Function

    '! assertEquals, assertNotEquals �̃��W�b�N����
    '!
    '! @param  theExpectedCount  integer ��r���鐔��
    '! @param  theHaystack       array �Ώۂ̔z��
    '! @param  theFlg            boolean �����̎��̕Ԃ�l
    '! @return                   text ���(success, failure, error)
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

    '! assertEquals, assertNotEquals �̃��W�b�N����
    '!
    '! @see    isEqual
    '! @param  theExpected   mixed   �������鍶��
    '! @param  theActual     mixed   ��������E��
    '! @param  theFlg        boolean �����̎��̕Ԃ�l
    '! @return               text ���(success, failure, error)
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

    '! assertEquals, assertNotEquals �̃��W�b�N����
    '!
    '! @param  theCondition  mixed   ��������l
    '! @param  theFlg        boolean �����̎��̕Ԃ�l
    '! @return               text ���(success, failure, error)
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

    '! assertContains, assertNotContains �̃��W�b�N����
    '!
    '! @param  theNeedles     mixed �����ΏۂƂȂ�l
    '! @param  theHaystack    array ��������z��
    '! @param  theFlg         boolean �����̎��̕Ԃ�l
    '! @return                text ���(success, failure, error)
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

    '! assertStringStartsWith, assertStringStartsNotWith, assertStringEndsWith, assertStringEndsNotWith �̃��W�b�N����
    '!
    '! @param  theTarget      string �����ΏۂƂȂ�l
    '! @param  theString      string �������镶����
    '! @param  theReverse     boolean ������𔽓]����Ȃ� True
    '! @param  theFlg         boolean �����̎��̕Ԃ�l
    '! @return                text ���(success, failure, error)
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

    '! assertInternalType, assertNotInternalType �̃��W�b�N����
    '!
    '! @see    equalInternalType
    '! @param  theExpected   string  �^�C�v
    '! @param  theActual     mixed   ��������E��
    '! @param  theFlg        boolean �����̎��̕Ԃ�l
    '! @return               text ���(success, failure, error)
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

    '! assertStringEqualsFile, assertStringNotEqualsFile �̃��W�b�N����
    '!
    '| @see    doSearchText
    '! @param  thePrefix      string �����ΏۂƂȂ�t�@�C���̃p�X
    '! @param  theString      string �������镶����
    '! @param  theFlg         boolean �����̎��̕Ԃ�l
    '! @return                text ���(success, failure, error)
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

    '! assertFileEquals, assertFileNotEquals �̃��W�b�N����
    '!
    '| @see    doFileDiff
    '! @param  theExpected    string �����ΏۂƂȂ�t�@�C���̃p�X
    '! @param  theActual      string ��r�ΏۂƂȂ�t�@�C���̃p�X
    '! @param  theFlg         boolean �����̎��̕Ԃ�l
    '! @return                text ���(success, failure, error)
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

    '! assertGreaterThan, assertGreaterThanOrEqual, assertLessThan, assertLessThanOrEqual �̃��W�b�N����
    '!
    '! @see    is_numeric
    '! @param  theExpected    ��r���鐔�l�̍���
    '! @param  theActual      ��r���鐔�l�̉E��
    '! @param  theOperator    string ��r���Z�q
    '! @return                text ���(success, failure, error)
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

    '! ���O�ɓf���o����`���ɕϊ�����
    '!
    '! @param  theArg        mixed �ϊ��������
    '! @return               String �`��
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

    '! ��̈��������O�p���b�Z�[�W�ɉ��H����
    '!
    '! @param   theArg      mixed:����
    '! @return              String:���b�Z�[�W
    Private Function gen1ArgToMsg(theArg)
        gen1ArgToMsg = "Arg => " & toStringValue(theArg)
    End Function

    '! ��̈��������O�p���b�Z�[�W�ɉ��H����
    '!
    '! @param   theArg1     mixed:����
    '! @param   theArg2     mixed:����
    '! @return              String:���b�Z�[�W
    Private Function gen2ArgToMsg(theArg1, theArg2)
        gen2ArgToMsg = "Arg1 => " & toStringValue(theArg1) _
                     & ", Arg2 => " & toStringValue(theArg2)
    End Function

    '! �G���[���������Ă��鎞�ɃG���[���b�Z�[�W��Ԃ�
    '!
    '! @param  void
    '! @return             �G���[���b�Z�[�W
    Private Function genErrorMsg()
        If (Err.Number = 0) Then
            genErrorMsg = ""
            Exit Function
        End If
        genErrorMsg = Err.Description & "(" & Err.Number & ")"
    End Function

    '! �����̃J�E���^�[��i�߂�
    '!
    '! @param  theMsg     string �������b�Z�[�W
    '! @return void
    Private Function addSuccess(theMsg)
        With dic.item(name)
            .item("success").Add .item("cnt"), theMsg
            .item("cnt") = .item("cnt") + 1
        End With
        addSuccess = "success"
    End Function

    '! ���s�̃J�E���^�[��i�߂�
    '!
    '! @param  theMsg     string ���s���b�Z�[�W
    '! @return void
    Private Function addFailure(theMsg)
        With dic.item(name)
            .item("failure").Add .item("cnt"), theMsg
            .item("cnt") = .item("cnt") + 1
        End With
        addFailure = "failure"
    End Function

   '! �G���[�̃J�E���^�[��i�߂�
    '!
    '! @param  theMsg     string �G���[���b�Z�[�W
    '! @return void
    Private Function addError(theMsg)
        With dic.item(name)
            .item("error").Add .item("cnt"), theMsg
            .item("cnt") = .item("cnt") + 1
        End With
        addError = "error"
    End Function

    '! �R���X�g���N�^
    '!
    '! @param  void
    '! @return void
    Private Sub Class_Initialize()
        Set dic = CreateObject("Scripting.Dictionary")
    End Sub

    '! �f�X�g���N�^
    '!
    '! @param  void
    '! @return void
    Private Sub Class_Terminate()
        Set dic = Nothing
    End Sub

    '! 2�̕ϐ�(���l�A�����A�z��)����������r����
    '!
    '! @see    isValueEqual, isArrayEqual
    '! @param  theExpected   mixed �������鍶��
    '! @param  theActual     mixed ��������E��
    '! @return               string success:�������Afailure:�������Ȃ�, error: �G���[
    Private Function isEqual(theExpected, theActual)
        If IsArray(theExpected) = True And IsArray(theActual) = True Then
            isEqual = isArrayEqual(theExpected, theActual)
            Exit Function
        End If

        isEqual = isValueEqual(theExpected, theActual)
    End Function

    '! 2�̔z�񂪓�������r����
    '!
    '! @see    isValueEqual
    '! @param  theExpected   array �������鍶��
    '! @param  theActual     array ��������E��
    '! @return               boolean True:�������AFalse:�������Ȃ�
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

    '! 2�̕ϐ�(���l�A����)����������r����
    '!
    '! @param  theExpected   mixed �������鍶��
    '! @param  theActual     mixed ��������E��
    '! @return               boolean True:�������AFalse:�������Ȃ�
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


    '! �z�񂪎g�p�\���ǂ������肷��
    '!
    '! @param  theA     array ��������l
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

    '! �l����Ɠ����̒l���ǂ����m�F
    '!
    '! @param  theA     mixed ��������l
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

    '! ��2�����̔z��̒��ɁA��1�������܂܂�Ă��鎞�Ɏ��s��񍐂���
    '!
    '! @param  theNeedles     mixed �����ΏۂƂȂ�l
    '! @param  theHaystack    array ��������z��
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

    '!  ��2�����̌^����1�����̎��� True ��Ԃ�
    '!
    '! @param  theExpected    string �^�C�v
    '! @param  theActual      mixed  ��������l
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

    '! theExpectedFile �̏ꏊ�ɂ���e�L�X�g�t�@�C���̒����� theActualString ����������
    '!
    '! @param  thePrefix      string �����ΏۂƂȂ�t�@�C���̃p�X
    '! @param  theString      string �������镶����
    '! @return boolean        �����񂪃q�b�g������ True, ���Ȃ����� False
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

    '! ���������l�^�Ȃ� True ��Ԃ�
    '!
    '! @param  theValue    �^���m�F����l
    '! @return             True: ���l�AFalse: ���l�ȊO
    Private Function is_numeric(theValue)
        Select Case VarType(theValue)
            Case vbInteger, vbLong, vbSingle, vbDouble
                is_numeric = True
            Case Else
                is_numeric = False
        End Select
    End Function

    '! ��1�����̏ꏊ�ɂ���e�L�X�g�t�@�C���Ƒ�2�����̏ꏊ�ɂ���e�L�X�g�t�@�C�����r����
    '!
    '! @param  theExpected    string �����ΏۂƂȂ�t�@�C���̃p�X
    '! @param  theActual      string ��r�ΏۂƂȂ�t�@�C���̃p�X
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

    '! �t�@�C���̑��݊m�F
    '!
    '! @param  theFileName      string �����ΏۂƂȂ�t�@�C���̃p�X
    '! @return                  True: ���݁AFalse: �Ȃ��ANull: �����G���[
    Private Function file_exists(theFileName)
        If (VarType(theFileName) <> vbString) Then
            file_exists = Null
            Exit Function
        End If
        file_exists = CreateObject("Scripting.FileSystemObject").FileExists(theFileName)
    End Function


End Class
