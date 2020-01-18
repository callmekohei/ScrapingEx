Attribute VB_Name = "UtilUi"
Option Explicit

''' --------------------------------------------------------
'''  FILE    : UtilUI.bas
'''  AUTHOR  : callmekohei <callmekohei at gmail.com>
'''  License : MIT license
'''  Respect : @kinuasa https://www.ka-net.org/blog/?p=4855
''' --------------------------------------------------------

#If VBA7 Then
    Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
    Private Declare PtrSafe Sub SwitchToThisWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal bool As Boolean)
    Private Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Boolean
#Else
    Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare Sub SwitchToThisWindow Lib "user32" (ByVal hwnd As Long, ByVal bool As Boolean)
    Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Boolean
#End If

''' put InternetExplorer front
Public Sub IEForeGround(ByVal doc As Variant, ByVal ieTitelName As String)

    ''' IE���őO�ʂɏo��

    SwitchToThisWindow doc.IEObj.hwnd, False

    ''' isZoomed
    Do
        ShowWindow doc.IEObj.hwnd, 3       ''' IE���ő剻
        koffeetime.Wait 300
    Loop Until Not (IsZoomed(doc.IEObj.hwnd))

    SetForegroundWindow (doc.IEObj.hwnd)

End Sub

' ''' �ʒm�o�[/Internet Explorer�_�C�A���O�𑀍삵�ăt�@�C�����_�E�����[�h
' Public Sub DownloadFileNotificationBar(ByVal hIE As Long, ByVal SaveFilePath As String)

'   ''' �t�@�C�������O�ɍ폜
'   With CreateObject("Scripting.FileSystemObject")
'     If .FileExists(SaveFilePath) Then .DeleteFile SaveFilePath, True
'   End With

'   Dim uiAuto As CUIAutomation: Set uiAuto = New CUIAutomation

'   ''' �ʒm�o�[��[�ʖ��ŕۑ�]������
'   PressSaveAsMenuNotificationBar uiAuto, hIE
'   ''' [���O��t���ĕۑ�]�_�C�A���O����
'   SaveAsFileNameDialog uiAuto, SaveFilePath
'   ''' �_�E�����[�h������A�ʒm�o�[�����
'   ClosingNotificationBar uiAuto, hIE

'   Set uiAuto = Nothing

' End Sub

' Private Function PressSaveAsMenuNotificationBar(ByRef uiAuto As CUIAutomation, ByVal ieWnd As Long)

'     ''' �ʒm�o�[���擾
'     Dim hwnd As Long
'     Do
'         DoEvents
'         koffeetime.Wait 1&
'         hwnd = FindWindowEx(ieWnd, 0, "Frame Notification Bar", vbNullString)
'     Loop Until hwnd

'     Do
'         DoEvents
'         koffeetime.Wait 1&
'     Loop Until IsWindowVisible(hwnd)

'     Dim elmNotificationBar As IUIAutomationElement: Set elmNotificationBar = uiAuto.ElementFromHandle(ByVal hwnd)

'     ''' [�ۑ�] �X�v���b�g�{�^���擾
'     Dim elmSaveSplitButton As IUIAutomationElement
'     Do
'         DoEvents
'         koffeetime.Wait 1&
'         Set elmSaveSplitButton = GetElement(uiAuto, elmNotificationBar, UIA_NamePropertyId, "�ۑ�", UIA_SplitButtonControlTypeId)
'     Loop While elmSaveSplitButton Is Nothing

'     ''' [�ۑ�] �h���b�v�_�E���擾
'     Const ROLE_SYSTEM_BUTTONDROPDOWN = &H38&
'     Dim elmSaveDropDownButton As IUIAutomationElement
'     Do
'         DoEvents
'         koffeetime.Wait 1&
'         Set elmSaveDropDownButton = GetElement(uiAuto, elmNotificationBar, UIA_LegacyIAccessibleRolePropertyId, ROLE_SYSTEM_BUTTONDROPDOWN, UIA_SplitButtonControlTypeId)
'     Loop While elmSaveDropDownButton Is Nothing

'     '''[�ۑ�]�h���b�v�_�E������
'     Dim iptn As IUIAutomationInvokePattern
'     Set iptn = elmSaveDropDownButton.GetCurrentPattern(UIA_InvokePatternId)

'     ''' ���j���[�E�C���h�E���擾
'     Dim elmSaveMenu As IUIAutomationElement
'     Do
'       iptn.Invoke
'       Set elmSaveMenu = GetElement(uiAuto, uiAuto.GetRootElement, UIA_ClassNamePropertyId, "#32768", UIA_MenuControlTypeId)
'       DoEvents
'       koffeetime.Wait 1&
'     Loop While elmSaveMenu Is Nothing

'     ''' [���O��t���ĕۑ�(A)]�{�^������
'     Dim hWndSaveMenu As Long
'     hWndSaveMenu = FindWindow("#32768", vbNullString)
'     PostMessage hWndSaveMenu, &H106, vbKeyA, 0&   ' SYSCHAR=0x106

' End Function

' Private Function SaveAsFileNameDialog(ByRef uiAuto As CUIAutomation, ByVal SaveFilePath As String)

'     '''[���O��t���ĕۑ�]�_�C�A���O�擾
'     Dim elmSaveAsWindow As IUIAutomationElement
'     Do
'       Set elmSaveAsWindow = GetElement(uiAuto, uiAuto.GetRootElement, UIA_NamePropertyId, "���O��t���ĕۑ�", UIA_WindowControlTypeId)
'         DoEvents
'         koffeetime.Wait 1&
'     Loop While elmSaveAsWindow Is Nothing

'     '[�t�@�C����]���擾 -> �t�@�C���p�X����
'     Dim elmFileNameEdit As IUIAutomationElement: Set elmFileNameEdit = GetElement(uiAuto, elmSaveAsWindow, UIA_NamePropertyId, "�t�@�C����:", UIA_EditControlTypeId)
'     Dim vptn As IUIAutomationValuePattern: Set vptn = elmFileNameEdit.GetCurrentPattern(UIA_ValuePatternId)
'     vptn.SetValue SaveFilePath

'     '[�ۑ�(S)]�{�^������
'     Dim elmSaveButton As IUIAutomationElement
'     Do
'     Set elmSaveButton = _
'       GetElement(uiAuto, elmSaveAsWindow, UIA_NamePropertyId, "�ۑ�(S)", UIA_ButtonControlTypeId)
'     Loop While elmSaveButton Is Nothing

'     Dim iptn As IUIAutomationInvokePattern: Set iptn = elmSaveButton.GetCurrentPattern(UIA_InvokePatternId)
'     iptn.Invoke

' End Function

' Private Function ClosingNotificationBar(ByRef uiAuto As CUIAutomation, ByVal ieWnd As Long)

'     ''' �ʒm�o�[���擾
'     Dim hwnd As Long
'     Do
'         DoEvents
'         koffeetime.Wait 1&
'         hwnd = FindWindowEx(ieWnd, 0, "Frame Notification Bar", vbNullString)
'     Loop Until hwnd

'     Do
'         DoEvents
'         koffeetime.Wait 1&
'     Loop Until IsWindowVisible(hwnd)

'     Dim elmNotificationBar As IUIAutomationElement: Set elmNotificationBar = uiAuto.ElementFromHandle(ByVal hwnd)


'     ''' [�ʒm�o�[�̃e�L�X�g]�擾
'     Dim elmNotificationText As IUIAutomationElement
'     Do
'         DoEvents
'         koffeetime.Wait 1&
'         Set elmNotificationText = GetElement(uiAuto, elmNotificationBar, UIA_NamePropertyId, "�ʒm�o�[�̃e�L�X�g", UIA_TextControlTypeId)
'     Loop While elmNotificationText Is Nothing

'     ''' [����]�{�^���擾
'     Dim elmCloseButton As IUIAutomationElement
'     Do
'         DoEvents
'         koffeetime.Wait 1&
'         Set elmCloseButton = GetElement(uiAuto, elmNotificationBar, UIA_NamePropertyId, "����", UIA_ButtonControlTypeId)
'     Loop While elmCloseButton Is Nothing


'     ''' [����]�{�^������
'     Do
'       DoEvents
'       koffeetime.Wait 1&
'     Loop Until InStr(elmNotificationText.GetCurrentPropertyValue(UIA_ValueValuePropertyId), "�_�E�����[�h���������܂���") > 0
'     Dim iptn As IUIAutomationInvokePattern: Set iptn = elmCloseButton.GetCurrentPattern(UIA_InvokePatternId)
'     iptn.Invoke

' End Function

' Private Function GetElement(ByVal uiAuto As CUIAutomation, _
'                             ByVal elmParent As IUIAutomationElement, _
'                             ByVal propertyId As Long, _
'                             ByVal propertyValue As Variant, _
'                             Optional ByVal ctrlType As Long = 0) As IUIAutomationElement
'     Dim cndFirst As IUIAutomationCondition
'     Dim cndSecond As IUIAutomationCondition

'     Set cndFirst = uiAuto.CreatePropertyCondition(propertyId, propertyValue)
'     If ctrlType <> 0 Then
'         Set cndSecond = uiAuto.CreatePropertyCondition(UIA_ControlTypePropertyId, ctrlType)
'         Set cndFirst = uiAuto.CreateAndCondition(cndFirst, cndSecond)
'     End If
'     Set GetElement = elmParent.FindFirst(TreeScope_Subtree, cndFirst)
' End Function
