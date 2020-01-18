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

    ''' IEを最前面に出す

    SwitchToThisWindow doc.IEObj.hwnd, False

    ''' isZoomed
    Do
        ShowWindow doc.IEObj.hwnd, 3       ''' IEを最大化
        koffeetime.Wait 300
    Loop Until Not (IsZoomed(doc.IEObj.hwnd))

    SetForegroundWindow (doc.IEObj.hwnd)

End Sub

' ''' 通知バー/Internet Explorerダイアログを操作してファイルをダウンロード
' Public Sub DownloadFileNotificationBar(ByVal hIE As Long, ByVal SaveFilePath As String)

'   ''' ファイルを事前に削除
'   With CreateObject("Scripting.FileSystemObject")
'     If .FileExists(SaveFilePath) Then .DeleteFile SaveFilePath, True
'   End With

'   Dim uiAuto As CUIAutomation: Set uiAuto = New CUIAutomation

'   ''' 通知バーの[別名で保存]を押す
'   PressSaveAsMenuNotificationBar uiAuto, hIE
'   ''' [名前を付けて保存]ダイアログ操作
'   SaveAsFileNameDialog uiAuto, SaveFilePath
'   ''' ダウンロード完了後、通知バーを閉じる
'   ClosingNotificationBar uiAuto, hIE

'   Set uiAuto = Nothing

' End Sub

' Private Function PressSaveAsMenuNotificationBar(ByRef uiAuto As CUIAutomation, ByVal ieWnd As Long)

'     ''' 通知バーを取得
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

'     ''' [保存] スプリットボタン取得
'     Dim elmSaveSplitButton As IUIAutomationElement
'     Do
'         DoEvents
'         koffeetime.Wait 1&
'         Set elmSaveSplitButton = GetElement(uiAuto, elmNotificationBar, UIA_NamePropertyId, "保存", UIA_SplitButtonControlTypeId)
'     Loop While elmSaveSplitButton Is Nothing

'     ''' [保存] ドロップダウン取得
'     Const ROLE_SYSTEM_BUTTONDROPDOWN = &H38&
'     Dim elmSaveDropDownButton As IUIAutomationElement
'     Do
'         DoEvents
'         koffeetime.Wait 1&
'         Set elmSaveDropDownButton = GetElement(uiAuto, elmNotificationBar, UIA_LegacyIAccessibleRolePropertyId, ROLE_SYSTEM_BUTTONDROPDOWN, UIA_SplitButtonControlTypeId)
'     Loop While elmSaveDropDownButton Is Nothing

'     '''[保存]ドロップダウン押下
'     Dim iptn As IUIAutomationInvokePattern
'     Set iptn = elmSaveDropDownButton.GetCurrentPattern(UIA_InvokePatternId)

'     ''' メニューウインドウを取得
'     Dim elmSaveMenu As IUIAutomationElement
'     Do
'       iptn.Invoke
'       Set elmSaveMenu = GetElement(uiAuto, uiAuto.GetRootElement, UIA_ClassNamePropertyId, "#32768", UIA_MenuControlTypeId)
'       DoEvents
'       koffeetime.Wait 1&
'     Loop While elmSaveMenu Is Nothing

'     ''' [名前を付けて保存(A)]ボタン押下
'     Dim hWndSaveMenu As Long
'     hWndSaveMenu = FindWindow("#32768", vbNullString)
'     PostMessage hWndSaveMenu, &H106, vbKeyA, 0&   ' SYSCHAR=0x106

' End Function

' Private Function SaveAsFileNameDialog(ByRef uiAuto As CUIAutomation, ByVal SaveFilePath As String)

'     '''[名前を付けて保存]ダイアログ取得
'     Dim elmSaveAsWindow As IUIAutomationElement
'     Do
'       Set elmSaveAsWindow = GetElement(uiAuto, uiAuto.GetRootElement, UIA_NamePropertyId, "名前を付けて保存", UIA_WindowControlTypeId)
'         DoEvents
'         koffeetime.Wait 1&
'     Loop While elmSaveAsWindow Is Nothing

'     '[ファイル名]欄取得 -> ファイルパス入力
'     Dim elmFileNameEdit As IUIAutomationElement: Set elmFileNameEdit = GetElement(uiAuto, elmSaveAsWindow, UIA_NamePropertyId, "ファイル名:", UIA_EditControlTypeId)
'     Dim vptn As IUIAutomationValuePattern: Set vptn = elmFileNameEdit.GetCurrentPattern(UIA_ValuePatternId)
'     vptn.SetValue SaveFilePath

'     '[保存(S)]ボタン押下
'     Dim elmSaveButton As IUIAutomationElement
'     Do
'     Set elmSaveButton = _
'       GetElement(uiAuto, elmSaveAsWindow, UIA_NamePropertyId, "保存(S)", UIA_ButtonControlTypeId)
'     Loop While elmSaveButton Is Nothing

'     Dim iptn As IUIAutomationInvokePattern: Set iptn = elmSaveButton.GetCurrentPattern(UIA_InvokePatternId)
'     iptn.Invoke

' End Function

' Private Function ClosingNotificationBar(ByRef uiAuto As CUIAutomation, ByVal ieWnd As Long)

'     ''' 通知バーを取得
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


'     ''' [通知バーのテキスト]取得
'     Dim elmNotificationText As IUIAutomationElement
'     Do
'         DoEvents
'         koffeetime.Wait 1&
'         Set elmNotificationText = GetElement(uiAuto, elmNotificationBar, UIA_NamePropertyId, "通知バーのテキスト", UIA_TextControlTypeId)
'     Loop While elmNotificationText Is Nothing

'     ''' [閉じる]ボタン取得
'     Dim elmCloseButton As IUIAutomationElement
'     Do
'         DoEvents
'         koffeetime.Wait 1&
'         Set elmCloseButton = GetElement(uiAuto, elmNotificationBar, UIA_NamePropertyId, "閉じる", UIA_ButtonControlTypeId)
'     Loop While elmCloseButton Is Nothing


'     ''' [閉じる]ボタン押下
'     Do
'       DoEvents
'       koffeetime.Wait 1&
'     Loop Until InStr(elmNotificationText.GetCurrentPropertyValue(UIA_ValueValuePropertyId), "ダウンロードが完了しました") > 0
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
