''' --------------------------------------------------------
'''  FILE    : UtilScrapingEx.bas
'''  AUTHOR  : callmekohei <callmekohei at gmail.com>
'''  License : MIT license
'''  Respect : @kinuasa https://www.ka-net.org/blog/?p=4855
''' --------------------------------------------------------
Option Explicit

''' StdRegProv class (The StdRegProv class contains methods that manipulate system registry keys and values. )
''' https://docs.microsoft.com/en-us/previous-versions/windows/desktop/regprov/stdregprov
Private Enum HKeysEnum
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
End Enum

''' Internet Explorer security zones registry entries for advanced users
''' https://support.microsoft.com/en-us/help/182569/internet-explorer-security-zones-registry-entries-for-advanced-users
Public Enum ZoneEnum
    MyComputer = 0
    LocalIntranetZone = 1
    TrustedSitesZone = 2
    InternetZone = 3
    RestrictedSitesZone = 4
End Enum

''' ShowWindow
#If VBA7 Then
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Boolean
#Else
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Boolean
#End If

''' SetForegroundWindow
#If VBA7 Then
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If

''' IsZoomed
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Double) As LongPtr
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef procTime As Double) As LongPtr
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long) ''' param type is DWWORD
    Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long ''' C++ Bool is VBA's Long
#Else
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Double) As Long
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef procTime As Double) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
    Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
#End If

Private Declare PtrSafe Sub SwitchToThisWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal bool As Boolean)

#If VBA7 Then
    Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
    Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long


Public Sub BeforeScrapingWithIE()

    ''' make homepage blank page
    HomepageBlankOnly

    ''' ContinuousBrowsing, isolation, homepage tab
    prepareIE

    ''' clear all ie
    KillAllIE
    DoEvents
    Sleep 1500

    ''' only blank tab page
    BlankTab

End Sub

Private Sub prepareIE()

    Dim wmi As Object:    Set wmi = CreateObject("Wbemscripting.SWbemLocator")
    Dim wmiSrv As Object: Set wmiSrv = wmi.ConnectServer(".", "root\default")
    Dim oReg As Object:   Set oReg = wmiSrv.Get("StdRegProv")

'''HKEY_CURRENT_USER
    Const HKEY_CURRENT_USER As Long = HKeysEnum.HKEY_CURRENT_USER

    ''' ContinuousBrowsing
    Const strKeyPath_ContinuousBrowsing As String = "Software\Microsoft\Internet Explorer\ContinuousBrowsing"
    Const strValueName_ContinuousBrowsing As String = "Enabled"
    Const dwValue_ContinuousBrowsing As Long = 0  ''' 0 : Disabled , 1 : Enabled
    oReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath_ContinuousBrowsing, strValueName_ContinuousBrowsing, dwValue_ContinuousBrowsing

    ''' Isolation64Bit
    Const strKeyPath_Isolation64Bit As String = "Software\Microsoft\Internet Explorer\Main"
    Const strValueName_Isolation64Bit As String = "Isolation64Bit"
    Const dwValue_Isolation64Bit As Long = 0
    oReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath_Isolation64Bit, strValueName_Isolation64Bit, dwValue_Isolation64Bit

    ''' Isolation
    Const strKeyPath_Isolation As String = "Software\Microsoft\Internet Explorer\Main"
    Const strValueName_Isolation As String = "Isolation"
    Const strValue_Isolation As String = "PMIL"
    oReg.SetStringValue HKEY_CURRENT_USER, strKeyPath_Isolation, strValueName_Isolation, strValue_Isolation

    Set wmi = Nothing
    Set wmiSrv = Nothing
    Set oReg = Nothing

End Sub

Private Sub HomepageBlankOnly()

    Dim wmi As Object:    Set wmi = CreateObject("Wbemscripting.SWbemLocator")
    Dim wmiSrv As Object: Set wmiSrv = wmi.ConnectServer(".", "root\default")
    Dim oReg As Object:   Set oReg = wmiSrv.Get("StdRegProv")

    Const HKEY_CURRENT_USER As Long = HKeysEnum.HKEY_CURRENT_USER
    Const strKeyPath As String = "Software\Microsoft\Internet Explorer\Main"

    ''' delete second start page
    Const strValueName_SndStartPage As String = "Secondary Start Pages"
    Dim szValue_SndStartPage As Variant: oReg.GetMultiStringValue HKEY_CURRENT_USER, strKeyPath, strValueName_SndStartPage, szValue_SndStartPage
    If IsArray(szValue_SndStartPage) Then
        oReg.DeleteValue HKEY_CURRENT_USER, strKeyPath, strValueName_SndStartPage
    End If

    ''' make fist start page as blank page
    Const strValueName_FstStartPage As String = "Start Page"
    Dim szValue_FstStartPage As String
    oReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, strValueName_FstStartPage, szValue_FstStartPage

    If szValue_FstStartPage <> "about:blank" Then
        oReg.SetStringValue HKEY_CURRENT_USER, strKeyPath, strValueName_FstStartPage, "about:blank"
    End If

    Set wmi = Nothing
    Set wmiSrv = Nothing
    Set oReg = Nothing

End Sub

''' Don't prompt for client certificate selection when no certificates or only one certificate exists
Public Sub NotPromptClientCertificate(ByVal aZone As ZoneEnum)

    Dim wmi As Object:    Set wmi = CreateObject("Wbemscripting.SWbemLocator")
    Dim wmiSrv As Object: Set wmiSrv = wmi.ConnectServer(".", "root\default")
    Dim oReg As Object:   Set oReg = wmiSrv.Get("StdRegProv")

    Const HKEY_CURRENT_USER As Long = HKeysEnum.HKEY_CURRENT_USER

    Dim strKeyPath As String: strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\" & CStr(aZone)
    Const strValueName As String = "1A04"
    Const dwValue As Long = 0  ''' 0 : do not prompt , 3 : prompt
    oReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, strValueName, dwValue

End Sub

''' Prompt for client certificate selection when no certificates or only one certificate exists
Public Sub PromptClientCertificate(ByVal aZone As ZoneEnum)

    Dim wmi As Object:    Set wmi = CreateObject("Wbemscripting.SWbemLocator")
    Dim wmiSrv As Object: Set wmiSrv = wmi.ConnectServer(".", "root\default")
    Dim oReg As Object:   Set oReg = wmiSrv.Get("StdRegProv")

    Const HKEY_CURRENT_USER As Long = HKeysEnum.HKEY_CURRENT_USER

    Dim strKeyPath As String: strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\" & CStr(aZone)
    Const strValueName As String = "1A04"
    Const dwValue As Long = 3  ''' 0 : do not prompt , 3 : prompt
    oReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, strValueName, dwValue

End Sub

Public Sub AddURLOnTrustedSitesZone(ByVal aURL As String)

    Dim wmi As Object:    Set wmi = CreateObject("Wbemscripting.SWbemLocator")
    Dim wmiSrv As Object: Set wmiSrv = wmi.ConnectServer(".", "root\default")
    Dim oReg As Object:   Set oReg = wmiSrv.Get("StdRegProv")

    Const HKEY_CURRENT_USER As Long = HKeysEnum.HKEY_CURRENT_USER

    Dim strKeyPath As String: strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\" & aURL & "\www"
    oReg.CreateKey HKEY_CURRENT_USER, strKeyPath

    Const strValueName As String = "https"
    Const dwValue As Long = 2
    oReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, strValueName, dwValue

End Sub

Public Sub RemoveURLOnTrustedSitesZone(ByVal aURL As String)

    Dim wmi As Object:    Set wmi = CreateObject("Wbemscripting.SWbemLocator")
    Dim wmiSrv As Object: Set wmiSrv = wmi.ConnectServer(".", "root\default")
    Dim oReg As Object:   Set oReg = wmiSrv.Get("StdRegProv")

    Const HKEY_CURRENT_USER As Long = HKeysEnum.HKEY_CURRENT_USER
    Dim strKeyPath As String: strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\" & aURL
    oReg.DeleteKey HKEY_CURRENT_USER, strKeyPath

End Sub

Public Sub KillAllIE()
    Dim objShell As Object: Set objShell = CreateObject("WScript.Shell")
    Dim objExec As Object: Set objExec = objShell.Exec("taskkill.exe /F /IM iexplore.exe")
    Set objShell = Nothing
    Set objExec = Nothing
End Sub

Public Sub BlankTab()

    Dim blank_url As String: blank_url = "about:blank"
    Dim doc As ScrapingEx: Set doc = Nothing: Set doc = New ScrapingEx
    doc.GotoPage blank_url, True

    CloseTabsExceptBlanckTabs
    CloseTabsExceptBlanckTabs
    CloseTabsExceptBlanckTabs

    UniqBlankTab
    UniqBlankTab
    UniqBlankTab

    doc.Quit
    Set doc = Nothing

End Sub

Private Sub CloseTabsExceptBlanckTabs()

    On Error Resume Next

    Dim obj As Object
    For Each obj In CreateObject("Shell.Application").Windows
        If IsWindowVisible(obj.hwnd) Then
            If obj.Name = "Internet Explorer" Then
                If obj.LocationURL <> "about:blank" Then
                    obj.Quit
                    DoEvents
                    Sleep 100
                End If
            End If
        End If
    Next obj

End Sub

Private Sub UniqBlankTab()

    On Error Resume Next

    Dim blankPage_i As Long

    Dim obj As Object
    For Each obj In CreateObject("Shell.Application").Windows
        If IsWindowVisible(obj.hwnd) Then
            If obj.Name = "Internet Explorer" Then
                If obj.LocationURL = "about:blank" Then
                    blankPage_i = blankPage_i + 1
                End If
            End If
        End If
    Next obj

    Set obj = Nothing
    For Each obj In CreateObject("Shell.Application").Windows
        If IsWindowVisible(obj.hwnd) Then
            If obj.Name = "Internet Explorer" Then
                If obj.LocationURL = "about:blank" Then
                    If blankPage_i <> 1 Then
                        obj.Quit
                        DoEvents
                        Sleep 100
                        blankPage_i = blankPage_i - 1
                    End If
                End If
            End If
        End If
    Next obj

End Sub

Private Sub PushWindowsSecurityOKButton(ByVal specificURL As String, Optional ByVal wait_ms As Long = 2000)

    Dim obj As Object
    For Each obj In CreateObject("Shell.Application").Windows
        If IsWindowVisible(obj.hwnd) Then
            If obj.Name = "Internet Explorer" Then
                If obj.LocationURL = specificURL Then
                    SetForegroundWindow obj.hwnd
                    Sleep wait_ms
                    SendKeys "{TAB}", True
                    Sleep wait_ms
                    SendKeys "{ENTER}", True
                End If
            End If
        End If
    Next obj

End Sub

''' @param table As Object(Of MSHTML.HTMLTable)
''' @return As Variant(Of Array(Of Array (Of Array (Of HTMLTableCell Or String))))
Public Function ArrTable(ByVal table As Object _
    , Optional ByVal asInnerText As Boolean = False) As Variant

    Dim arr() As Variant: ReDim arr(0 To table.children.Length - 1)

    Dim i As Long
    For i = 0 To UBound(arr)
        arr(i) = ArrTableSection(table.children.Item(i), asInnerText)
    Next i

    ArrTable = arr

End Function

''' @param tblSct As Object(Of MSHTML.HTMLTableSection)
''' @return As Variant(Of Array(Of Array (Of HTMLTableCell Or String)))
Public Function ArrTableSection(ByVal tblSct As Object _
    , Optional ByVal asInnerText As Boolean = False) As Variant

    If TypeName(tblSct) = "object HTMLTableSectionElement" Then Err.Raise 13
    Dim n As Long: n = tblSct.children.Length
    Dim arr As Variant: arr = Array(): ReDim arr(0 To n - 1)

    Dim i As Long
    For i = 0 To n - 1
        If TypeName(tblSct.children.Item(i)) = "HTMLTableRow" Then
            arr(i) = ArrTableSection(tblSct.children.Item(i), asInnerText)
        Else '''HTMLTableCell
            If asInnerText Then
                arr(i) = tblSct.children.Item(i).innerText
            Else
                Set arr(i) = tblSct.children.Item(i)
            End If
        End If
    Next i

    ArrTableSection = arr

End Function

Public Sub AddReference()

    Const MSHTML_HTMLDocument As String = "{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}"
    AddReferenceImpl MSHTML_HTMLDocument, 4, 0

    Const SHDocVw_InternetExplorer As String = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}"
    AddReferenceImpl SHDocVw_InternetExplorer, 1, 1

End Sub

Private Sub AddReferenceImpl( _
      ByVal aGuid As String _
    , ByVal aMajor As Long _
    , ByVal aMinor As Long _
    , Optional ByVal additional_ms As Long = 500)

    On Error GoTo Catch
    Application.VBE.ActiveVBProject.References.AddFromGuid GUID:=aGuid, Major:=aMajor, Minor:=aMinor

Catch:
    Select Case Err.Number
        Case 0
            Sleep additional_ms
            Exit Sub
        Case 32813  ' already set path
            Exit Sub
        Case Else
            MsgBox Err.Description & vbCrLf & Err.Number
    End Select
End Sub

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


''' 通知バー/Internet Explorerダイアログを操作してファイルをダウンロード
Public Sub DownloadFileNotificationBar(ByVal hIE As Long, ByVal SaveFilePath As String)

  ''' ファイルを事前に削除
  With CreateObject("Scripting.FileSystemObject")
    If .FileExists(SaveFilePath) Then .DeleteFile SaveFilePath, True
  End With
  
  Dim uiAuto As CUIAutomation: Set uiAuto = New CUIAutomation
  
  ''' 通知バーの[別名で保存]を押す
  PressSaveAsMenuNotificationBar uiAuto, hIE
  ''' [名前を付けて保存]ダイアログ操作
  SaveAsFileNameDialog uiAuto, SaveFilePath
  ''' ダウンロード完了後、通知バーを閉じる
  ClosingNotificationBar uiAuto, hIE
  
  Set uiAuto = Nothing

End Sub

Private Function PressSaveAsMenuNotificationBar(ByRef uiAuto As CUIAutomation, ByVal ieWnd As Long)

    ''' 通知バーを取得
    Dim hwnd As Long
    Do
        DoEvents
        koffeetime.Wait 1&
        hwnd = FindWindowEx(ieWnd, 0, "Frame Notification Bar", vbNullString)
    Loop Until hwnd

    Do
        DoEvents
        koffeetime.Wait 1&
    Loop Until IsWindowVisible(hwnd)
    
    Dim elmNotificationBar As IUIAutomationElement: Set elmNotificationBar = uiAuto.ElementFromHandle(ByVal hwnd)
       
    ''' [保存] スプリットボタン取得
    Dim elmSaveSplitButton As IUIAutomationElement
    Do
        DoEvents
        koffeetime.Wait 1&
        Set elmSaveSplitButton = GetElement(uiAuto, elmNotificationBar, UIA_NamePropertyId, "保存", UIA_SplitButtonControlTypeId)
    Loop While elmSaveSplitButton Is Nothing
    
    ''' [保存] ドロップダウン取得
    Const ROLE_SYSTEM_BUTTONDROPDOWN = &H38&
    Dim elmSaveDropDownButton As IUIAutomationElement
    Do
        DoEvents
        koffeetime.Wait 1&
        Set elmSaveDropDownButton = GetElement(uiAuto, elmNotificationBar, UIA_LegacyIAccessibleRolePropertyId, ROLE_SYSTEM_BUTTONDROPDOWN, UIA_SplitButtonControlTypeId)
    Loop While elmSaveDropDownButton Is Nothing

    '''[保存]ドロップダウン押下
    Dim iptn As IUIAutomationInvokePattern
    Set iptn = elmSaveDropDownButton.GetCurrentPattern(UIA_InvokePatternId)

    ''' メニューウインドウを取得
    Dim elmSaveMenu As IUIAutomationElement
    Do
      iptn.Invoke
      Set elmSaveMenu = GetElement(uiAuto, uiAuto.GetRootElement, UIA_ClassNamePropertyId, "#32768", UIA_MenuControlTypeId)
      DoEvents
      koffeetime.Wait 1&
    Loop While elmSaveMenu Is Nothing
    
    ''' [名前を付けて保存(A)]ボタン押下
    Dim hWndSaveMenu As Long
    hWndSaveMenu = FindWindow("#32768", vbNullString)
    PostMessage hWndSaveMenu, &H106, vbKeyA, 0&   ' SYSCHAR=0x106
    
End Function

Private Function SaveAsFileNameDialog(ByRef uiAuto As CUIAutomation, ByVal SaveFilePath As String)

    '''[名前を付けて保存]ダイアログ取得
    Dim elmSaveAsWindow As IUIAutomationElement
    Do
      Set elmSaveAsWindow = GetElement(uiAuto, uiAuto.GetRootElement, UIA_NamePropertyId, "名前を付けて保存", UIA_WindowControlTypeId)
        DoEvents
        koffeetime.Wait 1&
    Loop While elmSaveAsWindow Is Nothing
     
    '[ファイル名]欄取得 -> ファイルパス入力
    Dim elmFileNameEdit As IUIAutomationElement: Set elmFileNameEdit = GetElement(uiAuto, elmSaveAsWindow, UIA_NamePropertyId, "ファイル名:", UIA_EditControlTypeId)
    Dim vptn As IUIAutomationValuePattern: Set vptn = elmFileNameEdit.GetCurrentPattern(UIA_ValuePatternId)
    vptn.SetValue SaveFilePath
     
    '[保存(S)]ボタン押下
    Dim elmSaveButton As IUIAutomationElement
    Do
    Set elmSaveButton = _
      GetElement(uiAuto, elmSaveAsWindow, UIA_NamePropertyId, "保存(S)", UIA_ButtonControlTypeId)
    Loop While elmSaveButton Is Nothing
    
    Dim iptn As IUIAutomationInvokePattern: Set iptn = elmSaveButton.GetCurrentPattern(UIA_InvokePatternId)
    iptn.Invoke

End Function

Private Function ClosingNotificationBar(ByRef uiAuto As CUIAutomation, ByVal ieWnd As Long)

    ''' 通知バーを取得
    Dim hwnd As Long
    Do
        DoEvents
        koffeetime.Wait 1&
        hwnd = FindWindowEx(ieWnd, 0, "Frame Notification Bar", vbNullString)
    Loop Until hwnd

    Do
        DoEvents
        koffeetime.Wait 1&
    Loop Until IsWindowVisible(hwnd)
    
    Dim elmNotificationBar As IUIAutomationElement: Set elmNotificationBar = uiAuto.ElementFromHandle(ByVal hwnd)
    
     
    ''' [通知バーのテキスト]取得
    Dim elmNotificationText As IUIAutomationElement
    Do
        DoEvents
        koffeetime.Wait 1&
        Set elmNotificationText = GetElement(uiAuto, elmNotificationBar, UIA_NamePropertyId, "通知バーのテキスト", UIA_TextControlTypeId)
    Loop While elmNotificationText Is Nothing
     
    ''' [閉じる]ボタン取得
    Dim elmCloseButton As IUIAutomationElement
    Do
        DoEvents
        koffeetime.Wait 1&
        Set elmCloseButton = GetElement(uiAuto, elmNotificationBar, UIA_NamePropertyId, "閉じる", UIA_ButtonControlTypeId)
    Loop While elmCloseButton Is Nothing
    
    
    ''' [閉じる]ボタン押下
    Do
      DoEvents
      koffeetime.Wait 1&
    Loop Until InStr(elmNotificationText.GetCurrentPropertyValue(UIA_ValueValuePropertyId), "ダウンロードが完了しました") > 0
    Dim iptn As IUIAutomationInvokePattern: Set iptn = elmCloseButton.GetCurrentPattern(UIA_InvokePatternId)
    iptn.Invoke

End Function

Private Function GetElement(ByVal uiAuto As CUIAutomation, _
                            ByVal elmParent As IUIAutomationElement, _
                            ByVal propertyId As Long, _
                            ByVal propertyValue As Variant, _
                            Optional ByVal ctrlType As Long = 0) As IUIAutomationElement
    Dim cndFirst As IUIAutomationCondition
    Dim cndSecond As IUIAutomationCondition
    
    Set cndFirst = uiAuto.CreatePropertyCondition(propertyId, propertyValue)
    If ctrlType <> 0 Then
        Set cndSecond = uiAuto.CreatePropertyCondition(UIA_ControlTypePropertyId, ctrlType)
        Set cndFirst = uiAuto.CreateAndCondition(cndFirst, cndSecond)
    End If
    Set GetElement = elmParent.FindFirst(TreeScope_Subtree, cndFirst)
End Function
