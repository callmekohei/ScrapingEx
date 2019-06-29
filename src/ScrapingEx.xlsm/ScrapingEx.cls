VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScrapingEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' --------------------------------------------------------
'''  FILE    : ScrapingEX.cls
'''  AUTHOR  : callmekohei <callmekohei at gmail.com>
'''  License : MIT license
'''  Respect : Victor Zevallos - https://github.com/vba-dev/vba-Scraping
'''  Respect : kumatti1 - https://gist.github.com/kumatti1/6b68ea65fdfc9ecf727f
''' --------------------------------------------------------
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare PtrSafe _
Function IIDFromString Lib "ole32.dll" ( _
    ByVal lpsz As LongPtr, _
    ByVal lpiid As LongPtr _
) As Long

Private Declare PtrSafe _
Function DispCallFunc Lib "OleAut32.dll" ( _
    ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc_ As Long, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByRef prgvt As Integer, _
    ByRef prgpvarg As LongPtr, _
    ByRef pvargResult As Variant _
    ) As Long
Const CC_STDCALL = 4&

Private Declare PtrSafe _
Function CoGetObject Lib "Ole32" ( _
        ByVal pszName As LongPtr, _
        ByVal pBindOptions As LongPtr, _
        ByRef riid As GUID, _
        ByRef ppv As Any) As Long

Private Declare PtrSafe _
Function CoCreateInstance Lib "Ole32" ( _
        ByRef rclsid As GUID, _
        ByVal pUnkOuter As LongPtr, _
        ByVal dwClsContext As Long, _
        ByRef riid As GUID, _
        ByRef ppv As Any) As Long


''' QueryPerformanceCounter: https://docs.microsoft.com/en-us/windows/desktop/api/profileapi/nf-profileapi-queryperformancecounter
''' QueryPerformanceFrequency: https://docs.microsoft.com/en-us/windows/desktop/api/profileapi/nf-profileapi-queryperformancefrequency
''' Sleep: https://docs.microsoft.com/en-us/windows/desktop/api/synchapi/nf-synchapi-sleep
''' SetForegroundWindow : https://msdn.microsoft.com/ja-jp/windows/ms633539(v=vs.80)
''' IsWindowVisible: https://docs.microsoft.com/en-us/windows/desktop/api/winuser/nf-winuser-iswindowvisible

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Double) As LongPtr
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef procTime As Double) As LongPtr
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long) ''' param type is DWWORD
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long ''' C++ Bool is VBA's Long
#Else
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Double) As Long
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef procTime As Double) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
#End If

#Const HAS_REF = True

#If HAS_REF Then
    Private WithEvents gIE  As SHDocVw.InternetExplorer
Attribute gIE.VB_VarHelpID = -1
    Private WithEvents gDoc As MSHTML.HTMLDocument
Attribute gDoc.VB_VarHelpID = -1
    Private gElement        As MSHTML.IHTMLElement
    Private gChildren       As MSHTML.IHTMLDOMChildrenCollection
#Else
    Private gIE       As Object
    Private gDoc      As Object
    Private gElement  As Object
    Private gChildren As Object
#End If

Public Enum SpecMethodEnum
    sm_CSS = 0
    sm_AtCss = 1
    sm_ID = 2
End Enum

Public Enum MyErrEnum
    myErr_Timeout = 11111
End Enum

Private TimeOut_ms_ As Long


''' ----- Util functions

''' https://docs.microsoft.com/ja-jp/previous-versions/office/developer/office-2007/aa730921(v=office.12)

Private Function MilliSecondsTimer() As Double
    MilliSecondsTimer = 0
    Dim ticks As Double:     QueryPerformanceCounter ticks
    Dim frequency As Double: QueryPerformanceFrequency frequency
    If frequency Then MilliSecondsTimer = (ticks / frequency) * 1000
End Function


''' ----- Getter, Setter

Public Property Let TimeOut_milliSeconds(ByVal ms As Long)
    TimeOut_ms_ = ms
End Property


''' ----- Event listener

Private Sub Class_Initialize()

    TimeOut_ms_ = 30000  ''' 30s

    'https://msdn.microsoft.com/en-us/library/hh995094%28v=vs.85%29.aspx
    Dim unk As IUnknown 'IInternetExplorerManager
    Dim hr As Long
    Dim IID_IInternetExplorerManager As GUID
    hr = IIDFromString(StrPtr("{ACC84351-04FF-44F9-B23F-655ED168C6D5}"), VarPtr(IID_IInternetExplorerManager))
    'Debug.Print Hex$(hr)

    'CLSID_InternetExplorerManager
    hr = CoGetObject(StrPtr("new:DF4FCC34-067A-4E0A-8352-4A1A5095346E"), 0, IID_IInternetExplorerManager, unk)
    'Debug.Print Hex$(hr), unk Is Nothing

    Dim iid As GUID 'IID_IWebBrowser2
    hr = IIDFromString(StrPtr("{D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}"), VarPtr(iid))
    'Debug.Print Hex$(hr)

    Set gIE = New InternetExplorer
    Set gDoc = New MSHTML.HTMLDocument

    Dim Vnt(0 To 3) As Variant
    Vnt(0) = 1&
    Vnt(1) = StrPtr(vbNullString)
    Vnt(2) = VarPtr(iid)
    Vnt(3) = VarPtr(gIE)

    Dim pArgs(0 To 3) As LongPtr
    Dim i As Long
    Dim vt(0 To 3) As Integer
    For i = 0 To 3
        pArgs(i) = VarPtr(Vnt(i))
        vt(i) = VarType(Vnt(i))
    Next

    Dim VTBLIndex As Long
    VTBLIndex = 3
#If Win64 Then
    VTBLIndex = VTBLIndex * 8
#Else
    VTBLIndex = VTBLIndex * 4
#End If

    Dim vntResult As Variant
    hr = DispCallFunc(ObjPtr(unk), VTBLIndex, _
                        CC_STDCALL, vbLong, _
                        4, vt(0), pArgs(0), vntResult)

End Sub

Private Sub Class_Terminate()
    TerminateImpl
End Sub

Private Sub TerminateImpl()
    On Error GoTo Ending
        If Not gIE.Visible Then gIE.Quit
        Set gIE = Nothing
        Set gDoc = Nothing
Ending:
End Sub

''' DocumentComplete event
''' https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa768282(v%3Dvs.85)
#If HAS_REF Then
    Private Sub gIE_DocumentComplete(ByVal pDisp As Object, url As Variant)
        Sleep 3: Set gDoc = gIE.document
    End Sub
#End If

''' Reguler instance
Private Sub Class_Initialize2()

    TimeOut_ms_ = 30000  ''' 30s

    #If HAS_REF Then
        Set gIE = New InternetExplorer
        Set gDoc = New MSHTML.HTMLDocument
    #Else
        Set gIE = CreateObject("InternetExplorer.Application")
        Set gDoc = CreateObject("HtmlFile")
    #End If

End Sub


''' ----- Basic methods

Public Sub Quit()
    gIE.Quit
End Sub

Public Sub GotoPage(ByVal url As String, Optional ByVal visibleBrowser As Boolean = True)
    gIE.Visible = visibleBrowser
    gIE.navigate url
    LoadBrowser
    DoEvents: Sleep 3: Set gDoc = gIE.document
End Sub

Public Sub LoadBrowser()

    Set gDoc = Nothing

    Dim start As Double

    start = 0: start = MilliSecondsTimer()
    Do While gIE.ReadyState <> 4 Or gIE.Busy
        DoEvents
        Sleep 3
        If MilliSecondsTimer() - start > TimeOut_ms_ Then GoTo Timeout
    Loop

    start = 0: start = MilliSecondsTimer()
    Do While gIE.document.ReadyState <> "complete"
        DoEvents
        Sleep 3
        If MilliSecondsTimer() - start > TimeOut_ms_ Then GoTo Timeout
    Loop

    Sleep 3: Set gDoc = gIE.document

    GoTo Ending

Timeout:
    Err.Raise MyErrEnum.myErr_Timeout
Ending:
End Sub

''' https://developer.mozilla.org/en-US/docs/Web/API/EventTarget/dispatchEvent
''' https://developer.mozilla.org/en-US/docs/Web/Events#Standard_events
Public Function Fire(Optional ByVal aEventName As String = "change") As Boolean

    ''' Create event object
    Dim objEvent As Object: Set objEvent = gDoc.createEvent("HTMLEvents")
    objEvent.initEvent eventType:=aEventName, canBubble:=True, cancelable:=False

    ''' Event is fired for gElement
    Fire = gElement.dispatchEvent(objEvent)

End Function

Public Function IEObj() As Object
    Set IEObj = gIE
End Function

Public Function DocObj() As Object
    Set DocObj = gDoc
End Function

''' Select an item from an elements collection
Public Function Index(ByVal i As Long) As ScrapingEx
    Set gElement = gChildren.Item(i)
    Set Index = Me
End Function

''' Return to an gElement As MSHTML.IHTMLElement
Public Function Element() As Object
    Set Element = gElement
End Function

''' Return to an object to work with tables
''' @return MSHTML.HTMLTableRow
#If HAS_REF Then
Public Function RowTable() As MSHTML.HTMLTableRow
#Else
Public Function RowTable() As Object
#End If
    Set RowTable = gElement
End Function

''' Return the number of items in the collection
Public Function Count() As Integer
    Count = gChildren.Length
End Function

''' Return the text of the selected gElement
Public Function Text() As String
    Text = gElement.innerText
End Function

''' Return the value of the selected gElement
Public Function Value() As String
    Value = gElement.Value
End Function

''' Return the html of the selected gElement
Public Function Html() As String
    Html = gElement.innerHtml
End Function

''' Return the value of the specified attribute
Public Function Attr(ByVal atributeName As String) As String
    Attr = gElement.getAttribute(atributeName)
End Function

''' Assigns a value to a text or select gElement
Public Function FieldValue(ByVal txt As String) As ScrapingEx
    gElement.Value = txt
    Set FieldValue = Me
End Function


''' ----- options

''' Set an gElement to Selected True
Public Sub Selected()
    gElement.Selected = True
End Sub

''' Set an gElement to Checked Checked
Public Sub RadioChecked()
    gElement.Checked = "Checked"
End Sub

''' Set an gElement to Checked EmptyStrings
Public Sub RadioUnChecked()
    gElement.Checked = ""
End Sub

Public Sub SelectOptionDirect(ByVal selector As String)
    Me.At_CSS(selector).Selected
End Sub

Public Function SelectOptionTextValueIs(ByVal aSelector As String, ByVal txt As String, Optional aTimeOut_ms As Long = 3000, Optional ByVal optionAttr As String = "Value") As Boolean

    On Error GoTo Ending

        Until_ElementLocated sm_CSS, aSelector

        Dim i As Long
        Dim start As Double: start = MilliSecondsTimer()

        Do

            On Error Resume Next
                Me.CSS aSelector

            For i = 0 To (Me.Count - 1)
                If Me.Index(i).Attr(optionAttr) = txt Then
                    SelectOptionTextValueIs = True
                    Me.Index(i).Selected
                    GoTo Ending
                End If
            Next i
            If MilliSecondsTimer() - start > aTimeOut_ms Then GoTo Ending

            DoEvents: Sleep 3

        Loop While Not SelectOptionTextValueIs

Ending:
    If Err.Number <> 0 Then Err.clear
End Function

Public Function SelectOptionTextValueContains(ByVal aSelector As String, ByVal txt As String, Optional aTimeOut_ms As Long = 3000, Optional ByVal optionAttr As String = "Value") As Boolean

    On Error GoTo Ending

        Until_ElementLocated sm_CSS, aSelector

        Dim i As Long
        Dim start As Double: start = MilliSecondsTimer()

        Do

            On Error Resume Next
                Me.CSS aSelector

            For i = 0 To (Me.Count - 1)
                If InStr(Me.Index(i).Attr(optionAttr), txt) > 0 Then
                    SelectOptionTextValueContains = True
                    Me.Index(i).Selected
                    GoTo Ending
                End If
            Next i
            If MilliSecondsTimer() - start > aTimeOut_ms Then GoTo Ending

            DoEvents: Sleep 3

        Loop While Not SelectOptionTextValueContains

Ending:
    If Err.Number <> 0 Then Err.clear
End Function

Public Function SelectOptionTextValueMatches(ByVal aSelector As String, ByVal strPtrn As String, Optional aTimeOut_ms As Long = 3000, Optional ByVal optionAttr As String = "Value") As Boolean

    On Error GoTo Ending

        Until_ElementLocated sm_CSS, aSelector

        Dim regx As Object: Set regx = CreateObject("VBScript.RegExp")
        regx.Pattern = strPtrn: regx.ignorecase = True: regx.Global = True
        Dim i As Long
        Dim start As Double: start = MilliSecondsTimer()

        Do

            On Error Resume Next
                Me.CSS aSelector

            For i = 0 To (Me.Count - 1)
                If regx.test(Me.Index(i).Attr(optionAttr)) Then
                    SelectOptionTextValueMatches = True
                    Me.Index(i).Selected
                    GoTo Ending
                End If
            Next i
            If MilliSecondsTimer() - start > aTimeOut_ms Then GoTo Ending

            DoEvents: Sleep 3

        Loop While Not SelectOptionTextValueMatches

Ending:
    If Err.Number <> 0 Then Err.clear
    Set regx = Nothing
End Function


''' ----- a tags

Public Function ClickATagTextIs(ByVal aSelector As String, ByVal txt As String, Optional aTimeOut_ms As Long = 3000) As Boolean

    On Error GoTo Ending

        Until_ElementLocated sm_CSS, aSelector

        Dim i As Long
        Dim start As Double: start = MilliSecondsTimer()

        Do

            On Error Resume Next
                Me.CSS aSelector

            For i = 0 To (Me.Count - 1)
                If Me.Index(i).Text = txt Then
                    ClickATagTextIs = True
                    Me.Index(i).Click
                    GoTo Ending
                End If
            Next i
            If MilliSecondsTimer() - start > aTimeOut_ms Then GoTo Ending

            DoEvents: Sleep 3

        Loop While Not ClickATagTextIs

Ending:
    If Err.Number <> 0 Then Err.clear
End Function

Public Function ClickATagTextContains(ByVal aSelector As String, ByVal txt As String, Optional aTimeOut_ms As Long = 3000) As Boolean

    On Error GoTo Ending

        Until_ElementLocated sm_CSS, aSelector

        Dim i As Long
        Dim start As Double: start = MilliSecondsTimer()

        Do

            On Error Resume Next
                Me.CSS aSelector

            For i = 0 To (Me.Count - 1)
                If InStr(Me.Index(i).Text, txt) > 0 Then
                    ClickATagTextContains = True
                    Me.Index(i).Click
                    GoTo Ending
                End If
            Next i
            If MilliSecondsTimer() - start > aTimeOut_ms Then GoTo Ending

            DoEvents: Sleep 3

        Loop While Not ClickATagTextContains

Ending:
    If Err.Number <> 0 Then Err.clear
End Function

Public Function ClickATagTextMatches(ByVal aSelector As String, ByVal strPtrn As String, Optional aTimeOut_ms As Long = 3000) As Boolean

    On Error GoTo Ending

        Until_ElementLocated sm_CSS, aSelector

        Dim regx As Object: Set regx = CreateObject("VBScript.RegExp")
        regx.Pattern = strPtrn: regx.ignorecase = True: regx.Global = True
        Dim i As Long
        Dim start As Double: start = MilliSecondsTimer()

        Do

            On Error Resume Next
                Me.CSS aSelector

            For i = 0 To (Me.Count - 1)
                If regx.test(Me.Index(i).Text) Then
                    ClickATagTextMatches = True
                    Me.Index(i).Click
                    GoTo Ending
                End If
            Next i
            If MilliSecondsTimer() - start > aTimeOut_ms Then GoTo Ending

            DoEvents: Sleep 3

        Loop While Not ClickATagTextMatches

Ending:
    If Err.Number <> 0 Then Err.clear
    Set regx = Nothing
End Function


''' ----- pick up element(s)

' Example selector  id: "#nameId" | class: ".nameClas" | tag: "nameTag"
Public Function CSS(ByVal selector As String) As ScrapingEx
    Set gChildren = gDoc.querySelectorAll(selector)
    Set CSS = Me
End Function

Public Function At_CSS(ByVal selector As String) As ScrapingEx
    Set gElement = gDoc.querySelector(selector)
    Set At_CSS = Me
End Function

Public Function ID(ByVal idName As String) As ScrapingEx
    Set gElement = gDoc.getElementById(idName)
    Set ID = Me
End Function

Public Function Until_TextContains(ByVal aSelector As String, ByVal txt As String, Optional aTimeOut_ms As Long = 3000) As Boolean

    On Error GoTo Ending

        Until_ElementLocated sm_CSS, aSelector

        Dim i As Long
        Dim start As Double: start = MilliSecondsTimer()

        Do

            On Error Resume Next
                Me.CSS aSelector

            For i = 0 To (Me.Count - 1)
                If InStr(Me.Index(i).Text, txt) > 0 Then
                    Until_TextContains = True
                    GoTo Ending
                End If
            Next i
            If MilliSecondsTimer() - start > aTimeOut_ms Then GoTo Ending

            DoEvents: Sleep 3

        Loop While Not Until_TextContains

Ending:
    If Err.Number <> 0 Then Err.clear
End Function

Public Function Until_TextIs(ByVal aSelector As String, ByVal txt As String, Optional aTimeOut_ms As Long = 3000) As Boolean

    On Error GoTo Ending

        Until_ElementLocated sm_CSS, aSelector

        Dim i As Long
        Dim start As Double: start = MilliSecondsTimer()

        Do
            On Error Resume Next
                Me.CSS aSelector

            For i = 0 To (Me.Count - 1)
                If Me.Index(i).Text = txt Then
                    Until_TextIs = True
                    GoTo Ending
                End If
            Next i
            If MilliSecondsTimer() - start > aTimeOut_ms Then GoTo Ending

            DoEvents: Sleep 3

        Loop While Not Until_TextIs

Ending:
    If Err.Number <> 0 Then Err.clear
End Function

Public Function Until_TextMatches(ByVal aSelector As String, ByVal strPtrn As String, Optional aTimeOut_ms As Long = 3000) As Boolean

    On Error GoTo Ending

        Until_ElementLocated sm_CSS, aSelector

        Dim regx As Object: Set regx = CreateObject("VBScript.RegExp")
        regx.Pattern = strPtrn: regx.ignorecase = True: regx.Global = True
        Dim i As Long
        Dim start As Double: start = MilliSecondsTimer()

        Do
            On Error Resume Next
                Me.CSS aSelector

            For i = 0 To (Me.Count - 1)
                If regx.test(Me.Index(i).Text) Then
                    Until_TextMatches = True
                    GoTo Ending
                End If
            Next i
            If MilliSecondsTimer() - start > aTimeOut_ms Then GoTo Ending

            DoEvents: Sleep 3

        Loop While Not Until_TextMatches

Ending:
    If Err.Number <> 0 Then Err.clear
    Set regx = Nothing
End Function

Public Function Until_ElementLocated(ByVal aSpecMethod As SpecMethodEnum, ByVal aSelector As String, Optional aTimeOut_ms As Long = 3000) As ScrapingEx

    Dim start As Double: start = MilliSecondsTimer()

    On Error Resume Next
        Do
            Err.clear

            Select Case aSpecMethod
                Case SpecMethodEnum.sm_CSS:   Me.CSS aSelector
                Case SpecMethodEnum.sm_AtCss: Me.At_CSS aSelector
                Case SpecMethodEnum.sm_ID:    Me.ID aSelector
            End Select

            If Err.Number = 0 Then GoTo Ending
            If MilliSecondsTimer() - start > aTimeOut_ms Then GoTo Timeout

            DoEvents: Sleep 3

        Loop While Err.Number <> 0
    On Error GoTo 0

    GoTo Ending

Timeout:
    Err.Raise MyErrEnum.myErr_Timeout
Ending:
    Set Until_ElementLocated = Me
End Function


''' ----- click

Public Sub Click()
    gElement.Click
    LoadBrowser
    DoEvents: Sleep 3: Set gDoc = gIE.document
End Sub

Public Sub ClickTableCell(ByVal tblCell As Object)
    tblCell.children.Item(0).Click
    LoadBrowser
    DoEvents: Sleep 3: Set gDoc = gIE.document
End Sub
