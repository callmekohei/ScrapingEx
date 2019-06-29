VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScrapingEx_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' --------------------------------------------------------
'''  FILE    : ScrapingEX.cls
'''  AUTHOR  : callmekohei <callmekohei at gmail.com>
'''  License : MIT license
''' --------------------------------------------------------
Option Explicit

'Assert.RunTestOf New ScrapingEx_Test

Public Sub InputText_Test()
    Dim doc As ScrapingEx: Set doc = New ScrapingEx
    doc.GotoPage "http://www.htmq.com/html/input_text.shtml"
    Dim selector As String: selector = "#content_left > div.siyorei_browse > form > p:nth-child(1) > input[type=text]"
    doc.At_CSS(selector).FieldValue "foo bar baz"

    Assert.AreEqual "foo bar baz", doc.At_CSS(selector).Value

    doc.Quit
    Set doc = Nothing
End Sub

Public Sub InputPassword_Test()
    Dim doc As ScrapingEx: Set doc = New ScrapingEx
    doc.GotoPage "http://www.htmq.com/html/input_password.shtml"
    Dim selector As String: selector = "#content_left > div.siyorei_browse > form > p:nth-child(1) > input[type=password]:nth-child(1)"
    doc.At_CSS(selector).FieldValue "123456789"

    Assert.AreEqual "123456789", doc.At_CSS(selector).Value

    doc.Quit
    Set doc = Nothing
End Sub

Public Sub InputForm_Test()
    Dim doc As ScrapingEx: Set doc = New ScrapingEx
    doc.GotoPage "http://www.htmq.com/html/form.shtml"
    Dim selector As String: selector = "textarea"
    doc.At_CSS(selector).FieldValue "foo" & vbNewLine & "bar" & vbNewLine & "baz"

    Assert.AreEqual "foobarbaz", Core.ReReplace(doc.At_CSS(selector).Value, "\n|\r\n|\r", "", "g")

    doc.Quit
    Set doc = Nothing
End Sub

Public Sub CheckBox_Test()
    Dim doc As ScrapingEx: Set doc = New ScrapingEx
    doc.GotoPage "http://www.htmq.com/html/input_checkbox.shtml"
    Dim Checkbox As ScrapingEx: Set Checkbox = doc.CSS("form > p > input")
    Checkbox.Index(0).RadioUnChecked
    Checkbox.Index(1).RadioChecked
    Checkbox.Index(2).RadioUnChecked

    Assert.AreEqual False, Checkbox.Index(0).Element.Checked
    Assert.AreEqual True, Checkbox.Index(1).Element.Checked
    Assert.AreEqual False, Checkbox.Index(2).Element.Checked

    doc.Quit
    Set doc = Nothing
End Sub

Public Sub RadioBox_Test()
    Dim doc As ScrapingEx: Set doc = New ScrapingEx
    doc.GotoPage "https://www.tagindex.com/html_tag/form/input_radio.html"
    Dim selector As String: selector = " #content > section.example > div > figure > dl > dd > p:nth-child(1) > input[type=radio]:nth-child(3)"
    doc.At_CSS(selector).RadioChecked

    Assert.AreEqual True, doc.At_CSS(selector).Element.Checked

    doc.Quit
    Set doc = Nothing
End Sub

Public Sub SelectBox_Test()
    Dim doc As ScrapingEx: Set doc = New ScrapingEx
    doc.GotoPage "http://www.htmq.com/html/option.shtml"
    Dim selector As String: selector = "select > option"
    doc.SelectOptionTextValueIs selector, "Virgo"

    Assert.AreEqual True, doc.At_CSS("#content_left > div.siyorei_browse > form > p > select > option:nth-child(6)").Element.Selected

    doc.Quit
    Set doc = Nothing
End Sub

Public Sub A_Tag_Test()
    Dim doc As ScrapingEx: Set doc = New ScrapingEx
    doc.GotoPage "http://www.htmq.com/html/a.shtml"
    doc.ClickATagTextMatches "a", "ひとつ上の.*トリへ"
    doc.ClickATagTextContains "a", "シートリファレンス(目的"

    Assert.AreEqual "スタイルシートリファレンス（目的別）", doc.IEObj.LocationName

    doc.Quit
    Set doc = Nothing
End Sub

Public Sub ButtonYahoo_Test()
    Dim doc As ScrapingEx: Set doc = New ScrapingEx
    doc.GotoPage "https://www.yahoo.co.jp/"
    doc.ID("srchtxt").FieldValue "VBA"
    doc.ID("srchbtn").Click
    doc.LoadBrowser

    Assert.AreEqual "「VBA」の検索結果 - Yahoo!検索", doc.IEObj.LocationName

    doc.Quit
    Set doc = Nothing
End Sub

Public Sub ButtonGoogle_Test()
    Dim doc As ScrapingEx: Set doc = New ScrapingEx
    doc.GotoPage "https://www.google.com/"
    doc.At_CSS("#tsf > div:nth-child(2) > div > div.RNNXgb > div > div.a4bIc > input").FieldValue "VBA"
    doc.At_CSS("#tsf > div:nth-child(2) > div > div.FPdoLc.VlcLAe > center > input.gNO89b").Click

    Assert.AreEqual "VBA - Google 検索", doc.IEObj.LocationName

    doc.Quit
    Set doc = Nothing
End Sub

Public Sub Table_Test()
    Dim doc As ScrapingEx: Set doc = New ScrapingEx
    doc.GotoPage "http://www.htmq.com/html/table.shtml"

    Dim selector As String: selector = "#content_left > div.siyorei_browse > table > tbody > tr:nth-child(3) > td:nth-child(3)"
    Assert.AreEqual True, doc.Until_TextMatches(selector, "^ペスカトーレ")

    Dim tableArr As Variant: tableArr = ArrTable(doc.CSS("#content_left > div.siyorei_browse > table").Index(0).RowTable, True)
    Assert.AreEqualArr Array("メニュー", "説明", "豆知識"), tableArr(0)(0)
    Assert.AreEqualArr Array("カルボナーラ", "玉子とベーコンとクリームソースのパスタ", "カルボナーラとは炭焼き職人という意味"), tableArr(0)(1)
    Assert.AreEqualArr Array("ペスカトーレ", "エビとあさりの漁師風パスタ", "ペスカトーレは漁師風という意味"), tableArr(0)(2)

    doc.Quit
    Set doc = Nothing
End Sub

Public Sub Loto6Table_Test()
    Dim doc As ScrapingEx: Set doc = New ScrapingEx
    doc.GotoPage "https://www.mizuhobank.co.jp/retail/takarakuji/loto/loto6/index.html"

    Dim selector As String: selector = "#mainCol > article > section > section > section > div > div.sp-none > table:nth-child(1) > tbody > tr:nth-child(10) > td > strong"
    Assert.AreEqual True, doc.Until_TextMatches(selector, "[^ \t\n\r\f]")

    Dim tableArr As Variant: tableArr = ArrTable(doc.CSS("table.typeTK").Index(0).RowTable, True)(1)
    Assert.AreEqual "キャリーオーバー", tableArr(9)(0)
    Assert.AreEqual True, InStr(tableArr(9)(1), "円") > 0

    doc.Quit
    Set doc = Nothing
End Sub

Public Sub YahooCalendar_Test()
    Dim doc As ScrapingEx: Set doc = New ScrapingEx
    doc.GotoPage "https://www.yahoo.co.jp/"
    Dim tableArr As Variant: tableArr = ArrTable(doc.CSS("#pbcNx > table").Index(0).RowTable, True)(1)

    Assert.AreEqualArr Array("日", "月", "火", "水", "木", "金", "土"), tableArr(0)

    doc.Quit
    Set doc = Nothing
End Sub

Public Sub ScrapingRepeated_Test()

    ''' 8 s/term ( callmekohei's pc )
    ''' 100 ---> 13min, 1000 ---> 2.1h, 3000 ---> 6.5h

    Dim doc As ScrapingEx: Set doc = New ScrapingEx
    Dim selector As String
    Dim tables As ScrapingEx
    Dim tableArr As Variant

    Dim v As Variant
    For Each v In ArrRange(1, 3)

        '' HTML Quick Reference
        doc.GotoPage "http://www.htmq.com/html/table.shtml"
        selector = "#content_left > div.siyorei_browse > table > tbody > tr:nth-child(3) > td:nth-child(3)"
        Assert.AreEqual True, doc.Until_TextMatches(selector, "^ペスカトーレ")

        tableArr = ArrTable(doc.CSS("#content_left > div.siyorei_browse > table").Index(0).RowTable, True)
        Assert.AreEqualArr Array("メニュー", "説明", "豆知識"), tableArr(0)(0)
        Assert.AreEqualArr Array("カルボナーラ", "玉子とベーコンとクリームソースのパスタ", "カルボナーラとは炭焼き職人という意味"), tableArr(0)(1)
        Assert.AreEqualArr Array("ペスカトーレ", "エビとあさりの漁師風パスタ", "ペスカトーレは漁師風という意味"), tableArr(0)(2)

        ''' Yahoo
        doc.GotoPage "https://www.yahoo.co.jp/"
        selector = "#pbcNx > table"
        tableArr = ArrTable(doc.CSS("#pbcNx > table").Index(0).RowTable, True)(1)
        Assert.AreEqualArr Array("日", "月", "火", "水", "木", "金", "土"), tableArr(0)

        ''' Loto6
        doc.GotoPage "https://www.mizuhobank.co.jp/retail/takarakuji/loto/loto6/index.html"
        selector = "table.typeTK"
        selector = "#mainCol > article > section > section > section > div > div.sp-none > table:nth-child(1) > tbody > tr:nth-child(10) > td > strong"
        Assert.AreEqual True, doc.Until_TextMatches(selector, "[^ \t\n\r\f]")
        tableArr = ArrTable(doc.CSS("table.typeTK").Index(0).RowTable, True)(1)
        Assert.AreEqual "キャリーオーバー", tableArr(9)(0)
        Assert.AreEqual True, InStr(tableArr(9)(1), "円") > 0

    Next v

    doc.Quit
    Set doc = Nothing
End Sub

