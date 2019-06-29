Attribute VB_Name = "Sample"
''' --------------------------------------------------------
'''  FILE    : Sample.bas
'''  AUTHOR  : callmekohei <callmekohei at gmail.com>
'''  License : MIT license
''' --------------------------------------------------------
Option Explicit

''' ヤホーで検索する
Public Sub sample_yahoo()

    ''' スクレイピングＥｘを使えるようにします
    Dim doc As ScrapingEx: Set doc = New ScrapingEx

    ''' ヤホーのホームページを開きます
    doc.GotoPage "https://www.yahoo.co.jp/"

    ''' 検索窓に VBA と入力します
    doc.ID("srchtxt").FieldValue "VBA"

    ''' 検索ボタンを押します
    doc.ID("srchbtn").Click

End Sub

''' グーグルで検索する
Public Sub sample_google()

    ''' スクレイピングＥｘを使えるようにします
    Dim doc As ScrapingEx: Set doc = New ScrapingEx

    ''' グーグルのホームページを開きます
    doc.GotoPage "https://www.google.com/"

    ''' 検索窓に VBA と入力します
    doc.At_CSS("#tsf > div:nth-child(2) > div > div.RNNXgb > div > div.a4bIc > input").FieldValue "VBA"

    ''' 検索ボタンを押します
    doc.At_CSS("#tsf > div:nth-child(2) > div > div.FPdoLc.VlcLAe > center > input.gNO89b").Click

End Sub

''' ロト６の最新の結果を取得する
Public Sub Sample_Loto6()

    ''' スクレイピングＥｘを使えるようにします
    Dim doc As ScrapingEx: Set doc = New ScrapingEx

    ''' ロト６のホームページを開きます
    doc.GotoPage "https://www.mizuhobank.co.jp/retail/takarakuji/loto/loto6/index.html"

    ''' ロト６の最新の結果表のキャリーオーバーの金額のセルが空白でない状態になるまで待ちます
    Dim selector As String: selector = "#mainCol > article > section > section > section > div > div.sp-none > table:nth-child(1) > tbody > tr:nth-child(10) > td > strong"
    doc.Until_TextMatches selector, "[^ \t\n\r\f]"

    ''' ロト６の最新の結果表を配列にします
    Dim tableArr As Variant
    tableArr = ArrTable(doc.CSS("table.typeTK").Index(0).RowTable, True)(1)

    ''' イミディエイトウィンドウにて取得したデータを表示します
    Dim v
    For Each v In tableArr
        Debug.Print Join(v, " ")
    Next v

    ''' ブラウザ（ＩＥ）を片付けます
    doc.Quit
    Set doc = Nothing

End Sub
