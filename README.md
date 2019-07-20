[![MIT-LICENSE](http://img.shields.io/badge/license-MIT-blue.svg?style=flat)]( https://github.com/callmekohei/ScrapingEx/blob/master/LICENSE)


# ScrapingEx

`VBA`と`IE(Internet Explore)`でホームページからデータを取得する便利ライブラリーです

## License
MIT license

## Respect
[Victor Zevallos](https://github.com/vba-dev/vba-Scraping)  
[kumatti1](https://gist.github.com/kumatti1/6b68ea65fdfc9ecf727f)


## 外部ライブラリ

VBEのツール > 参照設定から下記のライブラリ（標準）を選択してください
```
Microsoft Internet Control
Microsoft HTML ObjectLibrary
```

## サンプルコード

### ヤホーで検索する
```vb
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
```

### グーグルで検索する
```vb
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
```

### ロト６の最新の結果を取得する
```vb
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
```
結果
```text
抽せん日 2019年6月27日
本数字 06 09 12 19 40 42
ボーナス数字 (07)
1等 該当なし 該当なし
2等 10口 7,036,900円
3等 213口 356,700円
4等 12,206口 6,500円
5等 191,383口 1,000円
販売実績額 1,450,251,200円
キャリーオーバー 234,558,536円
```

## その他

何かあれば `issue` になんでも書き込んでくださいっ  
もしくはツイッター @callmekohei でも結構です。。。
