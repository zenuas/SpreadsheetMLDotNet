# SpreadsheetMLDotNet [![MIT License](https://img.shields.io/badge/license-MIT-blue.svg?style=flat)](LICENSE)

SpreadsheetMLDotNet is Office Open XML SpreadsheetML library  

## Description

SpreadsheetMLDotNet is a poor sample program.  

SpreadsheetMLDotNetは貧弱なサンプルプログラムである。  
簡易の読み込みと書き込みをサポートする。  

## Usage

WorkbookReaderはセルの値のみ取得する。
式の評価やスタイル、フォーマットは取得しない。

```cs
using var workbook = SpreadsheetML.CreateWorkbookReader("Test.xlsx");
foreach (var name in workbook.WorkSheetNames)
{
    Console.WriteLine(name);
    using var worksheet = workbook.OpenWorksheet(name);
    foreach (var (cell, value) in worksheet)
    {
        Console.WriteLine($"{cell} = {value}");
    }
}
```

Workbookは一部の編集とインポート、エクスポートを行う。

```cs
var book = SpreadsheetML.Import("Base.xlsx");
var sheet = book.Worksheets[0];
sheet.SetCell("A1", "cell text");
SpreadsheetML.Export("Test.xlsx", book);
```

## Support

次の機能はサポートしていない。上にあるものほど対応優先度は高い。
* オートフィルタ
* 行、列の非表示
* シート非表示
* SUM以外の関数
* 行、列の挿入操作(式のセル参照調整)
* 行、列の削除操作(式のセル参照調整)
* 印刷関連
* 行の高さ計算
* 列の幅計算
* 名前参照
* パスワード
* 塗りつぶしのグラディエーション
* ピボットテーブル
* リンク
* テーマ

次の機能をサポートする計画はない。
* エクスポート時のShared Stringテーブル利用
* Calculation Chain
* 外部スプレッドシート参照
* コメント
* メタデータ
* クエリ
* 図
* グラフ
* オートシェイプ
* フォーム
* アプリケーション定義ファイル
* コアファイル
* それ以外のパート(12.3 Part Summary参照)
* VBA
