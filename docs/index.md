# はじめに

Office Open XML ファイルフォーマットのうち、SpreadsheetMLについて部分的に抜粋、翻訳したものである。
翻訳にあたり省略や言い換えが多くあるので、原文を参照すること。
主に ECMA-376-1:2016 より引用する。

* [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/)

## Overview

> This clause is informative.
> This clause contains an overview of Office Open XML.

この節には、Office Open XML の概要が含まれている。

### 8.1 Content Overview

> This standard contains predominantly the following three types of information:

この規格には主に次の 3種類の情報が含まれている。
本件ではこれらのうち一部を抜粋する。

1. > Normative W3C XML Schemas, informative RELAX NG schemas and an associated validation procedure
   > for validating document syntax against those schemas (Annex A and Annex B)

   標準的な W3C XML スキーマ、RELAX NG スキーマおよび関連する検証手順

2. > Descriptions of XML element semantics. The semantics of an XML element refers to its intended
   > interpretation by a human being (chiefly in §11, §12, §13, and §14)

   XML 要素の意味の詳細

3. > Additional syntax constraints in written form

   追加構文

### 8.2 Packages and Parts

> An Office Open XML document is represented as a series of related parts that are stored in a container called a package.
> Information about the relationships between a package and its parts is stored in the package's package-relationship ZIP item.
> Information about the relationships between two parts is stored in the part-relationship ZIP item for the source part.
> A package is an ordinary ZIP archive, which contains that package's content-type item, relationship items, and parts.
> (Packages are discussed further in ECMA-376-2.)

Office Open XML ドキュメントは、パッケージと呼ばれるコンテナに格納されるパーツとして表す。
パッケージとそのパーツ間の関係に関する情報は、パッケージリレーションシップというファイルに保存される。
パーツ同士の関係に関する情報は、パーツリレーションシップというファイルに保存される。
パッケージは通常の ZIP アーカイブであり、そのパッケージのコンテンツ タイプ、リレーションシップ、およびパーツが含まれる。
(パッケージについては ECMA-376-2 で詳しく説明されている)

> A WordprocessingML document contains a part for the body of the text;
> it might also contain a part for an image referenced by that text, and parts defining document characteristics, styles, and fonts.

WordprocessingML ドキュメントには、テキストの本文が含まれる。
また、参照される画像や、文書の特性、スタイル、フォントを定義も含まれる。
本件では解説しない。

> A SpreadsheetML document contains a separate part for each worksheet;
> it might also contain parts for images.

SpreadsheetML ドキュメントにはワークシートごとに格納される。
画像も含まれる。

> A PresentationML document contains a separate part for each slide.

PresentationML ドキュメントにはスライドごとに格納される。
本件では解説しない。

### 8.5 SpreadsheetML

> This subclause introduces the overall form of a SpreadsheetML package, and identifies some of its main components.
> (See Annex L for a more detailed introduction.)

この節では SpreadsheetML パッケージの全体像を紹介し、その主要コンポーネントのいくつかを紹介する。
(より詳細な概要については、付録 L を参照)

> A SpreadsheetML package has a relationship of type officeDocument, which specifies the location of the main part in the package.
> For a SpreadsheetML document, that part contains the workbook definition.

SpreadsheetML パッケージには、officeDocument タイプと関係があり、パッケージ内の主要部分の場所を指す。
SpreadsheetML ドキュメントの場合、ワークブック定義が含まれる。

> A SpreadsheetML package’s main part starts with a spreadsheet root element.
> That element is a workbook, which refers to one or more worksheets, which, in turn, contain the data.
> A worksheet is a two-dimensional grid of cells that are organized into rows and columns.

SpreadsheetML パッケージの主要部分は、スプレッドシートのルート要素で始まる。
その要素はワークブックであり、1 つ以上のワークシートを参照し、ワークシートにはデータが含まれる。
ワークシートは行と列によって編成されたセルの 2 次元グリッドである。

> The cell is the primary place in which data is stored and operated on.
> A cell can have a number of characteristics, such as numeric, text, date, or time formatting; alignment; font; color; and a border.

セルはデータが保存・操作される。
セルには数値、テキスト、日付、時刻の書式設定などを含めることができる。
アライメント、フォント、色、そして罫線などのフォーマット情報も含まれる。

> Each cell is identified by a cell reference, a combination of its column and row headings.
> Each horizontal set of cells in a worksheet is called a row, and each row has a heading numbered sequentially, starting at 1.
> Each vertical set of cells in a worksheet is called a column, and each column has an alphabetic heading named sequentially from A–Z, then AA–AZ, BA–BZ, and so on.

セルは行と列によって識別される。
ワークシート内の水平方向のセルは行と呼ばれ、各行には 1 から始まる連番が付けられる。
ワークシート内の縦方向のセルは列と呼ばれ、各列には A ～ Z、AA ～ AZ、BA ～ BZ などの順に名前が付けられる。

> Instead of data, a cell can contain a formula, which is a recipe for calculating a value.
> Some formulas—called functions—are predefined, while others are user-defined.
> Examples of predefined formulas are AVERAGE, MAX, MIN, and SUM.
> A function takes one or more arguments on which it operates, producing a result.
> For example, in the formula SUM(B1:B4), there is one argument, B1:B4, which is the range of cells B1–B4, inclusive.

セルには数式を含めることができる。
関数と呼ばれるいくつかの式は事前に定義されているが、その他はユーザー定義である。
関数は 1 つ以上の引数を受け取り、戻り値を返す。
たとえば、数式 SUM(B1:B4) には、引数 B1:B4 があり、これはセル B1 ～ B4 の範囲である。

> Other features that a SpreadsheetML document can contain include the following: comments, hyperlinks, images, and sorted and filtered tables.

SpreadsheetML ドキュメントに含めることができるその他の機能には、コメント、ハイパーリンク、画像、並べ替えおよびフィルター処理されたテーブルなどがある。

> A SpreadsheetML document is not stored as one large body in a single part;
> instead, the elements that implement certain groupings of functionality are stored in separate parts.
> For example, all the data for a worksheet is stored in that worksheet's part, all string literals from all worksheets are stored in a single shared string part, and each worksheet having comments has its own comments part.

SpreadsheetML ドキュメントは、1 つのファイルとして保存されるわけではない。
機能毎に個別のファイルとして格納される。
たとえば、ワークシートのデータにおいても、文字列リテラルは単一の共有文字列パーツに保存され、コメントはコメントパーツに保存される。
