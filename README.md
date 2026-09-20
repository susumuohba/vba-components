# vba-components
A collection of VBA components and snippets for numerical computation and business automation.

数値計算やビジネス自動化に役立つ、再利用可能なVBAのコンポーネント（クラスモジュール・標準モジュール）集です。

## フォルダ構成

- `math/Fraction.cls` （分数クラス・自動約分・四則演算対応）
- `math/TestFraction.bas` （分数クラスのテスト用モジュール）
- `testing/modAssert.bas` （自作の軽量単体テストフレームワーク。アドイン不要でモジュールをコピーするだけで動作。`AssertEqual` / `AssertTrue` / `AssertFalse` 等を提供し、結果はイミディエイトウィンドウに出力）
- `business/aggregate-workbooks/` （フォルダ内の複数Excelブックを1冊に自動集計するツール一式。詳細は下記「Business Tools」参照）
- `business/`配下は今後追加予定

## 使い方（Excel VBEへのインポート）

1. このリポジトリから必要な `.cls` や `.bas` ファイルをダウンロードします。
2. Excelを開き、Alt + F11 キーを押して VBE（Visual Basic Editor）を起動します。
3. メニューの [ファイル] ＞ [ファイルのインポート] から、取り込んでご利用ください。

## サンプルコード

```vb
Sub TestFraction()
    Dim f1 As New Fraction
    Dim f2 As New Fraction
    Dim ans As New Fraction
    
    f1.Init 2, 6
    f2.Init 1, 2
    
    Set ans = f1.Add(f2)
    
    Debug.Print f1.ToString() & " + " & f2.ToString() & " = " & ans.ToString()
End Sub
```

## Business Tools

### aggregate-workbooks

フォルダ内の同一フォーマットの複数Excelブックを、1冊の集計ブックへ自動集計するツールです。

- 単体テストコード付き（`testing/modAssert.bas`を使用、8件の自動テストを収録）
- 空行を含むデータのスキップ、対象ブック自身の自動除外に対応
- 仕様書・操作マニュアル・コード規約・引き継ぎメモ・テストケース一覧をまとめたサンプルドキュメントも同梱（`business/aggregate-workbooks/docs/複数ブック集計_5点セット雛形.docx`）

使い方: `modMain.bas` `modFileUtil.bas` `modAggregate.bas` `modAggregateTests.bas` の4ファイルを同一ブック（.xlsm）にインポートし、シート上のボタンに `modMain.RunAggregate` を登録して実行します。詳細な導入手順は同梱のドキュメントを参照してください。

## ライセンス

このリポジトリのコードは MIT License の下で公開されています。自由にご利用・改変いただけます。
