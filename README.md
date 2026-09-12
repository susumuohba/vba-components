# vba-components
A collection of VBA components and snippets for numerical computation and business automation.

数値計算やビジネス自動化に役立つ、再利用可能なVBAのコンポーネント（クラスモジュール・標準モジュール）集です。

## フォルダ構成

- math/Fraction.cls (分数クラス・自動約分・四則演算対応)
- math/TestFraction.bas (分数クラスのテスト用モジュール)
- business/ (今後追加予定)

## 使い方（Excel VBEへのインポート）

1. このリポジトリから必要な .cls や .bas ファイルをダウンロードします。
2. Excelを開き、Alt + F11 キーを押して VBE（Visual Basic Editor）を起動します。
3. メニューの [ファイル] ＞ [ファイルのインポート] から、取り込んでご利用ください。

## サンプルコード

Sub TestFraction()
    Dim f1 As New Fraction
    Dim f2 As New Fraction
    Dim ans As New Fraction
    
    f1.Init 2, 6
    f2.Init 1, 2
    
    Set ans = f1.Add(f2)
    
    Debug.Print f1.ToString() & " + " & f2.ToString() & " = " & ans.ToString()
End Sub

## ライセンス

このリポジトリのコードは MIT License の下で公開されています。自由にご利用・改変いただけます。
