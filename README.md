# VSIX

## 準備

* アプリと機能 - Visual Studio 2017 - Visual Studio 拡張機能の開発をインストールしておく

## 作成

* ファイル - 新規作成 - プロジェクト - 他の言語 - Visual C# - Extensibility - VSIX Project
    * 参照 - 参照の追加 - アセンブリ - 拡張 - MicrosoftVisualStudio.VCCodeModel にチェック
    * 追加 - 新しい項目 - Extensibility - Custom Command
    * MenuItemCallback() を編集する

## 実行

* 実行すると Visual Studio がもう一つ起動するので適当なソリューションを開いて読み込むのを待つ
    * ツール - Invoke Command とすると MenuItemCallback() がコールバックされる