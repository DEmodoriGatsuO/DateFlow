# DateFlow

## 概要

**DateFlow**は、PowerShellを使用してユーザーからのインタラクティブな日付入力を促し、その日付情報を使ってExcelおよびPowerPointファイルの処理を自動化するツールです。このプロジェクトは、SharePointに格納されたファイルの操作や、PowerPointファイル内の特定のテキストをユーザーが選択した日付で置き換える作業を効率化します。

## 特徴

- ユーザーフレンドリーな日付ピッカーインターフェース
- SharePoint上のExcelファイルの自動操作
- PowerPointファイルの複製とテキスト置換
- 複数のユーザー環境での実行を想定した設計

## システム要件

- Windows 10またはそれ以降
- PowerShell 5.0またはそれ以降
- Microsoft Office 2016またはそれ以降
- .NET Framework 4.5以上（Windowsフォームの使用に必要）

## インストール

1. リポジトリをクローンするか、ZIPファイルをダウンロードして展開します。

   ```bash
   git clone https://github.com/yourusername/DateFlow.git
   cd DateFlow
   ```

2. スクリプトファイルを任意のディレクトリに配置します。

## 使用方法

### 1. スクリプトの実行

`DateFlow.ps1`をPowerShellで実行します。

```powershell
.\DateFlow.ps1
```

### 2. 日付の選択

スクリプトを実行すると、日付ピッカーが表示されます。目的の日付を選択して「OK」ボタンを押します。

### 3. ファイル処理

選択された日付は自動的にPowerPointファイルのテキストボックスに反映されます。スクリプトは、以下の処理を行います:

- SharePointに格納されたExcelファイルを開きます。
- ローカルディレクトリに保存されたPowerPointファイルを複製します。
- 複製されたPowerPointファイル内の特定のテキスト（例: 「月」）を選択した日付で置き換えます。

### 4. 出力

PowerPointファイルが正常に更新されたことを確認するメッセージが表示されます。

## 注意事項

- スクリプト実行時には、Officeアプリケーション（Excel, PowerPoint）がインストールされていることを確認してください。
- スクリプトは、ユーザーのプロファイルパスやExcelのインストールパスを自動的に検出しますが、環境依存の問題が発生した場合は適宜修正が必要です。

## トラブルシューティング

### スクリプトが動作しない場合

- PowerShellの実行ポリシーを確認してください。未署名スクリプトがブロックされている場合は、以下のコマンドでポリシーを変更できます。
  
  ```powershell
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

- .NET Frameworkのバージョンが適切であることを確認してください。

## ライセンス

このプロジェクトはMITライセンスのもとで提供されています。詳細は[LICENSE](LICENSE)ファイルを参照してください。