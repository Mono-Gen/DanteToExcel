# Dante プリセット XML → Excel 変換ツール マニュアル

## ■ 概要
Dante Controller で書き出したプリセットファイル（XML）を Excel ファイルに変換するツールです。
Go言語で実装されており、Microsoft Excelのインストールを必要とせず、一瞬で `.xlsx` ファイルを直接生成します。
デバイス一覧、パッチマトリクス、フロー情報などを自動生成し、AES67の設定情報や詳細なPTP v2関連パラメータの解析にも対応しています。

## ■ 必要な環境
- **Windows**: Windows 10 / 11 (64-bit)
- **macOS**: macOS 10.15 以降 (Intel & Apple Silicon)
- **スプレッドシート閲覧ソフト**: 生成された `.xlsx` ファイルを開くためのソフト（Microsoft Excel、LibreOffice、Google スプレッドシートなど）

## ■ ファイル構成
- `DanteToExcel_windows_x64.exe` : Windows用実行ファイル
- `DanteToExcel_macOS_Intel` : macOS (Intel) 用実行ファイル
- `DanteToExcel_macOS_AppleSilicon` : macOS (Apple Silicon) 用実行ファイル

## ■ 使い方
1. お使いのOSに対応する実行ファイルを、変換したい Dante プリセット XML ファイルがあるフォルダに配置します。
2. 実行ファイルをダブルクリックして起動します（またはターミナル等コマンドラインから起動します）。
3. 同じフォルダに XML ファイルが複数存在する場合は、選択画面が表示されるので、目的のファイルの番号を入力して `Enter` キーを押します。
4. メニューが表示されるので、出力モードを選択します（通常は `1` ＝ Default を選択）。
5. 処理が完了すると、同じフォルダに Excel ファイル（`.xlsx`）が生成され、「Press Enter to exit...」と表示されます。`Enter` キーを押してコンソールを終了してください。

## ■ 出力モード
- **Default (1)**: 重要な項目に絞った概要データ。以下のシートを出力します。
  - `Devices`（基本的なプロパティ一覧）
  - `Patch Matrix`（ルーティングマトリクス）
  - `TX Flows`（送信フロー）
- **Detail (2)**: 以下の詳細情報を含む、すべてのデータを出力します。
  - `Devices` シートへの詳細情報追加（Pull Up値、詳細IP、PTP v2詳細、優先度など）
  - `TX Channels` シート（送信チャンネル一覧）
  - `RX Channels` シート（受信チャンネル一覧）
  - `Subscriptions` シート（接続関係一覧）

## ■ 注意事項
- 同名の `.xlsx` ファイルがすでに存在する場合は上書きされますのでご注意ください。
- 以前のPowerShell版とは異なり、Excelをバックグラウンドで起動しないため、変換中も他の作業を妨げません。また、処理速度が大幅に向上（1秒未満）しています。
