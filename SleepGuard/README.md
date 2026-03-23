# SleepGuard

**SleepGuard** は、指定したプロセスが起動している間は Windows のスリープを自動的に抑止し、全プロセスが終了してから設定した猶予時間が経過するとスリープを解除する、軽量なタスクトレイ常駐ツールです。

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-blue.svg)
![.NET](https://img.shields.io/badge/.NET-8.0-purple.svg)

---

## スクリーンショット

> ※ 起動後、タスクトレイに常駐します。ダブルクリックで設定画面を開けます。

---

## 機能

- **プロセス監視によるスリープ防止**  
  登録したプロセスが1つでも起動中であれば、Windows の電源設定によるスリープを自動的にブロックします。

- **猶予時間設定**  
  全監視プロセスが停止してから、設定した時間（1〜60分）が経過した後にスリープを解除します。

- **実行中プロセスから直接登録**  
  現在起動中のアプリを一覧から選んで即座に監視対象に追加できます。

- **タスクトレイ常駐**  
  最小化・閉じるとトレイに格納。右クリックメニューから監視状態の確認・一時停止・終了が可能です。

- **Windows スタートアップ登録**  
  PC 起動時に自動起動するよう設定できます。

- **動作ログ**  
  スリープ防止の有効化・解除をログファイルに記録します（自動ローテーション対応）。

---

## 動作原理

スリープ防止には Windows API の `SetThreadExecutionState` を使用します。

- 監視プロセス起動中 → `ES_CONTINUOUS | ES_SYSTEM_REQUIRED` でスリープをブロック
- 全プロセス停止＋猶予時間経過 → `ES_CONTINUOUS` でブロックを解除し、Windows の電源設定に制御を返す

SleepGuard 自身がスリープを発動させるわけではなく、Windows の電源設定を「一時的に無効化する」動作です。

---

## 必要環境

| 項目 | バージョン |
|---|---|
| OS | Windows 10 / 11 (x64) |
| .NET Runtime | 8.0 以上 |

> .NET 8 ランタイムは [Microsoft 公式サイト](https://dotnet.microsoft.com/download/dotnet/8.0) からダウンロードできます。

---

## インストール

インストール不要です。EXE ファイルをダブルクリックするだけで起動します。

1. [Releases](../../releases) から最新の `SleepGuard.zip` をダウンロード
2. 任意のフォルダに展開
3. `SleepGuard.exe` を実行

---

## ビルド方法

### 必要なツール

- Visual Studio 2022 (17.x 以上)
- .NET SDK 8.0 以上

### 手順

```bash
git clone https://github.com/YOUR_USERNAME/SleepGuard.git
cd SleepGuard
dotnet restore
dotnet build -c Release
```

または Visual Studio 2022 で `SleepGuard.sln` を開き、`Release | x64` でビルドしてください。

---

## 設定項目

| 設定 | 説明 | デフォルト |
|---|---|---|
| スリープ猶予時間 | 全プロセス停止後にスリープを許可するまでの時間 | 5分 |
| 確認間隔 | プロセスの死活監視間隔 | 10秒 |
| ディスプレイのスリープも防止 | モニターのスリープも同時にブロック | OFF |
| タスクトレイに常駐する | ×ボタンで閉じてもトレイに残る | ON |
| 起動時にタスクトレイに格納 | 起動時にウィンドウを表示しない | OFF |
| 動作ログを記録する | ログファイルへの書き出し | ON |
| Windows 起動時に自動起動 | スタートアップへの登録 | OFF |

---

## データ保存場所

設定・ログは EXE とは別の場所に保存されます。EXE を移動・削除しても設定は消えません。

```
%APPDATA%\SleepGuard\
  ├── settings.json   # 監視プロセス一覧・設定
  └── sleepguard.log  # 動作ログ（1MB超で自動ローテーション）
```

---

## 使用ライブラリ

| ライブラリ | バージョン | ライセンス |
|---|---|---|
| [Hardcodet.NotifyIcon.Wpf](https://github.com/hardcodet/wpf-notifyicon) | 1.1.0 | CPOL |
| [System.Text.Json](https://github.com/dotnet/runtime) | 8.0.5 | MIT |
| .NET 8 | 8.0 | MIT |

---

## ライセンス

[MIT License](LICENSE)

---

## 注意事項

- 本ソフトウェアは現状有姿（AS IS）で提供されます。
- スリープ制御はシステムの電源設定に依存します。一部の環境では期待通りに動作しない場合があります。
- 本ソフトウェアの使用によって生じたいかなる損害についても、作者は責任を負いません。
