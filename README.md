# SleepGuard

## v1.4.4 cleanup
- 「実行中プロセスから登録」の検索テキストボックスを削除
- 追加導線をボタンと実行中プロセス一覧に整理
- 状態同期版に合わせて未使用のカウントダウンUIを削除
- 実行中プロセス一覧の更新ロジックを簡素化

WPF (.NET 8 / VS2022) で作成した、監視対象プロセスの実行中だけ Windows の自動スリープを抑止するツールです。

## 構成
- `SleepGuard.sln` : Visual Studio 2022 ソリューション
- `SleepGuard/` : 本体プロジェクト
- `SleepGuard/Resources/` : アイコン
- `SleepGuard/Themes/` : 色・スタイル定義
- `SleepGuard/Models/` : 設定モデル
- `SleepGuard/Services/` : 監視サービス

## v1.4.3 の要点
- `Path` の曖昧参照を解消して Release x64 のビルドエラーを修正
- 監視対象追加 UI を `監視プロセスを追加` ボタン 1 個に整理
- 参照ダイアログで EXE を選んで `開く` を押した時点で即追加
- 旧バージョンごとの README を整理し、ルートをシンプル化

## GitHub へ上げる時の最小構成
この ZIP を展開したら、そのままリポジトリ直下に置いて問題ありません。
不要であれば下記はコミットしなくても構いません。
- `SleepGuard/bin/`
- `SleepGuard/obj/`
- `.vs/`

`.gitignore` を同梱しています。
