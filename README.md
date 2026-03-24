# AdoSQL

Windows 向けの **コマンドライン対話型 Access SQL クライアント**（実行イメージ: `adosql.exe`）。ADO 経由で `.accdb`（主対象）および環境に応じ `.mdb` に接続し、Access SQL を対話実行する。

リポジトリ: [github.com/kuwa2005/AdoSQL](https://github.com/kuwa2005/AdoSQL)

## ステータス

- **仕様**: [仕様書.md](./仕様書.md) に集約
- **実装**: **初版あり**（CMake + C++ / ADO）。**Tab 補完・行エディタの自前実装は未着手**（Issue #13〜#15）

## 実装方針（要約）

- **ネイティブ（C++ 想定）**・**.NET 不要**
- **32bit / 64bit で別 EXE**（ACE のビット数と揃える）
- **cmd.exe / PowerShell** から利用
- `adosql.exe` をフォルダにコピーするだけで使う想定。`adosql.ini` は必要に応じて**同じフォルダに自動生成**
- 対話入力では **Tab 補完**（初版: 対話コマンド・`set` 系・`@` パス、接続後はテーブル／保存クエリ名。詳細は仕様書 §3.4）

詳細は仕様書 **§0 実装方針** および **§0.1 実装前決定事項（技術方針）** を参照。

## 動作要件（利用者向け）

- **Windows**（コンソールから起動）
- **Microsoft Access Database Engine（ACE）** 等、OLE DB 経由で Access ファイルに接続できるプロバイダーがインストールされていること
- 32bit 版 Office / ACE のみの環境では **32bit 版 EXE**、64bit 版では **64bit 版 EXE** を使う（混在環境では状況に合わせて選択）

## 起動イメージ（仕様）

```text
adosql.exe <データベースファイルのパス> [パスワード]
```

空白を含むパスはシェルでクォートする。パスワードを引数で渡すとプロセス一覧や履歴に残る可能性がある（仕様書 §10）。

## セキュリティ・運用

- `@` スクリプトは信頼できるローカルパスのみを想定（仕様書 §4, §10）
- ローカル／信頼境界内での利用を想定

## ライセンス

[MIT License](./LICENSE)（Copyright (c) 2026 kuwa2005）

## 実装タスク（GitHub Issues）

実装は [Issues](https://github.com/kuwa2005/AdoSQL/issues) にチェックリスト付きで分割してある。各 Issue の元テキストはリポジトリの [`.github/issue-bodies/`](./.github/issue-bodies/) にも置いてある（再作成・オフライン参照用）。

## ビルド（開発者向け）

前提: **Visual Studio 2022**（「C++ によるデスクトップ開発」）、**CMake**（3.20+、PATH に `cmake`）。実行時は対象ビット数に合った **ACE** が必要。

```powershell
cmake -B build -G "Visual Studio 17 2022" -A x64
cmake --build build --config Release
# 出力: build\Release\adosql.exe

# 32bit (Win32)
cmake -B build32 -G "Visual Studio 17 2022" -A Win32
cmake --build build32 --config Release
```

`#import` は既定で `C:\Program Files\Common Files\System\ado\`（x64）または `Program Files (x86)\...`（Win32）の `msado15.dll` / `msadox.dll` を参照する。標準以外の配置なら `ado_session.cpp` のパスを環境に合わせて変更すること。

## 貢献

Issue ベースで進めています（上記「実装タスク」参照）。
