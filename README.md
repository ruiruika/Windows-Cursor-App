# Windows Mouse Cursor Manager

[日本語](#日本語) | [English](#english)

---

## 日本語

マウスカーソルを一括で簡単に変更するためのツールです。

### 特徴
- **一括適用**: フォルダ内のカーソルを解析し、Windowsの全ポインタをテーマに沿って一括変更します。
- **グローバル対応**: 日本語だけでなく、英語のファイル名（Normal, Busy, Link など）も自動判別します。
- **直感的な操作**: 適用したいフォルダ内のファイルを一つ選ぶだけで、そのフォルダ全体が適用されます。
- **多言語対応**: 日本語と英語をサポートしています。
- **配布に最適**: インストール不要の `.exe` 形式です。
- **ソースコード同梱**: 透明性のため、元コード（`cursor_manager_gui.py`）を同梱しています。

### 使い方
1. `CursorApp.exe` を実行します。
2. 「カーソルセットを適用する」ボタンを押します。
3. エクスプローラーが開くので、適用したいフォルダの中に入り、**中にある適当なカーソルファイルを一つ選択**してください。
4. そのフォルダにあるすべての対応カーソルが自動的に設定されます。

---

## English

A simple tool to change Windows mouse cursors in bulk.

### Features
- **Bulk Apply**: Analyzes files in a folder and updates all Windows pointers at once.
- **Global Support**: Automatically detects both Japanese and common English filenames (Normal, Busy, Link, etc.).
- **Intuitive**: Just pick any cursor file in your target folder to apply the entire set.
- **Multi-language**: Supports Japanese and English.
- **Portable**: Standalone `.exe` format (no Python installation required).
- **Source Code Included**: The original script (`cursor_manager_gui.py`) is included for transparency.

### How to Use
1. Run `CursorApp.exe`.
2. Click the "Apply Cursor Set" button.
3. When the explorer opens, navigate to your target folder and **select any cursor file inside**.
4. All corresponding cursors in that folder will be automatically applied.
