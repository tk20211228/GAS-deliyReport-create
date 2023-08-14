# GAS-deliyReport-create
初期コードをマージ

## ブランチ命名ルール

### 1. 機能やタスクベースの命名

- 新しい機能やタスクごとにブランチを作成する
- 命名規則: `feature/機能名` または `task/タスク名`

**例**:

```plaintext
feature/csv-import
feature/ui-redesign
task/bugfix-issue123
```

### 2. チケットや課題トラッキングシステムベースの命名

- トラッキングシステムと連動してブランチを作成する
- 命名規則: `issue/チケット番号`

**例**:

```plaintext
issue/CSV-101
issue/123
```

### 3. ユーザー名や作業者ベースの命名

- 作業者の名前を用いてブランチを作成する
- 命名規則: `username/機能名`

**例**:

```plaintext
john/csv-import
```

### 4. その他の命名規則

- リリースやホットフィックス関連のブランチ名
- 命名規則: `release/バージョン番号` 、 `hotfix/修正名`

**例**:

```plaintext
release/v1.2.0
hotfix/fix-login-issue
```
