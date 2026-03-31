# コピーAI - あなたの分身AIチャットボット

Claude APIを使って、あなたの口調・知識・個性を再現するCLIチャットボットです。

## セットアップ

### 1. 依存パッケージのインストール

```bash
pip install -r ai_copy/requirements.txt
```

### 2. APIキーの設定

[Anthropic Console](https://console.anthropic.com/) からAPIキーを取得し、環境変数に設定します。

```bash
export ANTHROPIC_API_KEY='your-api-key-here'
```

### 3. プロフィールのカスタマイズ

`ai_copy/profile.yaml` を編集して、あなたの個性を設定してください。

設定できる項目：
- **name** — 名前
- **about** — 職業・興味関心
- **speaking_style** — 口調・一人称・語尾の特徴
- **expertise** — 専門知識
- **introduction** — 自己紹介文
- **additional_traits** — その他の性格・特徴

## 使い方

```bash
python ai_copy/main.py
```

カスタムプロフィールを指定する場合：

```bash
python ai_copy/main.py path/to/my_profile.yaml
```

終了するには `quit`、`exit`、または `Ctrl+C` を入力してください。
