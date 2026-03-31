#!/usr/bin/env python3
"""コピーAI - あなたの分身AIチャットボット (CLI版)"""

import os
import sys
import yaml
from anthropic import Anthropic


def load_profile(profile_path: str = None) -> dict:
    """プロフィール設定を読み込む"""
    if profile_path is None:
        profile_path = os.path.join(os.path.dirname(__file__), "profile.yaml")
    with open(profile_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def build_system_prompt(profile: dict) -> str:
    """プロフィールからシステムプロンプトを生成する"""
    name = profile.get("name", "ユーザー")
    style = profile.get("speaking_style", {})
    expertise = profile.get("expertise", [])
    introduction = profile.get("introduction", "")
    traits = profile.get("additional_traits", [])
    about = profile.get("about", {})

    first_person = style.get("first_person", "僕")
    tone = style.get("tone", "丁寧")
    endings = style.get("sentence_endings", [])
    characteristics = style.get("characteristics", [])

    prompt = f"""あなたは「{name}」のコピーAIです。{name}本人になりきって会話してください。

## 基本ルール
- 一人称は「{first_person}」を使ってください
- 話し方は「{tone}」な口調で話してください
- AIであることを自分から明かさないでください
- {name}本人として自然に会話してください

## 話し方の特徴
{chr(10).join(f'- {c}' for c in characteristics)}

## 語尾の例
{chr(10).join(f'- {e}' for e in endings)}

## 職業・所属
- 仕事: {about.get("job", "未設定")}
- 場所: {about.get("location", "未設定")}

## 興味・関心
{chr(10).join(f'- {i}' for i in about.get("interests", []))}

## 専門知識
{chr(10).join(f'- {e}' for e in expertise)}

## 性格・特徴
{chr(10).join(f'- {t}' for t in traits)}

## 自己紹介
{introduction}
"""
    return prompt


def main():
    """メインのチャットループ"""
    # APIキーの確認
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("エラー: ANTHROPIC_API_KEY 環境変数を設定してください。")
        print("  export ANTHROPIC_API_KEY='your-api-key-here'")
        sys.exit(1)

    # プロフィール読み込み
    profile_path = sys.argv[1] if len(sys.argv) > 1 else None
    profile = load_profile(profile_path)
    name = profile.get("name", "コピーAI")

    # Claude クライアント初期化
    client = Anthropic(api_key=api_key)
    system_prompt = build_system_prompt(profile)
    messages = []

    print(f"{'=' * 50}")
    print(f"  {name} のコピーAI へようこそ！")
    print(f"  （終了するには 'quit' または 'exit' と入力）")
    print(f"{'=' * 50}")
    print()

    # 初回の挨拶
    introduction = profile.get("introduction", "").strip()
    if introduction:
        print(f"{name}: {introduction}")
        print()

    while True:
        try:
            user_input = input("あなた: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nまたね！")
            break

        if not user_input:
            continue
        if user_input.lower() in ("quit", "exit", "q"):
            print(f"\n{name}: またね！いつでも話しかけてね。")
            break

        messages.append({"role": "user", "content": user_input})

        try:
            response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=1024,
                system=system_prompt,
                messages=messages,
            )
            assistant_message = response.content[0].text
            messages.append({"role": "assistant", "content": assistant_message})
            print(f"\n{name}: {assistant_message}\n")
        except Exception as e:
            print(f"\nエラーが発生しました: {e}\n")
            messages.pop()  # 失敗したメッセージを除去


if __name__ == "__main__":
    main()
