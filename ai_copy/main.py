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


def _bullet_list(items: list) -> str:
    """リストを箇条書き文字列に変換する"""
    return "\n".join(f"- {item}" for item in items) if items else ""


def _format_dict_list(items: list, key_field: str, detail_fields: list) -> str:
    """辞書リストを読みやすい文字列に変換する"""
    lines = []
    for item in items:
        parts = [f"**{item.get(key_field, '不明')}**"]
        for field in detail_fields:
            if field in item:
                parts.append(f"{item[field]}")
        lines.append("- " + " / ".join(parts))
    return "\n".join(lines)


def build_system_prompt(profile: dict) -> str:
    """プロフィールからシステムプロンプトを生成する"""
    name = profile.get("name", "ユーザー")
    nickname = profile.get("nickname", "")
    style = profile.get("speaking_style", {})
    about = profile.get("about", {})
    core_values = profile.get("core_values", {})
    thinking = profile.get("thinking_patterns", {})
    career = profile.get("career", {})
    expertise = profile.get("expertise", {})
    business_ideas = profile.get("business_ideas", [])
    family = profile.get("family", {})
    relationships = profile.get("relationships", [])
    life_exp = profile.get("life_experiences", {})
    philosophy = profile.get("philosophy", {})
    guardrails = profile.get("guardrails", [])
    keywords = profile.get("personality_keywords", [])
    introduction = profile.get("introduction", "")

    first_person = style.get("first_person", "僕")
    tone = style.get("tone", "丁寧")
    endings = style.get("sentence_endings", [])
    characteristics = style.get("characteristics", [])

    # キャリア情報の構築
    current = career.get("current_role", {})
    side_biz = career.get("side_businesses", [])
    vision = career.get("vision", [])

    prompt = f"""あなたは「{name}」（ニックネーム：{nickname}）のコピーAIです。
{name}本人になりきって会話してください。

## 絶対ルール
{_bullet_list(guardrails)}

## 基本プロフィール
- 氏名: {name}（{profile.get("reading", "")}）
- ニックネーム: {nickname}
- 生年月日: {about.get("birthday", "")} {about.get("birth_time", "")}
- 年齢: {about.get("age", "")}歳
- MBTI: {about.get("mbti", "")}
- 居住地: {about.get("location", "")}
- 一人称: 「{first_person}」

## 話し方・コミュニケーション
- トーン: {tone}
- 語尾の例: {", ".join(endings)}
{_bullet_list(characteristics)}

## コアバリュー
{_bullet_list(core_values.get("primary", []))}

### 信念
{_bullet_list(core_values.get("beliefs", []))}

### 嫌いなこと
{_bullet_list(core_values.get("dislikes", []))}

## 思考パターン
### 意思決定の基準（上から優先）
{_bullet_list(thinking.get("decision_criteria", []))}

### 思考の傾向
{_bullet_list(thinking.get("tendencies", []))}

### モチベーションの源泉
{_bullet_list(thinking.get("motivation_sources", []))}

## キャリア
### 本業
- 会社: {current.get("company", "")} {current.get("division", "")}
- 役職: {current.get("position", "")}
- 主要プロジェクト:
{_bullet_list(current.get("projects", []))}

### 副業
{_format_dict_list(side_biz, "client", ["role", "work"])}

### キャリアビジョン
{_bullet_list(vision)}

## スキル・専門領域
### 技術スキル
{_bullet_list(expertise.get("technical", []))}

### ビジネススキル
{_bullet_list(expertise.get("business", []))}

## ビジネス構想
{_bullet_list(business_ideas)}

## 家族
- 妻: {family.get("wife", {}).get("name", "")} — {family.get("wife", {}).get("traits", "")}
- 結婚記念日: {family.get("wedding_anniversary", "")}
- 母: {family.get("mother", {}).get("name", "")}（{family.get("mother", {}).get("age", "")}歳、{family.get("mother", {}).get("location", "")}）
- 父: {family.get("father", {}).get("note", "")}
- 妹: {family.get("sister", {}).get("name", "")}
- 甥: {family.get("nephew", {}).get("name", "")}

## 主要な人間関係
{_format_dict_list(relationships, "name", ["relation"])}

## 人生経験
### 成功体験
{_bullet_list(life_exp.get("achievements", []))}

### 乗り越えた困難
{_bullet_list(life_exp.get("hardships", []))}

### 現在の課題
{_bullet_list(life_exp.get("current_challenges", []))}

## 人生観・死生観
- {philosophy.get("views_on_death", "")}
### 理想の未来
{_bullet_list(philosophy.get("ideal_future", []))}

## 趣味・興味
{_bullet_list(about.get("interests", []))}

## 人格の核キーワード
{", ".join(keywords)}

## 自己紹介（参考）
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
    nickname = profile.get("nickname", name)

    # Claude クライアント初期化
    client = Anthropic(api_key=api_key)
    system_prompt = build_system_prompt(profile)
    messages = []

    print(f"{'=' * 50}")
    print(f"  {name}（{nickname}）のコピーAI へようこそ！")
    print(f"  （終了するには 'quit' または 'exit' と入力）")
    print(f"{'=' * 50}")
    print()

    # 初回の挨拶
    introduction = profile.get("introduction", "").strip()
    if introduction:
        print(f"{nickname}: {introduction}")
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
            print(f"\n{nickname}: またね！いつでも話しかけてね。")
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
            print(f"\n{nickname}: {assistant_message}\n")
        except Exception as e:
            print(f"\nエラーが発生しました: {e}\n")
            messages.pop()  # 失敗したメッセージを除去


if __name__ == "__main__":
    main()
