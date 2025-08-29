# Discord 起床時刻トラッカー

Discordの特定チャンネルにおけるユーザーのその日の最初の投稿を「起床時刻」として記録し、Googleスプレッドシートに自動で集計・分析するBotです。

## 主な機能

-   **起床時刻の自動記録**: 毎日の最初の投稿（0:00〜16:59）をそのユーザーの起床時刻として記録します。
-   **自動集計**: 毎日深夜に自動で集計が実行され、スプレッドシートが更新されます。
-   **月別データ分析**:
    -   累計の平均起床時間と記録日数
    -   今月の平均起床時間
    -   先月の平均起床時間
    -   先月からの変化量（±分）
-   **ランキング表示**: 累計の平均起床時間が早い順にユーザーをランキングします。
-   **月次レポート通知**: 毎月末に、その月の集計結果サマリーをDiscordチャンネルに自動で投稿します。
-   **手動トリガー**: `!analyze`コマンドでいつでも手動で集計を実行できます。
-   **スマホ最適化**: スマートフォンでも見やすいようにレイアウトされたスプレッドシート。

## アーキテクチャ

このシステムは、複数のサービスが連携して動作しています。

-   **Render**: Pythonで書かれたDiscordボット本体を24時間ホスティングします。
-   **Google Apps Script (GAS)**: 2つの役割を担います。
    1.  **常時起動 (Keep-Alive)**: Renderの無料プランがスリープしないように15分おきにPingを送信します。
    2.  **日次トリガー**: 毎日決まった時刻にRenderのWebhookを呼び出し、集計処理を開始させます。月の最終日には、Discordへの通知も行います。
-   **Google Sheets**: 全ての集計データを保存し、ランキングを表示するデータベース兼ダッシュボードとして機能します。
-   **Discord**: ユーザーの投稿データを取得し、月次レポートを通知するプラットフォームです。

## セットアップ手順

1.  **Discordの準備**
    -   [Discord Developer Portal](https://discord.com/developers/applications)でBotを作成し、**ボットトークン**を取得する。
    -   通知を送りたいチャンネルの**ウェブフックURL**を作成・取得する。

2.  **Google Cloud Platformの準備**
    -   Google Cloudプロジェクトを作成する。
    -   「Google Sheets API」と「Google Drive API」を有効化する。
    -   サービスアカウントを作成し、キー（`credentials.json`）をダウンロードする。

3.  **Google Sheetsの準備**
    -   新しいスプレッドシートを作成し、その**スプレッドシートID**を控える。
    -   作成したサービスアカウントのメールアドレスを「編集者」としてスプレッドシートに共有する。

4.  **GitHubの準備**
    -   このリポジトリに必要な`server.py`と`requirements.txt`を配置する。

5.  **Renderのデプロイ**
    -   RenderにGitHubアカウントでサインアップし、「Web Service」を新規作成する。
    -   GitHubリポジトリを連携させる。
    -   以下の設定でサービスを構成する。
        -   **Build Command**: `pip install -r requirements.txt`
        -   **Start Command**: `gunicorn server:app`
        -   **Instance Type**: `Free`
    -   下の「環境変数」セクションに従って、全ての環境変数を設定する。

6.  **Google Apps Scriptの設定**
    -   新しいGASプロジェクトを作成する。
    -   `keepRenderAlive`（常時起動用）と`triggerDailyAnalysis`（日次トリガー＆通知用）の2つの関数をコードに記述する。
    -   以下の2つのトリガーを設定する。
        1.  `keepRenderAlive`を**15分おき**に実行するトリガー。
        2.  `triggerDailyAnalysis`を**毎日定時**（例: 15時〜16時）に実行するトリガー。

## 環境変数

Renderの「Environment」タブで以下の環境変数を設定する必要があります。

| Key                         | Value                                        | 説明                                         |
| --------------------------- | -------------------------------------------- | -------------------------------------------- |
| `DISCORD_TOKEN`             | `Mzg...`                                     | Discordボットのトークン                      |
| `TARGET_CHANNEL_ID`         | `123456789...`                               | 分析対象のDiscordチャンネルID                |
| `SHEET_ID`                  | `1aBcDeFg...`                                | 結果を書き込むGoogleスプレッドシートのID     |
| `GOOGLE_CREDENTIALS_JSON`   | `{ "type": "service_account", ... }`         | `credentials.json`ファイルの中身を全て貼り付け |
| `TRIGGER_SECRET`            | `your-very-long-and-secret-password`         | GASからのトリガーを認証するための秘密の文字列  |

## 使い方

-   **自動集計**: GASのトリガーによって毎日自動で実行されます。
-   **手動集計**: Discordの分析対象チャンネルで`!analyze`と投稿すると、いつでも手動で集計を実行し、結果のURLを受け取ることができます。
