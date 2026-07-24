# 小規模団体向け イベント受付管理システム


## 概要


このリポジトリは、Raspberry Pi 3 と Gmail、Excel、Dropbox を使って、
小規模団体のイベント申込を受付台帳へ自動登録するための実践サンプルです。


東京四万十会オンライン交流会の受付管理をもとに作成しています。


対象規模は、最大100名程度の単発イベントを想定しています。


## 主な機能


- Gmail通知メールの取得
- 簡易HPフォームから届く申込メールの分類
- Excel受付台帳への自動登録
- GmailメッセージIDによる重複登録防止
- Dropboxへの台帳コピー
- 毎日17時の Daily Report メール送信
- cronによる定期実行


## 動作確認環境


- Raspberry Pi 3
- Raspberry Pi OS
- Python 3.11
- Gmail API
- openpyxl
- rclone
- Dropbox


Raspberry Pi Zero W / Raspberry Pi Zero 2 W については検討しましたが、
レスポンス面で不安があり、本番テストは実施していません。


そのため、現時点では Raspberry Pi 3 のみを動作対象とします。


## 主なスクリプト


- scripts/import_online_to_excel.py
- scripts/daily_report_mail.py
- scripts/view_excel.py
- scripts/test_gmail_read.py


## cron設定例


5分ごとに受付処理を実行します。


*/5 * * * * /home/pi/gmail_sorter/venv/bin/python /home/pi/gmail_sorter/import_online_to_excel.py >> /home/pi/gmail_sorter/gmail_sorter.log 2>&1


毎日17時にDaily Reportメールを送信します。


0 17 * * * /home/pi/gmail_sorter/venv/bin/python /home/pi/gmail_sorter/daily_report_mail.py >> /home/pi/gmail_sorter/daily_report_mail.log 2>&1


## GitHubに登録しないファイル


以下のファイルは認証情報や個人情報を含むため、GitHubには登録しません。


- credentials.json
- token.json
- mail_config.py
- 受付台帳.xlsx
- gmail_sorter.log
- daily_report_mail.log


## 今後の展開


現在は Excel + Dropbox 版です。


今後は Googleスプレッドシート連携版へ拡張し、
より汎用的な単発イベント受付管理システムとして整理する予定です。
