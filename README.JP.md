[Click here for the English version](./README.md)

# 電子書籍データ管理とメールマージ自動化

このプロジェクトは、Googleスプレッドシートで電子書籍データを整理し、[グループマージアドオン](https://www.scriptable-assets.page/add-ons/group-merge/)を使用して個別のメールを自動送信するGoogle Apps Scriptを使用しています。このスクリプトは、確認メールとリンク共有メールの2種類のメールを処理するよう特別に設計されています。

- 確認メール: スクリプトは、受信者の名前、コース情報、書籍の詳細など、各確認メールに固有の項目を生成するためのデータを処理します。その後、[グループマージアドオン](https://www.scriptable-assets.page/add-ons/group-merge/)を使用して、受信者にパーソナライズされた確認メールを送信します。

- リンク共有メール: このスクリプトは、リンク共有メールのデータも処理します。受信者の名前、コースの詳細、共有リンクなどの必要な情報を整理します。[グループマージアドオン](https://www.scriptable-assets.page/add-ons/group-merge/)を使用して、電子書籍またはリソースへのアクセスに関連するリンクを提供する、パーソナライズされたリンク共有メールを受信者に送信します。

これらのプロセスを自動化することで、Google Apps Scriptは電子書籍データの管理と配布を合理化し、大量の受信者にパーソナライズされたメールを簡単に送信することができます。

## 前提条件

- GoogleシートにアクセスできるGoogleアカウント。
- GoogleシートとGoogle Apps Scriptの基本的な理解。
- [グループマージアドオン](https://www.scriptable-assets.page/add-ons/group-merge/)を使用するために別のスプレッドシートを設定する。

## 設定

1. **Googleシートを開く**: [サンプルGoogleシート](https://docs.google.com/spreadsheets/d/1mMuQSK06hIcAUcI1qW4cgD2_IKOXU9_DAR2CTj2a-a8/edit#gid=1834592607)にアクセスします。
2. **GAS認証を実施**: `Initial Setting`シートにアクセスし、初期設定ボタンをクリックします。これにより、Google Apps Scriptの認証ページに移動します。

   ![初期設定の画像](docs/assets/images/initial-setting.png)

3. **組み込み関数の定数変数をカスタマイズ**: Apps Scriptページに移動し、`variables.gs`の定数変数を必要に応じて調整します。

   ![変数カスタマイズの画像](docs/assets/images/custom-variables.png)

## 使用方法

1. **参照データを入力**: 参照シート(`Confirm Mail-Ref Data` / `Link Share Mail-Ref Data`)の指定されたオレンジ色の範囲に参照データを入力します。

   ![参照データ入力の画像](docs/assets/images/input-ref-data.png)

2. **データを処理**: GoogleシートのUIにあるカスタムメニューから関数を実行します。

   ![データ処理の画像](docs/assets/images/output-data.png)

   - `確認メール用のメールマージ