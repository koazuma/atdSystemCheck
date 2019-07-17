# 勤怠システムチェックツール

## 概要

アマノ勤怠システムの登録内容をチェックし、チェック結果を本人および上長にメール通知するためのツールです。
チェックの種類は以下の通りです。

* 残業時間
締め日を20日として閾値以上の残業をしているメンバーおよびその上長にメールで通知します。
閾値は25時間、エスカレーションレベルは無制限です。
* 申請漏れ
打ち忘れチェックリストに挙げられている項目を対象メンバーおよびその上長にメールで通知します。
エスカレーションレベルは1(対象者の1つ上まで)です。

## 環境要件

* pyhon 3.7以上
* selenium webdriver
* relativedelta

## 事前準備

以下はwindows7での設定内容です。

### 環境設定

setting.iniを編集して環境設定をします。

* chromedriver設定

ダウンロードしたchromedriverのフルパスを記載します。

```ini
[environment]
chromedriver = c:/driver/chromedriver.exe
```

* ログイン情報

cyberxeedのログイン情報を記載します。

```ini
[siteinfo]
url = https://cxg8.i-abs.co.jp/cyberx/login.asp

# login info
cp = xxx    ←変更(自分の会社コード)
id = xxx    ←変更(自分の個人コード)
pw = xxx    ←変更(自分のパスワード)
```

* gmail設定

送信用に使用するgmailアカウントとパスワードを記載します。

```ini
[mail]
smtp_host = smtp.gmail.com
smtp_port = 587
username = xxx@gmail.com    ←変更(gmailアドレス)
password = xxx              ←変更(googleアカウントのパスワード)
```

### googleアカウント設定

pythonからのメール送信を可能にするために、<https://myaccount.google.com/security>からgoogleアカウントの「安全性の低いアプリのアクセス」をONにしておく必要があります。(デフォルトはOFF)

### 社員情報設定

members.jsonを編集して社員情報の設定をします。
ある社員の社員番号をkeyとして、氏名、グループ、メールアドレス、上長(の社員番号)のオブジェクトがvalueという形の配列になっています。
上長は空、もしくはカンマ区切りでの複数設定も可能です。

```json
{
    "100":{
        "name":"野比のび太",
        "group":"1",
        "mail":"nobitanokuseninamaikida@gmail.com",
        "boss":"200,201"
    },
    "200":{
        "name":"骨川スネ夫",
        "group":"1",
        "mail":"sunechama@gmail.com",
        "boss":"300"
    },
    "201":{
        "name":"源静香",
        "group":"1",
        "mail":"sizuka@gmail.com",
        "boss":""
    },
    "300":{
        "name":"剛田武",
        "group":"",
        "mail":"gian@gmail.com",
        "boss":""
    },
}
```

## 使用方法

まずは以下usage。

```batch
> python checkOverWork.py -h
usage: checkOverWork.py [-h] -m {1,2} [-d DATE] [-e]

optional arguments:
  -h, --help            show this help message and exit
  -m {1,2}, --mode {1,2}
                        実行モード 1:残業時間チェック 2:打ち忘れチェック
  -d DATE, --date DATE  yyyymmdd形式で日を指定すると、その日に実行した仮定で実行される。
  -e, --exholiday       土日祝日の場合は何もせず終了する。
```

### -mオプション

* 1 : 残業時間確認
残業時間チェックは-mオプションに"1"を指定して下さい。
コマンドラインから以下を実行すると現在の実行者が確認可能なメンバー全員分の実行当日を含む期間の残業時間をチェックし、本人および上長にメールします。
エスカレーションレベルは無制限です。

```batch
python checkOverWork.py -m 1
```

* 2 : 打ち忘れチェックリスト確認
打ち忘れチェックは-mオプションに"2"を指定して下さい。
コマンドラインから以下を実行すると現在の実行者が確認可能な当月分のメンバー全員分の打ち忘れチェックリスト内容を取得し、本人および上長にメールします。
エスカレーションレベルは1です。(1つ上の上長まで)

```batch
python checkOverWork.py -m 2
```

### -dオプション

yyyymmddの日付つきでdオプションを指定すると、指定日に実行した仮定でコマンドが実行されます。前月分の状況確認等に。

```batch
python checkOverWork.py -m 1 -d 20190401
```

### -eオプション

eオプションを指定すると土日祝日は実行しなくなる。
祝日は以下からDL可能な内閣府配布の祝日リストを同ディレクトリに格納している事を前提とする。
<https://www8.cao.go.jp/chosei/shukujitsu/gaiyou.html#syukujitu>

```batch
python checkOverWork.py -m 2 -e
```

## 著作者

[koazuma](https://github.com/koazuma)
