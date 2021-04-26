# 勤怠システムチェックツール

## 概要

アマノ勤怠システムの登録内容をチェックし、チェック結果をメール通知またはCSV出力するためのツールです。

* 残業時間チェック
* 申請忘れチェック
* 工数登録チェック

## 環境要件

* python3.8以上
* selenium chrome webdriver
* dateutil
* syukujitsu.csv

## 事前準備

以下はwindows7での設定内容です。

### 環境構築

* python3系
  インストールする。
* 必要ライブラリ
  pipで以下ライブラリをインストールする。
  * selenium
  * relativedelta(dateutil)
* chrome webdriver
  以下より使用バージョンに合ったchromedriverをダウンロードしておく。
  ChromeDriver <https://chromedriver.chromium.org/downloads>
* 本プロジェクト
  Git CLONEもしくはダウンロードしておく。
* 祝日リスト
  内閣府HPより国民の祝日csvをダウンロードしてプロジェクトと同じディレクトリに配置しておく。
  内閣府HP <https://www8.cao.go.jp/chosei/shukujitsu/gaiyou.html#syukujitu>

### アプリケーション環境設定

setting.iniを編集して環境設定をします。

* chromedriver設定
  ダウンロードしたchromedriverのフルパスを記載します。
  ※ディレクトリの区切りは\ではなく/を使用する

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
ある社員の社員番号をkeyとして"氏名、グループ、メールアドレス、上長(の社員番号)、無視設定"のオブジェクトがvalueという形の配列になっています。
上長は空、もしくはカンマ区切りでの複数設定も可能です。
無視設定を0以外にすると通知されなくなります。

```json
{
    "100":{
        "name":"野比のび太",
        "group":"1",
        "mail":"nobitanokuseninamaikida@gmail.com",
        "boss":"200,201",
        "ignore":"0"
    },
    "200":{
        "name":"骨川スネ夫",
        "group":"1",
        "mail":"sunechama@gmail.com",
        "boss":"300",
        "ignore":"1"
    },
    "201":{
        "name":"源静香",
        "group":"1",
        "mail":"sizuka@gmail.com",
        "boss":"",
        "ignore":"0"
    },
    "300":{
        "name":"剛田武",
        "group":"",
        "mail":"gian@gmail.com",
        "boss":"",
        "ignore":"0"
    },
}
```

## 使用方法

### usage

```batch
> python atdSystemCheck.py -h
usage: atdSystemCheck.py [-h] -m {1,2,3} -o {1,2} [-d DATE] [-e]

optional arguments:
  -h, --help            show this help message and exit
  -m {1,2,3}, --mode {1,2,3}
                        チェック種別 1:残業時間 2:打ち忘れ 3:工数登録
  -o {1,2}, --output {1,2}
                        出力タイプ 1:メール送信 2:CSVファイル出力
  -d DATE, --date DATE  yyyymmdd形式で日を指定すると、その日に実行した仮定で実行される。
  -e, --exholiday       土日祝日の場合はチェックをしない。
```

### オプション

#### -m --mode

チェック種別を指定する必須オプション。

* 1 : 残業時間チェック
  残業時間チェックは-mオプションに"1"を指定して下さい。
  指定日(デフォルトは前日)を含む期間の残業時間をチェックします。

  ```batch
  python atdSystemCheck.py -m 1 -o 1
  ```

* 2 : 打ち忘れチェック
  打ち忘れチェックは-mオプションに"2"を指定して下さい。
  指定日(デフォルトは前日)を含む期間の打ち忘れチェックをします。

  ```batch
  python atdSystemCheck.py -m 2 -o 1
  ```

* 3 : 工数登録チェック
  残業時間のファイル出力は-mオプションに"3"を指定して下さい。
  指定日(デフォルトは前日)を含む期間の工数登録チェックをします。

  ```batch
  python atdSystemCheck.py -m 3 -o 1
  ```

#### -o --output

チェック結果をメール通知またはCSVファイル出力のいずれかを選択できます。

* 1 : メール通知
  メール通知の場合は-oオプションに1を指定して下さい。チェック結果を対象者にメール通知します。

  ```batch
  python atdSystemCheck.py -m 1 -o 1
  ```

  * エスカレーションレベル
    mオプションのチェック種別によってエスカレーションレベルが変わります。
    1(残業時間チェック)は無制限、2(打ち忘れチェック),3(工数登録チェック)は1つ上までとなります。エスカレーションレベルはsetting.iniのMAIL_ESC_LEVELで変更可能です。

    ```ini
    MAIL_ESC_LEVEL = -1
    ```

  * 残業時間閾値
    mオプションが1(残業時間チェック)の場合、合計残業時間が指定閾値を超過しているメンバーのみにメールが送信されます。閾値は25時間となります。
    閾値はsetting.iniのOVERWORK_THRESHOLDで変更可能です。

    ```ini
    OVERWORK_THRESHOLD = 25
    ```

* 2 : CSV出力
  CSV出力の場合は-oオプションに2を指定して下さい。同一ディレクトリにチェック結果をCSV形式で出力します。
  ファイル名は'resultAtdCheck_m{mode}_{yyyymmdd-yyyymmdd(対象期間)}_{yyyymmdd-hhmmss(実行日時)}.csv'となります。

  ```batch
  python atdSystemCheck.py -m 1 -o 2
  -> 'resultAtdCheck_m02_20190801-20190831_20190814-090042.csv'
  ```

#### -d --date

yyyymmddの日付つきでdオプションを指定すると、指定日に実行した仮定でコマンドが実行されます。前月分の状況確認等に。

```batch
python atdSystemCheck.py -m 1 -o 1 -d 20190401
```

#### -e --exholiday

eオプションを指定すると土日祝日は実行しなくなります。
eオプションを指定する場合は以下からDL可能な内閣府配布の祝日リストを同ディレクトリに格納している事が前提となります。
<https://www8.cao.go.jp/chosei/shukujitsu/gaiyou.html#syukujitu>

```batch
python atdSystemCheck.py -m 2 -o 1 -e
```

#### -c --cmpcodefilter

社員番号付きで指定すると、指定した社員のみチェックされます。ブランク区切りで複数指定可能。mode1,3のみ効果があり、mode2は指定しても挙動は変わりません。

```batch
python atdSystemCheck.py -m 1 -o 1 -c 111 286
```

## 著作者

[koazuma](https://github.com/koazuma)
