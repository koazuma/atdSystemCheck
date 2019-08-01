from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from selenium.common import exceptions
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
import logging
import sys
import os
import csv
import configparser
import traceback
import json
import argparse

# scriptフォルダ
parentdir = os.path.dirname(__file__)
# configファイル名
CONFIGFILE = os.path.join(parentdir, 'setting.ini')
# 社員リスト
EMPLOYEE_LIST = os.path.join(parentdir, 'members.json')
# 残業時間閾値(H)
OVERWORK_THRESHOLD = 25
# メール通知エスカレーションレベル(上限なし:-1 本人まで:0 直上長まで:1 ...)
# 残業時間チェック
MAIL_ESC_OVERWORK = -1
# 打ち忘れチェック
MAIL_ESC_STAMPMISS = 1

# log設定
logging.basicConfig(level=logging.INFO,
    format="%(asctime)s %(process)d %(name)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

# ログファイル出力用
handler = logging.FileHandler(filename=__file__ + ".log")
handler.setLevel(logging.INFO)
handler.setFormatter(logging.Formatter("%(asctime)s %(process)d %(name)s %(levelname)s %(message)s"))
logger.addHandler(handler)

logger.info("---- START "+__file__+" ----")

####################################
# 期間検索用関数
####################################
def getSpan(nowDate, type):
    """
    Overview:
        対象期間(開始日/終了日)を取得するための関数
    Args:
        nowDate (datetime): 対象期間取得のための基幹日
        type (int): 1:20日締め月間(残業時間等), 2:月末締め月間 のいずれか
    Returns:
        fromDate(relativedelta), toDate(relativedelta): 開始日,終了日
    """
    logger.info("START function getSpan nowDate:" + nowDate.strftime('%Y%m%d') + " type:" + str(type))

    try:
        # 締め期間ルール設定
        if type == 1:
            FROM_DAY = 21
        elif type == 2:
            FROM_DAY = 1
        else:
            logger.error("function getSpan: input type is nether 1 nor 2")
            raise ValueError

        # 当月開始日より前
        if nowDate.day < FROM_DAY:
            toDate = date(nowDate.year, nowDate.month, FROM_DAY) - relativedelta(days=1)
            fromDate = toDate - relativedelta(months=1) + relativedelta(days=1)
        # 当月開始日以降
        else :
            fromDate = date(nowDate.year, nowDate.month, FROM_DAY)
            toDate = fromDate + relativedelta(months=1) - relativedelta(days=1)

    except ValueError as e :
        logger.error("function getSpan: " + str(e))
        raise (e)
    
    return fromDate, toDate

####################################
# メニュークリック関数
####################################
def menuClick(titleStr):
    """
    Overvew:
        左メニューより指定のタイトルリンクをクリックする
    Args:
        titleStr (string):titleタグの文字列
    Returns:
        None
    Raises:
        NoSuchElementException: 親windowにハンドルがある状態で使用しないと発生
    """
    logger.info('START function menuClick titleStr: ' + titleStr)
    try:
        # フレーム指定
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[0])

        # リンククリック
        driver.find_element_by_xpath("//a[@title='" + titleStr + "']").click()

    except exceptions.NoSuchElementException as e :
        logger.error("function menuClick: " + str(e))
        raise (e)
    else:
        return True

####################################
# 就業週報月報画面 : 期間検索
####################################
def getOverWork():
    #----- 就業週報月報のtdタグidルール -----
    # 各tdタグは以下のような命名規則になっている
    # grdXyw1500g-rc-{日付-1}-{列ID}
    DAILYID_F = "grdXyw1500g-rc-" # 日毎のid先頭部分
    # 列ID(列数-1)
    itemids = {
        "法定外勤":"13",
        "深夜残業":"15",
        "休日勤務":"16",
        "休日深夜":"17"
    }
    NAME = "氏名"
    CMPID = "社員番号"
    TOTALTIME = "残業合計"
    WORKDAYS = "出勤日数"
    FDATE_ID = "grdXyw1500g-rc-0-0" # 1日目のid
    EMPTY_MARK = "----" # 稼働時間ゼロ表示
    #----------------------------------------
    logger.info('START function getOverWork')
    # フレーム指定
    driver.switch_to.parent_frame()
    frames = driver.find_elements_by_xpath("//frame")
    driver.switch_to.frame(frames[1])

    # 表示月、取得データを初期化
    dispmonth = 0
    rets = []

    while True:

        try:
            # 氏名、社員番号取得
            name = driver.find_element_by_xpath("//*[@id='formshow']/table/tbody/tr[4]/td/table/tbody/tr/td[7]").text
            cmpid = driver.find_element_by_xpath("//*[@id='formshow']/table/tbody/tr[4]/td/table/tbody/tr/td[6]").text

            # 指定日、表示月、稼働時間を初期化
            curdate = startdate
            wt = {}
            wt[NAME] = name
            wt[CMPID] = cmpid
            for v in itemids.keys():
                wt[v] = relativedelta()
            wt[TOTALTIME] = relativedelta()
            wt[WORKDAYS] = 0

            # 終了日まで指定日をインクリメントしながらデータ取得
            while curdate <= enddate:
                # 月を指定して月報を表示(初回および月が変わった時のみ)
                if curdate.month != dispmonth:
                    dtElm = driver.find_element_by_id("CmbYM")
                    Select(dtElm).select_by_value(curdate.strftime("%Y%m"))
                    driver.find_element_by_name("srchbutton").click()
                    dispmonth = curdate.month
                    
                    try:
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_all_elements_located((By.XPATH, "//td[@id='" + FDATE_ID + "']"))
                        )
                    except exceptions.TimeoutException as e:
                        logger.error('画面表示タイムアウトエラー')
                        logger.error(e)
                        raise(e)
                    except exceptions.UnexpectedAlertPresentException as e:
                        logger.warning('該当者不在のためスキップ')
                        #logger.warning(e)
                        Alert(driver).accept()
                        break
                    

                # 対象日の指定列のデータを取得
                for key in itemids.keys():
                    # 対象日のtdタグのidを作成
                    tgtid = DAILYID_F + str(int(curdate.strftime("%d")) -1) + "-" + itemids[key]
                    elm = driver.find_element_by_xpath("//td[@id='"+tgtid+"']")
                    workTime = elm.get_attribute("DefaultValue")
                    if workTime == EMPTY_MARK:
                        workTime = "00:00"
                    else:
                        wt[WORKDAYS] += 1
                    wt[key] += relativedelta(hours=int(workTime.split(":")[0]),minutes=int(workTime.split(":")[1]))
                    wt[TOTALTIME] += relativedelta(hours=int(workTime.split(":")[0]),minutes=int(workTime.split(":")[1]))

                # 対象日のデータ取得を終えたらインクリメントして翌日へ
                curdate += relativedelta(days=1)
            # reletivedelta型をHH:MM型の文字列に変換
            for key in wt.keys():
                if type(wt[key]) is relativedelta:
                    wt[key] = '{hour:02}:{min:02}'.format(hour=wt[key].hours+wt[key].days*24,min=wt[key].minutes)
            # 出勤日数をstring型に変換
            wt[WORKDAYS] = str(wt[WORKDAYS])
            # 合計時間が閾値より大きい場合、対象社員の結果をリストに保存
            if int(wt[TOTALTIME].split(":")[0]) >= OVERWORK_THRESHOLD:
                rets.append(wt)
                logger.info(wt)

        except exceptions.UnexpectedAlertPresentException as e:
            logger.warning('該当者不在のためスキップ')
            Alert(driver).accept()

        ####################################
        # 対象社員を変更
        ####################################
        # 次が選択可なら次社員を選択
        tgtname = "button4"
        if driver.find_element_by_name(tgtname).is_enabled() :
            driver.find_element_by_name(tgtname).click()
        # 次が選択不可ならループ終了
        else :
            break
    return rets

####################################
# 結果CSV出力
####################################
def csvOutput(rets,fpath):
    """
    Overview
        連想配列をcsvファイルに出力する。
    Args
        rets: 出力対象の連想配列
        fpath: 出力ファイルパス
    Return
        なし
    """
    logger.info('START function csvOutput')
    # ファイルオープン
    with open(fpath, 'w', newline='', encoding='utf_8_sig') as fp:
        writer = csv.writer(fp, lineterminator='\r\n')
        headerFlg = 0

        # データ行出力
        for ret in rets:
            if headerFlg == 0:
                writer.writerow(ret.keys())
                headerFlg = 1
            writer.writerow(ret.values())
        logger.info("Output '"+fpath+"'.")

####################################
# 打ち忘れチェックリスト取得
####################################
def checkStampMiss():
    """
    Overview
        打ち忘れチェックリスト取得
    Args
        なし
    Return
        打ち忘れチェックリストの社員番号、氏名、対象日、メッセージ項目1の連想配列
    """
    logger.info('START function checkStampMiss')
    try:
        # フレーム指定
        driver.switch_to.parent_frame()
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])
        # 個人選択ボタンクリック
        driver.find_element_by_xpath('//*[@id="Xyw1120g_form"]/table/tbody/tr[2]/td/table[2]/tbody/tr/td/table/tbody/tr/td[2]/input').click()
        
        # サブウインドウにフォーカス移動
        wh = driver.window_handles
        driver.switch_to.window(wh[1])
        logger.info('switch target window - '+str(wh[1]))
        # フレーム指定
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])

        # フレーム内描画待ち(タイムアウトが多いため追加)
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.ID, 'AllSel'))
            )
        except exceptions.TimeoutException as e:
            logger.error('画面表示タイムアウトエラー')
            logger.error(e)
            sys.exit()
    
        # 全選択ラジオボタンクリック
        driver.find_element_by_id('AllSel').click()
        # 確定ボタンクリック
        driver.find_element_by_id('buttonKAKUTEI').click()
        
        # メインウィンドウにフォーカス移動
        driver.switch_to.window(wh[0])
        logger.info('switch target window - '+str(wh[0]))

        # フレーム指定
        driver.switch_to.parent_frame()
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])

        # 期間指定
        startYMD = driver.find_element_by_name('StartYMD')
        endYMD = driver.find_element_by_name('EndYMD')
        if startYMD.get_attribute('value') != startdate.strftime('%Y%m%d'):
            startYMD.clear()
            startYMD.send_keys(startdate.strftime('%Y%m%d'))
            endYMD.clear()
            endYMD.send_keys(enddate.strftime('%Y%m%d'))
            driver.find_element_by_name('srchbutton').click()
        
        # テーブルのデータ取得
        # テーブル要素のid構成を利用して取得
        # テーブル行数 : row(１行目:0, 2行目:1, ...)
        # 行自体          : grdXyw1120G-r-{row}
        # 社員番号        : grdXyw1120G-rc-{row}-0
        # 氏名            : grdXyw1120G-rc-{row}-1
        # 対象日          : grdXyw1120G-rc-{row}-2
        # メッセージ項目1  : grdXyw1120G-rc-{row}-4
        TRIDBASE = 'grdXyw1120G-r-'
        TDIDBASE = 'grdXyw1120G-rc-'
        itemids = {
            "社員番号":"0",
            "氏名":"1",
            "対象日":"2",
            "メッセージ項目1":"4"
        }

        # 初期化
        row = 0
        rets = []

        while True:
            # 初期化
            ret = {}
            # 対象行の有無チェック
            rowid = TRIDBASE + str(row)
            try:
                driver.find_element_by_id(rowid)
            except exceptions.NoSuchElementException as e :
                logging.info('最終行到達')
                break
            except Exception as e :
                logging.error('想定外の例外エラー発生')
                logging.error(e)
                raise

            # 要素取得
            for key in itemids.keys():
                targetid = TDIDBASE + str(row) + '-' + itemids[key]
                ret[key] = driver.find_element_by_id(targetid).text

            # 次行にインクリメント
            logging.info(ret)
            rets.append(ret)
            row += 1

        return(rets)

    except Exception as e:
        logger.error(e)
        raise

####################################
# 上長のメールアドレス追加
####################################
def addBossMailRecursive(id, mail, members, levels, depth=0):
    """
    Overview
        指定idの社員番号の上司、さらにその上司、、、と配列に追加
    Args
        id(string): 社員番号
        mail(array): メールアドレス
        members(dictionary): 社員情報
        levels(int): エスカレーション階層(-1:無限)
        depth(int): 現在エスカレーション階層
    Return
        mail(array): メールアドレス配列 
    """
    if levels != depth:
        bossids = members[id]['boss'].split(',')
        for bossid in bossids:
            if bossid in members:
                bossmail = members[bossid]['mail']
                if bossmail not in mail:
                    mail.append(bossmail)
                if members[bossid]['boss'] != "":
                    addBossMailRecursive(bossid, mail, members, levels, depth +1)
            else:
                # memers.jsonにいない場合
                logger.error('Not found cmpcode: ' + cmpcode + ' in members.')
    return mail

####################################
# メール送信
####################################
def sendResultMail(rets, mailsub, mailstr, attaches, levels=-1):
    """
    Overview
        結果をメールで送信する
    Args
        rets: 社員番号キーを持つ辞書型オブジェクトのリスト
        mailsub: メール件名
        mailstr: メール本文の先頭文章。
        attaches: 添付ファイルパス
        levels: メール通知のエスカレーションレベル(-1は無限(デフォルト))
    Return
        なし
    """
    from email import message
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email.mime.multipart import MIMEMultipart
    from os.path import basename
    from email.header import Header
    from email.utils import formatdate
    from email import encoders

    logger.info('START function sendResultMail')
    # Gmail SMTP設定
    smtp_host = config.get('mail', 'smtp_host')
    smtp_port = config.get('mail', 'smtp_port')
    username = config.get('mail', 'username')
    password = config.get('mail', 'password')

    from_addr = 'noreply@gmail.com'

    # body初期化
    mail_body = ''

    for ret in rets:
        # タイトル行出力
        if mail_body == '':
            mail_body = '\t'.join(ret.keys())
        mail_body = mail_body + '\r\n' + '\t'.join(ret.values())
    mail_body = mailstr + mail_body

    # メンバーのjsonファイル読み込み
    with open(EMPLOYEE_LIST,'r',encoding='utf-8') as f:
        members = json.load(f)

        # 宛先設定
        mail_to = []
        mail_cc = []
        for ret in rets:
            cmpcode = str(int(ret['社員番号']))
            # TO設定
            if cmpcode in members:
                selfmail = members[cmpcode]['mail']
                if selfmail not in mail_to:
                    mail_to.append(selfmail)
            else:
                # memers.jsonにいない場合
                logger.error('Not found cmpcode: ' + cmpcode + ' in members.')
            # CC設定
            mail_cc = addBossMailRecursive(cmpcode, mail_cc, members, levels)

    # メールの内容を作成
    if attaches:
        mime = MIMEMultipart()
        mime.attach(MIMEText(_text=mail_body, _subtype='plain', _charset='utf-8'))
    else:
        mime = MIMEText(_text=mail_body, _subtype='plain', _charset='utf-8')

    mime['Subject'] = Header(mailsub, 'utf-8')
    mime['From'] = from_addr
    mime['To'] = ','.join(mail_to)
    mime['Cc'] = ','.join(mail_cc)
    mime['Date'] = formatdate()
    
    # ファイル添付
    if attaches:
        for attach in attaches:
            logger.info('attach file name {}'.format(attach))
            attachment = MIMEBase('application', 'csv')
            with open(attach, mode='rb') as f:
                attachment.set_payload(f.read())
            encoders.encode_base64(attachment)
            attachment.add_header("Content-Disposition","attachment", filename=basename(attach))
            mime.attach(attachment)

    # メールサーバー認証
    smtp = smtplib.SMTP(smtp_host, smtp_port)
    smtp.ehlo()
    smtp.starttls()
    smtp.ehlo()
    smtp.login(username, password)
    # メール送信
    smtp.sendmail(from_addr, mail_to + mail_cc, mime.as_string())
    smtp.quit()

####################################
# main
####################################
# コマンドライン引数定義
argparser = argparse.ArgumentParser()
argparser.add_argument('-m', '--mode', type=int, choices=[1,2], help='実行モード 1:残業時間チェック 2:打ち忘れチェック', required=True)
argparser.add_argument('-d', '--date', type=lambda s: datetime.strptime(s, '%Y%m%d'), help='yyyymmdd形式で日を指定すると、その日に実行した仮定で実行される。')
argparser.add_argument('-e', '--exholiday', action='store_true', help='土日祝日の場合はチェックをしない。')

# 引数パース
args = argparser.parse_args()

# 実行モード
mode = args.mode
# 指定日付(未指定時は前日)
if args.date:
    nowDate = date(args.date.year, args.date.month, args.date.day)
else:
    nowDate = date.today() - relativedelta(days=1)

# 休日未実行設定の場合は休日判定
if args.exholiday:
    from japan_holiday import JapanHoliday
    jpholiday = JapanHoliday(path=os.path.join(parentdir, 'syukujitsu.csv'))
    if jpholiday.is_holiday(date.today().strftime('%Y-%m-%d')):
        logger.info('End halfway, because of holiday.')
        sys.exit()

# config読み込み
try:
    config = configparser.ConfigParser()
    config.read(CONFIGFILE)
except Exception as e:
    logger.error('configファイル"'+CONFIGFILE+'"が見つかりません。')
    logger.error(e)
    sys.exit()

# 開始日、終了日を取得
startdate,enddate = getSpan(nowDate,mode)
logger.info("collectionTerm: "+str(startdate)+" - "+str(enddate))

# 結果CSVファイル名セット
CSVNAME = "OverWork"+startdate.strftime('_F%Y%m%d')+enddate.strftime('-T%Y%m%d') \
    +datetime.now().strftime('_@%Y%m%d-%H%M%S') +".csv"
CSVNAME = os.path.join(parentdir, CSVNAME)

# webDriver起動
try:
    chromedriver_path = config.get('environment', 'chromedriver')
    driver = webdriver.Chrome(chromedriver_path)
except Exception as e:
    logger.error("実行可能なWebDriver'"+chromedriver_path+"'が見つかりません。")
    logger.error(e)
    sys.exit()

####################################
# ログイン認証
####################################
logger.info('START login')
driver.get(config.get('siteinfo', 'url'))

# メンテナンス中
if driver.title == 'sorry page':
    logger.error('サーバメンテナンス中により処理中止')
    sys.exit()

# 表示待ち
try:
    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.NAME, "DataSource"))
    )
except exceptions.TimeoutException as e:
    logger.error('画面表示タイムアウトエラー')
    logger.error(e)
    sys.exit()

# ID/PW入力
driver.find_element_by_name("DataSource").send_keys(config.get('siteinfo', 'cp'))
driver.find_element_by_name("LoginID").send_keys(config.get('siteinfo', 'id'))
driver.find_element_by_name("PassWord").send_keys(config.get('siteinfo', 'pw'))

try:
    # ログインボタンクリック
    driver.find_element_by_name("LOGINBUTTON").click()
    driver.find_element_by_tag_name("title")
except exceptions.NoSuchElementException as e:
    logger.error('ログインに失敗しました。cp:'+config.get('siteinfo', 'cp')+' id:'+config.get('siteinfo', 'id')+' pw:'+config.get('siteinfo', 'pw'))
    logger.error(e)
    sys.exit()

####################################
# ホーム画面 : 指定画面へ遷移
####################################
if mode == 1:
    try:
        menuClick("就業週報月報")
    except Exception as e:
        logger.error('メニュークリック失敗')
        logger.error(str(e))
        sys.exit()
    # 残業時間取得
    rets = getOverWork()
    # 結果をメール送信
    if len(rets) > 0:
        #csvOutput(rets,CSVNAME)
        sendResultMail(rets,
            '残業時間チェック結果のお知らせ '+startdate.strftime('%m/%d-')+enddate.strftime('%m/%d'),
            '残業時間'+str(OVERWORK_THRESHOLD)+'H超過対象者およびその上長への通知です。\n対象者は45Hを超えないよう、計画的に稼働してください。\n\n',
            #[CSVNAME],
            False,
            MAIL_ESC_OVERWORK)

elif mode == 2:
    try:
        menuClick("打ち忘れﾁｪｯｸﾘｽﾄ")
    except Exception as e:
        logger.error('メニュークリック失敗')
        logger.error(str(e))
        sys.exit()
    # 打ち忘れチェックリスト結果取得
    rets = checkStampMiss()
    # 結果をメール送信
    if len(rets) > 0:
        sendResultMail(rets,
            '打ち忘れチェックリスト確認結果のお知らせ '+startdate.strftime('%m/%d-')+enddate.strftime('%m/%d'),
            '打ち忘れチェックリスト確認結果を連絡します。\n対象者は速やかに必要な申請を行ってください。\n\n',
            False,
            MAIL_ESC_STAMPMISS)

# 終了処理
driver.close()
logger.info("---- COMPLETE "+__file__+"  ----")
exit()