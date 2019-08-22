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
# 直近のチェック対象日取得
####################################
def getRecentTargetDate(targetDate):
    """
    Overview:
        直近のチェック対象日取得。
        基本は前日だが、休日除外設定の場合は直近の営業日。
    Args:
        targetDate (date): チェック日
    Returns:
        rtd(date): 直近チェック対象日
    """
    rtd = targetDate - relativedelta(days=1)
    if args.exholiday and isHoliday(rtd, os.path.join(parentdir, 'syukujitsu.csv')):
        rtd = getRecentTargetDate(rtd)
    return rtd

####################################
# 期間検索用関数
####################################
def getSpan(targetDate, type):
    """
    Overview:
        対象期間(開始日/終了日)を取得するための関数
    Args:
        targetDate (datetime): 対象期間取得のための基幹日
        type (int): 1:20日締め月間(残業時間等), 2:月末締め月間 のいずれか
    Returns:
        fromDate(relativedelta), toDate(relativedelta): 開始日,終了日
    """
    logger.info("START function getSpan targetDate:" + targetDate.strftime('%Y%m%d') + " type:" + str(type))

    try:
        # 締め期間ルール設定
        if type == 1:
            FROM_DAY = 21
        elif type == 2 or type == 3:
            FROM_DAY = 1
        else:
            logger.error("function getSpan: input type error. type:" + type)
            raise ValueError

        # 基準日を取得
        baseDate = getRecentTargetDate(targetDate)
        # 当月開始日より前
        if baseDate.day < FROM_DAY:
            toDate = date(baseDate.year, baseDate.month, FROM_DAY) - relativedelta(days=1)
            fromDate = toDate - relativedelta(months=1) + relativedelta(days=1)
        # 当月開始日以降
        else :
            fromDate = date(baseDate.year, baseDate.month, FROM_DAY)
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
                        logger.warning('該当者不在のためスキップ - 対象期間変更 氏名:'+name+' 対象日:'+curdate.strftime("%Y%m%d"))
                        continue
                    

                # 対象日の指定列のデータを取得
                try:
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
                except exceptions.NoSuchElementException as e:
                        logger.warning('該当者不在のためスキップ - 対象日データ取得 氏名:'+name+' 対象日:'+curdate.strftime("%Y%m%d"))
                # 対象日のデータ取得を終えたらインクリメントして翌日へ
                curdate += relativedelta(days=1)
            # reletivedelta型をHH:MM型の文字列に変換
            for key in wt.keys():
                if type(wt[key]) is relativedelta:
                    wt[key] = '{hour:02}:{min:02}'.format(hour=wt[key].hours+wt[key].days*24,min=wt[key].minutes)
            # 出勤日数をstring型に変換
            wt[WORKDAYS] = str(wt[WORKDAYS])
            # 集計結果を追加
            rets.append(wt)
            
        except exceptions.UnexpectedAlertPresentException as e:
            logger.warning('該当者不在のためスキップ - 対象者変更 氏名:'+name)
            continue

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
        if mail_body == '':
            # タイトル行出力
            mail_body = '\t'.join(ret.keys())
        # チェック結果出力
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
            if cmpcode in members:
                # TO設定
                selfmail = members[cmpcode]['mail']
                if selfmail not in mail_to:
                    mail_to.append(selfmail)
                # CC設定
                mail_cc = addBossMailRecursive(cmpcode, mail_cc, members, levels)
            else:
                # memers.jsonにいない場合
                logger.error('Not found cmpcode: ' + cmpcode + ' in members.')
            
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
# 休日判定
####################################
def isHoliday(targetDate, syukujitsuPath):
    """
    Overview
        祝休日かどうか判定する
    Args
        targetDate(date): 判定対象日
        syukujitsuPath(string): 国土交通省発行のsyukujitsu.csvの格納パス
    Return
        retTrue: 祝休日 False: 非休日
    Raises:
        TypeError: targetDateがdate型でない場合に発生
    """
    try:
        ret = False
        if targetDate.weekday() == 5 or targetDate.weekday() == 6:
            ret = True
        else:
            from japan_holiday import JapanHoliday
            jpholiday = JapanHoliday(path=syukujitsuPath)
            if jpholiday.is_holiday(targetDate.strftime('%Y-%m-%d')):
                ret = True
    
    except TypeError as e :
        logger.error("function isHoliday: " + str(e))
        raise (e)
    else:
        return ret

####################################
# 工数配分入力 - 個人選択
####################################
def selectMember(id):
    """
    Overview
        工数配分入力結果で個人選択する
    Args
        社員番号
    Return
        なし
    """
    logger.info('START function selectMember id:' + id)
    try:
        # フレーム指定
        driver.switch_to.parent_frame()
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])
        # 個人選択ボタンクリック
        driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[4]/input').click()
        
        # サブウインドウにフォーカス移動
        wh = driver.window_handles
        driver.switch_to.window(wh[1])
        # フレーム指定
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])

        # フレーム内描画待ち(タイムアウトが多いため追加)
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.NAME, 'lstSelemp'))
            )
        except exceptions.TimeoutException as e:
            logger.error('画面表示タイムアウトエラー')
            logger.error(e)
            sys.exit()
    
        # selectインスタンス作成
        memberSelect = Select(driver.find_element_by_name('lstSelemp'))
        # 指定のvalue値のoptionを選択
        memberSelect.select_by_value(id)
        # 確定ボタンクリック
        driver.find_element_by_id('buttonKAKUTEI').click()

        # メインウィンドウにフォーカス移動
        driver.switch_to.window(wh[0])

        # フレーム指定
        driver.switch_to.parent_frame()
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])

    except Exception as e:
        logger.error(e)
        raise

####################################
# 工数配分入力結果取得
####################################
def checkManHourRegist():
    """
    Overview
        工数配分入力結果での就業時間と合計が不一致な日をチェックする
    Args
        なし
    Return
        rets: 工数配分入力結果での社員番号、氏名、時間不一致日の連想配列
    """
    logger.info('START function checkManHourRegist')
    try:
        # フレーム指定
        driver.switch_to.parent_frame()
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])
        # 個人選択ボタンクリック
        driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[4]/input').click()
        
        # サブウインドウにフォーカス移動
        wh = driver.window_handles
        driver.switch_to.window(wh[1])
        # フレーム指定
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])

        # フレーム内描画待ち(タイムアウトが多いため追加)
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.NAME, 'lstSelemp'))
            )
        except exceptions.TimeoutException as e:
            logger.error('画面表示タイムアウトエラー')
            logger.error(e)
            raise
    
        # selectインスタンス作成、メンバーリスト取得
        ids = []
        memberSelect = Select(driver.find_element_by_name('lstSelemp'))
        members = memberSelect.options
        for member in members:
            ids.append(member.get_attribute("value"))
        memberSelect.select_by_index(0)
        # 確定ボタンクリック
        driver.find_element_by_id('buttonKAKUTEI').click()
        
        # メインウィンドウにフォーカス移動
        driver.switch_to.window(wh[0])

        # フレーム指定
        driver.switch_to.parent_frame()
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])

        # 初期化
        rets = []

        # メンバー指定
        for id in ids:
            selectMember(id)

            # 対象期間のスタート区間まで戻る
            nowTerm = driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[6]').text.split(' ')
            nowTermStart = datetime.strptime(nowTerm[0],'%Y/%m/%d')
            while startdate < datetime.date(nowTermStart):
                driver.find_element_by_name('PrevEmpCode').click()
                # フレーム指定
                driver.switch_to.parent_frame()
                frames = driver.find_elements_by_xpath("//frame")
                driver.switch_to.frame(frames[1])
                # 現在表示中の開始日を取得
                nowTerm = driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[6]').text.split(' ')
                nowTermStart = datetime.strptime(nowTerm[0],'%Y/%m/%d')
            while enddate >= datetime.date(nowTermStart):
                # 工数登録誤りチェック(対象期間終了まで繰り返し)
                # テーブル要素の構成「/table/tbody/tr[X]/td[Y]」が以下ルールになっている
                # X...1:日付, 8:就業時間, 16:合計
                # Y...3:1日目, 4:2日目,...,9:7日目
                for i in range(3,10):
                    ret = {}
                    wt = driver.find_element_by_xpath('//*[@id="xyw4100_form"]/table/tbody/tr[8]/td[' + str(i) +']/font').text
                    total = driver.find_element_by_xpath('//*[@id="xyw4100_form"]/table/tbody/tr[16]/td[' + str(i) +']/font').text
                    if wt != total:
                        ret['氏名'] = driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[3]').text
                        ret['社員番号'] = id
                        ret['日付'] = driver.find_element_by_xpath('//*[@id="xyw4100_form"]/table/tbody/tr[1]/td[' + str(i) +']').text
                        ret['就業時間'] = wt
                        ret['合計'] = total
                        rets.append(ret)
                        logging.info(ret)
                # 次期間に移動
                driver.find_element_by_name('NextEmpCode').click()
                # フレーム指定
                driver.switch_to.parent_frame()
                frames = driver.find_elements_by_xpath("//frame")
                driver.switch_to.frame(frames[1])
                # 現在表示中の開始日を取得
                nowTerm = driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[6]').text.split(' ')
                nowTermStart = datetime.strptime(nowTerm[0],'%Y/%m/%d')

        return(rets)

    except Exception as e:
        logger.error(e)
        raise

####################################
# main
####################################
# コマンドライン引数定義
argparser = argparse.ArgumentParser()
argparser.add_argument('-m', '--mode', type=int, choices=[1,2,3], help='チェック種別 1:残業時間 2:打ち忘れ 3:工数登録', required=True)
argparser.add_argument('-o', '--output', type=int, choices=[1,2], help='出力タイプ 1:メール送信 2:CSVファイル出力', required=True)
argparser.add_argument('-d', '--date', type=lambda s: datetime.strptime(s, '%Y%m%d'), help='yyyymmdd形式で日を指定すると、その日に実行した仮定で実行される。')
argparser.add_argument('-e', '--exholiday', action='store_true', help='土日祝日の場合はチェックをしない。')

# 引数パース
args = argparser.parse_args()

# チェック種別
mode = args.mode
# 指定日付(未指定時は当日)
if args.date:
    nowDate = date(args.date.year, args.date.month, args.date.day)
else:
    nowDate = date.today()

# 休日未実行設定の場合は休日判定
if args.exholiday and isHoliday(nowDate, os.path.join(parentdir, 'syukujitsu.csv')):
    logger.info('End halfway, because of holiday.')
    sys.exit()

# config読み込み
try:
    config = configparser.ConfigParser()
    config.read(CONFIGFILE, 'UTF-8')
except Exception as e:
    logger.error('configファイル"'+CONFIGFILE+'"が見つかりません。')
    logger.error(e)
    sys.exit()

# 開始日、終了日を取得
startdate,enddate = getSpan(nowDate,mode)
# 強制設定用
#startdate = date(2018,11,1)
#enddate = date(2019,8,31)
logger.info("collectionTerm: "+str(startdate)+" - "+str(enddate))

# 結果CSVファイル名セット
CSVNAME = 'resultAtdCheck_'+'m{:02}'.format(mode) + startdate.strftime('_%Y%m%d') + enddate.strftime('-%Y%m%d') \
    + datetime.now().strftime('_%Y%m%d-%H%M%S') + ".csv"
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
# チェック結果取得
####################################
# mode別チェック結果取得
try:
    menuClick(config.get('modeinfo_'+str(mode), 'CLICKMENU'))
except Exception as e:
    logger.error('メニュークリック失敗')
    logger.error(str(e))
    sys.exit()

# 残業時間取得
if mode == 1:
    retsOver = getOverWork()
    # メール送信の場合は合計時間が閾値以上のもののみリストに追加    
    if args.output == 1:
        OVERWORK_THRESHOLD = config.getint('modeinfo_'+str(mode), 'OVERWORK_THRESHOLD')
    else:
        OVERWORK_THRESHOLD = 0
    rets = []
    for ret in retsOver:
        if int(ret['残業合計'].split(":")[0]) >= OVERWORK_THRESHOLD:
            logger.info(ret)
            rets.append(ret)

# 打ち忘れチェックリスト取得
elif mode == 2:
    rets = checkStampMiss()

# 工数登録結果取得
elif mode == 3:
    rets = checkManHourRegist()

# 結果が1件以上あったらアウトプット
if len(rets) > 0:
    # メール送信
    if args.output == 1:
        sendResultMail(rets,
        config.get('modeinfo_'+str(mode), 'MAILTITLE')+' '+startdate.strftime('%m/%d-')+enddate.strftime('%m/%d'),
        config.get('modeinfo_'+str(mode), 'MAILBODY')+'\n\n',
        #[CSVNAME],
        False,
        config.get('modeinfo_'+str(mode), 'MAIL_ESC_LEVEL'))
    # CSVファイル出力
    elif args.output == 2:
        csvOutput(rets,CSVNAME)

# 終了処理
driver.close()
logger.info("---- COMPLETE "+__file__+"  ----")
exit()