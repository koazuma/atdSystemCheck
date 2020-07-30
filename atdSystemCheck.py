from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common import exceptions
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
from logging import getLogger, StreamHandler, FileHandler, Formatter, DEBUG, INFO, WARNING, ERROR, handlers
from selenium.webdriver.common.keys import Keys
import sys
import os
import csv
import configparser
import traceback
import json
import argparse
import inspect
import time

# scriptフォルダ
parentdir = os.path.dirname(__file__)
# configファイル名
CONFIGFILE = os.path.join(parentdir, 'setting.ini')
# 社員リスト
EMPLOYEE_LIST = os.path.join(parentdir, 'members.json')

# log設定
getLogger().setLevel(DEBUG)
logger = getLogger(__name__)

# ログファイル出力用
filehandler = handlers.RotatingFileHandler(filename=__file__ + '.log', maxBytes=1024*1024, backupCount=5)
filehandler.setLevel(DEBUG)
filehandler.setFormatter(Formatter("%(asctime)s %(process)d %(name)s %(levelname)s %(message)s"))
logger.addHandler(filehandler)

# コンソールログ出力用
streamhandler = StreamHandler()
streamhandler.setLevel(INFO)
streamhandler.setFormatter(Formatter("%(asctime)s %(process)d %(name)s %(levelname)s %(message)s"))
logger.addHandler(streamhandler)

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
    logger.info(str(getCurLineNo())+' START function targetDate:' + targetDate.strftime('%Y%m%d') + ' type:' + str(type))

    try:
        # 締め期間ルール設定
        if type == 1:
            FROM_DAY = 21
        elif type == 2 or type == 3:
            FROM_DAY = 1
        else:
            logger.error(str(getCurLineNo())+' input type error. type:' + type)
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

    except exceptions.ValueError as e :
        logger.error(str(getCurLineNo())+' '+ str(e))
        raise(e)
    
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
    logger.info(str(getCurLineNo())+' START function titleStr:' + titleStr)
    try:
        # フレーム指定
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[0])
        # リンククリック
        findElement('xpath', "//a[@title='" + titleStr + "']").click()
        # 表示待ち
        waitLocate()
        time.sleep(FORCESLEEPSEC)

    except (exceptions.NoSuchElementException, exceptions.TimeoutException) as e :
        raise(e)
    else:
        return True
####################################
# 要素の描画確認後取得
####################################
def findElement(method, target, state='locate'):
    """
    Overview
        指定要素を描画を待って取得する
    Args
        method(string): 要素確認方法(xpath/id/name)
        target(string): 対象の要素(id,xpath,name等)
        state(string): 確認内容(locate,click,select,text,frame等)
    Return
        対象の要素オブジェクト
    Raises:
        TimeoutException: 指定秒数待機しても要素が確認できなかった場合に発生
    """
    logger.debug(str(getCurLineNo())+' START function method:'+method+' target:'+target+' state:'+state)
    try:
        # state別に要素が確認できるまで待機
        waitDriver(method, target, state)
        # method別に要素取得
        if method == 'xpath':
            ret = driver.find_element_by_xpath(target)
        elif method == 'id':
            ret = driver.find_element_by_id(target)
        elif method == 'name':
            ret = driver.find_element_by_name(target)
        else:
            logger.error(str(getCurLineNo())+' 引数エラー method:'+method)
            raise ValueError

    except exceptions.UnexpectedAlertPresentException as e:
        logger.error(str(getCurLineNo())+' UnexpectedAlertPresentException発生 ' +str(e))
        raise(e)
    except Exception as e:
        logger.error(str(getCurLineNo())+' 例外エラー発生 ' +str(e))
        raise(e)
    else:
        return ret

####################################
# パターン別待機
####################################
def waitDriver(method, target, state):
    """
    Overview
        state別に指定要素の描画を待つ
    Args
        method(string): 要素確認方法(xpath/id/name)
        target(string): 対象の要素(id,xpath,name等)
        state(string): 確認内容(locate,click,select,frame等)
    Return
        なし
    Raises:
        TimeoutException: 指定秒数待機しても要素が確認できなかった場合に発生
    """
    bymethod = {
        'xpath' : By.XPATH,
        'id' : By.ID,
        'name' : By.NAME
    }
    try:
        if state == 'locate':
            WebDriverWait(driver, TIMEOUTSEC).until(EC.presence_of_element_located((bymethod[method],target)))
        elif state == 'click':
            WebDriverWait(driver, TIMEOUTSEC).until(EC.element_to_be_clickable((bymethod[method],target)))
        elif state == 'select':
            WebDriverWait(driver, TIMEOUTSEC).until(EC.element_to_be_selected((bymethod[method],target)))
        elif state == 'text':
            WebDriverWait(driver, TIMEOUTSEC).until(EC.text_to_be_present_in_element((bymethod[method],target)))
        elif state == 'vlocate':
            WebDriverWait(driver, TIMEOUTSEC).until(EC.visibility_of_element_located((bymethod[method],target)))
        elif state == 'frame':
            WebDriverWait(driver, TIMEOUTSEC).until(EC.frame_to_be_available_and_switch_to_it((bymethod[method],target)))
        else:
            logger.error(str(getCurLineNo())+' 引数エラー state:'+state)
            raise ValueError
    except exceptions.TimeoutException as e:
        logger.error(str(getCurLineNo())+' 画面表示タイムアウトエラー method:'+method+' target:'+target+' state:'+state)
        raise(e)
    except exceptions.UnexpectedAlertPresentException as e:
        logger.error(str(getCurLineNo())+' UnexpectedAlertPresentException発生 '+str(e))
        raise(e)
    except Exception as e:
        logger.error(str(getCurLineNo())+' 想定外の例外発生 '+str(e))
        raise(e)

####################################
# 要素描画待ち
####################################
def waitLocate():
    """
    Overview
        画面描画を待つ
    Args
        なし
    Return
        なし
    Raises:
        TimeoutException: 指定秒数待機しても要素が確認できなかった場合に発生
    """
    try:
        WebDriverWait(driver, TIMEOUTSEC).until(EC.presence_of_all_elements_located)
    except exceptions.TimeoutException as e:
        logger.error(str(getCurLineNo())+' 画面表示タイムアウトエラー')
        raise(e)

####################################
# メンバーリスト取得
####################################
def getMemberList(xp):
    """
    Overview
        サブウインドウからメンバー一覧を取得する
    Args
        xp: 個人選択のXPATH
    Return
        ids: 社員番号リスト
    """
    logger.info(str(getCurLineNo())+' START function')
    try:
        # フレーム指定
        driver.switch_to.parent_frame()
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])
        # 個人選択ボタンクリック
        findElement('xpath', xp).click()
        
        # サブウインドウにフォーカス移動
        wh = driver.window_handles
        driver.switch_to.window(wh[1])
        # フレーム指定
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])

        # selectインスタンス作成、メンバーリスト取得
        ids = []
        memberSelect = Select(findElement('name','lstSelemp'))
        members = memberSelect.options
        for member in members:
            # 対象者フィルタオプションなし、またはありでかつ指定されている場合、メンバーに追加
            if cmpcodefilter is None or int(member.get_attribute('value')) in cmpcodefilter:
                ids.append(member.get_attribute('value'))
        memberSelect.select_by_index(0)
        # 確定ボタンクリック
        findElement('id','buttonKAKUTEI').click()
        time.sleep(FORCESLEEPSEC)
        
        # メインウィンドウにフォーカス移動
        driver.switch_to.window(wh[0])

        # フレーム指定
        driver.switch_to.parent_frame()
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])

    except Exception as e:
        logger.error(str(getCurLineNo())+' 想定外の例外エラー発生')
        raise(e)

    return ids

####################################
# 就業週報月報画面 : 残業時間取得
####################################
def getOverWork():
    """
    Overview
        就業週報月報画面より残業時間を集計する
    Args
        なし
    Return
        rets: 氏名、社員番号、各種残業時間、合計残業時間の連想配列のリスト
    """
    logger.info(str(getCurLineNo())+' START function')
    try:
        # メンバー取得
        ids = getMemberList('/html/body/form/table/tbody/tr[4]/td/table/tbody/tr/td[8]/input')

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
        FDATE_ID = "grdXyw1500g-rc-0-0" # 1日目のid
        EMPTY_MARK = "----" # 稼働時間ゼロ表示
        #----------------------------------------

        # 表示月、取得データを初期化
        dispmonth = 0
        rets = []

        # メンバー指定
        for id in ids:
            try:
                selectMember(id, '/html/body/form/table/tbody/tr[4]/td/table/tbody/tr/td[8]/input')

                # 氏名、社員番号取得
                logger.debug(str(getCurLineNo())+' 氏名、社員番号取得')
                name = findElement('xpath',"//*[@id='formshow']/table/tbody/tr[4]/td/table/tbody/tr/td[7]").text
                cmpid = findElement('xpath',"//*[@id='formshow']/table/tbody/tr[4]/td/table/tbody/tr/td[6]").text
                logger.info(str(getCurLineNo())+' 氏名:'+name+' 社員番号:'+cmpid)

                # 指定日、表示月、稼働時間を初期化
                curdate = startdate
                wt = {}
                wt[NAME] = name
                wt[CMPID] = cmpid
                for v in itemids.keys():
                    wt[v] = relativedelta()
                wt[TOTALTIME] = relativedelta()

                # 終了日まで指定日をインクリメントしながらデータ取得
                while curdate <= enddate:
                    logger.debug(str(getCurLineNo())+' 対象日:'+str(curdate))
                    # 月を指定して月報を表示(初回および月が変わった時のみ)
                    if curdate.month != dispmonth:
                        try:
                            dtElm = findElement('id','CmbYM')
                            Select(dtElm).select_by_value(curdate.strftime('%Y%m'))
                            logger.debug(str(getCurLineNo())+' 指定月変更:'+curdate.strftime('%Y%m'))
                            findElement('name','srchbutton','click').click()
                            waitLocate()
                            # WebDriverWaitで例外エラーを止める事が出来なかったため、止む無くsleepを使用
                            time.sleep(FORCESLEEPSEC)
                            dispmonth = curdate.month
                        except exceptions.TimeoutException as e:
                            logger.error(str(getCurLineNo())+' 画面表示タイムアウトエラー')
                            raise(e)
                        except exceptions.UnexpectedAlertPresentException as e:
                            # 入社月の前月対策
                            logger.warning(str(getCurLineNo())+' 該当者不在のためスキップ - 対象期間変更 氏名:'+name+' 対象日:'+curdate.strftime("%Y%m%d"))
                            continue

                    # 対象日の指定列のデータを取得
                    try:
                        # 対象日を確認
                        tgtdateid = DAILYID_F + str(int(curdate.strftime('%d')) -1) + '-0'
                        tgtdate = findElement('xpath',"//td[@id='"+tgtdateid+"']").get_attribute("DefaultValue")
                        if tgtdate != curdate.strftime('%m/%d') :
                            logger.error(str(getCurLineNo())+' 対象日相違エラー。 tgtdate:'+tgtdate+' curdate:'+curdate.strftime('%m/%d'))
                            raise(ValueError)

                        for key in itemids.keys():
                            # 対象日のtdタグのidを作成
                            tgtid = DAILYID_F + str(int(curdate.strftime("%d")) -1) + "-" + itemids[key]
                            # 対象日の稼働時間を取得
                            workTime = findElement('xpath',"//td[@id='"+tgtid+"']").get_attribute("DefaultValue")
                            if workTime == EMPTY_MARK:
                                workTime = "00:00"
                            wt[key] += relativedelta(hours=int(workTime.split(":")[0]),minutes=int(workTime.split(":")[1]))
                            wt[TOTALTIME] += relativedelta(hours=int(workTime.split(":")[0]),minutes=int(workTime.split(":")[1]))
                    except (exceptions.NoSuchElementException, exceptions.TimeoutException) as e:
                            logger.warning(str(getCurLineNo())+' 該当者不在のためスキップ - 対象日データ取得 氏名:'+name+' 対象日:'+curdate.strftime("%Y%m%d"))
                    # 対象日のデータ取得を終えたらインクリメントして翌日へ
                    curdate += relativedelta(days=1)
                # reletivedelta型をHH:MM型の文字列に変換
                for key in wt.keys():
                    if type(wt[key]) is relativedelta:
                        wt[key] = '{hour:02}:{min:02}'.format(hour=wt[key].hours+wt[key].days*24,min=wt[key].minutes)
                # 集計結果を追加
                rets.append(wt)
                logger.info(str(getCurLineNo())+' 集計結果追加 '+str(wt))
            
            except exceptions.UnexpectedAlertPresentException as e:
                # 対象社員が退職後等で不在の場合スキップ
                name = findElement('xpath',"//*[@id='formshow']/table/tbody/tr[4]/td/table/tbody/tr/td[7]").text
                logger.warning(str(getCurLineNo())+' 該当者不在のためスキップ - 対象者変更 氏名:'+name)
                # continue

    except Exception as e:
        logger.error(str(getCurLineNo())+' 想定外の例外エラー発生')
        raise(e)

    return rets

####################################
# カレント行取得
####################################
def getCurLineNo(depth=0):
  frame = inspect.currentframe().f_back
  return os.path.basename(frame.f_code.co_filename), frame.f_code.co_name, frame.f_lineno

####################################
# 結果CSV出力
####################################
def csvOutput(rets,fpath):
    """
    Overview
        リストをcsvファイルに出力する。
    Args
        rets(list): 出力対象のリスト
        fpath(string): 出力ファイルパス
    Return
        なし
    """
    logger.info(str(getCurLineNo())+' START function')
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
        logger.info(str(getCurLineNo())+' Output '+fpath)

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
    logger.info(str(getCurLineNo())+' START function')
    try:
        # フレーム指定
        driver.switch_to.parent_frame()
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])
        # 個人選択ボタンクリック
        findElement('xpath', '//*[@id="Xyw1120g_form"]/table/tbody/tr[2]/td/table[2]/tbody/tr/td/table/tbody/tr/td[2]/input').click()
        waitLocate()
        time.sleep(FORCESLEEPSEC)
        
        # サブウインドウにフォーカス移動
        wh = driver.window_handles
        driver.switch_to.window(wh[1])
        # フレーム指定
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])

        # 全選択ラジオボタンクリック
        findElement('id', 'AllSel').click()
        # 確定ボタンクリック
        findElement('id', 'buttonKAKUTEI').click()
        time.sleep(FORCESLEEPSEC)
        
        # メインウィンドウにフォーカス移動
        driver.switch_to.window(wh[0])

        # フレーム指定
        driver.switch_to.parent_frame()
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])

        # 期間指定
        startYMD = findElement('name', 'StartYMD')
        logger.info(str(getCurLineNo())+' startYMD:'+startYMD.get_attribute('value')+' startdate:'+startdate.strftime('%Y%m%d'))
        if startYMD.get_attribute('value') != startdate.strftime('%Y%m%d'):
            time.sleep(FORCESLEEPSEC)
            startYMD = findElement('name', 'StartYMD')
            startYMD.clear()
            startYMD.send_keys(startdate.strftime('%Y%m%d'))
            endYMD = findElement('name', 'EndYMD')
            endYMD.clear()
            endYMD.send_keys(enddate.strftime('%Y%m%d'))
            findElement('name', 'srchbutton').click()
            waitLocate()
            time.sleep(FORCESLEEPSEC)
        
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
        # スクロール用action生成
        actions = ActionChains(driver)

        while True:
            # 初期化
            ret = {}
            # 対象行の有無チェック
            rowid = TRIDBASE + str(row)
            try:
                driver.find_element_by_id(rowid)
            except exceptions.NoSuchElementException as e :
                logger.info(str(getCurLineNo())+' 最終行到達')
                break
            except Exception as e :
                logger.error(str(getCurLineNo())+' 対象行確認-想定外の例外エラー発生')
                raise(e)

            # 要素取得
            for key in itemids.keys():
                targetid = TDIDBASE + str(row) + '-' + itemids[key]
                while True:
                    ret[key] = findElement('id',targetid).text
                    # 取得できていない場合はスクロール
                    if ret[key] == '':
                        logger.debug(str(getCurLineNo())+' 要素取得失敗のためスクロール key:'+key)
                        driver.find_element_by_tag_name('body').send_keys(Keys.PAGE_DOWN)
                        # 横方向の移動のため要素移動
                        actions.move_to_element(driver.find_element_by_id(targetid))
                        actions.perform()
                    else:
                        break

            # 次行にインクリメント
            logger.info(str(getCurLineNo())+' '+str(ret))
            rets.append(ret)
            row += 1

        return(rets)

    except Exception as e:
        logger.error(str(getCurLineNo())+' 想定外の例外エラー発生')
        raise(e)

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
    # エスカレーションレベルに達したら終了
    if levels != depth:
        # デリミタ(,)で分割して上司ID取得
        bossids = members[id]['boss'].split(',')
        for bossid in bossids:
            if bossid in members:
                bossmail = members[bossid]['mail']
                # 上司のメールアドレスを宛先に追加(既に存在する場合は除く)
                if bossmail not in mail:
                    mail.append(bossmail)
                # 更に上司がいる場合は再帰的にエスカレーション
                if members[bossid]['boss'] != "":
                    addBossMailRecursive(bossid, mail, members, levels, depth +1)
            else:
                # memers.jsonにいない場合
                logger.error(str(getCurLineNo())+' Not found cmpcode:' + cmpcode + ' in members.')
    return mail

####################################
# 無視メンバー削除
####################################
def deleteIgnoreMember(rets):
    """
    Overview
        無視設定メンバーを削除する。
    Args
        rets(list): 社員番号キーを持つ辞書型オブジェクトのリスト
    Return
        newRets(list): 無視設定メンバーを除いたリスト
    """
    newRets = []
    # メンバーのjsonファイル読み込み
    with open(EMPLOYEE_LIST,'r',encoding='utf-8') as f:
        members = json.load(f)
        for ret in rets:
            cmpcode = str(int(ret['社員番号']))
            if members[cmpcode]['ignore'] == "0":
                newRets.append(ret)
    return(newRets)

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

    logger.info(str(getCurLineNo())+' START function')
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
                logger.error(str(getCurLineNo())+' Not found cmpcode:' + cmpcode + ' in members.')
            
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
            logger.info(str(getCurLineNo())+' attach file name {}'.format(attach))
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
    
    except exceptions.TypeError as e :
        logger.error(str(getCurLineNo())+' '+ str(e))
        raise(e)
    else:
        return ret

####################################
# 工数配分入力 - 個人選択
####################################
def selectMember(id, xp):
    """
    Overview
        工数配分入力結果で個人選択する
    Args
        id: 社員番号
        xp: 個人選択ボタンパス
    Return
        なし
    """
    logger.info(str(getCurLineNo())+' START function id:' + id)
    try:
        # フレーム指定
        driver.switch_to.parent_frame()
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])
        # 個人選択ボタンクリック
        findElement('xpath',xp).click()
        
        # サブウインドウにフォーカス移動
        wh = driver.window_handles
        driver.switch_to.window(wh[1])
        # フレーム指定
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])

        # selectインスタンス作成
        memberSelect = Select(findElement('name','lstSelemp'))
        # 指定のvalue値のoptionを選択
        memberSelect.select_by_value(id)
        # 確定ボタンクリック
        findElement('id','buttonKAKUTEI').click()
        time.sleep(FORCESLEEPSEC)

        # メインウィンドウにフォーカス移動
        driver.switch_to.window(wh[0])

        # フレーム指定
        driver.switch_to.parent_frame()
        frames = driver.find_elements_by_xpath("//frame")
        driver.switch_to.frame(frames[1])

    except Exception as e:
        logger.error(str(getCurLineNo())+' '+ str(e))
        raise(e)

####################################
# 期間チェック
####################################
def isContainDate(mmdd, startdate, enddate):
    """
    Overview
        対象日付が対象期間内かチェックする。start,endと同日は期間内とする。
    Args
        mmdd string : MM/DD形式の文字列
        startdate date
        enddate date
    Return
        ret: 期間内ならtrue、期間外ならfalse
    """
    logger.debug(str(getCurLineNo())+' START function mmdd:'+mmdd+' startdate:'+str(startdate)+' enddate:'+str(enddate))
    try:
        if startdate.strftime('%m/%d') <= mmdd and enddate.strftime('%m/%d') >= mmdd:
            ret = True
        else:
            ret = False
        return ret

    except Exception as e:
        logger.error(str(getCurLineNo())+' '+str(e))
        raise(e)

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
    logger.info(str(getCurLineNo())+' START function')
    try:
        # メンバー取得
        ids = getMemberList('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[4]/input')

        # 初期化
        rets = []

        # メンバー指定
        for id in ids:
            selectMember(id, '/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[4]/input')

            # 対象期間のスタート区間まで戻る
            nowTerm = findElement('xpath','/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[6]').text.split(' ')
            nowTermStart = datetime.strptime(nowTerm[0],'%Y/%m/%d')
            while startdate < datetime.date(nowTermStart):
                findElement('name','PrevEmpCode','click').click()
                waitLocate()
                time.sleep(FORCESLEEPSEC)
                # フレーム指定
                driver.switch_to.parent_frame()
                frames = driver.find_elements_by_xpath("//frame")
                driver.switch_to.frame(frames[1])
                # 現在表示中の開始日を取得
                nowTerm = findElement('xpath', '/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[6]').text.split(' ')
                nowTermStart = datetime.strptime(nowTerm[0],'%Y/%m/%d')

            while enddate >= datetime.date(nowTermStart):
                # 工数登録誤りチェック(対象期間終了まで繰り返し)
                # テーブル要素の構成「/table/tbody/tr[X]/td[Y]」が以下ルールになっている
                # X...1:日付, 8:就業時間, 16:合計
                # Y...3:1日目, 4:2日目,...,9:7日目
                for i in range(3,10):
                    ret = {}
                    mmdd = findElement('xpath', '//*[@id="xyw4100_form"]/table/tbody/tr[1]/td[' + str(i) +']').text
                    wt = findElement('xpath', '//*[@id="xyw4100_form"]/table/tbody/tr[8]/td[' + str(i) +']/font').text
                    total = findElement('xpath', '//*[@id="xyw4100_form"]/table/tbody/tr[16]/td[' + str(i) +']/font').text
                    if wt != total and isContainDate(mmdd, startdate, enddate):
                        ret['氏名'] = findElement('xpath', '/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[3]').text
                        ret['社員番号'] = id
                        ret['日付'] = findElement('xpath', '//*[@id="xyw4100_form"]/table/tbody/tr[1]/td[' + str(i) +']').text
                        ret['就業時間'] = wt
                        ret['合計'] = total
                        rets.append(ret)
                        logger.info(ret)
                # 次期間に移動
                findElement('name', 'NextEmpCode', 'click').click()
                waitLocate()
                time.sleep(FORCESLEEPSEC)
                # フレーム指定
                driver.switch_to.parent_frame()
                frames = driver.find_elements_by_xpath("//frame")
                driver.switch_to.frame(frames[1])
                # 現在表示中の開始日を取得
                nowTerm = findElement('xpath', '/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[6]').text.split(' ')
                nowTermStart = datetime.strptime(nowTerm[0],'%Y/%m/%d')

        return(rets)

    except Exception as e:
        logger.error(str(getCurLineNo())+' '+str(e))
        raise(e)

####################################
# 最終エラー処理
####################################
def cleanUpAfterError(error=None, webdriver=None):
    """
    Overview
        例外エラー後の一連処理を行う。
    Args
        webdriver: 終了させるwebdriverインスタンス。未起動時は引数なし
    Return
        なし
    """
    if webdriver is not None:
        webdriver.close()
    if error is not None:
        logger.exception(str(getCurLineNo())+' '+str(error))
    sys.exit()

####################################
# main
####################################
# コマンドライン引数定義
argparser = argparse.ArgumentParser()
argparser.add_argument('-m', '--mode', type=int, choices=[1,2,3], help='チェック種別 1:残業時間 2:打ち忘れ 3:工数登録', required=True)
argparser.add_argument('-o', '--output', type=int, choices=[1,2], help='出力タイプ 1:メール送信 2:CSVファイル出力', required=True)
argparser.add_argument('-d', '--date', type=lambda s: datetime.strptime(s, '%Y%m%d'), help='yyyymmdd形式で日を指定すると、その日に実行した仮定で実行される。')
argparser.add_argument('-e', '--exholiday', action='store_true', help='土日祝日の場合はチェックをしない。')
argparser.add_argument('-c', '--cmpcodefilter', type=int, nargs='*', help='対象の社員番号を指定。ブランク区切りで複数指定可能。')

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
    logger.info(str(getCurLineNo())+' 祝休日のため処理終了')
    cleanUpAfterError()

# 社員番号フィルタ取得
cmpcodefilter = args.cmpcodefilter

# config読み込み
try:
    config = configparser.ConfigParser()
    config.read(CONFIGFILE, 'UTF-8')
except Exception as e:
    logger.error(str(getCurLineNo())+' configファイル"'+CONFIGFILE+'"が見つかりません。')
    cleanUpAfterError(e)

# タイムアウト設定を取得
TIMEOUTSEC = int(config.get('environment', 'TIMEOUTSEC'))
FORCESLEEPSEC = int(config.get('environment', 'FORCESLEEPSEC'))
ERRORRETRYCOUNT = int(config.get('environment', 'ERRORRETRYCOUNT'))

# 開始日、終了日を取得
startdate,enddate = getSpan(nowDate,mode)
# 強制設定用
#startdate = date(2018,11,1)
#enddate = date(2019,8,31)
logger.info(str(getCurLineNo())+" collectionTerm: "+str(startdate)+" - "+str(enddate))

# 結果CSVファイル名セット
CSVNAME = 'resultAtdCheck_'+'m{:02}'.format(mode) + startdate.strftime('_%Y%m%d') + enddate.strftime('-%Y%m%d') \
    + datetime.now().strftime('_%Y%m%d-%H%M%S') + ".csv"
CSVNAME = os.path.join(parentdir, CSVNAME)

# webDriver起動
try:
    chromedriver_path = config.get('environment', 'chromedriver')
    driver = webdriver.Chrome(chromedriver_path)
    driver.implicitly_wait(TIMEOUTSEC)
except Exception as e:
    logger.error(str(getCurLineNo())+" 実行可能なWebDriver'"+chromedriver_path+"'が見つかりません。")
    cleanUpAfterError(e)

####################################
# ログイン認証
####################################
logger.info(str(getCurLineNo())+' START login')
driver.get(config.get('siteinfo', 'url'))
# 表示待ち
try:
    waitLocate()

    # メンテナンス中
    if driver.title == 'sorry page':
        logger.error(str(getCurLineNo())+' サーバメンテナンス中により処理中止')
        cleanUpAfterError(None,driver)

    # ID/PW入力
    findElement('name','DataSource').send_keys(config.get('siteinfo', 'cp'))
    findElement('name','LoginID').send_keys(config.get('siteinfo', 'id'))
    findElement('name','PassWord').send_keys(config.get('siteinfo', 'pw'))

    # ログインボタンクリック
    findElement('name','LOGINBUTTON').click()
    waitLocate()
    time.sleep(1)
except exceptions.TimeoutException as e:
    cleanUpAfterError(e,driver)
except exceptions.UnexpectedAlertPresentException as e:
    logger.error(str(getCurLineNo())+' ログインに失敗しました。cp:'+config.get('siteinfo', 'cp')+' id:'+config.get('siteinfo', 'id')+' pw:'+config.get('siteinfo', 'pw'))
    cleanUpAfterError(e,driver)
except Exception as e:
    cleanUpAfterError(e,driver)

####################################
# チェック結果取得
####################################
# mode別チェック結果取得
try:
    menuClick(config.get('modeinfo_'+str(mode), 'CLICKMENU'))
except Exception as e:
    logger.error(str(getCurLineNo())+' メニュークリック失敗')
    cleanUpAfterError(e,driver)

try:
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
                logger.info(str(getCurLineNo())+' '+str(ret))
                rets.append(ret)

    # 打ち忘れチェックリスト取得
    elif mode == 2:
        rets = checkStampMiss()

    # 工数登録結果取得
    elif mode == 3:
        rets = checkManHourRegist()
except Exception as e:
    logger.error(str(getCurLineNo())+' 結果取得失敗')
    cleanUpAfterError(e,driver)

####################################
# チェック結果アウトプット
####################################
try:
    # 無視リスト対象者削除
    rets = deleteIgnoreMember(rets)

    # 結果が1件以上あったらアウトプット
    if len(rets) > 0:
        # メール送信
        if args.output == 1:
            sendResultMail(rets,
            config.get('modeinfo_'+str(mode), 'MAILTITLE')+' '+startdate.strftime('%m/%d-')+enddate.strftime('%m/%d'),
            config.get('modeinfo_'+str(mode), 'MAILBODY')+'\n\n',
            #[CSVNAME],
            False,
            int(config.get('modeinfo_'+str(mode), 'MAIL_ESC_LEVEL')))
        # CSVファイル出力
        elif args.output == 2:
            csvOutput(rets,CSVNAME)
except Exception as e:
    logger.error(str(getCurLineNo())+' 結果送信失敗')
    cleanUpAfterError(e,driver)

# 終了処理
logger.info("---- COMPLETE "+__file__+"  ----")
cleanUpAfterError(None,driver)