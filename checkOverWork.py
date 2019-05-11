from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
import logging
import sys
import os
import csv

# 環境情報
PATH_CHROMEDRIVER = "c:/driver/chromedriver.exe"

# サイト情報
URL_LOGIN = "https://cxg8.i-abs.co.jp/cyberx/login.asp"

# ログイン情報
LOGIN_ID = "111"
LOGIN_PW = "kouhei03"
COMPANY = "icd"

# USAGE
USAGE = "Usage: " + sys.argv[0] + " mode [ yyyy mm dd ]\n" \
        " - mode : 1(overwork check) or 2(stamp miss check)"

# log設定
logging.basicConfig(level=logging.INFO,
    format="%(asctime)s %(process)d %(name)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

# 標準出力用 -> 出力レベルを変更しようとしてもbasicConfigをDEBUGにしないと出ないためコメントアウト
#handler1 = logging.StreamHandler()
#handler1.setLevel(logging.INFO)
#handler1.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))

# ログファイル出力用
handler2 = logging.FileHandler(filename=__file__ + ".log")
handler2.setLevel(logging.DEBUG)
handler2.setFormatter(logging.Formatter("%(asctime)s %(process)d %(name)s %(levelname)s %(message)s"))

#logger.addHandler(handler1)
logger.addHandler(handler2)

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

    except NoSuchElementException as e :
        logger.error("function menuClick: " + str(e))
        raise (e)
    else:
        return True

####################################
# 初期処理
####################################
# 引数チェック
logger.info('START input parameter check')
args = sys.argv
try:
    # 引数4個
    if len(args) == 4 +1:
        nowDate = date(int(args[2]), int(args[3]), int(args[4]))
    # 引数1個
    elif len(args) == 1 +1:
        nowDate = date.today()
    else :
        raise SyntaxError
    
except SyntaxError as e:
    logger.error(e)
    logger.error(USAGE)
    logger.error("引数は1個または4個のみ設定可能です。")
    sys.exit()

except ValueError as e:
    logger.error(e)
    logger.error(USAGE)
    logger.error("引数は年月日に適した数値のみ設定可能です。")
    sys.exit()

try:
    mode = int(args[1])
    if not (mode == 1 or mode == 2):
        raise ValueError

except ValueError as e:
    logger.error(e)
    logger.error(USAGE)
    logger.error("modeは1または2のみ設定可能です。")
    sys.exit()

# 開始日、終了日を取得
startdate,enddate = getSpan(nowDate,mode)
logger.info("collectionTerm: "+str(startdate)+" - "+str(enddate))

# 結果CSVファイル名
CSVNAME = "OverWork"+startdate.strftime('_F%Y%m%d')+enddate.strftime('-T%Y%m%d') \
    +datetime.now().strftime('_@%Y%m%d-%H%M%S') +".csv"

# webDriver起動
if os.path.exists(PATH_CHROMEDRIVER):
    driver = webdriver.Chrome(PATH_CHROMEDRIVER)
else:    
    logger.error("実行可能なWebDriver'"+PATH_CHROMEDRIVER+"'が見つかりません。")
    sys.exit()

####################################
# ログイン認証
####################################
logger.info('START login')
driver.get(URL_LOGIN)

# 表示待ち
WebDriverWait(driver, 10).until(
    EC.presence_of_all_elements_located((By.NAME, "DataSource"))
)

# ID/PW入力
driver.find_element_by_name("DataSource").send_keys(COMPANY)
driver.find_element_by_name("LoginID").send_keys(LOGIN_ID)
driver.find_element_by_name("PassWord").send_keys(LOGIN_PW)

# ログインボタンクリック
driver.find_element_by_name("LOGINBUTTON").click()

####################################
# ホーム画面 : 就業週報月報画面へ遷移
####################################
if mode == 1:
    menuClick("就業週報月報")
elif mode == 2:
    menuClick("打ち忘れﾁｪｯｸﾘｽﾄ")

####################################
# 就業週報月報画面 : 期間検索
####################################
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
logger.info('START monthlyWorkReport display')
# フレーム指定
driver.switch_to.parent_frame()
frames = driver.find_elements_by_xpath("//frame")
driver.switch_to.frame(frames[1])

# 表示月、取得データを初期化
dispmonth = 0
rets = []

while True:

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

    # 終了日まで指定日をインクリメントしながらデータ取得
    while curdate <= enddate:
        # 表示月を指定して月報を表示(指定日が月を跨いだ場合も)
        if curdate.month != dispmonth:
            dtElm = driver.find_element_by_id("CmbYM")
            Select(dtElm).select_by_value(curdate.strftime("%Y%m"))
            driver.find_element_by_name("srchbutton").click()
            dispmonth = curdate.month
            
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//td[@id='" + FDATE_ID + "']"))
            )

        # 対象日の指定列のデータを取得
        for key in itemids.keys():
            # 対象日のtdタグのidを作成
            tgtid = DAILYID_F + str(int(curdate.strftime("%d")) -1) + "-" + itemids[key]
            elm = driver.find_element_by_xpath("//td[@id='"+tgtid+"']")
            workTime = elm.get_attribute("DefaultValue")
            if workTime == EMPTY_MARK:
                workTime = "00:00"
            wt[key] += relativedelta(hours=int(workTime.split(":")[0]),minutes=int(workTime.split(":")[1]))
            wt[TOTALTIME] += relativedelta(hours=int(workTime.split(":")[0]),minutes=int(workTime.split(":")[1]))

        # 対象日のデータ取得を終えたらインクリメントして翌日へ
        curdate += relativedelta(days=1)
    # 対象社員の結果をリストに保存
    rets.append(wt)

    ####################################
    # 対象社員を変更
    ####################################
    # 次が選択可なら次社員を選択
    if driver.find_element_by_name("button4").is_enabled() :
        driver.find_element_by_name("button4").click()
    # 次が選択不可ならループ終了
    else :
        break

####################################
# 結果CSV出力
####################################
logger.info('START csv report output')
# ファイルオープン
with open(CSVNAME, 'w', newline='', encoding='utf_8_sig') as fp:
    writer = csv.writer(fp, lineterminator='\r\n')
    
    # ヘッダ行出力
    writer.writerow(wt.keys())

    # データ行出力
    for ret in rets: # メンバー分ループ
        csvlist = []
        for key in ret.keys(): # 出力項目分ループ
            # 文字列型はそのまま出力
            if type(ret[key]) is str:
                csvlist.append(ret[key])
            # 時間(relativedelta)型はHH:MM書式で出力
            elif type(ret[key]) is relativedelta:
                csvlist.append('{hour:02}:{min:02}'.format(
                    hour=ret[key].hours+ret[key].days*24,min=ret[key].minutes))
            else:
                raise TypeError("type: "+type(ret[key]))

        writer.writerow(csvlist)
        logger.debug(csvlist)

# 終了処理
driver.close()
logger.info("Output '"+CSVNAME+"'.")
logger.info("---- COMPLETE "+__file__+"  ----")
exit()