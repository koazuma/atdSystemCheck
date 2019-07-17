import csv
from datetime import datetime


class JapanHoliday:
    """
    内閣府が公表している「平成29年（2017年）から平成31年（2019年）国民の祝日等
    （いわゆる振替休日等を含む）」のCSVを使用して
    入力された日付('2017-01-01' %Y-%m-%d形式)が土、日、休日か判定するクラス
    """

    #def __init__(self, path='syukujitsu_kyujitsu.csv', encoding='cp932'):
    def __init__(self, path='syukujitsu.csv', encoding='cp932'):
        self._holidays = self._read_holiday(path, encoding)

    def _read_holiday(self, path, encoding):
        """
        CSVファイルを読み込み、self.holidaysに以下の形式のdictをListに格納する
        {'2017-01-09': {'day': '2017-01-09', 'name': '成人の日'}}
        CSVファイルがないとIOErrorが発生する

        :param path: 祝日と祝日名が記入されたCSVファイル。ヘッダーが必要
        :param encoding: CSVファイルのエンコーディング
        :return: [{'2017-01-09': {'day': '2017-01-09', 'name': '成人の日'}...}]
        """
        holidays = {}
        with open(path, encoding=encoding, newline='') as f:
            reader = csv.reader(f)
            next(reader)  # CSVのヘッダーを飛ばす
            for row in reader:
                day, name = row[0], row[1]
                # 内閣府ファイルのフォーマット変更に合わせて追記
                day = datetime.strptime(day, '%Y/%m/%d').strftime('%Y-%m-%d')
                holidays[day] = {'day': day, 'name': name}
        return holidays

    def is_holiday(self, day_str):
        """
        土、日、祝日か判定する
        :param day_str: '2018-03-01'の%Y-%m-%dのstrを受け取る。
                        形式が違うとValueErrorが発生
        :return: 土、日、祝日ならTrue, 平日ならFalse
        """
        try:
            day = datetime.strptime(day_str, '%Y-%m-%d')
            if day.weekday() >= 5:
                return True
        except ValueError:
            print('日付は2018-03-01 %Y-%m-%dの形式で入力してください')
            raise ValueError
        if day_str in self._holidays:
            return True
        return False

    def get_holiday_dict(self):
        """
        祝日を一覧化したdictを返す
        {'2017-01-09': {'day': '2017-01-09', 'name': '成人の日'},
         '2017-02-11': {'day': '2017-02-11', 'name': '建国記念の日'},...}
        """
        return self._holidays