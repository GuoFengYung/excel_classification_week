import os
import pandas as pd
import argparse
from datetime import datetime
import holidays


def read_excel(path_to_excel: str, save_name: str) -> str:

    if os.path.join('output_data') is not None:
        os.makedirs('output_data', exist_ok=True)
    if os.path.join('temp') is not None:
        os.makedirs('temp', exist_ok=True)

    tw_holiday = holidays.country_holidays('TW')
    data = pd.read_csv(path_to_excel, encoding="utf_8_sig")

    # region ===== Classification a week =====
    time_list = data['時間'].tolist()
    week_list = ["1", "2", "3", "4", "5", "6", "7"]
    time_li = []
    week_li = []
    for time_bar in time_list:
        if not isinstance(time_bar, str):
            time_bar = datetime.strftime(time_bar, '%Y/%m/%d %H:%M')
            time_bar = datetime.strptime(time_bar, '%Y/%m/%d %H:%M')
            week_li.append(week_list[datetime.date(time_bar).weekday()])
            time_bar = datetime.strftime(time_bar, '%Y/%m/%d %H:%M')
            time_li.append(time_bar)
        else:
            time_bar = datetime.strptime(time_bar, '%Y/%m/%d %H:%M')
            week_li.append(week_list[datetime.date(time_bar).weekday()])
            time_bar = datetime.strftime(time_bar, '%Y/%m/%d %H:%M')
            time_li.append(time_bar)

    df = pd.DataFrame({'時間': time_li,
                       '星期': week_li})
    df.to_csv(os.getcwd() + os.path.sep + 'temp' + os.path.sep + 'all.csv', encoding="utf_8_sig")
    # endregion ==============================

    # region ===== Split holiday =============
    time_list = data['時間'].tolist()
    week_list = ["1", "2", "3", "4", "5", "6", "7"]
    time_li = []
    week_li = []
    for time_bar in time_list:
        if time_bar in tw_holiday:
            continue
        else:
            if not isinstance(time_bar, str):
                time_bar = datetime.strftime(time_bar, '%Y/%m/%d %H:%M')
                time_bar = datetime.strptime(time_bar, '%Y/%m/%d %H:%M')
                week_li.append(week_list[datetime.date(time_bar).weekday()])
                time_bar = datetime.strftime(time_bar, '%Y/%m/%d %H:%M')
                time_li.append(time_bar)
            else:
                time_bar = datetime.strptime(time_bar, '%Y/%m/%d %H:%M')
                week_li.append(week_list[datetime.date(time_bar).weekday()])
                time_bar = datetime.strftime(time_bar, '%Y/%m/%d %H:%M')
                time_li.append(time_bar)

    df = pd.DataFrame({'時間': time_li,
                       '星期': week_li})
    df.to_csv(os.getcwd() + os.path.sep + 'temp' + os.path.sep + 'output.csv', encoding="utf_8_sig")
    # endregion ==============================

    # region ===== Classification holiday ====
    data = pd.read_csv(os.getcwd() + os.path.sep + 'temp' + os.path.sep + 'all.csv', encoding="utf_8_sig")
    time_list = data['時間'].tolist()
    fes_li = []
    name_li = []
    for time_bar in time_list:
        if time_bar in tw_holiday:
            fes_li.append(time_bar)
            name_li.append('0')

    df = pd.DataFrame({
        '時間': fes_li,
        '星期': name_li
    })
    df.to_csv(os.getcwd() + os.path.sep + 'temp' + os.path.sep + 'output1.csv', encoding="utf_8_sig")


    # endregion ==============================

    # region ===== merge week and holiday to origin data ====
    df1 = pd.read_csv(os.getcwd() + os.path.sep + 'temp' + os.path.sep + 'output.csv', encoding="utf_8_sig")
    df2 = pd.read_csv(os.getcwd() + os.path.sep + 'temp' + os.path.sep + 'output1.csv', encoding="utf_8_sig")
    df3 = pd.concat([df1, df2])
    df3 = df3.drop('Unnamed: 0', axis=1)
    df3 = df3.sort_values(by='時間')
    df3.to_csv(os.getcwd() + os.path.sep + 'temp' + os.path.sep + "final.csv", index=False, encoding="utf_8_sig")
    df4 = pd.read_csv(os.getcwd() + os.path.sep + 'temp' + os.path.sep + 'final.csv', encoding="utf_8_sig")
    df5 = pd.read_csv(path_to_excel)
    df4['星期']
    df5['星期'] = df4['星期']
    df5.to_csv(save_name, encoding="utf_8_sig", index=False)
    print(f'Done! Csv has converted to {save_name}')
    # endregion ==============================




if __name__ == '__main__':
    def main():
        parser = argparse.ArgumentParser()

        # region ===== path arguments ==========
        parser.add_argument('-e', '--excel', type=str, required=True, help='path to excel')
        parser.add_argument('-s', '--save', type=str, required=True, help='save name')
        # endregion ============================
        args = parser.parse_args()
        path_to_excel = args.excel
        save_name = args.save
        read_excel(path_to_excel, save_name)

    main()
