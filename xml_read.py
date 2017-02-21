from xlrd import *
import calendar
import datetime
import time
rst_string = "txt已输出"


def not_null(lst):
    n_lst = [each for each in lst if each is not '']
    return n_lst


def get_year_month(_DATE_ORIGIN):
    _DATE = not_null(_DATE_ORIGIN)
    print(_DATE)
    date = _DATE[1].split(' ')[0]
    year_and_month = date[:-2]

    return year_and_month


def is_weekday(date):
    date = datetime.datetime.strptime(date, '%Y/%m/%d')
    day = date.weekday() + 1
    _is_or_not = True if day in range(1, 6) else False
    return _is_or_not


def return_is_weekday_list(year_month, _DAYS):
    is_weekday_list = []
    try:
        for each in _DAYS:
            day = str(int(each))
            date = year_month + day
            dict_tmp = {}
            dict_tmp[int(each)] = is_weekday(date)
            is_weekday_list.append(dict_tmp)
        print(year_month)
        mon = int(year_month.split('/')[-2])
        print(mon)

        holiday_file = open("holiday_2017.py", 'r')
        holiday = holiday_file.readlines()
        ##print(holiday)
        str_holi = ""
        for each in holiday:
            str_holi += each
        HOLIDAY_IS_WEEKDAY = eval(str_holi)
        for days in is_weekday_list:
            for day, t_or_f in days.items():
                for k, v in HOLIDAY_IS_WEEKDAY[mon].items():
                    if k == day:
                        if t_or_f != v:
                            days[day] = HOLIDAY_IS_WEEKDAY[mon][k]
    except Exception as e:
        print(Exception, ":", e)
    return is_weekday_list


def gen_data(row_list):
    try:
        new_row_list = []
        for row in row_list:
            if row_list.index(row) % 2 != 1:
                row = not_null(row)
            new_row_list.append(row)

        for row in new_row_list:
            for each in row:
                if each == '':
                    row[row.index(each)] = str(row.index(each) + 1)

        _paired = []
        for row in new_row_list:
            row_index = new_row_list.index(row)
            if row_index % 2 != 1:
                dict_tmp = {}
                dict_tmp["name"] = row[3]
                dict_tmp["record"] = new_row_list[row_index + 1]
                _paired.append(dict_tmp)

        for everyone_data in _paired:
            everyone_data["new_record"] = {}
            for _record in everyone_data["record"]:
                _index = everyone_data["record"].index(_record)
                _idx_record = {}
                _idx_record[_index + 1] = _record

                everyone_data["new_record"].update(_idx_record)

            del (everyone_data["record"])
    except Exception as e:
        print(Exception, ":", e)
    return _paired


def get_days(year_month):
    year = int(year_month.split('/')[0])
    day = int(year_month.split('/')[1])
    last_day = calendar.monthrange(year,
                        day)
    _DAYS = []
    for i in range(1, last_day[1]+1):
        _DAYS.append(i)
    return _DAYS


def main(fname):
    print("hello world")
    try:
        xls_data = open_workbook(fname)
        _TABLE = xls_data.sheet_by_name("刷卡记录")
        _ROWS = _TABLE.nrows
        if _ROWS % 2 == 1:
            _ROWS += 1
        _DATE_ORIGIN = _TABLE.row_values(2)
        year_month = get_year_month(_DATE_ORIGIN)
        _DAYS = get_days(year_month)
        is_weekday_list = return_is_weekday_list(year_month, _DAYS)
    except Exception as e:
        print(Exception, ":", e)

    try:
        row_list = []
        for row in range(4, _ROWS-1):
            row_data = _TABLE.row_values(row)
            row_list.append(row_data)
        row_list.append([])
        record_data = gen_data(row_list)
    except Exception as e:
        print(Exception, ":", e)

    for everyone in record_data:
        everyone["缺勤"] = []
        everyone["忘打卡"] = []
        everyone["加班"] = []
        for date, record in everyone["new_record"].items():
            for data in is_weekday_list:
                for day, t_or_f in data.items():
                    if day == date:
                        if '\n' not in record:
                            if t_or_f == True:
                                everyone["缺勤"].append(day)
                        if record.count('\n') != 0:
                            if t_or_f != True:
                                everyone["加班"].append(day)
                        if record.count('\n') == 1:
                            忘打卡 = {day: record}
                            everyone["忘打卡"].append(忘打卡)
        del (everyone["new_record"])
    ##    return record_data
    now_time = time.strftime("%m%d_%H%M")
    fname = fname.split('.')[0]
    f_name = fname + now_time + '.txt'
    fp = open(f_name, 'w', encoding="utf-8")
    for data in record_data:
        try:
            stri = 'name' + ':' + str(data["name"]) + '\n' \
                    + '忘打卡' + ":" + str(data['忘打卡']) + '\n' \
                    + '加班' + ":" + str(data['加班']) + '\n' \
                    + '缺勤' + ":" + str(data['缺勤']) + '\n'
            stri = stri.replace('\\n', '')
            fp.write(stri)
            fp.write('\n')
            fp.flush()
        except Exception:
            print('process except')
            print('to return in except')
    fp.close()
    return rst_string
if __name__ == "__main__":
    main(fname)
