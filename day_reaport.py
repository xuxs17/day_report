import xlrd
import math

class day_report():
#初版(已经优化，添加千分位)
    def gen_report_word(self):
        excel = xlrd.open_workbook('D:/日报/日报.xlsx')
        table = excel.sheet_by_index(0)
        for n in range(1, table.nrows):
            target = table.cell(n, 1).value
            today_value = table.cell(n, 2).value
            before_day_value = table.cell(n, 3).value
            day_ratio = format(abs(float(table.cell(n, 4).value)), '.2%')
            week_ratio = format(abs(float(table.cell(n, 5).value)), '.2%')
            # three_day_ratio = format(abs(float(table.cell(n, 4).value)),'.2%')
            three_day_ratio = table.cell(n, 6).value

            data = {'target': target, 'today_value': today_value, 'before_day_value': before_day_value,
                    'day_ratio': day_ratio, 'week_ratio': week_ratio, 'three_day_ratio': three_day_ratio}


            if data['three_day_ratio'] == '-':
            # 当连续三日累计的值不存在的时候
                if float(data['today_value']) > 1:
                # 判断指标不是百分比类型
                    if data['today_value'] - data['before_day_value'] >= 0:
                        print('{}日环比增长{}({:,}k),'.format(data['target'], data['day_ratio'],round((data['today_value'] - data['before_day_value']) / 1000)))
                    else:
                        print('{}日环比下降{}(-{:,}k),'.format(data['target'], data['day_ratio'], round(
                            abs((data['today_value'] - data['before_day_value']) / 1000))))
                else:
                # 指标是百分比类型
                    if data['today_value'] - data['before_day_value'] >= 0:
                        print('{}日环比增长{}({}->{}),'.format(data['target'], data['day_ratio'],
                                                          format(data['before_day_value'], '.2%'),
                                                          format(data['today_value'], '.2%')))
                    else:
                        print('{}日环比下降{}({}->{}),'.format(data['target'], data['day_ratio'],
                                                          format(data['before_day_value'], '.2%'),
                                                          format(data['today_value'], '.2%')))

            else:
            # 当连续三日累计的值存在的时候
                if float(data['today_value']) > 1:
                # 判断指标不是百分比类型
                    if data['today_value'] - data['before_day_value'] >= 0:
                        print('{}日环比增长{}({:,}k),{}连续三日累计增长{}({:,}k),'.format(data['target'], data['day_ratio'], round(
                            (data['today_value'] - data['before_day_value']) / 1000),
                                                                         data['target'],
                                                                         format(abs(float(data['three_day_ratio'] )),
                                                                                '.2%'), round(abs(
                                (data['today_value'] - (
                                data['today_value'] / (1 + (data['three_day_ratio'])))) / 1000))))
                    else:
                        print('{}日环比下降{}(-{:,}k),{}连续三日累计下降{}(-{:,}k),'.format(data['target'], data['day_ratio'], round(
                            abs((data['today_value'] - data['before_day_value']) / 1000)),
                                                                          data['target'],
                                                                          format(abs(float(data['three_day_ratio'] )),
                                                                                 '.2%'), round(abs(
                                (data['today_value'] - (
                                data['today_value'] / (1 + (data['three_day_ratio'])))) / 1000))))
                else:
                # 指标是百分比类型
                    if data['today_value'] - data['before_day_value'] >= 0:
                        print('{}日环比增长{}({}->{}),'.format(data['target'], data['day_ratio'],
                                                          format(data['before_day_value'], '.2%'),
                                                          format(data['today_value'], '.2%')))
                    else:
                        print('{}日环比下降{}({}->{}),'.format(data['target'], data['day_ratio'],
                                                          format(data['before_day_value'], '.2%'),
                                                          format(data['today_value'], '.2%')))



if __name__ == '__main__':
    a=day_report().gen_report_word()

