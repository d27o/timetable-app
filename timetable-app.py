from flask import Flask, request
import telebot 
import parser 
import xlrd3 as xlrd

# parser code
book = xlrd.open_workbook("./data/КБиСП 4 курс 2 сем-Д.xlsx")

# sheet type <class 'xlrd3.sheet.Sheet'>
sheet = book.sheet_by_index(0)
# 1-st: row
# 2-nd: col

class DayOfWeek(object):

    def __init__(self, name, first, second, third, fourth, fifth, sixth):
        self.name = name
        self.first = first
        self.second = second
        self.third = third
        self.fourth = fourth
        self.fifth = fifth
        self.sixth = sixth

def make_day_of_week(name, first, second, third, fourth, fifth, sixth):
    day = DayOfWeek(name, first, second, third, fourth, fifth, sixth)
    return day

def get_timetable_by_group(group):

    max_rows = sheet.nrows
    max_cols = sheet.ncols 

    days_of_week = []

    for col in range(max_cols):
        cell_value = str(sheet.cell(1, col).value)
        if cell_value != '' and group in cell_value:
            lessonStr_1_1 = f"| 1 {sheet.cell(3, col).value} {sheet.cell(3, col+1).value} {sheet.cell(3, col+2).value} {sheet.cell(3, col+3).value} \n|| {sheet.cell(4, col).value} {sheet.cell(4, col+1).value} {sheet.cell(4, col+2).value} {sheet.cell(4, col+3).value}\n" 
            lessonStr_1_2 = f"| 2 {sheet.cell(5, col).value} {sheet.cell(5, col+1).value} {sheet.cell(5, col+2).value} {sheet.cell(5, col+3).value} \n|| {sheet.cell(6, col).value} {sheet.cell(6, col+1).value} {sheet.cell(6, col+2).value} {sheet.cell(6, col+3).value}\n" 
            lessonStr_1_3 = f"| 3 {sheet.cell(7, col).value} {sheet.cell(7, col+1).value} {sheet.cell(7, col+2).value} {sheet.cell(7, col+3).value} \n|| {sheet.cell(8, col).value} {sheet.cell(8, col+1).value} {sheet.cell(8, col+2).value} {sheet.cell(8, col+3).value}\n" 
            lessonStr_1_4 = f"| 4 {sheet.cell(9, col).value} {sheet.cell(9, col+1).value} {sheet.cell(9, col+2).value} {sheet.cell(9, col+3).value} \n|| {sheet.cell(10, col).value} {sheet.cell(10, col+1).value} {sheet.cell(10, col+2).value} {sheet.cell(10, col+3).value}\n" 
            lessonStr_1_5 = f"| 5 {sheet.cell(11, col).value} {sheet.cell(11, col+1).value} {sheet.cell(11, col+2).value} {sheet.cell(11, col+3).value} \n|| {sheet.cell(12, col).value} {sheet.cell(12, col+1).value} {sheet.cell(12, col+2).value} {sheet.cell(12, col+3).value}\n" 
            lessonStr_1_6 = f"| 6 {sheet.cell(13, col).value} {sheet.cell(13, col+1).value} {sheet.cell(13, col+2).value} {sheet.cell(13, col+3).value} \n|| {sheet.cell(14, col).value} {sheet.cell(14, col+1).value} {sheet.cell(14, col+2).value} {sheet.cell(14, col+3).value}\n" 
            days_of_week.append(make_day_of_week('Понедельник', lessonStr_1_1, lessonStr_1_2, lessonStr_1_3, lessonStr_1_4, lessonStr_1_5, lessonStr_1_6))    

            lessonStr_2_1 = f"| {sheet.cell(15, col).value} {sheet.cell(15, col+1).value} {sheet.cell(15, col+2).value} {sheet.cell(15, col+3).value} \n|| {sheet.cell(16, col).value} {sheet.cell(16, col+1).value} {sheet.cell(16, col+2).value} {sheet.cell(16, col+3).value}\n" 
            lessonStr_2_2 = f"| {sheet.cell(17, col).value} {sheet.cell(17, col+1).value} {sheet.cell(17, col+2).value} {sheet.cell(17, col+3).value} \n|| {sheet.cell(18, col).value} {sheet.cell(18, col+1).value} {sheet.cell(18, col+2).value} {sheet.cell(18, col+3).value}\n" 
            lessonStr_2_3 = f"| {sheet.cell(19, col).value} {sheet.cell(19, col+1).value} {sheet.cell(19, col+2).value} {sheet.cell(19, col+3).value} \n|| {sheet.cell(20, col).value} {sheet.cell(20, col+1).value} {sheet.cell(20, col+2).value} {sheet.cell(20, col+3).value}\n" 
            lessonStr_2_4 = f"| {sheet.cell(21, col).value} {sheet.cell(21, col+1).value} {sheet.cell(21, col+2).value} {sheet.cell(21, col+3).value} \n|| {sheet.cell(22, col).value} {sheet.cell(22, col+1).value} {sheet.cell(22, col+2).value} {sheet.cell(22, col+3).value}\n" 
            lessonStr_2_5 = f"| {sheet.cell(23, col).value} {sheet.cell(23, col+1).value} {sheet.cell(23, col+2).value} {sheet.cell(23, col+3).value} \n|| {sheet.cell(24, col).value} {sheet.cell(24, col+1).value} {sheet.cell(24, col+2).value} {sheet.cell(24, col+3).value}\n" 
            lessonStr_2_6 = f"| {sheet.cell(25, col).value} {sheet.cell(25, col+1).value} {sheet.cell(25, col+2).value} {sheet.cell(25, col+3).value} \n|| {sheet.cell(26, col).value} {sheet.cell(26, col+1).value} {sheet.cell(26, col+2).value} {sheet.cell(26, col+3).value}\n" 
            days_of_week.append(make_day_of_week('Вторник', lessonStr_2_1, lessonStr_2_2, lessonStr_2_3, lessonStr_2_4, lessonStr_2_5, lessonStr_2_6))    

            lessonStr_3_1 = f"| {sheet.cell(27, col).value} {sheet.cell(27, col+1).value} {sheet.cell(27, col+2).value} {sheet.cell(27, col+3).value} \n|| {sheet.cell(28, col).value} {sheet.cell(28, col+1).value} {sheet.cell(28, col+2).value} {sheet.cell(28, col+3).value}\n" 
            lessonStr_3_2 = f"| {sheet.cell(29, col).value} {sheet.cell(29, col+1).value} {sheet.cell(29, col+2).value} {sheet.cell(29, col+3).value} \n|| {sheet.cell(30, col).value} {sheet.cell(30, col+1).value} {sheet.cell(30, col+2).value} {sheet.cell(30, col+3).value}\n" 
            lessonStr_3_3 = f"| {sheet.cell(31, col).value} {sheet.cell(31, col+1).value} {sheet.cell(31, col+2).value} {sheet.cell(31, col+3).value} \n|| {sheet.cell(32, col).value} {sheet.cell(32, col+1).value} {sheet.cell(32, col+2).value} {sheet.cell(32, col+3).value}\n" 
            lessonStr_3_4 = f"| {sheet.cell(33, col).value} {sheet.cell(33, col+1).value} {sheet.cell(33, col+2).value} {sheet.cell(33, col+3).value} \n|| {sheet.cell(34, col).value} {sheet.cell(34, col+1).value} {sheet.cell(34, col+2).value} {sheet.cell(34, col+3).value}\n" 
            lessonStr_3_5 = f"| {sheet.cell(35, col).value} {sheet.cell(35, col+1).value} {sheet.cell(35, col+2).value} {sheet.cell(35, col+3).value} \n|| {sheet.cell(36, col).value} {sheet.cell(36, col+1).value} {sheet.cell(36, col+2).value} {sheet.cell(36, col+3).value}\n" 
            lessonStr_3_6 = f"| {sheet.cell(37, col).value} {sheet.cell(37, col+1).value} {sheet.cell(37, col+2).value} {sheet.cell(37, col+3).value} \n|| {sheet.cell(38, col).value} {sheet.cell(38, col+1).value} {sheet.cell(38, col+2).value} {sheet.cell(38, col+3).value}\n"
            days_of_week.append(make_day_of_week('Среда', lessonStr_3_1, lessonStr_3_2, lessonStr_3_3, lessonStr_3_4, lessonStr_3_5, lessonStr_3_6))    

            lessonStr_4_1 = f"| {sheet.cell(39, col).value} {sheet.cell(39, col+1).value} {sheet.cell(39, col+2).value} {sheet.cell(39, col+3).value} \n|| {sheet.cell(40, col).value} {sheet.cell(40, col+1).value} {sheet.cell(40, col+2).value} {sheet.cell(40, col+3).value}"
            lessonStr_4_2 = f"| {sheet.cell(41, col).value} {sheet.cell(41, col+1).value} {sheet.cell(41, col+2).value} {sheet.cell(41, col+3).value} \n|| {sheet.cell(42, col).value} {sheet.cell(42, col+1).value} {sheet.cell(42, col+2).value} {sheet.cell(42, col+3).value}"
            lessonStr_4_3 = f"| {sheet.cell(43, col).value} {sheet.cell(43, col+1).value} {sheet.cell(43, col+2).value} {sheet.cell(43, col+3).value} \n|| {sheet.cell(44, col).value} {sheet.cell(44, col+1).value} {sheet.cell(44, col+2).value} {sheet.cell(44, col+3).value}"
            lessonStr_4_4 = f"| {sheet.cell(45, col).value} {sheet.cell(45, col+1).value} {sheet.cell(45, col+2).value} {sheet.cell(45, col+3).value} \n|| {sheet.cell(46, col).value} {sheet.cell(46, col+1).value} {sheet.cell(46, col+2).value} {sheet.cell(46, col+3).value}"
            lessonStr_4_5 = f"| {sheet.cell(47, col).value} {sheet.cell(47, col+1).value} {sheet.cell(47, col+2).value} {sheet.cell(47, col+3).value} \n|| {sheet.cell(48, col).value} {sheet.cell(48, col+1).value} {sheet.cell(48, col+2).value} {sheet.cell(48, col+3).value}"
            lessonStr_4_6 = f"| {sheet.cell(49, col).value} {sheet.cell(49, col+1).value} {sheet.cell(49, col+2).value} {sheet.cell(49, col+3).value} \n|| {sheet.cell(50, col).value} {sheet.cell(50, col+1).value} {sheet.cell(50, col+2).value} {sheet.cell(50, col+3).value}"
            days_of_week.append(make_day_of_week('Четверг', lessonStr_4_1, lessonStr_4_2, lessonStr_4_3, lessonStr_4_4, lessonStr_4_5, lessonStr_4_6))    
            
            lessonStr_5_1 = f"| {sheet.cell(51, col).value} {sheet.cell(51, col+1).value} {sheet.cell(51, col+2).value} {sheet.cell(51, col+3).value} \n|| {sheet.cell(52, col).value} {sheet.cell(52, col+1).value} {sheet.cell(52, col+2).value} {sheet.cell(52, col+3).value}"
            lessonStr_5_2 = f"| {sheet.cell(53, col).value} {sheet.cell(53, col+1).value} {sheet.cell(53, col+2).value} {sheet.cell(53, col+3).value} \n|| {sheet.cell(54, col).value} {sheet.cell(54, col+1).value} {sheet.cell(54, col+2).value} {sheet.cell(54, col+3).value}"
            lessonStr_5_3 = f"| {sheet.cell(55, col).value} {sheet.cell(55, col+1).value} {sheet.cell(55, col+2).value} {sheet.cell(55, col+3).value} \n|| {sheet.cell(56, col).value} {sheet.cell(56, col+1).value} {sheet.cell(56, col+2).value} {sheet.cell(56, col+3).value}"
            lessonStr_5_4 = f"| {sheet.cell(57, col).value} {sheet.cell(57, col+1).value} {sheet.cell(57, col+2).value} {sheet.cell(57, col+3).value} \n|| {sheet.cell(58, col).value} {sheet.cell(58, col+1).value} {sheet.cell(58, col+2).value} {sheet.cell(58, col+3).value}"
            lessonStr_5_5 = f"| {sheet.cell(59, col).value} {sheet.cell(59, col+1).value} {sheet.cell(59, col+2).value} {sheet.cell(59, col+3).value} \n|| {sheet.cell(60, col).value} {sheet.cell(60, col+1).value} {sheet.cell(60, col+2).value} {sheet.cell(60, col+3).value}"
            lessonStr_5_6 = f"| {sheet.cell(61, col).value} {sheet.cell(61, col+1).value} {sheet.cell(61, col+2).value} {sheet.cell(61, col+3).value} \n|| {sheet.cell(62, col).value} {sheet.cell(62, col+1).value} {sheet.cell(62, col+2).value} {sheet.cell(62, col+3).value}"
            days_of_week.append(make_day_of_week('Пятница', lessonStr_5_1, lessonStr_5_2, lessonStr_5_3, lessonStr_5_4, lessonStr_5_5, lessonStr_5_6))    

            lessonStr_6_1 = f"| {sheet.cell(63, col).value} {sheet.cell(63, col+1).value} {sheet.cell(63, col+2).value} {sheet.cell(63, col+3).value} \n|| {sheet.cell(64, col).value} {sheet.cell(64, col+1).value} {sheet.cell(64, col+2).value} {sheet.cell(64, col+3).value}"
            lessonStr_6_2 = f"| {sheet.cell(65, col).value} {sheet.cell(65, col+1).value} {sheet.cell(65, col+2).value} {sheet.cell(65, col+3).value} \n|| {sheet.cell(66, col).value} {sheet.cell(66, col+1).value} {sheet.cell(66, col+2).value} {sheet.cell(66, col+3).value}"
            lessonStr_6_3 = f"| {sheet.cell(67, col).value} {sheet.cell(67, col+1).value} {sheet.cell(67, col+2).value} {sheet.cell(67, col+3).value} \n|| {sheet.cell(68, col).value} {sheet.cell(68, col+1).value} {sheet.cell(68, col+2).value} {sheet.cell(68, col+3).value}"
            lessonStr_6_4 = f"| {sheet.cell(69, col).value} {sheet.cell(69, col+1).value} {sheet.cell(69, col+2).value} {sheet.cell(69, col+3).value} \n|| {sheet.cell(70, col).value} {sheet.cell(70, col+1).value} {sheet.cell(70, col+2).value} {sheet.cell(70, col+3).value}"
            lessonStr_6_5 = f"| {sheet.cell(71, col).value} {sheet.cell(71, col+1).value} {sheet.cell(71, col+2).value} {sheet.cell(71, col+3).value} \n|| {sheet.cell(72, col).value} {sheet.cell(72, col+1).value} {sheet.cell(72, col+2).value} {sheet.cell(72, col+3).value}"
            lessonStr_6_6 = f"| {sheet.cell(73, col).value} {sheet.cell(73, col+1).value} {sheet.cell(73, col+2).value} {sheet.cell(73, col+3).value} \n|| {sheet.cell(74, col).value} {sheet.cell(74, col+1).value} {sheet.cell(74, col+2).value} {sheet.cell(74, col+3).value}"
            days_of_week.append(make_day_of_week('Суббота', lessonStr_6_1, lessonStr_6_2, lessonStr_6_3, lessonStr_6_4, lessonStr_6_5, lessonStr_6_6))    

    return days_of_week

# bot code
bot = telebot.TeleBot('1818764636:AAGNmvmt4N02sCX5m_HJRlCMtMT0fJlAAZg')
bot.set_webhook(url="https://fb5022a7c0f5.ngrok.io")
app = Flask(__name__)


@app.route('/', methods=["POST"])
def webhook():
    bot.process_new_updates(
        [telebot.types.Update.de_json(request.stream.read().decode("utf-8"))]
    )
    return "ok"


@bot.message_handler(commands=['start'])
def start_command(message):
    bot.send_message(message.chat.id, 'Hi, i am timetable bot for RTU MIREA. Please enter you group by /enter_group command. ' + 
    'For example /enter_group ББСО-02-17')

@bot.message_handler(commands=['enter_group'])
def enter_group(message):
    group = message.text.split(" ")[1]


    res = get_timetable_by_group(str(group))
    resStr = ''
    for i in range(6):
        resStr += res[i].name + '\n'
        resStr += f"{res[i].first} '\n'"
        resStr += f"{res[i].second} '\n'"
        resStr += f"{res[i].third} '\n'"
        resStr += f"{res[i].fourth} '\n'"
        resStr += f"{res[i].fifth} '\n'"
        resStr += f"{res[i].sixth} '\n'"
        resStr += f"-----\n"

    bot.send_message(message.chat.id, resStr)
    
if __name__ == "__main__":
    app.run()