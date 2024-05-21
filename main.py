import openpyxl

import telebot
from telebot import types

from easydata2 import *

import datetime

DB_NAME = 'members'

#Forma_BI11_22042024

bot = telebot.TeleBot(get_id_data(DB_NAME, 'system')) 

def sync_first_biology_list(file, id):
    wb = openpyxl.load_workbook(filename=f'{file}.xlsx')
    active_sheet = wb['1']
    
    students_answers = get_item_data(DB_NAME, id, 'students_answers')
    students = list(students_answers.keys())
    print(f'Список студентов: {students}')
    
    props = {'D': '2', 'E': '2', 'F': '3', 'G': '4', 'H': '5', 'I': '9', 'J': '13'}
    
    for i in range(8, len(students) + 8):
        for l in ['D', 'E', 'F', 'G', 'H', 'I', 'J']: #2 2 3 4 5 9 13
            active_sheet[f'{l}{i}'] = students_answers[students[i - 8]][props[l]]
            print(f'Координаты: {l}{i}')
            print(f'Студент: {students[i - 8]}')
            print(f'Задание: {[props[l]]}')
            print(f'Ответ: {students_answers[students[i - 8]][props[l]]}')
    
    print('Сохраняем...')
    wb.save(filename=f'{file}.xlsx')
    print('Готово!')
    
def sync_students_list(file, id):
    wb = openpyxl.load_workbook(filename=f'{file}.xlsx')
    active_sheet = wb['Список']
    
    students_variants = get_item_data(DB_NAME, id, 'students_variants')
    students_notes = get_item_data(DB_NAME, id, 'students_notes')
    students = list(students_variants.keys())
    print(f'Список студентов: {students}')
    for i in range(8, len(students) + 8):
        print(f'Координаты: В{i}')
        print(f'Студент: {students[i - 8]}')
        active_sheet[f'B{i}'] = students[i - 8]
        
    for i in range(8, len(students) + 8):
        active_sheet[f'C{i}'] = int(students_variants[students[i - 8]])
        print(f'Координаты: C{i}')
        print(f'Студент: {students[i - 8]}')
        print(f'Вариант: {students_variants[students[i - 8]]}')
        
    for i in range(8, len(students) + 8):
        active_sheet[f'D{i}'] = int(students_notes[students[i - 8]])
        print(f'Координаты: D{i}')
        print(f'Студент: {students[i - 8]}')
        print(f'Оценка: {students_notes[students[i - 8]]}')
    
    print('Сохраняем...')
    wb.save(filename=f'{file}_py.xlsx')
    print('Готово!')
        
@bot.message_handler(commands=['test'])
def test(message):
    sync_students_list('Forma_BI11_22042024', '5776829003')
    sync_first_biology_list('Forma_BI11_22042024_py', '5776829003')
    
@bot.message_handler(commands=['start'])
def start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    btn1 = types.KeyboardButton('👤 Профиль')
    btn2 = types.KeyboardButton('📥 Начать работу')
    btn3 = types.KeyboardButton('💳 Подписка')
    btn4 = types.KeyboardButton('📕 Помощь')
    markup.add(btn1, btn2, btn3, btn4)
    
    bot.send_message(message.chat.id, text='Привет! Этот бот может помочь вам заполнять отчетные таблицы СтатГрад. Пожалуйста, следуйте инструкциям', reply_markup=markup)

@bot.message_handler(content_types=['text'])
def func(message):
    
    step = get_item_data(DB_NAME, str(message.from_user.id), 'step')
    license = get_item_data(DB_NAME, str(message.from_user.id), 'license')
        
    if message.text == '💳 Подписка':
        bot.send_message(message.chat.id, text=f'Из-за бета тестирования вам выдана подписка на `2024` год бесплатно', parse_mode='Markdown')
        
        give_item_data(DB_NAME, str(message.from_user.id), 'license', datetime.datetime.now().strftime('%Y'))
    
    if message.text == '📥 Начать работу' and license == datetime.datetime.now().strftime('%Y'):
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        btn1 = types.KeyboardButton('💂‍♂️ Английский язык')
        btn2 = types.KeyboardButton('🌿 Биология')
        markup.add(btn1, btn2)
        
        bot.send_message(message.chat.id, text='Выберите необходимый предмет', reply_markup=markup, parse_mode='Markdown')
        
        give_item_data(DB_NAME, str(message.from_user.id), 'step', 1)
        
    if message.text == '👤 Профиль':
        license = get_item_data(DB_NAME, str(message.from_user.id), 'license')
        help_count = get_item_data(DB_NAME, str(message.from_user.id), 'count')
        license_a = '🟢 Активна' if license != None else '🔴 Не активна'
        text = '**Профиль пользователя**\n\n'
        text += f'**Пользователь:** `{message.from_user.id}`\n'
        text += f'**Лицензия:** `{license_a}`\n\n'
        text += f'**Год активаци:** {license}\n'
        text += f'**Кол-во запросов:** {help_count}'
        
        
        bot.send_message(message.chat.id, text=text, parse_mode='Markdown')
    
    if message.text == '📕 Помощь':
        text = 'Привет! В данном боте вы можете ускорить заполнение таблиц СтатГрад.'
        text = text + '\nСледуйте инструкциям, а после заполнения нажмите на "📌 Следующий шаг"'
        text = text + '\nСписок учеников можно делать как и отдельными сообщениями, но и как одним огромным сообщением, где ученики идут в столбик'
        text = text + '\nЕсли в перечне первого шага нет нужного предмета - отправьте нам названием предмета и пустую таблицу, а так же комментарий, по типу критерий оценок или нюансов'
        
        bot.send_message(message.chat.id, text=text, parse_mode='Markdown')
        
    if step == 1 and message.text != '📌 Следующий шаг':
        if message.text in ['🌿 Биология', '💂‍♂️ Английский язык']:
            give_item_data(DB_NAME, str(message.from_user.id), 'subject', message.text.split(' ', 1)[1])
        
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        btn1 = types.KeyboardButton('📌 Следующий шаг')
        markup.add(btn1)
        
        bot.send_message(message.chat.id, text=f'Успешно выбран предмет `{message.text}`', parse_mode='Markdown', reply_markup=markup)
    if step == 2 and message.text != '📌 Следующий шаг':
        students = get_item_data(DB_NAME, str(message.from_user.id), 'students_variants')
        
        if not students:
            students = {}
            
        students_notes = get_item_data(DB_NAME, str(message.from_user.id), 'students_notes')
        
        if not students_notes:
            students_notes = {}
            
        for student in message.text.split('\n'):
            
            if len(student.split('-')[0].split()) > 3:
                bot.send_message(message.chat.id, text='Неверный формат ввода. Введите в формате ФИО - Вариант - Оценка за последний период')
                return
        
            try:    
                students_notes[student.split('-', 2)[0].strip()] = student.split('-', 2)[2].strip()
                students[student.split('-', 2)[0].strip()] = student.split('-', 2)[1].strip()
            except:
                bot.send_message(message.chat.id, text='Неверный формат ввода. Введите в формате ФИО - Вариант - Оценка за последний период')
                return
            
            give_item_data(DB_NAME, str(message.from_user.id), 'students_notes', students_notes)
            give_item_data(DB_NAME, str(message.from_user.id), 'students_variants', students)
        
        variants_count = max(list(students.values()))
        students_count = len(list(students.keys()))
        
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        btn1 = types.KeyboardButton('📌 Следующий шаг')
        markup.add(btn1)
        
        bot.send_message(message.chat.id, text=f'Получено учеников: {students_count}\nВариантов: {variants_count}', parse_mode='Markdown', reply_markup=markup)
    if step == 3 and message.text != '📌 Следующий шаг':
        answers = get_item_data(DB_NAME, str(message.from_user.id), 'answers')
        
        if answers == None:
            answers = {}
        
        answers_on_variant = {}
        
        answers_on_variant_list = message.text.split('\n')
        for answer in answers_on_variant_list:
            answers_on_variant[answers_on_variant_list.index(answer) + 1] = answer
            
        answers[len(list(answers.keys())) + 1] = answers_on_variant
        
        give_item_data(DB_NAME, str(message.from_user.id), 'answers', answers)
        
        
        variants_count = len(list(answers.keys()))

        bot.send_message(message.chat.id, text=f'Получено {variants_count} вариантов', parse_mode='Markdown')
    if step == 4 and message.text != '📌 Следующий шаг' and message.text != '✔️ Завершить':
        students_variants = get_item_data(DB_NAME, str(message.from_user.id), 'students_variants')
        
        students_answers = get_item_data(DB_NAME, str(message.from_user.id), 'students_answers')
        
        if students_answers == None:
            students_answers = {}
        
        students_answers_on_variant = {}
        
        students_answers_on_variant_list = message.text.split('\n')
        for answer in students_answers_on_variant_list:
            students_answers_on_variant[students_answers_on_variant_list.index(answer) + 1] = answer
            
        students = list(students_variants.keys())
        
        try:
            student = students[len(list(students_answers.keys()))]
        except IndexError:
            bot.send_message(message.chat.id, text=f'Количество ответов превысело количество введенных учеников')
            return
        
        students_answers[student] = students_answers_on_variant
        
        give_item_data(DB_NAME, str(message.from_user.id), 'students_answers', students_answers)
        
        
        students_count = len(list(students_answers.keys()))
        
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        btn1 = types.KeyboardButton('✔️ Завершить')
        markup.add(btn1)

        bot.send_message(message.chat.id, text=f'Получен ученик {student}.\nПолучено {students_count} учеников', parse_mode='Markdown', reply_markup=markup)
    if message.text == '📌 Следующий шаг' or message.text == '✔️ Завершить':
        
        if step == 1:
            bot.send_message(message.chat.id, text='Необходим список учеников, идущий в нужном порядке. Второе число после тире - это вариант, а второе - оценка за предыдущий период. В этом же порядке вы и будете выдавать мне их ответы.', parse_mode='Markdown')
            bot.send_message(message.chat.id, text='``Пример ввода``', parse_mode='Markdown')
            bot.send_message(message.chat.id, text='Петров Александр Иванович - 1 - 5', parse_mode='Markdown')
            bot.send_message(message.chat.id, text='Иванов Петр Николаевич - 2 - 3', parse_mode='Markdown')
        
        if step == 2:
            bot.send_message(message.chat.id, text='Необходимы ответы на варианты. Вводите ответы в столбик. Одно сообщение - один вариант. Для ввода в столбик на компьютере, используйте SHIFT+ENTER', parse_mode='Markdown')
            bot.send_message(message.chat.id, text='``Пример ввода``', parse_mode='Markdown')
            bot.send_message(message.chat.id, text='13\n232133\nдемократия\n122121', parse_mode='Markdown')
            bot.send_message(message.chat.id, text='475\n212233\nсоциализм\n212121', parse_mode='Markdown')
        
        if step == 3:
            bot.send_message(message.chat.id, text='Необходимы ответы учеников. Вводите ответы в столбик. Ввод учеников строго в том порядке, в котором вы их изначально вводили. Одно сообщение - один ученкик. Для ввода в столбик на компьютере, используйте SHIFT+ENTER', parse_mode='Markdown')
            bot.send_message(message.chat.id, text='``Пример ввода``', parse_mode='Markdown')
            bot.send_message(message.chat.id, text='13\n232133\nдемократия\n122121', parse_mode='Markdown')
            bot.send_message(message.chat.id, text='475\n212233\nсоциализм\n212121', parse_mode='Markdown')
            
            
        if step == 4:
            
            bot.send_message(message.chat.id, text='Ввод данных завершен. Начинаю проверку ответов и выставление баллов', parse_mode='Markdown')
            give_item_data(DB_NAME, str(message.from_user.id), 'step', 0)
            
            give_item_data(DB_NAME, str(message.from_user.id), 'students_variants', None)
            give_item_data(DB_NAME, str(message.from_user.id), 'answers', None)
            give_item_data(DB_NAME, str(message.from_user.id), 'students_answers', None)
            give_item_data(DB_NAME, str(message.from_user.id), 'students_notes', None)
            
            start(message)
            return
            
        give_item_data(DB_NAME, str(message.from_user.id), 'step', int(step) + 1)

while True:
    print('Telegram бот активен')
    bot.polling(none_stop=True, interval=0)