import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
import pandas as pd
import random
import os

df = None
file_path = None
sheet_name = None
list_files = []
list_file_paths = []
list_file_paths_from_log = []
selected_file_path = None
is_opened_from_list = False

def open_tab(mode):
    """Комментарий"""
    global gen_tab_counter
    global race
    global df
    global file_path
    global list_files
    global list_file_paths
    global selected_file_path
    global spare_param_pts
    global spare_skill_pts
    global is_opened_from_list

    race = None
    name = ''
    list_races = ['Люди', 'Эльфы', 'Гномы', 'Орки', 'Нежить']
    list_human_names = ['Моусабель', 'Тасселандра', 'Ната Серпент', 'Хукир Наламбар', 'Мадислак Улмокина']
    list_elf_names = ['Рилеф Селеварун', 'Меларю Летний Бриз', 'Шиидра Ильфелкиир', 'Вин Экорн', 'Сейлас Катдейрин']
    list_dwarf_names = ['Димбль \"Пивохлёб\"', 'Эллижобель', 'Глим \"Работяга\"', 'Келлен', 'Бимпноттин']
    list_orc_names = ['Хенк', 'Шамп', 'Фенг', 'Оунка', 'Вола']
    list_undead_names = ['Морелл Клемонс', 'Маурус Линдсей', 'Камрен Кервин', 'Джастиш Гнаш', 'Берти Арк']
    list_param_pts = [0] * 11
    list_param = ['Сила', 'Тело', 'Точность','Подвижность', 'Интеллект', 'Сообразительность', 'Восприятие', 'Фантазия', 
                  'Воля', 'Харизма', 'Внешность']
    list_skill = ['Риторика', 'Дипломатия', 'Хитрость', 'Эмпатия', 'Внимательность', 'Техника', 'Вычисление', 'Анализ',
                  'Моделирование', 'Искусство', 'Исполнение', 'Мелкая моторика', 'Ремесло', 'Лабораторная работа', 
                  'Рукопашный бой', 'Холодное оружие', 'Огнестрельное оружие', 'Атлетика', 'Акробатика', 'Ловкость рук', 
                  'Скрытность', 'Интуиция', 'Медитация', 'Ментальный контроль', 'Стойкость', 'Мыслеобраз']
    list_skill_for_display = []
    list_skill_pts_for_display = []
    list_skill_pts = [0] * 26
    spare_param_pts = 0
    spare_skill_pts = 0

    labels_skill = []
    labels_pts = []
    buttons_del = []
    buttons_minus_skill = []
    buttons_plus_skill = []
    labels_min_skill_pts = []
    labels_max_skill_pts = []
    
    if mode == 1:
        def read_name():
            nonlocal name
            name = entry_name.get()

        def add_pts(index):
            global spare_param_pts
            nonlocal list_param_pts
            nonlocal list_max_param
            nonlocal list_min_param
            if spare_param_pts > 0 and list_param_pts[index] < list_max_param[index]:
                spare_param_pts -= 1
                list_param_pts[index] += 1
                update_pts()
            if list_param_pts[index] == list_max_param[index]:
                if index == 0:
                    btn_plus_param1.config(state="disabled")
                if index == 1:
                    btn_plus_param2.config(state="disabled")
                if index == 2:
                    btn_plus_param3.config(state="disabled")
                if index == 3:
                    btn_plus_param4.config(state="disabled")
                if index == 4:
                    btn_plus_param5.config(state="disabled")
                if index == 5:
                    btn_plus_param6.config(state="disabled")
                if index == 6:
                    btn_plus_param7.config(state="disabled")
                if index == 7:
                    btn_plus_param8.config(state="disabled")
                if index == 8:
                    btn_plus_param9.config(state="disabled")
                if index == 9:
                    btn_plus_param10.config(state="disabled")
                if index == 10:
                    btn_plus_param11.config(state="disabled")
            if list_param_pts[index] > list_min_param[index]:
                if index == 0:
                    btn_minus_param1.config(state="normal")
                if index == 1:
                    btn_minus_param2.config(state="normal")
                if index == 2:
                    btn_minus_param3.config(state="normal")
                if index == 3:
                    btn_minus_param4.config(state="normal")
                if index == 4:
                    btn_minus_param5.config(state="normal")
                if index == 5:
                    btn_minus_param6.config(state="normal")
                if index == 6:
                    btn_minus_param7.config(state="normal")
                if index == 7:
                    btn_minus_param8.config(state="normal")
                if index == 8:
                    btn_minus_param9.config(state="normal")
                if index == 9:
                    btn_minus_param10.config(state="normal")
                if index == 10:
                    btn_minus_param11.config(state="normal")

        def subtract_pts(index):
            global spare_param_pts
            nonlocal list_param_pts
            nonlocal list_min_param
            nonlocal list_max_param
            if list_param_pts[index] > list_min_param[index]:
                spare_param_pts += 1
                list_param_pts[index] -= 1
            update_pts()
            if list_param_pts[index] == list_min_param[index]:
                if index == 0:
                    btn_minus_param1.config(state="disabled")
                if index == 1:
                    btn_minus_param2.config(state="disabled")
                if index == 2:
                    btn_minus_param3.config(state="disabled")
                if index == 3:
                    btn_minus_param4.config(state="disabled")
                if index == 4:
                    btn_minus_param5.config(state="disabled")
                if index == 5:
                    btn_minus_param6.config(state="disabled")
                if index == 6:
                    btn_minus_param7.config(state="disabled")
                if index == 7:
                    btn_minus_param8.config(state="disabled")
                if index == 8:
                    btn_minus_param9.config(state="disabled")
                if index == 9:
                    btn_minus_param10.config(state="disabled")
                if index == 10:
                    btn_minus_param11.config(state="disabled")
            if list_param_pts[index] < list_max_param[index]:
                if index == 0:
                    btn_plus_param1.config(state="normal")
                if index == 1:
                    btn_plus_param2.config(state="normal")
                if index == 2:
                    btn_plus_param3.config(state="normal")
                if index == 3:
                    btn_plus_param4.config(state="normal")
                if index == 4:
                    btn_plus_param5.config(state="normal")
                if index == 5:
                    btn_plus_param6.config(state="normal")
                if index == 6:
                    btn_plus_param7.config(state="normal")
                if index == 7:
                    btn_plus_param8.config(state="normal")
                if index == 8:
                    btn_plus_param9.config(state="normal")
                if index == 9:
                    btn_plus_param10.config(state="normal")
                if index == 10:
                    btn_plus_param11.config(state="normal")

        def save_file():
            nonlocal list_param
            nonlocal list_skill_for_display
            nonlocal list_skill_pts_for_display
            global selected_file_path
            global race
            nonlocal name
            global spare_param_pts
            global spare_skill_pts
            read_name()
            df = pd.DataFrame(index=range(33), columns=['A', 'B'])
            df['A'][0] = race
            df['A'][1:12] = list_param
            df['A'][12] = 'Свободные очки:'
            df['A'][14:14 + len(list_skill_for_display)] = list_skill_for_display
            df['A'][14 + len(list_skill_for_display)] = 'Свободные очки:'
            df['B'][0] = name
            df['B'][1:12] = list_param_pts
            df['B'][12] = spare_param_pts
            df['B'][14:14 + len(list_skill_for_display)] = list_skill_pts_for_display
            df['B'][14 + len(list_skill_for_display)] = spare_skill_pts
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            selected_file_path = file_path
            if file_path:
                df.to_excel(file_path, index=False, header=False)
                update_log()
            else:
                pass
            
        new_tab = Frame(tab_control)
        tab_control.add(new_tab, text=f"Ген. персонажа {gen_tab_counter}")
        tabmenu.add_command(label=f"Ген. персонажа {gen_tab_counter}")

        last_selected = None
        label_race = Label(new_tab, text="Раса:")
        combobox = ttk.Combobox(new_tab, values=list_races, width=15, state="readonly")
        label_race.place(x=110, y=105)
        combobox.place(x=150, y=105)

        label_name = Label(new_tab, text="Имя:")
        entry_name = Entry(new_tab, width=40)
        label_name.place(x=430, y=105)
        entry_name.place(x=470, y=105)

        race = list_races[random.randint(1, 5) - 1]
        if race == 'Люди':
            spare_param_pts = 12
            spare_skill_pts = 15
            list_min_param = [1] * 11
            list_max_param = [5, 5, 5, 5, 5, 6, 5, 6, 5, 5, 5]
            for index in range(11):
                list_param_pts[index] = list_min_param[index]
            for index in range(26):
                list_skill_pts[index] = 0
            entry_name.delete(0, END)
            entry_name.insert(0, list_human_names[random.randint(1, 5) - 1])
            combobox.set('Люди')
            last_selected = 'Люди'
        if race == 'Эльфы':
            spare_param_pts = 9
            spare_skill_pts = 12
            list_min_param = [1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 2]
            list_max_param = [5, 5, 5, 5, 5, 5, 6, 5, 5, 5, 7]
            for index in range(11):
                list_param_pts[index] = list_min_param[index]
            for index in range(26):
                list_skill_pts[index] = 0
            entry_name.delete(0, END)
            entry_name.insert(0, list_elf_names[random.randint(1, 5) - 1])
            combobox.set('Эльфы')
            last_selected = 'Эльфы'
        if race == 'Гномы':
            spare_param_pts = 10
            spare_skill_pts = 12
            list_min_param = [2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1]
            list_max_param = [6, 6, 5, 5, 6, 5, 5, 5, 5, 5, 5]
            for index in range(11):
                list_param_pts[index] = list_min_param[index]
            for index in range(26):
                list_skill_pts[index] = 0
            entry_name.delete(0, END)
            entry_name.insert(0, list_dwarf_names[random.randint(1, 5) - 1])
            combobox.set('Гномы')
            last_selected = 'Гномы'
        if race == 'Орки':
            spare_param_pts = 10
            spare_skill_pts = 12
            list_min_param = [2, 2, 1, 1, 1, 1, 1, 1, 2, 1, 1]
            list_max_param = [7, 6, 5, 5, 5, 5, 5, 4, 7, 4, 4]
            for index in range(11):
                list_param_pts[index] = list_min_param[index]
            for index in range(26):
                list_skill_pts[index] = 0
            entry_name.delete(0, END)
            entry_name.insert(0, list_orc_names[random.randint(1, 5) - 1])
            combobox.set('Орки')
            last_selected = 'Орки'
        if race == 'Нежить':
            spare_param_pts = 9
            spare_skill_pts = 6
            list_min_param = [1] * 11
            list_max_param = [9, 3, 4, 4, 5, 4, 7, 3, 4, 3, 3]
            for index in range(11):
                list_param_pts[index] = list_min_param[index]
            for index in range(26):
                list_skill_pts[index] = 0
            entry_name.delete(0, END)
            entry_name.insert(0, list_undead_names[random.randint(1, 5) - 1])
            combobox.set('Нежить')
            last_selected = 'Нежить'
       
        label_spare_param = Label(new_tab, text=f'Свободные очки:  {spare_param_pts}')
        label_spare_skill = Label(new_tab, text=f'Свободные очки:  {spare_skill_pts}')
        label_spare_param.place(x=40, y=490)
        label_spare_skill.place(x=420, y=490)

        label_param_pt1 = Label(new_tab, text=f'{list_param_pts[0]}', justify='left')
        label_param_pt2 = Label(new_tab, text=f'{list_param_pts[1]}', justify='left')
        label_param_pt3 = Label(new_tab, text=f'{list_param_pts[2]}', justify='left')
        label_param_pt4 = Label(new_tab, text=f'{list_param_pts[3]}', justify='left')
        label_param_pt5 = Label(new_tab, text=f'{list_param_pts[4]}', justify='left')
        label_param_pt6 = Label(new_tab, text=f'{list_param_pts[5]}', justify='left')
        label_param_pt7 = Label(new_tab, text=f'{list_param_pts[6]}', justify='left')
        label_param_pt8 = Label(new_tab, text=f'{list_param_pts[7]}', justify='left')
        label_param_pt9 = Label(new_tab, text=f'{list_param_pts[8]}', justify='left')
        label_param_pt10 = Label(new_tab, text=f'{list_param_pts[9]}', justify='left')
        label_param_pt11 = Label(new_tab, text=f'{list_param_pts[10]}', justify='left')

        label_min_param1 = Label(new_tab, text=f'{list_min_param[0]}', justify='left')
        label_min_param2 = Label(new_tab, text=f'{list_min_param[1]}', justify='left')
        label_min_param3 = Label(new_tab, text=f'{list_min_param[2]}', justify='left')
        label_min_param4 = Label(new_tab, text=f'{list_min_param[3]}', justify='left')
        label_min_param5 = Label(new_tab, text=f'{list_min_param[4]}', justify='left')
        label_min_param6 = Label(new_tab, text=f'{list_min_param[5]}', justify='left')
        label_min_param7 = Label(new_tab, text=f'{list_min_param[6]}', justify='left')
        label_min_param8 = Label(new_tab, text=f'{list_min_param[7]}', justify='left')
        label_min_param9 = Label(new_tab, text=f'{list_min_param[8]}', justify='left')
        label_min_param10 = Label(new_tab, text=f'{list_min_param[9]}', justify='left')
        label_min_param11 = Label(new_tab, text=f'{list_min_param[10]}', justify='left')

        label_max_param1 = Label(new_tab, text=f'{list_max_param[0]}', justify='left')
        label_max_param2 = Label(new_tab, text=f'{list_max_param[1]}', justify='left')
        label_max_param3 = Label(new_tab, text=f'{list_max_param[2]}', justify='left')
        label_max_param4 = Label(new_tab, text=f'{list_max_param[3]}', justify='left')
        label_max_param5 = Label(new_tab, text=f'{list_max_param[4]}', justify='left')
        label_max_param6 = Label(new_tab, text=f'{list_max_param[5]}', justify='left')
        label_max_param7 = Label(new_tab, text=f'{list_max_param[6]}', justify='left')
        label_max_param8 = Label(new_tab, text=f'{list_max_param[7]}', justify='left')
        label_max_param9 = Label(new_tab, text=f'{list_max_param[8]}', justify='left')
        label_max_param10 = Label(new_tab, text=f'{list_max_param[9]}', justify='left')
        label_max_param11 = Label(new_tab, text=f'{list_max_param[10]}', justify='left')

        def update_pts():
            global spare_param_pts
            nonlocal label_spare_param
            label_param_pt1.config(text=f'{list_param_pts[0]}')
            label_param_pt2.config(text=f'{list_param_pts[1]}')
            label_param_pt3.config(text=f'{list_param_pts[2]}')
            label_param_pt4.config(text=f'{list_param_pts[3]}')
            label_param_pt5.config(text=f'{list_param_pts[4]}')
            label_param_pt6.config(text=f'{list_param_pts[5]}')
            label_param_pt7.config(text=f'{list_param_pts[6]}')
            label_param_pt8.config(text=f'{list_param_pts[7]}')
            label_param_pt9.config(text=f'{list_param_pts[8]}')
            label_param_pt10.config(text=f'{list_param_pts[9]}')
            label_param_pt11.config(text=f'{list_param_pts[10]}')
            label_spare_param.config(text=f'Свободные очки:  {spare_param_pts}')

            label_min_param1.config(text=f'{list_min_param[0]}')
            label_min_param2.config(text=f'{list_min_param[1]}')
            label_min_param3.config(text=f'{list_min_param[2]}')
            label_min_param4.config(text=f'{list_min_param[3]}')
            label_min_param5.config(text=f'{list_min_param[4]}')
            label_min_param6.config(text=f'{list_min_param[5]}')
            label_min_param7.config(text=f'{list_min_param[6]}')
            label_min_param8.config(text=f'{list_min_param[7]}')
            label_min_param9.config(text=f'{list_min_param[8]}')
            label_min_param10.config(text=f'{list_min_param[9]}')
            label_min_param11.config(text=f'{list_min_param[10]}')

            label_max_param1.config(text=f'{list_max_param[0]}')
            label_max_param2.config(text=f'{list_max_param[1]}')
            label_max_param3.config(text=f'{list_max_param[2]}')
            label_max_param4.config(text=f'{list_max_param[3]}')
            label_max_param5.config(text=f'{list_max_param[4]}')
            label_max_param6.config(text=f'{list_max_param[5]}')
            label_max_param7.config(text=f'{list_max_param[6]}')
            label_max_param8.config(text=f'{list_max_param[7]}')
            label_max_param9.config(text=f'{list_max_param[8]}')
            label_max_param10.config(text=f'{list_max_param[9]}')
            label_max_param11.config(text=f'{list_max_param[10]}')

        def update_points(event):
            global spare_param_pts
            global spare_skill_pts
            nonlocal last_selected
            nonlocal list_min_param
            nonlocal list_max_param
            nonlocal labels_skill
            nonlocal labels_pts
            nonlocal buttons_del
            nonlocal buttons_minus_skill
            nonlocal buttons_plus_skill
            nonlocal labels_min_skill_pts
            nonlocal labels_max_skill_pts
            nonlocal list_skill_for_display
            nonlocal list_skill_pts_for_display
            current_selection = combobox.get()
            if current_selection == 'Люди' and last_selected != 'Люди':
                spare_param_pts = 12
                spare_skill_pts = 15
                list_min_param = [1] * 11
                list_max_param = [5, 5, 5, 5, 5, 6, 5, 6, 5, 5, 5]
                for index in range(11):
                    list_param_pts[index] = list_min_param[index]
                for index in range(26):
                    list_skill_pts[index] = 0
                entry_name.delete(0, END)
            if current_selection == 'Эльфы' and last_selected != 'Эльфы':
                spare_param_pts = 9
                spare_skill_pts = 12
                list_min_param = [1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 2]
                list_max_param = [5, 5, 5, 5, 5, 5, 6, 5, 5, 5, 7]
                for index in range(11):
                    list_param_pts[index] = list_min_param[index]
                for index in range(26):
                    list_skill_pts[index] = 0
                entry_name.delete(0, END)
            if current_selection == 'Гномы' and last_selected != 'Гномы':
                spare_param_pts = 10
                spare_skill_pts = 12
                list_min_param = [2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1]
                list_max_param = [6, 6, 5, 5, 6, 5, 5, 5, 5, 5, 5]
                for index in range(11):
                    list_param_pts[index] = list_min_param[index]
                for index in range(26):
                    list_skill_pts[index] = 0
                entry_name.delete(0, END)
            if current_selection == 'Орки' and last_selected != 'Орки':
                spare_param_pts = 10
                spare_skill_pts = 12
                list_min_param = [2, 2, 1, 1, 1, 1, 1, 1, 2, 1, 1]
                list_max_param = [7, 6, 5, 5, 5, 5, 5, 4, 7, 4, 4]
                for index in range(11):
                    list_param_pts[index] = list_min_param[index]
                for index in range(26):
                    list_skill_pts[index] = 0
                entry_name.delete(0, END)
            if current_selection == 'Нежить' and last_selected != 'Нежить':
                spare_param_pts = 9
                spare_skill_pts = 6
                list_min_param = [1] * 11
                list_max_param = [9, 3, 4, 4, 5, 4, 7, 3, 4, 3, 3]
                for index in range(11):
                    list_param_pts[index] = list_min_param[index]
                for index in range(26):
                    list_skill_pts[index] = 0
                entry_name.delete(0, END)
            update_pts()
            
            last_selected = current_selection
            btn_minus_param1.config(state="disabled")
            btn_minus_param2.config(state="disabled")
            btn_minus_param3.config(state="disabled")
            btn_minus_param4.config(state="disabled")
            btn_minus_param5.config(state="disabled")
            btn_minus_param6.config(state="disabled")
            btn_minus_param7.config(state="disabled")
            btn_minus_param8.config(state="disabled")
            btn_minus_param9.config(state="disabled")
            btn_minus_param10.config(state="disabled")
            btn_minus_param11.config(state="disabled")

            while labels_skill:
                labels_skill[0].destroy()
                labels_pts[0].destroy()
                buttons_del[0].destroy()
                buttons_minus_skill[0].destroy()
                buttons_plus_skill[0].destroy()
                labels_min_skill_pts[0].destroy()
                labels_max_skill_pts[0].destroy()

                label_spare_skill.config(text=f'Свободные очки:  {spare_skill_pts}')
                del labels_skill[0]
                del labels_pts[0]
                del buttons_del[0]
                del buttons_minus_skill[0]
                del buttons_plus_skill[0]
                del labels_min_skill_pts[0]
                del labels_max_skill_pts[0]
            list_skill_for_display = []
            list_skill_pts_for_display = []
            label_spare_skill.config(text=f'Свободные очки:  {spare_skill_pts}')
        
        def generate_params():
            global spare_param_pts
            nonlocal list_param_pts
            nonlocal list_min_param
            nonlocal list_max_param
            param_pts = spare_param_pts
            i = 1
            while i <= param_pts:
                random_index = random.randint(1, 11) - 1
                if list_param_pts[random_index] < list_min_param[random_index] + 2 and list_param_pts[random_index] < list_max_param[random_index]:
                    list_param_pts[random_index] += 1
                    spare_param_pts -= 1
                    i += 1
            update_pts()

        def generate_skills():
            global spare_skill_pts
            nonlocal list_skill
            nonlocal list_skill_pts
            nonlocal list_skill_for_display
            nonlocal list_skill_pts_for_display
            skill_pts = spare_skill_pts
            i = 1
            while i <= skill_pts:
                random_index = random.randint(1, 26) - 1
                if list_skill_pts[random_index] < 1 + 3:
                    list_skill_pts[random_index] += 1
                    spare_skill_pts -= 1
                    i += 1
            for index, value in enumerate(list_skill_pts):
                if value > 0:
                    list_skill_for_display.append(list_skill[index])
                    list_skill_pts_for_display.append(value)
            list_min_skill = [1] * 15
            list_max_skill = [5] * 15
            y_pos = 170

            for index, value in enumerate(list_skill_pts_for_display):
                label_skill = Label(new_tab, text=f'{list_skill_for_display[index]}', justify='left')
                label_pts = Label(new_tab, text=f'{value}', justify='left')
                label_skill.place(x=420, y=y_pos)
                label_pts.place(x=560, y=y_pos)

                labels_skill.append(label_skill)
                labels_pts.append(label_pts)

                btn_del = Button(new_tab, text="X", command=lambda idx=index: delete_skill(idx))
                btn_del.place(x=400, y=y_pos, width=19, height=19)
                buttons_del.append(btn_del)

                btn_minus_skill = Button(new_tab, text="-", command=lambda idx=index: subtract_skill_pts(idx))
                btn_plus_skill = Button(new_tab, text="+", command=lambda idx=index: add_skill_pts(idx))
                btn_minus_skill.place(x=580, y=y_pos, width=19, height=19)
                btn_plus_skill.place(x=610, y=y_pos, width=19, height=19)
                if value == 1:
                    btn_minus_skill.config(state="disabled")
                buttons_minus_skill.append(btn_minus_skill)
                buttons_plus_skill.append(btn_plus_skill)

                label_min_skill_pts = Label(new_tab, text=f'{list_min_skill[index]}', justify='left')
                label_max_skill_pts = Label(new_tab, text=f'{list_max_skill[index]}', justify='left')
                label_min_skill_pts.place(x=660, y=y_pos)
                label_max_skill_pts.place(x=730, y=y_pos)
                labels_min_skill_pts.append(label_min_skill_pts)
                labels_max_skill_pts.append(label_max_skill_pts)
                y_pos += 20

            def update_label(index, new_value):
                global spare_skill_pts
                nonlocal label_spare_skill
                labels_pts[index].config(text=f'{new_value}')
                label_spare_skill.config(text=f'Свободные очки:  {spare_skill_pts}')

            def delete_skill(index):
                global spare_skill_pts
                nonlocal label_spare_skill
                nonlocal y_pos
                labels_skill[index].destroy()
                labels_pts[index].destroy()
                buttons_del[index].destroy()
                buttons_minus_skill[index].destroy()
                buttons_plus_skill[index].destroy()
                labels_min_skill_pts[index].destroy()
                labels_max_skill_pts[index].destroy()
                spare_skill_pts += list_skill_pts_for_display[index]
                label_spare_skill.config(text=f'Свободные очки:  {spare_skill_pts}')

                del labels_skill[index]
                del labels_pts[index]
                del buttons_del[index]
                del buttons_minus_skill[index]
                del buttons_plus_skill[index]
                del labels_min_skill_pts[index]
                del labels_max_skill_pts[index]
                del list_skill_pts_for_display[index]
                del list_skill_for_display[index]

                for idx in range(index, len(labels_skill)):
                    labels_skill[idx].place_configure(y=labels_skill[idx].winfo_y() - 20)
                    labels_pts[idx].place_configure(y=labels_pts[idx].winfo_y() - 20)
                    buttons_del[idx].place_configure(y=buttons_del[idx].winfo_y() - 20)
                    buttons_minus_skill[idx].place_configure(y=buttons_minus_skill[idx].winfo_y() - 20)
                    buttons_plus_skill[idx].place_configure(y=buttons_plus_skill[idx].winfo_y() - 20)
                    labels_min_skill_pts[idx].place_configure(y=labels_min_skill_pts[idx].winfo_y() - 20)
                    labels_max_skill_pts[idx].place_configure(y=labels_max_skill_pts[idx].winfo_y() - 20)

                for i, button in enumerate(buttons_del):
                    button.config(command=lambda idx=i: delete_skill(idx))
                for i, button in enumerate(buttons_minus_skill):
                    button.config(command=lambda idx=i: subtract_skill_pts(idx))
                for i, button in enumerate(buttons_plus_skill):
                    button.config(command=lambda idx=i: add_skill_pts(idx))

                y_pos -= 20
                update_btn_state()

            def subtract_skill_pts(index):
                global spare_skill_pts
                if list_skill_pts_for_display[index] > 1:
                    list_skill_pts_for_display[index] -= 1
                    spare_skill_pts += 1
                    update_label(index, list_skill_pts_for_display[index])
                    update_btn_state()
                if list_skill_pts_for_display[index] == 1:
                    buttons_minus_skill[index].config(state="disabled")
                if list_skill_pts_for_display[index] < 5:
                    buttons_plus_skill[index].config(state="normal")

            def add_skill_pts(index):
                global spare_skill_pts
                if list_skill_pts_for_display[index] < 5 and spare_skill_pts > 0:
                    list_skill_pts_for_display[index] += 1
                    spare_skill_pts -= 1
                    update_label(index, list_skill_pts_for_display[index])
                    update_btn_state()
                    if list_skill_pts_for_display[index] == 5:
                        buttons_plus_skill[index].config(state="disabled")
                    if list_skill_pts_for_display[index] > 1:
                        buttons_minus_skill[index].config(state="normal")

            def add_skill():
                skill_window = tk.Toplevel(window)
                skill_window.title("Добавление навыка")
                window_width = 300
                window_height = 220

                screen_width = window.winfo_screenwidth()
                screen_height = window.winfo_screenheight()
                
                position_top = int(screen_height / 2 - window_height / 2)
                position_right = int(screen_width / 2 - window_width / 2)

                skill_window.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

                skill_window.resizable(False, False)

                skill_window.transient(window)
                skill_window.grab_set()

                tk.Label(skill_window, text="Выберите навык из списка:").pack(pady=10)
                skill_category = ttk.Combobox(skill_window, values=[''] + list_skill, state="readonly")
                skill_category.pack(pady=5)

                tk.Label(skill_window, text="ИЛИ").pack(pady=5)
                tk.Label(skill_window, text="Введите название навыка:").pack()
                skill_entry = tk.Entry(skill_window, width=25)
                skill_entry.pack(pady=5)
                
                def on_continue():
                    nonlocal labels_skill
                    nonlocal labels_pts
                    nonlocal buttons_del
                    nonlocal buttons_minus_skill
                    nonlocal buttons_plus_skill
                    nonlocal labels_min_skill_pts
                    nonlocal labels_max_skill_pts
                    nonlocal y_pos
                    nonlocal list_skill_for_display
                    nonlocal list_skill_pts_for_display
                    global spare_skill_pts
                    skill_name = skill_entry.get()
                    if skill_name:
                        list_skill_for_display.append(skill_name)
                        list_skill_pts_for_display.append(1)
                        label_skill = Label(new_tab, text=f'{skill_name}', justify='left')
                        label_pts = Label(new_tab, text='1', justify='left')
                        label_skill.place(x=420, y=y_pos)
                        label_pts.place(x=560, y=y_pos)

                        labels_skill.append(label_skill)
                        labels_pts.append(label_pts)

                        local_idx = len(list_skill_pts_for_display) - 1
                        btn_del = Button(new_tab, text="X", command=lambda idx=local_idx: delete_skill(idx))
                        btn_del.place(x=400, y=y_pos, width=19, height=19)
                        buttons_del.append(btn_del)

                        btn_minus_skill = Button(new_tab, text="-", command=lambda idx=local_idx: subtract_skill_pts(idx))
                        btn_plus_skill = Button(new_tab, text="+", command=lambda idx=local_idx: add_skill_pts(idx))
                        btn_minus_skill.place(x=580, y=y_pos, width=19, height=19)
                        btn_plus_skill.place(x=610, y=y_pos, width=19, height=19)
                        btn_minus_skill.config(state="disabled")
                        buttons_minus_skill.append(btn_minus_skill)
                        buttons_plus_skill.append(btn_plus_skill)

                        label_min_skill_pts = Label(new_tab, text=f'{list_min_skill[local_idx]}', justify='left')
                        label_max_skill_pts = Label(new_tab, text=f'{list_max_skill[local_idx]}', justify='left')
                        label_min_skill_pts.place(x=660, y=y_pos)
                        label_max_skill_pts.place(x=730, y=y_pos)
                        labels_min_skill_pts.append(label_min_skill_pts)
                        labels_max_skill_pts.append(label_max_skill_pts)
                        y_pos += 20
                    else:
                        skill_name = skill_category.get()
                        if skill_name:
                            list_skill_for_display.append(skill_name)
                            label_skill = Label(new_tab, text=f'{skill_name}', justify='left')
                            label_pts = Label(new_tab, text='1', justify='left')
                            label_skill.place(x=420, y=y_pos)
                            label_pts.place(x=560, y=y_pos)
                            labels_skill.append(label_skill)
                            labels_pts.append(label_pts)

                            local_idx = len(list_skill_pts_for_display) - 1
                            btn_del = Button(new_tab, text="X", command=lambda idx=local_idx: delete_skill(idx))
                            btn_del.place(x=400, y=y_pos, width=19, height=19)
                            buttons_del.append(btn_del)

                            btn_minus_skill = Button(new_tab, text="-", command=lambda idx=local_idx: subtract_skill_pts(idx))
                            btn_plus_skill = Button(new_tab, text="+", command=lambda idx=local_idx: add_skill_pts(idx))
                            btn_minus_skill.place(x=580, y=y_pos, width=19, height=19)
                            btn_plus_skill.place(x=610, y=y_pos, width=19, height=19)
                            buttons_minus_skill.append(btn_minus_skill)
                            buttons_plus_skill.append(btn_plus_skill)

                            label_min_skill_pts = Label(new_tab, text=f'{list_min_skill[local_idx]}', justify='left')
                            label_max_skill_pts = Label(new_tab, text=f'{list_max_skill[local_idx]}', justify='left')
                            label_min_skill_pts.place(x=660, y=y_pos)
                            label_max_skill_pts.place(x=730, y=y_pos)
                            labels_min_skill_pts.append(label_min_skill_pts)
                            labels_max_skill_pts.append(label_max_skill_pts)
                            y_pos += 20
                    spare_skill_pts -= 1
                    update_btn_state()
                    label_spare_skill.config(text=f'Свободные очки:  {spare_skill_pts}')
                    skill_window.destroy()
                
                def on_cancel():
                    """Comment"""
                    skill_window.destroy()
                
                button_frame = tk.Frame(skill_window)
                button_frame.pack(pady=10)
                tk.Button(button_frame, text="Отмена", command=on_cancel).pack(side=tk.LEFT, padx=5)
                tk.Button(button_frame, text="Продолжить", command=on_continue).pack(side=tk.LEFT, padx=5)
                window.wait_window(skill_window)

            def update_btn_state():
                if spare_skill_pts > 0:
                    btn_add_skill.config(state="normal")
                else:
                    btn_add_skill.config(state="disabled")

            btn_add_skill = Button(new_tab, text="Добавить навык", command=add_skill)
            btn_add_skill.place(x=600, y=490)
            label_spare_skill.config(text=f'Свободные очки:  {spare_skill_pts}')
            update_btn_state()

        combobox.bind("<<ComboboxSelected>>", update_points)
        generate_params()
        generate_skills()

        separator = tk.Frame(new_tab, height=2, bd=1, relief=tk.SUNKEN)
        separator.place(relx=0, rely=0.15, relwidth=1)

        btn_close = Button(new_tab, text = "Закрыть вкладку", command=close_tab)
        btn_save = Button(new_tab, text = "Сохранить", command=save_file)

        label_params = Label(new_tab, text=f"Параметры")
        label_param_values = Label(new_tab, text=f"Значение")
        label_min_param = Label(new_tab, text=f"Мин. знач")
        label_max_param = Label(new_tab, text=f"Макс. знач")
        label_skills = Label(new_tab, text=f"Навыки")
        label_skill_values = Label(new_tab, text=f"Значение")
        label_min_skill = Label(new_tab, text=f"Мин. знач")
        label_max_skill = Label(new_tab, text=f"Макс. знач")

        label_param1 = Label(new_tab, text='Сила', justify='left')
        label_param2 = Label(new_tab, text='Тело', justify='left')
        label_param3 = Label(new_tab, text='Точность', justify='left')
        label_param4 = Label(new_tab, text='Подвижность', justify='left')
        label_param5 = Label(new_tab, text='Интеллект', justify='left')
        label_param6 = Label(new_tab, text='Сообразительность', justify='left')
        label_param7 = Label(new_tab, text='Восприятие', justify='left')
        label_param8 = Label(new_tab, text='Фантазия', justify='left')
        label_param9 = Label(new_tab, text='Воля', justify='left')
        label_param10 = Label(new_tab, text='Харизма', justify='left')
        label_param11 = Label(new_tab, text='Внешность', justify='left')

        btn_minus_param1 = Button(new_tab, text="-", command=lambda: subtract_pts(0))
        btn_plus_param1 = Button(new_tab, text="+", command=lambda: add_pts(0))
        btn_minus_param2 = Button(new_tab, text="-", command=lambda: subtract_pts(1))
        btn_plus_param2 = Button(new_tab, text="+", command=lambda: add_pts(1))
        btn_minus_param3 = Button(new_tab, text="-", command=lambda: subtract_pts(2))
        btn_plus_param3 = Button(new_tab, text="+", command=lambda: add_pts(2))
        btn_minus_param4 = Button(new_tab, text="-", command=lambda: subtract_pts(3))
        btn_plus_param4 = Button(new_tab, text="+", command=lambda: add_pts(3))
        btn_minus_param5 = Button(new_tab, text="-", command=lambda: subtract_pts(4))
        btn_plus_param5 = Button(new_tab, text="+", command=lambda: add_pts(4))
        btn_minus_param6 = Button(new_tab, text="-", command=lambda: subtract_pts(5))
        btn_plus_param6 = Button(new_tab, text="+", command=lambda: add_pts(5))
        btn_minus_param7 = Button(new_tab, text="-", command=lambda: subtract_pts(6))
        btn_plus_param7 = Button(new_tab, text="+", command=lambda: add_pts(6))
        btn_minus_param8 = Button(new_tab, text="-", command=lambda: subtract_pts(7))
        btn_plus_param8 = Button(new_tab, text="+", command=lambda: add_pts(7))    
        btn_minus_param9 = Button(new_tab, text="-", command=lambda: subtract_pts(8))
        btn_plus_param9 = Button(new_tab, text="+", command=lambda: add_pts(8))
        btn_minus_param10 = Button(new_tab, text="-", command=lambda: subtract_pts(9))
        btn_plus_param10 = Button(new_tab, text="+", command=lambda: add_pts(9))
        btn_minus_param11 = Button(new_tab, text="-", command=lambda: subtract_pts(10))
        btn_plus_param11 = Button(new_tab, text="+", command=lambda: add_pts(10))

        btn_minus_param1.place(x=190, y=170, width=19, height=19)
        btn_plus_param1.place(x=220, y=170, width=19, height=19)
        btn_minus_param2.place(x=190, y=190, width=19, height=19)
        btn_plus_param2.place(x=220, y=190, width=19, height=19)
        btn_minus_param3.place(x=190, y=210, width=19, height=19)
        btn_plus_param3.place(x=220, y=210, width=19, height=19)
        btn_minus_param4.place(x=190, y=230, width=19, height=19)
        btn_plus_param4.place(x=220, y=230, width=19, height=19)
        btn_minus_param5.place(x=190, y=250, width=19, height=19)
        btn_plus_param5.place(x=220, y=250, width=19, height=19)
        btn_minus_param6.place(x=190, y=270, width=19, height=19)
        btn_plus_param6.place(x=220, y=270, width=19, height=19)
        btn_minus_param7.place(x=190, y=290, width=19, height=19)
        btn_plus_param7.place(x=220, y=290, width=19, height=19)
        btn_minus_param8.place(x=190, y=310, width=19, height=19)
        btn_plus_param8.place(x=220, y=310, width=19, height=19)
        btn_minus_param9.place(x=190, y=330, width=19, height=19)
        btn_plus_param9.place(x=220, y=330, width=19, height=19)
        btn_minus_param10.place(x=190, y=350, width=19, height=19)
        btn_plus_param10.place(x=220, y=350, width=19, height=19)
        btn_minus_param11.place(x=190, y=370, width=19, height=19)
        btn_plus_param11.place(x=220, y=370, width=19, height=19)

        label_params.place(x=40, y=140)
        label_param_values.place(x=145, y=140)
        label_min_param.place(x=245, y=140)
        label_max_param.place(x=315, y=140)
        label_skills.place(x=420, y=140)
        label_skill_values.place(x=535, y=140)
        label_min_skill.place(x=635, y=140)
        label_max_skill.place(x=705, y=140)

        label_param1.place(x=40, y=170)
        label_param2.place(x=40, y=190)
        label_param3.place(x=40, y=210)
        label_param4.place(x=40, y=230)
        label_param5.place(x=40, y=250)
        label_param6.place(x=40, y=270)
        label_param7.place(x=40, y=290)
        label_param8.place(x=40, y=310)
        label_param9.place(x=40, y=330)
        label_param10.place(x=40, y=350)
        label_param11.place(x=40, y=370)

        label_param_pt1.place(x=170, y=170)
        label_param_pt2.place(x=170, y=190)
        label_param_pt3.place(x=170, y=210)
        label_param_pt4.place(x=170, y=230)
        label_param_pt5.place(x=170, y=250)
        label_param_pt6.place(x=170, y=270)
        label_param_pt7.place(x=170, y=290)
        label_param_pt8.place(x=170, y=310)
        label_param_pt9.place(x=170, y=330)
        label_param_pt10.place(x=170, y=350)
        label_param_pt11.place(x=170, y=370)

        label_min_param1.place(x=270, y=170)
        label_min_param2.place(x=270, y=190)
        label_min_param3.place(x=270, y=210)
        label_min_param4.place(x=270, y=230)
        label_min_param5.place(x=270, y=250)
        label_min_param6.place(x=270, y=270)
        label_min_param7.place(x=270, y=290)
        label_min_param8.place(x=270, y=310)
        label_min_param9.place(x=270, y=330)
        label_min_param10.place(x=270, y=350)
        label_min_param11.place(x=270, y=370)

        label_max_param1.place(x=340, y=170)
        label_max_param2.place(x=340, y=190)
        label_max_param3.place(x=340, y=210)
        label_max_param4.place(x=340, y=230)
        label_max_param5.place(x=340, y=250)
        label_max_param6.place(x=340, y=270)
        label_max_param7.place(x=340, y=290)
        label_max_param8.place(x=340, y=310)
        label_max_param9.place(x=340, y=330)
        label_max_param10.place(x=340, y=350)
        label_max_param11.place(x=340, y=370)

        def gen_random():
            global spare_param_pts
            global spare_skill_pts
            nonlocal list_min_param
            nonlocal list_max_param
            global race
            nonlocal list_param_pts
            nonlocal list_skill_pts
            nonlocal last_selected
            global spare_skill_pts
            nonlocal label_spare_skill
            nonlocal labels_skill
            nonlocal labels_pts
            nonlocal buttons_del
            nonlocal buttons_minus_skill
            nonlocal buttons_plus_skill
            nonlocal labels_min_skill_pts
            nonlocal labels_max_skill_pts
            nonlocal list_skill_for_display
            nonlocal list_skill_pts_for_display
            race = list_races[random.randint(1, 5) - 1]
            if race == 'Люди':
                spare_param_pts = 12
                spare_skill_pts = 15
                list_min_param = [1] * 11
                list_max_param = [5, 5, 5, 5, 5, 6, 5, 6, 5, 5, 5]
                for index in range(11):
                    list_param_pts[index] = list_min_param[index]
                for index in range(26):
                    list_skill_pts[index] = 0
                entry_name.delete(0, END)
                entry_name.insert(0, list_human_names[random.randint(1, 5) - 1])
                combobox.set('Люди')
                last_selected = 'Люди'
            if race == 'Эльфы':
                spare_param_pts = 9
                spare_skill_pts = 12
                list_min_param = [1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 2]
                list_max_param = [5, 5, 5, 5, 5, 5, 6, 5, 5, 5, 7]
                for index in range(11):
                    list_param_pts[index] = list_min_param[index]
                for index in range(26):
                    list_skill_pts[index] = 0
                entry_name.delete(0, END)
                entry_name.insert(0, list_elf_names[random.randint(1, 5) - 1])
                combobox.set('Эльфы')
                last_selected = 'Эльфы'
            if race == 'Гномы':
                spare_param_pts = 10
                spare_skill_pts = 12
                list_min_param = [2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1]
                list_max_param = [6, 6, 5, 5, 6, 5, 5, 5, 5, 5, 5]
                for index in range(11):
                    list_param_pts[index] = list_min_param[index]
                for index in range(26):
                    list_skill_pts[index] = 0
                entry_name.delete(0, END)
                entry_name.insert(0, list_dwarf_names[random.randint(1, 5) - 1])
                combobox.set('Гномы')
                last_selected = 'Гномы'
            if race == 'Орки':
                spare_param_pts = 10
                spare_skill_pts = 12
                list_min_param = [2, 2, 1, 1, 1, 1, 1, 1, 2, 1, 1]
                list_max_param = [7, 6, 5, 5, 5, 5, 5, 4, 7, 4, 4]
                for index in range(11):
                    list_param_pts[index] = list_min_param[index]
                for index in range(26):
                    list_skill_pts[index] = 0
                entry_name.delete(0, END)
                entry_name.insert(0, list_orc_names[random.randint(1, 5) - 1])
                combobox.set('Орки')
                last_selected = 'Орки'
            if race == 'Нежить':
                spare_param_pts = 9
                spare_skill_pts = 6
                list_min_param = [1] * 11
                list_max_param = [9, 3, 4, 4, 5, 4, 7, 3, 4, 3, 3]
                for index in range(11):
                    list_param_pts[index] = list_min_param[index]
                for index in range(26):
                    list_skill_pts[index] = 0
                entry_name.delete(0, END)
                entry_name.insert(0, list_undead_names[random.randint(1, 5) - 1])
                combobox.set('Нежить')
                last_selected = 'Нежить'
            generate_params()
            while labels_skill:
                labels_skill[0].destroy()
                labels_pts[0].destroy()
                buttons_del[0].destroy()
                buttons_minus_skill[0].destroy()
                buttons_plus_skill[0].destroy()
                labels_min_skill_pts[0].destroy()
                labels_max_skill_pts[0].destroy()

                label_spare_skill.config(text=f'Свободные очки:  {spare_skill_pts}')
                del labels_skill[0]
                del labels_pts[0]
                del buttons_del[0]
                del buttons_minus_skill[0]
                del buttons_plus_skill[0]
                del labels_min_skill_pts[0]
                del labels_max_skill_pts[0]
            list_skill_for_display = []
            list_skill_pts_for_display = []
            generate_skills()
            label_spare_skill.config(text=f'Свободные очки:  {spare_skill_pts}')
            

        btn_random = Button(new_tab, text="Случайный персонаж", command=gen_random)
        btn_random.place(x=325, y=20)
        label_random = Label(new_tab, text="При нажатии на кнопку \"Случайный персонаж\" раса, имя, параметры и навыки генерируются автоматически")
        label_random.place(x=83, y=55)

        btn_save.place(x=540, y=530)
        btn_close.place(x=160, y=530)
        gen_tab_counter += 1
        update_tabmenu()

        tab_control.select(new_tab)
    
    else:
        def read_file():
            nonlocal list_param
            nonlocal list_param_pts
            nonlocal list_skill_for_display
            nonlocal list_skill_pts_for_display
            global selected_file_path
            global race
            nonlocal name
            global spare_param_pts
            global spare_skill_pts
            global gen_tab_counter
            nonlocal labels_skill
            nonlocal labels_pts
            nonlocal buttons_del
            nonlocal buttons_minus_skill
            nonlocal buttons_plus_skill
            nonlocal labels_min_skill_pts
            nonlocal labels_max_skill_pts
            nonlocal list_min_param
            nonlocal list_max_param
            global is_opened_from_list
            nonlocal combobox
            nonlocal entry_name

            if not is_opened_from_list:
                file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
                selected_file_path = file_path
            if selected_file_path:
                new_tab = Frame(tab_control)
                tab_control.add(new_tab, text=f"Ген. персонажа {gen_tab_counter}")
                tabmenu.add_command(label=f"Ген. персонажа {gen_tab_counter}")

                last_selected = None
                label_race = Label(new_tab, text="Раса:")
                combobox = ttk.Combobox(new_tab, values=list_races, width=15, state="readonly")
                combobox.set("Выберите расу")
                label_race.place(x=110, y=105)
                combobox.place(x=150, y=105)

                label_name = Label(new_tab, text="Имя:")
                entry_name = Entry(new_tab, width=40)
                label_name.place(x=430, y=105)
                entry_name.place(x=470, y=105)

                df_new = pd.read_excel(selected_file_path, header=None)
                race = df_new.iloc[0, 0] 
                name =  df_new.iloc[0, 1]
                list_param = df_new.iloc[1:12, 0].tolist()
                list_param_pts = df_new.iloc[1:12, 1].tolist()
                spare_param_pts = df_new.iloc[12, 1]
                list_combined = df_new.iloc[14:, 0].dropna().tolist()
                list_skill_for_display = list_combined[:-1]
                list_pts_combined = df_new.iloc[14:, 1].dropna().tolist()
                list_skill_pts_for_display = list_pts_combined[:-1]
                spare_skill_pts = list_pts_combined[-1]
                if race == 'Люди':
                    list_min_param = [1] * 11
                    list_max_param = [5, 5, 5, 5, 5, 6, 5, 6, 5, 5, 5]
                    combobox.set('Люди')
                    last_selected = 'Люди'
                    entry_name.delete(0, END)
                    entry_name.insert(0, name)
                if race == 'Эльфы':
                    list_min_param = [1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 2]
                    list_max_param = [5, 5, 5, 5, 5, 5, 6, 5, 5, 5, 7]
                    combobox.set('Эльфы')
                    last_selected = 'Эльфы'
                    entry_name.delete(0, END)
                    entry_name.insert(0, name)
                if race == 'Гномы':
                    list_min_param = [2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1]
                    list_max_param = [6, 6, 5, 5, 6, 5, 5, 5, 5, 5, 5]
                    combobox.set('Гномы')
                    last_selected = 'Гномы'
                    entry_name.delete(0, END)
                    entry_name.insert(0, name)
                if race == 'Орки':
                    list_min_param = [2, 2, 1, 1, 1, 1, 1, 1, 2, 1, 1]
                    list_max_param = [7, 6, 5, 5, 5, 5, 5, 4, 7, 4, 4]
                    combobox.set('Орки')
                    last_selected = 'Орки'
                    entry_name.delete(0, END)
                    entry_name.insert(0, name)
                if race == 'Нежить':
                    list_min_param = [1] * 11
                    list_max_param = [9, 3, 4, 4, 5, 4, 7, 3, 4, 3, 3]
                    combobox.set('Нежить')
                    last_selected = 'Нежить'
                    entry_name.delete(0, END)
                    entry_name.insert(0, name)
                update_log()
            else:
                pass

            def read_name():
                nonlocal name
                name = entry_name.get()

            def add_pts(index):
                global spare_param_pts
                nonlocal list_param_pts
                nonlocal list_max_param
                nonlocal list_min_param
                if spare_param_pts > 0 and list_param_pts[index] < list_max_param[index]:
                    spare_param_pts -= 1
                    list_param_pts[index] += 1
                    update_pts()
                if list_param_pts[index] == list_max_param[index]:
                    if index == 0:
                        btn_plus_param1.config(state="disabled")
                    if index == 1:
                        btn_plus_param2.config(state="disabled")
                    if index == 2:
                        btn_plus_param3.config(state="disabled")
                    if index == 3:
                        btn_plus_param4.config(state="disabled")
                    if index == 4:
                        btn_plus_param5.config(state="disabled")
                    if index == 5:
                        btn_plus_param6.config(state="disabled")
                    if index == 6:
                        btn_plus_param7.config(state="disabled")
                    if index == 7:
                        btn_plus_param8.config(state="disabled")
                    if index == 8:
                        btn_plus_param9.config(state="disabled")
                    if index == 9:
                        btn_plus_param10.config(state="disabled")
                    if index == 10:
                        btn_plus_param11.config(state="disabled")
                if list_param_pts[index] > list_min_param[index]:
                    if index == 0:
                        btn_minus_param1.config(state="normal")
                    if index == 1:
                        btn_minus_param2.config(state="normal")
                    if index == 2:
                        btn_minus_param3.config(state="normal")
                    if index == 3:
                        btn_minus_param4.config(state="normal")
                    if index == 4:
                        btn_minus_param5.config(state="normal")
                    if index == 5:
                        btn_minus_param6.config(state="normal")
                    if index == 6:
                        btn_minus_param7.config(state="normal")
                    if index == 7:
                        btn_minus_param8.config(state="normal")
                    if index == 8:
                        btn_minus_param9.config(state="normal")
                    if index == 9:
                        btn_minus_param10.config(state="normal")
                    if index == 10:
                        btn_minus_param11.config(state="normal")

            def subtract_pts(index):
                global spare_param_pts
                nonlocal list_param_pts
                nonlocal list_min_param
                nonlocal list_max_param
                if list_param_pts[index] > list_min_param[index]:
                    spare_param_pts += 1
                    list_param_pts[index] -= 1
                update_pts()
                if list_param_pts[index] == list_min_param[index]:
                    if index == 0:
                        btn_minus_param1.config(state="disabled")
                    if index == 1:
                        btn_minus_param2.config(state="disabled")
                    if index == 2:
                        btn_minus_param3.config(state="disabled")
                    if index == 3:
                        btn_minus_param4.config(state="disabled")
                    if index == 4:
                        btn_minus_param5.config(state="disabled")
                    if index == 5:
                        btn_minus_param6.config(state="disabled")
                    if index == 6:
                        btn_minus_param7.config(state="disabled")
                    if index == 7:
                        btn_minus_param8.config(state="disabled")
                    if index == 8:
                        btn_minus_param9.config(state="disabled")
                    if index == 9:
                        btn_minus_param10.config(state="disabled")
                    if index == 10:
                        btn_minus_param11.config(state="disabled")
                if list_param_pts[index] < list_max_param[index]:
                    if index == 0:
                        btn_plus_param1.config(state="normal")
                    if index == 1:
                        btn_plus_param2.config(state="normal")
                    if index == 2:
                        btn_plus_param3.config(state="normal")
                    if index == 3:
                        btn_plus_param4.config(state="normal")
                    if index == 4:
                        btn_plus_param5.config(state="normal")
                    if index == 5:
                        btn_plus_param6.config(state="normal")
                    if index == 6:
                        btn_plus_param7.config(state="normal")
                    if index == 7:
                        btn_plus_param8.config(state="normal")
                    if index == 8:
                        btn_plus_param9.config(state="normal")
                    if index == 9:
                        btn_plus_param10.config(state="normal")
                    if index == 10:
                        btn_plus_param11.config(state="normal")

            def save_file():
                nonlocal list_param
                nonlocal list_skill_for_display
                nonlocal list_skill_pts_for_display
                global selected_file_path
                global race
                nonlocal name
                global spare_param_pts
                global spare_skill_pts
                read_name()
                df = pd.DataFrame(index=range(33), columns=['A', 'B'])
                df['A'][0] = race
                df['A'][1:12] = list_param
                df['A'][12] = 'Свободные очки:'
                df['A'][14:14 + len(list_skill_for_display)] = list_skill_for_display
                df['A'][14 + len(list_skill_for_display)] = 'Свободные очки:'
                df['B'][0] = name
                df['B'][1:12] = list_param_pts
                df['B'][12] = spare_param_pts
                df['B'][14:14 + len(list_skill_for_display)] = list_skill_pts_for_display
                df['B'][14 + len(list_skill_for_display)] = spare_skill_pts
                file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
                selected_file_path = file_path
                if file_path:
                    df.to_excel(file_path, index=False, header=False)
                    update_log()
                else:
                    pass

            label_spare_param = Label(new_tab, text=f'Свободные очки:  {spare_param_pts}')
            label_spare_skill = Label(new_tab, text=f'Свободные очки:  {spare_skill_pts}')
            label_spare_param.place(x=40, y=490)
            label_spare_skill.place(x=420, y=490)

            label_param_pt1 = Label(new_tab, text=f'{list_param_pts[0]}', justify='left')
            label_param_pt2 = Label(new_tab, text=f'{list_param_pts[1]}', justify='left')
            label_param_pt3 = Label(new_tab, text=f'{list_param_pts[2]}', justify='left')
            label_param_pt4 = Label(new_tab, text=f'{list_param_pts[3]}', justify='left')
            label_param_pt5 = Label(new_tab, text=f'{list_param_pts[4]}', justify='left')
            label_param_pt6 = Label(new_tab, text=f'{list_param_pts[5]}', justify='left')
            label_param_pt7 = Label(new_tab, text=f'{list_param_pts[6]}', justify='left')
            label_param_pt8 = Label(new_tab, text=f'{list_param_pts[7]}', justify='left')
            label_param_pt9 = Label(new_tab, text=f'{list_param_pts[8]}', justify='left')
            label_param_pt10 = Label(new_tab, text=f'{list_param_pts[9]}', justify='left')
            label_param_pt11 = Label(new_tab, text=f'{list_param_pts[10]}', justify='left')

            label_min_param1 = Label(new_tab, text=f'{list_min_param[0]}', justify='left')
            label_min_param2 = Label(new_tab, text=f'{list_min_param[1]}', justify='left')
            label_min_param3 = Label(new_tab, text=f'{list_min_param[2]}', justify='left')
            label_min_param4 = Label(new_tab, text=f'{list_min_param[3]}', justify='left')
            label_min_param5 = Label(new_tab, text=f'{list_min_param[4]}', justify='left')
            label_min_param6 = Label(new_tab, text=f'{list_min_param[5]}', justify='left')
            label_min_param7 = Label(new_tab, text=f'{list_min_param[6]}', justify='left')
            label_min_param8 = Label(new_tab, text=f'{list_min_param[7]}', justify='left')
            label_min_param9 = Label(new_tab, text=f'{list_min_param[8]}', justify='left')
            label_min_param10 = Label(new_tab, text=f'{list_min_param[9]}', justify='left')
            label_min_param11 = Label(new_tab, text=f'{list_min_param[10]}', justify='left')

            label_max_param1 = Label(new_tab, text=f'{list_max_param[0]}', justify='left')
            label_max_param2 = Label(new_tab, text=f'{list_max_param[1]}', justify='left')
            label_max_param3 = Label(new_tab, text=f'{list_max_param[2]}', justify='left')
            label_max_param4 = Label(new_tab, text=f'{list_max_param[3]}', justify='left')
            label_max_param5 = Label(new_tab, text=f'{list_max_param[4]}', justify='left')
            label_max_param6 = Label(new_tab, text=f'{list_max_param[5]}', justify='left')
            label_max_param7 = Label(new_tab, text=f'{list_max_param[6]}', justify='left')
            label_max_param8 = Label(new_tab, text=f'{list_max_param[7]}', justify='left')
            label_max_param9 = Label(new_tab, text=f'{list_max_param[8]}', justify='left')
            label_max_param10 = Label(new_tab, text=f'{list_max_param[9]}', justify='left')
            label_max_param11 = Label(new_tab, text=f'{list_max_param[10]}', justify='left')

            def update_pts():
                global spare_param_pts
                nonlocal label_spare_param
                label_param_pt1.config(text=f'{list_param_pts[0]}')
                label_param_pt2.config(text=f'{list_param_pts[1]}')
                label_param_pt3.config(text=f'{list_param_pts[2]}')
                label_param_pt4.config(text=f'{list_param_pts[3]}')
                label_param_pt5.config(text=f'{list_param_pts[4]}')
                label_param_pt6.config(text=f'{list_param_pts[5]}')
                label_param_pt7.config(text=f'{list_param_pts[6]}')
                label_param_pt8.config(text=f'{list_param_pts[7]}')
                label_param_pt9.config(text=f'{list_param_pts[8]}')
                label_param_pt10.config(text=f'{list_param_pts[9]}')
                label_param_pt11.config(text=f'{list_param_pts[10]}')
                label_spare_param.config(text=f'Свободные очки:  {spare_param_pts}')

                label_min_param1.config(text=f'{list_min_param[0]}')
                label_min_param2.config(text=f'{list_min_param[1]}')
                label_min_param3.config(text=f'{list_min_param[2]}')
                label_min_param4.config(text=f'{list_min_param[3]}')
                label_min_param5.config(text=f'{list_min_param[4]}')
                label_min_param6.config(text=f'{list_min_param[5]}')
                label_min_param7.config(text=f'{list_min_param[6]}')
                label_min_param8.config(text=f'{list_min_param[7]}')
                label_min_param9.config(text=f'{list_min_param[8]}')
                label_min_param10.config(text=f'{list_min_param[9]}')
                label_min_param11.config(text=f'{list_min_param[10]}')

                label_max_param1.config(text=f'{list_max_param[0]}')
                label_max_param2.config(text=f'{list_max_param[1]}')
                label_max_param3.config(text=f'{list_max_param[2]}')
                label_max_param4.config(text=f'{list_max_param[3]}')
                label_max_param5.config(text=f'{list_max_param[4]}')
                label_max_param6.config(text=f'{list_max_param[5]}')
                label_max_param7.config(text=f'{list_max_param[6]}')
                label_max_param8.config(text=f'{list_max_param[7]}')
                label_max_param9.config(text=f'{list_max_param[8]}')
                label_max_param10.config(text=f'{list_max_param[9]}')
                label_max_param11.config(text=f'{list_max_param[10]}')

            def update_points(event):
                global spare_param_pts
                global spare_skill_pts
                nonlocal last_selected
                nonlocal list_min_param
                nonlocal list_max_param
                nonlocal labels_skill
                nonlocal labels_pts
                nonlocal buttons_del
                nonlocal buttons_minus_skill
                nonlocal buttons_plus_skill
                nonlocal labels_min_skill_pts
                nonlocal labels_max_skill_pts
                nonlocal list_skill_for_display
                nonlocal list_skill_pts_for_display
                current_selection = combobox.get()
                if current_selection == 'Люди' and last_selected != 'Люди':
                    spare_param_pts = 12
                    spare_skill_pts = 15
                    list_min_param = [1] * 11
                    list_max_param = [5, 5, 5, 5, 5, 6, 5, 6, 5, 5, 5]
                    for index in range(11):
                        list_param_pts[index] = list_min_param[index]
                    for index in range(26):
                        list_skill_pts[index] = 0
                    entry_name.delete(0, END)
                if current_selection == 'Эльфы' and last_selected != 'Эльфы':
                    spare_param_pts = 9
                    spare_skill_pts = 12
                    list_min_param = [1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 2]
                    list_max_param = [5, 5, 5, 5, 5, 5, 6, 5, 5, 5, 7]
                    for index in range(11):
                        list_param_pts[index] = list_min_param[index]
                    for index in range(26):
                        list_skill_pts[index] = 0
                    entry_name.delete(0, END)
                if current_selection == 'Гномы' and last_selected != 'Гномы':
                    spare_param_pts = 10
                    spare_skill_pts = 12
                    list_min_param = [2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1]
                    list_max_param = [6, 6, 5, 5, 6, 5, 5, 5, 5, 5, 5]
                    for index in range(11):
                        list_param_pts[index] = list_min_param[index]
                    for index in range(26):
                        list_skill_pts[index] = 0
                    entry_name.delete(0, END)
                if current_selection == 'Орки' and last_selected != 'Орки':
                    spare_param_pts = 10
                    spare_skill_pts = 12
                    list_min_param = [2, 2, 1, 1, 1, 1, 1, 1, 2, 1, 1]
                    list_max_param = [7, 6, 5, 5, 5, 5, 5, 4, 7, 4, 4]
                    for index in range(11):
                        list_param_pts[index] = list_min_param[index]
                    for index in range(26):
                        list_skill_pts[index] = 0
                    entry_name.delete(0, END)
                if current_selection == 'Нежить' and last_selected != 'Нежить':
                    spare_param_pts = 9
                    spare_skill_pts = 6
                    list_min_param = [1] * 11
                    list_max_param = [9, 3, 4, 4, 5, 4, 7, 3, 4, 3, 3]
                    for index in range(11):
                        list_param_pts[index] = list_min_param[index]
                    for index in range(26):
                        list_skill_pts[index] = 0
                    entry_name.delete(0, END)
                update_pts()
                last_selected = current_selection
                btn_minus_param1.config(state="disabled")
                btn_minus_param2.config(state="disabled")
                btn_minus_param3.config(state="disabled")
                btn_minus_param4.config(state="disabled")
                btn_minus_param5.config(state="disabled")
                btn_minus_param6.config(state="disabled")
                btn_minus_param7.config(state="disabled")
                btn_minus_param8.config(state="disabled")
                btn_minus_param9.config(state="disabled")
                btn_minus_param10.config(state="disabled")
                btn_minus_param11.config(state="disabled")

                while labels_skill:
                    labels_skill[0].destroy()
                    labels_pts[0].destroy()
                    buttons_del[0].destroy()
                    buttons_minus_skill[0].destroy()
                    buttons_plus_skill[0].destroy()
                    labels_min_skill_pts[0].destroy()
                    labels_max_skill_pts[0].destroy()

                    label_spare_skill.config(text=f'Свободные очки:  {spare_skill_pts}')
                    del labels_skill[0]
                    del labels_pts[0]
                    del buttons_del[0]
                    del buttons_minus_skill[0]
                    del buttons_plus_skill[0]
                    del labels_min_skill_pts[0]
                    del labels_max_skill_pts[0]
                list_skill_for_display = []
                list_skill_pts_for_display = []
                label_spare_skill.config(text=f'Свободные очки:  {spare_skill_pts}')

            def generate_params():
                global spare_param_pts
                nonlocal list_param_pts
                nonlocal list_min_param
                nonlocal list_max_param
                param_pts = spare_param_pts
                i = 1
                while i <= param_pts:
                    random_index = random.randint(1, 11) - 1
                    if list_param_pts[random_index] < list_min_param[random_index] + 2 and list_param_pts[random_index] < list_max_param[random_index]:
                        list_param_pts[random_index] += 1
                        spare_param_pts -= 1
                        i += 1
                update_pts()

            def generate_skills():
                global spare_skill_pts
                nonlocal list_skill
                nonlocal list_skill_pts
                nonlocal list_skill_for_display
                nonlocal list_skill_pts_for_display
                skill_pts = spare_skill_pts
                i = 1
                while i <= skill_pts:
                    random_index = random.randint(1, 26) - 1
                    if list_skill_pts[random_index] < 1 + 3:
                        list_skill_pts[random_index] += 1
                        spare_skill_pts -= 1
                        i += 1
                for index, value in enumerate(list_skill_pts):
                    if value > 0:
                        list_skill_for_display.append(list_skill[index])
                        list_skill_pts_for_display.append(value)
                list_min_skill = [1] * 15
                list_max_skill = [5] * 15
                y_pos = 170

                for index, value in enumerate(list_skill_pts_for_display):
                    label_skill = Label(new_tab, text=f'{list_skill_for_display[index]}', justify='left')
                    label_pts = Label(new_tab, text=f'{value}', justify='left')
                    label_skill.place(x=420, y=y_pos)
                    label_pts.place(x=560, y=y_pos)

                    labels_skill.append(label_skill)
                    labels_pts.append(label_pts)

                    btn_del = Button(new_tab, text="X", command=lambda idx=index: delete_skill(idx))
                    btn_del.place(x=400, y=y_pos, width=19, height=19)
                    buttons_del.append(btn_del)

                    btn_minus_skill = Button(new_tab, text="-", command=lambda idx=index: subtract_skill_pts(idx))
                    btn_plus_skill = Button(new_tab, text="+", command=lambda idx=index: add_skill_pts(idx))
                    btn_minus_skill.place(x=580, y=y_pos, width=19, height=19)
                    btn_plus_skill.place(x=610, y=y_pos, width=19, height=19)
                    if value == 1:
                        btn_minus_skill.config(state="disabled")
                    buttons_minus_skill.append(btn_minus_skill)
                    buttons_plus_skill.append(btn_plus_skill)

                    label_min_skill_pts = Label(new_tab, text=f'{list_min_skill[index]}', justify='left')
                    label_max_skill_pts = Label(new_tab, text=f'{list_max_skill[index]}', justify='left')
                    label_min_skill_pts.place(x=660, y=y_pos)
                    label_max_skill_pts.place(x=730, y=y_pos)
                    labels_min_skill_pts.append(label_min_skill_pts)
                    labels_max_skill_pts.append(label_max_skill_pts)
                    y_pos += 20

                def update_label(index, new_value):
                    global spare_skill_pts
                    nonlocal label_spare_skill
                    labels_pts[index].config(text=f'{new_value}')
                    label_spare_skill.config(text=f'Свободные очки:  {spare_skill_pts}')

                def delete_skill(index):
                    global spare_skill_pts
                    nonlocal label_spare_skill
                    nonlocal y_pos
                    labels_skill[index].destroy()
                    labels_pts[index].destroy()
                    buttons_del[index].destroy()
                    buttons_minus_skill[index].destroy()
                    buttons_plus_skill[index].destroy()
                    labels_min_skill_pts[index].destroy()
                    labels_max_skill_pts[index].destroy()
                    spare_skill_pts += list_skill_pts_for_display[index]
                    label_spare_skill.config(text=f'Свободные очки:  {spare_skill_pts}')
                    del labels_skill[index]
                    del labels_pts[index]
                    del buttons_del[index]
                    del buttons_minus_skill[index]
                    del buttons_plus_skill[index]
                    del labels_min_skill_pts[index]
                    del labels_max_skill_pts[index]
                    del list_skill_pts_for_display[index]
                    del list_skill_for_display[index]

                    for idx in range(index, len(labels_skill)):
                        labels_skill[idx].place_configure(y=labels_skill[idx].winfo_y() - 20)
                        labels_pts[idx].place_configure(y=labels_pts[idx].winfo_y() - 20)
                        buttons_del[idx].place_configure(y=buttons_del[idx].winfo_y() - 20)
                        buttons_minus_skill[idx].place_configure(y=buttons_minus_skill[idx].winfo_y() - 20)
                        buttons_plus_skill[idx].place_configure(y=buttons_plus_skill[idx].winfo_y() - 20)
                        labels_min_skill_pts[idx].place_configure(y=labels_min_skill_pts[idx].winfo_y() - 20)
                        labels_max_skill_pts[idx].place_configure(y=labels_max_skill_pts[idx].winfo_y() - 20)

                    for i, button in enumerate(buttons_del):
                        button.config(command=lambda idx=i: delete_skill(idx))
                    for i, button in enumerate(buttons_minus_skill):
                        button.config(command=lambda idx=i: subtract_skill_pts(idx))
                    for i, button in enumerate(buttons_plus_skill):
                        button.config(command=lambda idx=i: add_skill_pts(idx))

                    y_pos -= 20
                    update_btn_state()

                def subtract_skill_pts(index):
                    global spare_skill_pts
                    if list_skill_pts_for_display[index] > 1:
                        list_skill_pts_for_display[index] -= 1
                        spare_skill_pts += 1
                        update_label(index, list_skill_pts_for_display[index])
                        update_btn_state()
                    if list_skill_pts_for_display[index] == 1:
                        buttons_minus_skill[index].config(state="disabled")
                    if list_skill_pts_for_display[index] < 5:
                        buttons_plus_skill[index].config(state="normal")

                def add_skill_pts(index):
                    global spare_skill_pts
                    if list_skill_pts_for_display[index] < 5 and spare_skill_pts > 0:
                        list_skill_pts_for_display[index] += 1
                        spare_skill_pts -= 1
                        update_label(index, list_skill_pts_for_display[index])
                        update_btn_state()
                        if list_skill_pts_for_display[index] == 5:
                            buttons_plus_skill[index].config(state="disabled")
                        if list_skill_pts_for_display[index] > 1:
                            buttons_minus_skill[index].config(state="normal")

                def add_skill():
                    skill_window = tk.Toplevel(window)
                    skill_window.title("Добавление навыка")

                    window_width = 300
                    window_height = 220

                    screen_width = window.winfo_screenwidth()
                    screen_height = window.winfo_screenheight()
                    
                    position_top = int(screen_height / 2 - window_height / 2)
                    position_right = int(screen_width / 2 - window_width / 2)
                    
                    skill_window.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

                    skill_window.resizable(False, False)

                    skill_window.transient(window)
                    skill_window.grab_set()

                    tk.Label(skill_window, text="Выберите навык из списка:").pack(pady=10)
                    skill_category = ttk.Combobox(skill_window, values=[''] + list_skill, state="readonly")
                    skill_category.pack(pady=5)

                    tk.Label(skill_window, text="ИЛИ").pack(pady=5)
                    tk.Label(skill_window, text="Введите название навыка:").pack()
                    skill_entry = tk.Entry(skill_window, width=25)
                    skill_entry.pack(pady=5)
                    
                    def on_continue():
                        nonlocal labels_skill
                        nonlocal labels_pts
                        nonlocal buttons_del
                        nonlocal buttons_minus_skill
                        nonlocal buttons_plus_skill
                        nonlocal labels_min_skill_pts
                        nonlocal labels_max_skill_pts
                        nonlocal y_pos
                        nonlocal list_skill_for_display
                        nonlocal list_skill_pts_for_display
                        global spare_skill_pts
                        skill_name = skill_entry.get()
                        if skill_name:
                            list_skill_for_display.append(skill_name)
                            list_skill_pts_for_display.append(1)
                            label_skill = Label(new_tab, text=f'{skill_name}', justify='left')
                            label_pts = Label(new_tab, text='1', justify='left')
                            label_skill.place(x=420, y=y_pos)
                            label_pts.place(x=560, y=y_pos)

                            labels_skill.append(label_skill)
                            labels_pts.append(label_pts)

                            local_idx = len(list_skill_pts_for_display) - 1
                            btn_del = Button(new_tab, text="X", command=lambda idx=local_idx: delete_skill(idx))
                            btn_del.place(x=400, y=y_pos, width=19, height=19)
                            buttons_del.append(btn_del)

                            btn_minus_skill = Button(new_tab, text="-", command=lambda idx=local_idx: subtract_skill_pts(idx))
                            btn_plus_skill = Button(new_tab, text="+", command=lambda idx=local_idx: add_skill_pts(idx))
                            btn_minus_skill.place(x=580, y=y_pos, width=19, height=19)
                            btn_plus_skill.place(x=610, y=y_pos, width=19, height=19)
                            btn_minus_skill.config(state="disabled")
                            buttons_minus_skill.append(btn_minus_skill)
                            buttons_plus_skill.append(btn_plus_skill)

                            label_min_skill_pts = Label(new_tab, text=f'{list_min_skill[local_idx]}', justify='left')
                            label_max_skill_pts = Label(new_tab, text=f'{list_max_skill[local_idx]}', justify='left')
                            label_min_skill_pts.place(x=660, y=y_pos)
                            label_max_skill_pts.place(x=730, y=y_pos)
                            labels_min_skill_pts.append(label_min_skill_pts)
                            labels_max_skill_pts.append(label_max_skill_pts)
                            y_pos += 20
                        else:
                            skill_name = skill_category.get()
                            if skill_name:
                                list_skill_for_display.append(skill_name)
                                list_skill_pts_for_display.append(1)
                                label_skill = Label(new_tab, text=f'{skill_name}', justify='left')
                                label_pts = Label(new_tab, text='1', justify='left')
                                label_skill.place(x=420, y=y_pos)
                                label_pts.place(x=560, y=y_pos)
                                labels_skill.append(label_skill)
                                labels_pts.append(label_pts)

                                local_idx = len(list_skill_pts_for_display) - 1
                                btn_del = Button(new_tab, text="X", command=lambda idx=local_idx: delete_skill(idx))
                                btn_del.place(x=400, y=y_pos, width=19, height=19)
                                buttons_del.append(btn_del)

                                btn_minus_skill = Button(new_tab, text="-", command=lambda idx=local_idx: subtract_skill_pts(idx))
                                btn_plus_skill = Button(new_tab, text="+", command=lambda idx=local_idx: add_skill_pts(idx))
                                btn_minus_skill.place(x=580, y=y_pos, width=19, height=19)
                                btn_plus_skill.place(x=610, y=y_pos, width=19, height=19)
                                buttons_minus_skill.append(btn_minus_skill)
                                buttons_plus_skill.append(btn_plus_skill)

                                label_min_skill_pts = Label(new_tab, text=f'{list_min_skill[local_idx]}', justify='left')
                                label_max_skill_pts = Label(new_tab, text=f'{list_max_skill[local_idx]}', justify='left')
                                label_min_skill_pts.place(x=660, y=y_pos)
                                label_max_skill_pts.place(x=730, y=y_pos)
                                labels_min_skill_pts.append(label_min_skill_pts)
                                labels_max_skill_pts.append(label_max_skill_pts)
                                y_pos += 20
                        spare_skill_pts -= 1
                        update_btn_state()
                        label_spare_skill.config(text=f'Свободные очки:  {spare_skill_pts}')
                        skill_window.destroy()
                    
                    def on_cancel():
                        """Comment"""
                        skill_window.destroy()
                    
                    button_frame = tk.Frame(skill_window)
                    button_frame.pack(pady=10)
                    tk.Button(button_frame, text="Отмена", command=on_cancel).pack(side=tk.LEFT, padx=5)
                    tk.Button(button_frame, text="Продолжить", command=on_continue).pack(side=tk.LEFT, padx=5)
                    window.wait_window(skill_window)

                def update_btn_state():
                    if spare_skill_pts > 0:
                        btn_add_skill.config(state="normal")
                    else:
                        btn_add_skill.config(state="disabled")

                btn_add_skill = Button(new_tab, text="Добавить навык", command=add_skill)
                btn_add_skill.place(x=600, y=490)
                label_spare_skill.config(text=f'Свободные очки:  {spare_skill_pts}')
                update_btn_state()

            combobox.bind("<<ComboboxSelected>>", update_points)
            generate_params()
            generate_skills()

            separator = tk.Frame(new_tab, height=2, bd=1, relief=tk.SUNKEN)
            separator.place(relx=0, rely=0.15, relwidth=1)

            btn_close = Button(new_tab, text = "Закрыть вкладку", command=close_tab)
            btn_save = Button(new_tab, text = "Сохранить", command=save_file)

            label_params = Label(new_tab, text=f"Параметры")
            label_param_values = Label(new_tab, text=f"Значение")
            label_min_param = Label(new_tab, text=f"Мин. знач")
            label_max_param = Label(new_tab, text=f"Макс. знач")
            label_skills = Label(new_tab, text=f"Навыки")
            label_skill_values = Label(new_tab, text=f"Значение")
            label_min_skill = Label(new_tab, text=f"Мин. знач")
            label_max_skill = Label(new_tab, text=f"Макс. знач")

            label_param1 = Label(new_tab, text='Сила', justify='left')
            label_param2 = Label(new_tab, text='Тело', justify='left')
            label_param3 = Label(new_tab, text='Точность', justify='left')
            label_param4 = Label(new_tab, text='Подвижность', justify='left')
            label_param5 = Label(new_tab, text='Интеллект', justify='left')
            label_param6 = Label(new_tab, text='Сообразительность', justify='left')
            label_param7 = Label(new_tab, text='Восприятие', justify='left')
            label_param8 = Label(new_tab, text='Фантазия', justify='left')
            label_param9 = Label(new_tab, text='Воля', justify='left')
            label_param10 = Label(new_tab, text='Харизма', justify='left')
            label_param11 = Label(new_tab, text='Внешность', justify='left')

            btn_minus_param1 = Button(new_tab, text="-", command=lambda: subtract_pts(0))
            btn_plus_param1 = Button(new_tab, text="+", command=lambda: add_pts(0))
            btn_minus_param2 = Button(new_tab, text="-", command=lambda: subtract_pts(1))
            btn_plus_param2 = Button(new_tab, text="+", command=lambda: add_pts(1))
            btn_minus_param3 = Button(new_tab, text="-", command=lambda: subtract_pts(2))
            btn_plus_param3 = Button(new_tab, text="+", command=lambda: add_pts(2))
            btn_minus_param4 = Button(new_tab, text="-", command=lambda: subtract_pts(3))
            btn_plus_param4 = Button(new_tab, text="+", command=lambda: add_pts(3))
            btn_minus_param5 = Button(new_tab, text="-", command=lambda: subtract_pts(4))
            btn_plus_param5 = Button(new_tab, text="+", command=lambda: add_pts(4))
            btn_minus_param6 = Button(new_tab, text="-", command=lambda: subtract_pts(5))
            btn_plus_param6 = Button(new_tab, text="+", command=lambda: add_pts(5))
            btn_minus_param7 = Button(new_tab, text="-", command=lambda: subtract_pts(6))
            btn_plus_param7 = Button(new_tab, text="+", command=lambda: add_pts(6))
            btn_minus_param8 = Button(new_tab, text="-", command=lambda: subtract_pts(7))
            btn_plus_param8 = Button(new_tab, text="+", command=lambda: add_pts(7))    
            btn_minus_param9 = Button(new_tab, text="-", command=lambda: subtract_pts(8))
            btn_plus_param9 = Button(new_tab, text="+", command=lambda: add_pts(8))
            btn_minus_param10 = Button(new_tab, text="-", command=lambda: subtract_pts(9))
            btn_plus_param10 = Button(new_tab, text="+", command=lambda: add_pts(9))
            btn_minus_param11 = Button(new_tab, text="-", command=lambda: subtract_pts(10))
            btn_plus_param11 = Button(new_tab, text="+", command=lambda: add_pts(10))

            btn_minus_param1.place(x=190, y=170, width=19, height=19)
            btn_plus_param1.place(x=220, y=170, width=19, height=19)
            btn_minus_param2.place(x=190, y=190, width=19, height=19)
            btn_plus_param2.place(x=220, y=190, width=19, height=19)
            btn_minus_param3.place(x=190, y=210, width=19, height=19)
            btn_plus_param3.place(x=220, y=210, width=19, height=19)
            btn_minus_param4.place(x=190, y=230, width=19, height=19)
            btn_plus_param4.place(x=220, y=230, width=19, height=19)
            btn_minus_param5.place(x=190, y=250, width=19, height=19)
            btn_plus_param5.place(x=220, y=250, width=19, height=19)
            btn_minus_param6.place(x=190, y=270, width=19, height=19)
            btn_plus_param6.place(x=220, y=270, width=19, height=19)
            btn_minus_param7.place(x=190, y=290, width=19, height=19)
            btn_plus_param7.place(x=220, y=290, width=19, height=19)
            btn_minus_param8.place(x=190, y=310, width=19, height=19)
            btn_plus_param8.place(x=220, y=310, width=19, height=19)
            btn_minus_param9.place(x=190, y=330, width=19, height=19)
            btn_plus_param9.place(x=220, y=330, width=19, height=19)
            btn_minus_param10.place(x=190, y=350, width=19, height=19)
            btn_plus_param10.place(x=220, y=350, width=19, height=19)
            btn_minus_param11.place(x=190, y=370, width=19, height=19)
            btn_plus_param11.place(x=220, y=370, width=19, height=19)

            label_params.place(x=40, y=140)
            label_param_values.place(x=145, y=140)
            label_min_param.place(x=245, y=140)
            label_max_param.place(x=315, y=140)
            label_skills.place(x=420, y=140)
            label_skill_values.place(x=535, y=140)
            label_min_skill.place(x=635, y=140)
            label_max_skill.place(x=705, y=140)

            label_param1.place(x=40, y=170)
            label_param2.place(x=40, y=190)
            label_param3.place(x=40, y=210)
            label_param4.place(x=40, y=230)
            label_param5.place(x=40, y=250)
            label_param6.place(x=40, y=270)
            label_param7.place(x=40, y=290)
            label_param8.place(x=40, y=310)
            label_param9.place(x=40, y=330)
            label_param10.place(x=40, y=350)
            label_param11.place(x=40, y=370)

            label_param_pt1.place(x=170, y=170)
            label_param_pt2.place(x=170, y=190)
            label_param_pt3.place(x=170, y=210)
            label_param_pt4.place(x=170, y=230)
            label_param_pt5.place(x=170, y=250)
            label_param_pt6.place(x=170, y=270)
            label_param_pt7.place(x=170, y=290)
            label_param_pt8.place(x=170, y=310)
            label_param_pt9.place(x=170, y=330)
            label_param_pt10.place(x=170, y=350)
            label_param_pt11.place(x=170, y=370)

            label_min_param1.place(x=270, y=170)
            label_min_param2.place(x=270, y=190)
            label_min_param3.place(x=270, y=210)
            label_min_param4.place(x=270, y=230)
            label_min_param5.place(x=270, y=250)
            label_min_param6.place(x=270, y=270)
            label_min_param7.place(x=270, y=290)
            label_min_param8.place(x=270, y=310)
            label_min_param9.place(x=270, y=330)
            label_min_param10.place(x=270, y=350)
            label_min_param11.place(x=270, y=370)

            label_max_param1.place(x=340, y=170)
            label_max_param2.place(x=340, y=190)
            label_max_param3.place(x=340, y=210)
            label_max_param4.place(x=340, y=230)
            label_max_param5.place(x=340, y=250)
            label_max_param6.place(x=340, y=270)
            label_max_param7.place(x=340, y=290)
            label_max_param8.place(x=340, y=310)
            label_max_param9.place(x=340, y=330)
            label_max_param10.place(x=340, y=350)
            label_max_param11.place(x=340, y=370)

            def gen_random():
                global spare_param_pts
                global spare_skill_pts
                nonlocal list_min_param
                nonlocal list_max_param
                global race
                nonlocal list_param_pts
                nonlocal list_skill_pts
                nonlocal last_selected
                global spare_skill_pts
                nonlocal label_spare_skill
                nonlocal labels_skill
                nonlocal labels_pts
                nonlocal buttons_del
                nonlocal buttons_minus_skill
                nonlocal buttons_plus_skill
                nonlocal labels_min_skill_pts
                nonlocal labels_max_skill_pts
                nonlocal list_skill_for_display
                nonlocal list_skill_pts_for_display
                race = list_races[random.randint(1, 5) - 1]
                if race == 'Люди':
                    spare_param_pts = 12
                    spare_skill_pts = 15
                    list_min_param = [1] * 11
                    list_max_param = [5, 5, 5, 5, 5, 6, 5, 6, 5, 5, 5]
                    for index in range(11):
                        list_param_pts[index] = list_min_param[index]
                    for index in range(26):
                        list_skill_pts[index] = 0
                    entry_name.delete(0, END)
                    entry_name.insert(0, list_human_names[random.randint(1, 5) - 1])
                    combobox.set('Люди')
                    last_selected = 'Люди'
                if race == 'Эльфы':
                    spare_param_pts = 9
                    spare_skill_pts = 12
                    list_min_param = [1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 2]
                    list_max_param = [5, 5, 5, 5, 5, 5, 6, 5, 5, 5, 7]
                    for index in range(11):
                        list_param_pts[index] = list_min_param[index]
                    for index in range(26):
                        list_skill_pts[index] = 0
                    entry_name.delete(0, END)
                    entry_name.insert(0, list_elf_names[random.randint(1, 5) - 1])
                    combobox.set('Эльфы')
                    last_selected = 'Эльфы'
                if race == 'Гномы':
                    spare_param_pts = 10
                    spare_skill_pts = 12
                    list_min_param = [2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1]
                    list_max_param = [6, 6, 5, 5, 6, 5, 5, 5, 5, 5, 5]
                    for index in range(11):
                        list_param_pts[index] = list_min_param[index]
                    for index in range(26):
                        list_skill_pts[index] = 0
                    entry_name.delete(0, END)
                    entry_name.insert(0, list_dwarf_names[random.randint(1, 5) - 1])
                    combobox.set('Гномы')
                    last_selected = 'Гномы'
                if race == 'Орки':
                    spare_param_pts = 10
                    spare_skill_pts = 12
                    list_min_param = [2, 2, 1, 1, 1, 1, 1, 1, 2, 1, 1]
                    list_max_param = [7, 6, 5, 5, 5, 5, 5, 4, 7, 4, 4]
                    for index in range(11):
                        list_param_pts[index] = list_min_param[index]
                    for index in range(26):
                        list_skill_pts[index] = 0
                    entry_name.delete(0, END)
                    entry_name.insert(0, list_orc_names[random.randint(1, 5) - 1])
                    combobox.set('Орки')
                    last_selected = 'Орки'
                if race == 'Нежить':
                    spare_param_pts = 9
                    spare_skill_pts = 6
                    list_min_param = [1] * 11
                    list_max_param = [9, 3, 4, 4, 5, 4, 7, 3, 4, 3, 3]
                    for index in range(11):
                        list_param_pts[index] = list_min_param[index]
                    for index in range(26):
                        list_skill_pts[index] = 0
                    entry_name.delete(0, END)
                    entry_name.insert(0, list_undead_names[random.randint(1, 5) - 1])
                    combobox.set('Нежить')
                    last_selected = 'Нежить'
                generate_params()
                while labels_skill: 
                    labels_skill[0].destroy()
                    labels_pts[0].destroy()
                    buttons_del[0].destroy()
                    buttons_minus_skill[0].destroy()
                    buttons_plus_skill[0].destroy()
                    labels_min_skill_pts[0].destroy()
                    labels_max_skill_pts[0].destroy()

                    label_spare_skill.config(text=f'Свободные очки:  {spare_skill_pts}')
                    del labels_skill[0]
                    del labels_pts[0]
                    del buttons_del[0]
                    del buttons_minus_skill[0]
                    del buttons_plus_skill[0]
                    del labels_min_skill_pts[0]
                    del labels_max_skill_pts[0]
                list_skill_for_display = []
                list_skill_pts_for_display = []
                generate_skills()
                label_spare_skill.config(text=f'Свободные очки:  {spare_skill_pts}')
                

            btn_random = Button(new_tab, text="Случайный персонаж", command=gen_random)
            btn_random.place(x=325, y=20)
            label_random = Label(new_tab, text="При нажатии на кнопку \"Случайный персонаж\" раса, имя, характеристики и навыки генерируются автоматически")
            label_random.place(x=75, y=55)

            btn_save.place(x=540, y=530)
            btn_close.place(x=160, y=530)
            gen_tab_counter += 1
            update_tabmenu()

            tab_control.select(new_tab)

        read_file()

            
def close_tab():
    """Коммнетарий"""
    current_tab_index = tab_control.index("current")
    tab_control.forget(current_tab_index)
    update_tabmenu()

def update_tabmenu():
    """Коммнетарий"""
    tabmenu.delete(0, tk.END)
    tab_list = [tab_control.tab(i, option='text') for i in range(tab_control.index('end'))]
    for tab in tab_list:
        tabmenu.add_command(label=tab, command=lambda t=tab: go_to_tab(t))

def go_to_tab(tab):
    """Коммнетарий"""
    index = [tab_control.tab(i, option='text') for i in range(tab_control.index('end'))].index(tab)
    tab_control.select(index)

def update_log():
    """Коммнетарий"""
    global list_file_paths
    global list_files
    global list_file_paths_from_log
    list_files = []
    list_file_paths = []
    list_file_paths_from_log = []
    file_path = selected_file_path
    with open('log.txt', 'a+') as file:
        file.write(f'{file_path}\n')
    with open('log.txt', 'r') as file:
        lines = file.readlines()
        for line in reversed(lines):
            list_file_paths_from_log.append(line.strip())
    seen_files = set()
    for item in list_file_paths_from_log:
        file_name = os.path.basename(item)
        if file_name not in seen_files:
            list_files.append(file_name)
            list_file_paths.append(item)
            seen_files.add(file_name)
    for item in tree.get_children():
        tree.delete(item)
    for item in list_files:
        tree.insert("", tk.END, text=item)

def on_double_click(event):
    global selected_file_path
    global list_file_paths
    global tree
    global is_opened_from_list
    item_index = tree.index(tree.selection()[0])
    selected_file_path = list_file_paths[item_index]
    if not os.path.exists(selected_file_path):
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Ошибка", "Файл не найден")
        root.destroy()
    else:
        is_opened_from_list = True
        open_tab(0)
        is_opened_from_list = False

def show_help():
    help_window = tk.Toplevel()
    help_window.title("Помощь")
    help_window.geometry("400x200")
    help_text = "Для создания случайного персонажа нажмите кнопку \"Создать персонажа\".\n" \
        "Для загрузки персонажа из файла нажмите кнопку \"Загрузить\" и выберите нужный файл.\n" \
        "Для открытия изменённого ранее файла дважды нажмите на его имя в списке.\n" \
        "Во вкладке генерации персонажа Вы можете изменять его расу, имя и характеристики."
    tk.Label(help_window, text=help_text, wraplength=380, justify='left').pack(padx=10, pady=10)
    tk.Button(help_window, text="Закрыть", command=help_window.destroy).pack(pady=15)

def show_info():
    info_window = tk.Toplevel()
    info_window.title("Информация о приложении")
    info_window.geometry("400x200")
    help_text = "Десктопное приложение для генерации персонажей настольной игры DnD.\nВерсия 0.1."
    tk.Label(info_window, text=help_text, wraplength=380, justify='left').pack(padx=10, pady=10)
    tk.Button(info_window, text="Закрыть", command=info_window.destroy).pack(pady=25)


window = Tk()
window.title("Генератор")
window.geometry("800x600")
window.resizable(False, False)
mainmenu = Menu(window)
window.config(menu = mainmenu)

gen_tab_counter = 1

tab_control = ttk.Notebook(window)
tab_control.pack(expand=1, fill='both')

start_tab = ttk.Frame(tab_control)
tab_control.add(start_tab, text='Главная страница')

filemenu = Menu(mainmenu, tearoff=0)
filemenu.add_command(label="Открыть файл", command=lambda: open_tab(0))
filemenu.add_command(label="Выход", command=exit)

tabmenu = Menu(mainmenu, tearoff=0)

helpmenu = Menu(mainmenu, tearoff=0)
helpmenu.add_command(label="Помощь", command=show_help)
helpmenu.add_command(label="О приложении", command=show_info)

mainmenu.add_cascade(label="Файл", menu=filemenu)
mainmenu.add_cascade(label="Вкладка", menu=tabmenu)
mainmenu.add_cascade(label="Справка", menu=helpmenu)

btn_random = Button(start_tab, text="Создать песонажа", command=lambda: open_tab(1))
btn_random.place(x=140, y=200)

btn_manual = Button(start_tab, text="Загрузить", command=lambda: open_tab(0))
btn_manual.place(x=160, y=280)

label0 = Label(start_tab, text="Недавно созданные/открытые файлы")
label0.place(x=310, y=25)

tree = ttk.Treeview(start_tab, show='tree')
tree.place(x=310, y=50, width=350, height=450)
tree.bind("<Double-1>", on_double_click)
labe_tree = Label(start_tab, text="Кликните дважды по названию файла, чтобы открыть его")
labe_tree.place(x=310, y=500)

with open('log.txt', 'r') as file:
    lines = file.readlines()
    for line in reversed(lines):
        list_file_paths_from_log.append(line.strip())

seen_files = set()
for item in list_file_paths_from_log:
    file_name = os.path.basename(item)
    if file_name not in seen_files:
        list_files.append(file_name)
        list_file_paths.append(item)
        seen_files.add(file_name)

for item in list_files:
    tree.insert("", tk.END, text=item)


window.mainloop()