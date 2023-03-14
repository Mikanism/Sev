import streamlit as st
import datetime as dt
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import docx
import io
import json
import pandas as pd
import sqlite3

material_data = ['Панель', 'Кирпич', 'Монолит']
perecrit_data = ['ж/б', 'Деревянные', 'Смешанные']
klass_data = ['Эконом', 'Бизнес', 'Элит']
vid_data = ['На улицу', 'Во двор', 'Во двор и на улицу', 'Панорамный']
status_data = ['Без отделки / требуется капитальный ремонт', 'Под чистовую отделку', 'Среднее жилое состояние / требуется косметический ремонт', 'Хорошее состояние', 'Отличное (евроремонт)', 'Ремонт премиум класса']
obrem_data = ['Без обременений', 'Ипотека в силу закона', 'Залог в силу закона', 'Рента в силу закона', 'Не зарегистрировано']

st.title('Отчет')
add_selectbox = st.sidebar.selectbox(
    "Выбор",
    ("ОПЕКА", "Загрузить")
)
if add_selectbox == "ОПЕКА":
    with st.form("my_form", clear_on_submit = False):
        tab1, tab2, tab3, tab4 = st.tabs(["Задание на оценку", "Объект оценки", "Здание и подъезд", 'Фото'])
        with tab1:
            # Задание на оценку
            number = st.text_input('Номер отчета', '')
            c1, c2, c3 = st.columns(3)
            with c1:
                date_otchet = st.date_input("Дата составления отчета", dt.date.today())
            with c2:
                date_ocenki = st.date_input("Дата Оценки", dt.date.today())
            with c3:
                date_osmotr = st.date_input("Дата Осмотра", dt.date.today())

            # Заказчик
            ww1,ww2 = st.columns(2)
            with ww1:
                zakazchik = st.text_input('Заказчик ФИО', '')
                zakazchik_pasport_number = st.text_input('Паспорт Cерия/Номер', '')
                zakazchik_pasport_date = st.text_input('Дата выдачи', '')
                zakazchik_pasport_kem = st.text_input('Кем выдан', '')
            with ww2:
                prava = st.selectbox('Оцениваемые права', ['Право собственности', 'Право требования'])
                obrem = st.selectbox('Обременения', obrem_data)
                docs = st.text_input('Правоустанавливающие документы','Выписка из ЕГРН, Свидетельство о государственной регистрации права')
                docs2 = st.text_input('Документы технической инвентаризации', 'Кадастровый паспорт помещения, Технический паспорт жилого помещения (квартиры), Поэтажный план, Экспликация')


        with tab2:
            # Объект оценки
            prav_oblad = st.text_input('Наименование правообладателей')
            adres = st.text_input('Адрес')
            area = st.text_input('Площадь квартиры')
            kadast_number = st.text_input('Кадастровый номер')

            pr1, pr2 = st.columns(2)
            with pr1:
                price = st.number_input('Рыночная стоимость', format = '%d', value = 1)
            with pr2:
                pr_propis = st.text_input('Цена прописью')

            metro = st.text_input('Ближайшее метро/Ж.Д. Станция/Остановка общественного транспорта')

            tm1, tm2 = st.columns(2)
            with tm1:
                time_metro = st.number_input('Растояние до Метро', format = '%d', value = 1)
            with tm2:
                way = st.radio(
                "Способ",
                ('пешком', 'на общественном транспорте'))
            # Инфа о Квартире
            flor = st.text_input('Этаж расположения квартиры')#Этаж
            rooms = st.text_input('Количество комнат в квартире')#Комнаты
            status = st.selectbox('Общее состояние помещений', status_data) #Общее состояние помещений
            kv1, kv2, kv3 = st.columns(3)
            with kv1:
                area_balcon = st.text_input('Площадь квартиры с учетом площади лоджии/балкона')# Площадь_Балкон
                potolki = st.text_input('Высота потолков')# Потолки
            with kv2:
                Liv_area = st.text_input('Жилая площадь')# Жилая 
                kitchen = st.text_input('Площадь кухни')# Кухня
                vid = st.selectbox('Вид из окна', vid_data)# Вид
            with kv3:
                balkon = st.radio(
                "Наличие балкона, лоджии",
                ('Да', 'Нет'), horizontal = True) # Балкон
                steklo = st.radio(
                "Наличие остекления",
                ('Да', 'Нет'), horizontal = True) # Стекло
                Wcc = st.radio(
                "Тип санузла",
                ('Совмещенный', 'Раздельный', 'два санузла')) # WC 
            st.subheader('Инженерные коммуникации')
            #tech = st.text_input('Техническое обеспечение квартиры','холодное и горячее водоснабжение от городской сети, электроснабжение, центральное отопление, канализация.')
            tech = st.multiselect('Техническое обеспечение квартиры',['холодное и горячее водоснабжение от городской сети','холодное водоснабжение', 'электроснабжение', 'центральное отопление', 'газ', 'газовая колонка', 'канализация'])
            obor = st.radio(
                "Необходимое оборудование для инженерных коммуникаций",
                ('Имеются', 'Не имеются'), horizontal = True)
            
        with tab3:
            # Здание и подъезд
            yyear = st.text_input('Год постройки дома')
            iznos = st.text_input('Физический износ жилого дома')
            facade = st.text_input('Внешний вид фасада дома')
            flors = st.text_input('Этажность жилого дома')
            materual = st.selectbox('Материал стен дома', material_data)
            perecrit = st.selectbox('Характеристика (вид) перекрытий', perecrit_data)   
            klass = st.selectbox('Класс жилого дома', klass_data)
            
            podz1, podz2, podz3 = st.columns(3)
            with podz1:
                q1 = st.radio(
                "Наличие подземных этажей",
                ('Да', 'Нет'), horizontal = True)

                q2 = st.radio(
                "Наличие подземной парковки",
                ('Да', 'Нет'), horizontal = True)
            with podz2:
                q3 = st.radio(
                "Ограждение по периметру",
                ('Да', 'Нет'), horizontal = True)

                q4 = st.radio(
                "Физическая охрана",
                ('Да', 'Нет'), horizontal = True)
            with podz3:
                park = st.text_input('Наличие парковки(охраняемой/стихийной)')
            q5 = st.radio(
            "?",
            ('Домофон', 'Кодовый замок', 'Консьерж в подъезде'), horizontal = True)
        with tab4:
            pic = st.file_uploader('Карта Области')
            pic1 = st.file_uploader('Карта Метро')
        upload = st.form_submit_button("Сохранить")

        
    if upload:
        with st.spinner('Wait for it...'):
            # Правки
            price = f'{int(price):,}'.replace(',', ' ')
            tech = str(tech).replace("'", '')[1:-1]

            doc = DocxTemplate('Макет ОПЕКА.docx')
            map_obl = ''
            map_metro = ''
            if pic != None:
                map_obl = InlineImage(doc, pic, width=Mm(180), height=Mm(73))
            if pic1 != None:
                map_metro = InlineImage(doc, pic1, width=Mm(180), height=Mm(73))

            context = { 'Номер': number,    # Задание на оценку
                        'Дата_Оценки': date_ocenki.strftime(format = "%d.%m.%Y"), 
                        'Дата': date_otchet.strftime(format = "%d.%m.%Y"),
                        'Дата_Осмотр': date_osmotr.strftime(format = "%d.%m.%Y"),
                        'Заказчик': zakazchik,  # Заказчик
                        'Серия_номер':zakazchik_pasport_number, 
                        'Дата_выдачи':zakazchik_pasport_date,   
                        'Кем': zakazchik_pasport_kem,
                        'Право': prava,
                        'Обрем': obrem,
                        'Док1': docs,
                        'Док2': docs2,
                        'Доки': docs + '\n' + docs2,
                        'Адрес' : adres,    # Объект оценки
                        'Обладать' : prav_oblad,
                        'Кадастр': kadast_number,  
                        'Площадь': area,    
                        'Стоимость': price,
                        'Стоим_Пропись': pr_propis, 
                        'Метро': metro,
                        'Растояние_Метро': time_metro,
                        'как' : way,
                        'Этаж': flor,   # Инфа о Квартире
                        'Комнаты': rooms,
                        'Площадь_Летних': area_balcon,
                        'Кухня': kitchen,
                        'Жилая': Liv_area,
                        'Потолки': potolki,
                        'Балкон': balkon,
                        'Стекло': steklo,
                        'WC': Wcc,
                        'Вид': vid,
                        'Состояние': status,
                        'Тех': tech,
                        'Обор': obor,
                        'Год_дома': yyear,  # Здание и подъезд
                        'Износ': iznos,
                        'Фасад': facade,    
                        'Этажность': flors, 
                        'Материал': materual, 
                        'Перек': perecrit, 
                        'Класс': klass,
                        'Под_эт': q1, 
                        'Парк': q2,
                        'Паркк': park,
                        'Перим': q3, 
                        'Охрана': q4, 
                        'Домофон': q5,
                        'map_obl': map_obl,
                        'map_metro': map_metro
                        }
            
            
            con = sqlite3.connect('Testdb.sqlite')
            cur = con.cursor()
            pd.DataFrame([dict(list(context.items())[:-2])]).to_sql('otchet', con, if_exists = 'append', index = False)
            con.close()

            doc.render(context)
            #doc.render(context_pic)
            doc.save('Макет_ОПЕКА1.docx')
            st.success('Done!')


        doc_download = docx.Document('Макет_ОПЕКА1.docx')
        bio = io.BytesIO()
        doc_download.save(bio)
        
        if doc_download:
            st.download_button(
                label="Скачать",
                data=bio.getvalue(),
                file_name="Макет_ОПЕКА1.docx",
                mime="docx"
            )


if add_selectbox == 'Загрузить':
    slct1, slct2 = st.columns([1, 1])
    selectform = slct1.selectbox(
        "Выберите Макет",
        ['ОПЕКА', 'База'], 
    )

    if selectform == 'База':
        con = sqlite3.connect('Testdb.sqlite')
        cur = con.cursor()
        sql = '''Select * from otchet'''
        st.dataframe(pd.read_sql(sql, con))
        con.close()
    
    if selectform == 'ОПЕКА':
        shablon = slct2.text_input('Ввдедите номер отчета')
        con = sqlite3.connect('Testdb.sqlite')
        cur = con.cursor()
        sql = 'Select * from otchet Where Номер = ?'
        df = pd.read_sql(sql, con, params = (shablon,))
        con.close()

        if not df.empty:
            st.dataframe(df)
            with st.form("my_form", clear_on_submit = False):
                tab1, tab2, tab3, tab4 = st.tabs(["Задание на оценку", "Объект оценки", "Здание и подъезд", 'Фото'])
                with tab1:
                    # Задание на оценку
                    number = st.text_input('Номер отчета', df['Номер'][0])
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        date_otchet = st.date_input("Дата составления отчета", dt.date.today())
                    with c2:
                        date_ocenki = st.date_input("Дата Оценки", dt.date.today())
                    with c3:
                        date_osmotr = st.date_input("Дата Осмотра", dt.date.today())

                    # Заказчик
                    ww1, ww2 = st.columns(2)
                    with ww1:
                        zakazchik = st.text_input('Заказчик ФИО', df['Заказчик'][0])
                        zakazchik_pasport_number = st.text_input('Паспорт Cерия/Номер', df['Серия_номер'][0])
                        zakazchik_pasport_date = st.text_input('Дата выдачи', df['Дата_выдачи'][0])
                        zakazchik_pasport_kem = st.text_input('Кем выдан', df['Кем'][0])
                    with ww2:
                        prava = st.selectbox('Оцениваемые права', ['Право собственности', 'Право требования'], ['Право собственности', 'Право требования'].index(df['Право'][0]))
                        obrem = st.selectbox('Обременения', obrem_data, obrem_data.index(df['Обрем'][0]))
                        docs = st.text_input('Правоустанавливающие документы', df['Док1'][0])
                        docs2 = st.text_input('Документы технической инвентаризации', df['Док2'][0])


                with tab2:
                    # Объект оценки
                    prav_oblad = st.text_input('Наименование правообладателей', df['Обладать'][0])
                    adres = st.text_input('Адрес', df['Адрес'][0])
                    area = st.text_input('Площадь квартиры', df['Площадь'][0])
                    kadast_number = st.text_input('Кадастровый номер', df['Кадастр'][0])

                    pr1, pr2 = st.columns(2)
                    with pr1:
                        price = st.number_input('Рыночная стоимость', format = '%d', value = int(str(df['Стоимость'][0]).replace(' ', '')))
                    with pr2:
                        pr_propis = st.text_input('Цена прописью', df['Стоим_Пропись'][0])

                    metro = st.text_input('Ближайшее метро/Ж.Д. Станция/Остановка общественного транспорта', df['Метро'][0])

                    tm1, tm2 = st.columns(2)
                    with tm1:
                        time_metro = st.number_input('Растояние до Метро', format = '%d', value = df['Растояние_Метро'][0])
                    with tm2:
                        way = st.radio(
                        "Способ",
                        ('пешком', 'на общественном транспорте'), ('пешком', 'на общественном транспорте').index(df['как'][0]))
                    # Инфа о Квартире
                    flor = st.text_input('Этаж расположения квартиры', df['Этаж'][0])#Этаж
                    rooms = st.text_input('Количество комнат в квартире', df['Комнаты'][0])#Комнаты
                    status = st.selectbox('Общее состояние помещений', status_data, status_data.index(df['Состояние'][0])) #Общее состояние помещений
                    kv1, kv2, kv3 = st.columns(3)
                    with kv1:
                        area_balcon = st.text_input('Площадь квартиры с учетом площади лоджии/балкона', df['Площадь_Летних'][0])# Площадь_Балкон
                        potolki = st.text_input('Высота потолков', df['Потолки'][0])# Потолки
                    with kv2:
                        Liv_area = st.text_input('Жилая площадь', df['Жилая'][0])# Жилая 
                        kitchen = st.text_input('Площадь кухни', df['Кухня'][0])# Кухня
                        vid = st.selectbox('Вид из окна', vid_data, vid_data.index(df['Вид'][0]))# Вид
                    with kv3:
                        balkon = st.radio(
                        "Наличие балкона, лоджии",
                        ('Да', 'Нет'), horizontal = True, index = ('Да', 'Нет').index(df['Балкон'][0])) # Балкон
                        steklo = st.radio(
                        "Наличие остекления",
                        ('Да', 'Нет'), horizontal = True, index = ('Да', 'Нет').index(df['Стекло'][0])) # Стекло
                        Wcc = st.radio(
                        "Тип санузла",
                        ('Совмещенный', 'Раздельный', 'два санузла'), index = ('Совмещенный', 'Раздельный', 'два санузла').index(df['WC'][0])) # WC 
                    st.subheader('Инженерные коммуникации')
                    #tech = st.text_input('Техническое обеспечение квартиры','холодное и горячее водоснабжение от городской сети, электроснабжение, центральное отопление, канализация.')
                    tech_d = df['Тех'][0].replace(', ', ',').split(',')
                    if tech_d == ['']:
                        tech_d = None
                    
                    tech = st.multiselect('Техническое обеспечение квартиры',
                    ['холодное и горячее водоснабжение от городской сети','холодное водоснабжение', 'электроснабжение', 'центральное отопление', 'газ', 'газовая колонка', 'канализация'],
                    tech_d)
                    obor = st.radio(
                        "Необходимое оборудование для инженерных коммуникаций",
                        ('Имеются', 'Не имеются'), horizontal = True, index = ('Имеются', 'Не имеются').index(df['Обор'][0]))
                with tab3:
                    # Здание и подъезд
                    yyear = st.text_input('Год постройки дома', df['Год_дома'][0])
                    iznos = st.text_input('Физический износ жилого дома', df['Износ'][0])
                    facade = st.text_input('Внешний вид фасада дома', df['Фасад'][0])
                    flors = st.text_input('Этажность жилого дома', df['Этажность'][0])
                    materual = st.selectbox('Материал стен дома', material_data, index = material_data.index(df['Материал'][0]))
                    perecrit = st.selectbox('Характеристика (вид) перекрытий', perecrit_data, index = perecrit_data.index(df['Перек'][0]))  
                    klass = st.selectbox('Класс жилого дома', klass_data, index = klass_data.index(df['Класс'][0]))
                    
                    podz1, podz2, podz3 = st.columns(3)
                    with podz1:
                        q1 = st.radio(
                        "Наличие подземных этажей",
                        ('Да', 'Нет'), horizontal = True, index = ('Да', 'Нет').index(df['Под_эт'][0]))

                        q2 = st.radio(
                        "Наличие подземной парковки",
                        ('Да', 'Нет'), horizontal = True, index = ('Да', 'Нет').index(df['Парк'][0]))
                    with podz2:
                        q3 = st.radio(
                        "Ограждение по периметру",
                        ('Да', 'Нет'), horizontal = True, index = ('Да', 'Нет').index(df['Перим'][0]))

                        q4 = st.radio(
                        "Физическая охрана",
                        ('Да', 'Нет'), horizontal = True, index = ('Да', 'Нет').index(df['Охрана'][0]))
                    with podz3:
                        park = st.text_input('Наличие парковки(охраняемой/стихийной)', df['Паркк'][0])
                    q5 = st.radio(
                    "?",
                    ('Домофон', 'Кодовый замок', 'Консьерж в подъезде'), horizontal = True, index = ('Домофон', 'Кодовый замок', 'Консьерж в подъезде').index(df['Домофон'][0]))
                with tab4:
                    pic = st.file_uploader('Карта Области')
                    pic1 = st.file_uploader('Карта Метро')
                upload = st.form_submit_button("Сохранить")
   
            if upload:
                with st.spinner('Wait for it...'):
                    # Правки
                    price = f'{int(price):,}'.replace(',', ' ')
                    tech = str(tech).replace("'", '')[1:-1]

                    doc = DocxTemplate('Макет ОПЕКА.docx')
                    map_obl = ''
                    map_metro = ''
                    if pic != None:
                        map_obl = InlineImage(doc, pic, width=Mm(180), height=Mm(73))
                    if pic1 != None:
                        map_metro = InlineImage(doc, pic1, width=Mm(180), height=Mm(73))

                    context = { 'Номер': number,    # Задание на оценку
                                'Дата_Оценки': date_ocenki.strftime(format = "%d.%m.%Y"), 
                                'Дата': date_otchet.strftime(format = "%d.%m.%Y"),
                                'Дата_Осмотр': date_osmotr.strftime(format = "%d.%m.%Y"),
                                'Заказчик': zakazchik,  # Заказчик
                                'Серия_номер':zakazchik_pasport_number, 
                                'Дата_выдачи':zakazchik_pasport_date,   
                                'Кем': zakazchik_pasport_kem,
                                'Право': prava,
                                'Обрем': obrem,
                                'Док1': docs,
                                'Док2': docs2,
                                'Доки': docs + '\n' + docs2,
                                'Адрес' : adres,    # Объект оценки
                                'Обладать' : prav_oblad,
                                'Кадастр': kadast_number,  
                                'Площадь': area,    
                                'Стоимость': price,
                                'Стоим_Пропись': pr_propis, 
                                'Метро': metro,
                                'Растояние_Метро': time_metro,
                                'как' : way,
                                'Этаж': flor,   # Инфа о Квартире
                                'Комнаты': rooms,
                                'Площадь_Летних': area_balcon,
                                'Кухня': kitchen,
                                'Жилая': Liv_area,
                                'Потолки': potolki,
                                'Балкон': balkon,
                                'Стекло': steklo,
                                'WC': Wcc,
                                'Вид': vid,
                                'Состояние': status,
                                'Тех': tech,
                                'Обор': obor,
                                'Год_дома': yyear,  # Здание и подъезд
                                'Износ': iznos,
                                'Фасад': facade,    
                                'Этажность': flors, 
                                'Материал': materual, 
                                'Перек': perecrit, 
                                'Класс': klass,
                                'Под_эт': q1, 
                                'Парк': q2,
                                'Паркк': park,
                                'Перим': q3, 
                                'Охрана': q4, 
                                'Домофон': q5,
                                'map_obl': map_obl,
                                'map_metro': map_metro
                                }
                    
                    con = sqlite3.connect('Testdb.sqlite')
                    cur = con.cursor()
                    cur.execute("DELETE FROM otchet WHERE Номер = ?", (context['Номер'],))
                    pd.DataFrame([dict(list(context.items())[:-2])]).to_sql('otchet', con, if_exists = 'append', index = False)
                    con.close()

                    doc.render(context)
                    #doc.render(context_pic)
                    doc.save('Макет_ОПЕКА1.docx')
                    st.success('Done!')


                doc_download = docx.Document('Макет_ОПЕКА1.docx')
                bio = io.BytesIO()
                doc_download.save(bio)
                
                if doc_download:
                    st.download_button(
                        label="Скачать",
                        data=bio.getvalue(),
                        file_name="Макет_ОПЕКА1.docx",
                        mime="docx"
                    ) 


