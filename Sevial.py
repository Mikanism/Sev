import streamlit as st
import datetime as dt
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import docx
import io

material_data = ['Панель', 'Кирпич', 'Монолит']
perecrit_data = ['ж/б', 'Деревянные', 'Смешанные']
klass_data = ['Эконом', 'Бизнес', 'Элит']
vid_data = ['На улицу', 'Во двор', 'Во двор и на улицу', 'Панорамный']
status_data = ['Без отделки / требуется капитальный ремонт', 'Под чистовую отделку', 'Среднее жилое состояние / требуется косметический ремонт', 'Хорошее состояние', 'Отличное (евроремонт)', 'Ремонт премиум класса']
obrem_data = ['Без обременений', 'Ипотека в силу закона', 'Залог в силу закона', 'Рента в силу закона', 'Не зарегистрировано']
metro_data = ['Авиамоторная', 'Автозаводская', 'Академическая', 'Александровский сад', 'Алексеевская', 'Алма-Атинская', 'Алтуфьево', 'Аминьевская', 'Андроновка', 'Аникеевка', 
'Аннино', 'Арбатская', 'Аэропорт', 'Бабушкинская', 'Багратионовская', 'Баковка', 'Балтийская', 'Баррикадная', 'Бауманская', 'Беговая', 'Белокаменная', 'Беломорская', 'Белорусская', 
'Беляево', 'Бескудниково', 'Бибирево', 'Библиотека им. Ленина', 'Битца', 'Битцевский парк', 'Борисово', 'Боровицкая', 'Боровское шоссе', 'Ботанический сад', 'Братиславская', 
'Бульвар Адмирала Ушакова', 'Бульвар Дмитрия Донского', 'Бульвар Рокоссовского', 'Бунинская аллея', 'Бутово', 'Бутырская', 'Варшавская', 'ВДНХ', 'Верхние котлы', 'Верхние Лихоборы', 
'Владыкино', 'Водники', 'Водный стадион', 'Войковская', 'Волгоградский проспект', 'Волжская', 'Волоколамская', 'Воробьевы горы', 'Воронцовская', 'Выставочная', 'Выставочный центр', 
'Выхино', 'Говорово', 'Гражданская', 'Давыдково', 'Дегунино', 'Деловой центр', 'Депо', 'Динамо', 'Дмитровская', 'Добрынинская', 'Долгопрудная', 'Домодедовская', 'Достоевская', 'Дубровка', 
'Жулебино', 'ЗИЛ', 'Зорге', 'Зюзино', 'Зябликово', 'Измайлово', 'Измайловская', 'Калитники', 'Калужская', 'Кантемировская', 'Карамышевская', 'Каховская', 'Каширская', 'Киевская', 
'Китай-город', 'Кленовый бульвар', 'Кожуховская', 'Коломенская', 'Коммунарка', 'Комсомольская', 'Коньково', 'Коптево', 'Косино', 'Котельники', 'Красногвардейская', 'Красногорская', 
'Краснопресненская', 'Красносельская', 'Красные ворота', 'Красный Балтиец', 'Красный Строитель', 'Крестьянская застава', 'Кропоткинская', 'Крылатское', 'Крымская', 'Кубанская', 
'Кузнецкий мост', 'Кузьминки', 'Кунцевская', 'Курская', 'Курьяново', 'Кутузовская', 'Ленинский проспект', 'Лермонтовский проспект', 'Лесопарковая', 'Лефортово', 'Лианозово', 
'Лихоборы', 'Лобня', 'Локомотив', 'Ломоносовский проспект', 'Лубянка', 'Лужники', 'Лухмановская', 'Люблино', 'Марк', 'Марксистская', 'Марьина роща', 'Марьино', 'Маяковская', 
'Медведково', 'Международная', 'Менделеевская', 'Минская', 'Митино', 'Мичуринский проспект', 'Мневники', 'Можайская', 'Молодежная', 'Москва-Товарная', 'Москворечье', 'Мякинино', 
'Нагатинская', 'Нагатинский Затон', 'Нагорная', 'Народное ополчение', 'Нахабино', 'Нахимовский проспект', 'Некрасовка', 'Немчиновка', 'Нижегородская', 'Новаторская', 'Новогиреево', 
'Новодачная', 'Новокосино', 'Новокузнецкая', 'Новопеределкино', 'Новослободская', 'Новохохловская', 'Новоясеневская', 'Новые Черемушки', 'Одинцово', 'Озерная', 'Окружная', 'Окская', 
'Октябрьская', 'Октябрьское поле', 'Ольховая', 'Опалиха', 'Орехово', 'Остафьево', 'Отрадное', 'Охотный ряд', 'Павелецкая', 'Павшино', 'Панфиловская', 'Парк Культуры', 'Парк Победы', 
'Партизанская', 'Пенягино', 'Первомайская', 'Перерва', 'Перово', 'Петровский Парк', 'Петровско-Разумовская', 'Печатники', 'Пионерская', 'Планерная', 'Площадь Гагарина', 'Площадь Ильича', 
'Площадь Революции', 'Подольск', 'Покровское', 'Покровское-Стрешнево', 'Полежаевская', 'Полянка', 'Пражская', 'Преображенская площадь', 'Прокшино', 'Пролетарская', 'Проспект Вернадского', 
'Проспект Мира', 'Профсоюзная', 'Пушкинская', 'Пятницкое шоссе', 'Рабочий поселок', 'Раменки', 'Рассказовка', 'Речной вокзал', 'Ржевская', 'Рижская', 'Римская', 'Ростокино', 'Рубцовская', 
'Румянцево', 'Рязанский проспект', 'Савеловская', 'Саларьево', 'Свиблово', 'Севастопольская', 'Селигерская', 'Семеновская', 'Серпуховская', 'Сетунь', 'Силикатная', 'Сколково', 
'Славянский бульвар', 'Смоленская', 'Сокол', 'Соколиная гора', 'Сокольники', 'Солнцево', 'Спартак', 'Спортивная', 'Сретенский бульвар', 'Стахановская', 'Стрешнево', 'Строгино', 'Стромынка', 
'Студенческая', 'Сухаревская', 'Сходненская', 'Таганская', 'Тверская', 'Театральная', 'Текстильщики', 'Телецентр', 'Теплый Стан', 'Тестовская', 'Технопарк', 'Тимирязевская', 'Третьяковская', 
'Трикотажная', 'Тропарево', 'Трубная', 'Тульская', 'Тургеневская', 'Тушинская', 'Угрешская', 'Улица 1905 года', 'Улица 800-летия Москвы', 'Улица Академика Королева', 'Улица Академика Янгеля', 
'Улица Горчакова', 'Улица Дмитриевского', 'Улица Милашенкова', 'улица Народного ополчения', 'Улица Сергея Эйзенштейна', 'Улица Скобелевская', 'Улица Старокачаловская', 'Университет', 
'Филатов Луг', 'Филевский парк', 'Фили', 'Фонвизинская', 'Фрунзенская', 'Хлебниково', 'Ховрино', 'Хорошево', 'Хорошевская', 'Царицыно', 'Цветной бульвар', 'ЦСКА', 'Черкизовская', 'Чертановская', 
'Чеховская', 'Чистые пруды', 'Чкаловская', 'Шаболовская', 'Шелепиха', 'Шереметьевская', 'Шипиловская', 'Шоссе Энтузиастов', 'Щелковская', 'Щербинка', 'Щукинская', 'Электрозаводская', 
'Юго-Восточная', 'Юго-Западная', 'Южная', 'Ясенево']


st.title('Отчет')



with st.form("my_form", clear_on_submit = True):
    tab1, tab2, tab3, tab4 = st.tabs(["Задание на оценку", "Объект оценки", "Здание и подъезд",'фото'])
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
        price = st.text_input('Рыночная стоимость')
        metro = st.selectbox('Метро', metro_data)
        tm1, tm2 = st.columns(2)
        with tm1:
            time_metro = st.number_input('Растояние до Метро', format = '%d', value = 1)
        with tm2:
            way = st.radio(
            "",
            ('пешком', 'на общественном транспорте'))

        # Инфа о Квартире
        flor = st.text_input('Этаж расположения квартиры')#Этаж
        rooms = st.text_input('Количество комнат в квартире')#Комнаты
        status = st.selectbox('Общее состояние помещений', status_data) #Общее состояние помещений
        kv1, kv2, kv3= st.columns(3)
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
            ('Совмещенный', 'Раздельный')) # WC
        st.subheader('Инженерные коммуникации')
        #tech = st.text_input('Техническое обеспечение квартиры','холодное и горячее водоснабжение от городской сети, электроснабжение, центральное отопление, канализация.')
        tech = st.multiselect('Техническое обеспечение квартиры',['холодное и горячее водоснабжение от городской сети', 'электроснабжение', 'центральное отопление', 'канализация'])
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
        
        podz1, podz2, = st.columns(2)
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

        q5 = st.radio(
        "",
        ('Домофон', 'Кодовый замок', 'Консьерж в подъезде'), horizontal = True)
    with tab4:
        pic = st.file_uploader('Выберете фото')
        
    upload = st.form_submit_button("Сохранить")
    
if upload:
    with st.spinner('Wait for it...'):
        doc = DocxTemplate('Макет ОПЕКА.docx')
        imagen = InlineImage(doc, pic, width=Mm(150))

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
                    'Адрес' : adres,    # Объект оценки
                    'Обладать' : prav_oblad,
                    'Кадастр': kadast_number,  
                    'Площадь': area,    
                    'Стоимость': price, 
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
                    'Перим': q3, 
                    'Охрана': q4, 
                    'Домофон': q5,
                    'imagen': imagen
                    }
        
        #imagen = InlineImage(doc, pic, width=Mm(150))# width is in millimetres
        #context_pic = {'imagen': imagen}
        


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
