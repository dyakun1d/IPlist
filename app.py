import streamlit as st
import ipaddress
from math import ceil
from excel_report import create_excel

# Интерфейс Streamlit
st.title("creating Exel report")

# Интерфейс для ввода данных
author = st.text_input("Autor")
creation_date = st.date_input("Date")
project_name = st.text_input("Project name")
power_wp = st.number_input("write the power of the station (MWp)", min_value=1)

# Ввод количества каждого устройства
inverter_type = st.selectbox("inverter type", ["KACO", "Sungrow", "Huawei"])
inverter_count = st.number_input("how many inverters do we have ?", min_value=0)
logger_count = st.number_input("how many loggers do we have?", min_value=0) if inverter_type != "KACO" else 0
trafo_station_count = st.number_input("how many trafostation do we have ?", min_value=0)
scb_count = st.number_input("how many SCBs do we have?", min_value=0)
ncu_count = st.number_input("how many NCU do we have?", min_value=0)

# Логика для подключения инверторов и SCBs
connection_type_inverter = "Modbus" if inverter_type != "KACO" else "TCP/IP"
inverters_per_station = {}
if connection_type_inverter == "Modbus" and trafo_station_count > 0:
    for station_index in range(1, trafo_station_count + 1):
        station_letter = chr(64 + station_index)  # Получаем букву для трафостанции (A, B, C, ...)
        count = st.number_input(
            f"Сколько инверторов подключено к трафостанции {station_letter}?",
            min_value=0,
            max_value=inverter_count
        )
        inverters_per_station[station_letter] = count
        inverter_count -= count  # Уменьшаем количество оставшихся инверторов

# Установка начального IP и маски подсети в зависимости от мощности проекта
if power_wp > 20:
    start_ip_str = "172.24.0.11"
    subnet_mask = "255.255.0.0"
    gateway = "172.24.0.1"
    network = ipaddress.ip_network("172.24.0.0/16")
else:
    start_ip_str = "10.0.10.10"
    subnet_mask = "255.255.254.0"
    gateway = "10.0.10.1"
    network = ipaddress.ip_network("10.0.10.0/23")

start_ip = ipaddress.ip_address(start_ip_str)
ip_iterator = network.hosts()

# Пропускаем IP-адреса до start_ip
for ip in ip_iterator:
    if ip >= start_ip:
        break

# Проверка и создание отчетов при нажатии кнопки
if st.button("make a report"):
    # Собрать данные для отчета
    author_data = {
        'author': author,
        'creation_date': creation_date.strftime('%Y-%m-%d'),
        'project_name': project_name,
        'power_wp': power_wp
    }

    # Подготовка данных для таблицы
    headers = ['Device', 'type of connection', '№', 'IPs', 'subnet mask', 'Gateway', 'Modbus adr.', 'Type', 'Cable Type']
    
    # Добавляем столбец "Параметры подключения" для инверторов, если они не KACO
    if inverter_type != "KACO":
        headers.append('Device')
    
    main_table_data = [headers]

    letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    modbus_adr = 1  # Начало нумерации Modbus адресов для SCBs и инверторов
    sensor_modbus_adr = 100  # Начало нумерации Modbus адресов для Temperature Sensors и Pyranometers
    cable_type = "LiYCY (TP) 2x2x0,8"  # Тип кабеля для Modbus устройств

    # Генерация данных для таблицы
    for device_name, count, connection_type in [
        ('Loggers', logger_count, 'TCP/IP'),
        ('Trafostation', trafo_station_count, 'TCP/IP'),
        ('SCBs', scb_count, 'Modbus'),  # SCBs всегда Modbus
        ('NCU', ncu_count, 'TCP/IP'),
        ('Main cabinet', 1, 'TCP/IP'),  # Добавляем Main cabinet
        ('PQM', 1, 'TCP/IP'),  # Добавляем PQM
        ('POI', 2, 'TCP/IP'),  # Добавляем POI Main и Backup
        ('Q.Reader cabinet', 1, 'TCP/IP')  # Добавляем Q.Reader cabinet
    ]:
        if count > 0:
            for i in range(count):
                if device_name == 'Trafostation':
                    # Для трафостанций с IP-адресом тип подключения TCP/IP
                    if power_wp > 20:
                        ip_address = str(start_ip + i * 256)
                    else:
                        ip_address = str(next(ip_iterator))
                    device_id = letters[i]  # Обозначение A, B, C, ...
                    main_table_data.append([device_name, connection_type, device_id, ip_address, subnet_mask, gateway, '', '', ''])

                    # Получаем третий октет IP-адреса трафостанции
                    third_octet = ip_address.split('.')[2]

                    # Добавляем соответствующий шкаф мониторинга
                    for j in range(6):  # Выделяем 6 IP-адресов для каждого шкафа мониторинга
                        ip_monitor = f"{ip_address.split('.')[0]}.{ip_address.split('.')[1]}.{third_octet}.{int(ip_address.split('.')[3]) + j + 1}"
                        device_name_monitor = 'Monitoring cabinet'
                        main_table_data.append([device_name_monitor, connection_type, f"{device_id}-{j+1}", ip_monitor, subnet_mask, gateway, '', '', ''])

                elif device_name == 'SCBs':
                    # Для SCBs с Modbus уникальные адреса
                    ip_address = ''
                    main_table_data.append([device_name, connection_type, i + 1, ip_address, subnet_mask, '', modbus_adr, '', cable_type])
                    modbus_adr += 1
                elif device_name == 'Main cabinet':
                    # Для Main cabinet выделяем 10 IP-адресов
                    device_id = 1
                    for _ in range(10):
                        ip_address = str(next(ip_iterator))
                        main_table_data.append([device_name, connection_type, device_id, ip_address, subnet_mask, gateway, '', '', ''])
                elif device_name == 'PQM':
                    # Для PQM резервируем один IP-адрес
                    ip_address = str(next(ip_iterator))
                    main_table_data.append([device_name, connection_type, 1, ip_address, subnet_mask, gateway, '', '', ''])
                elif device_name == 'POI':
                    # Для POI резервируем два IP-адреса (Main и Backup)
                    ip_address = str(next(ip_iterator))
                    main_table_data.append([device_name, connection_type, 'Main' if i == 0 else 'Backup', ip_address, subnet_mask, gateway, '', '', ''])
                elif device_name == 'Q.Reader cabinet':
                    # Для Q.Reader cabinet выделяем 6 IP-адресов
                    for j in range(6):
                        ip_address = str(next(ip_iterator))
                        main_table_data.append([device_name, connection_type, j + 1, ip_address, subnet_mask, gateway, '', '', ''])

                else:
                    # Для других TCP/IP устройств
                    ip_address = str(next(ip_iterator))
                    main_table_data.append([device_name, connection_type, i + 1, ip_address, subnet_mask, gateway, '', '', ''])

    # Добавление информации о NCU_workstation если выбрана опция NCU
    if ncu_count > 0:
        for device_type in ['Main', 'Backup']:
            device_name = 'NCU_workstation'
            connection_type = 'TCP/IP'
            ip_address = str(next(ip_iterator))
            main_table_data.append([device_name, connection_type, device_type, ip_address, subnet_mask, gateway, '', '', ''])

    # Добавление информации о подключении инверторов к станциям для KACO
    if inverter_type == "KACO":
        for i in range(inverter_count):
            ip_address = str(next(ip_iterator))
            main_table_data.append([f"Inverter {i + 1}", 'TCP/IP', i + 1, ip_address, subnet_mask, gateway, '', '', ''])

    # Добавление информации о подключении инверторов к станциям для других типов (Modbus)
    if inverter_type != "KACO" and connection_type_inverter == "Modbus":
        for station_letter, count in inverters_per_station.items():
            modbus_adr = 1  # Начало нумерации Modbus адресов для каждого инвертора на станции
            for inv_index in range(1, count + 1):
                device_id = f"{station_letter}1.{inv_index}"  # Форматирование ID: A1.1, A1.2, ...
                row = [
                    f"Инвертор {device_id}",
                    connection_type_inverter,
                    device_id,
                    '',
                    '',
                    '',
                    modbus_adr,
                    '19200/8/N/1',  # Параметры подключения для Modbus инверторов
                    cable_type  # Тип кабеля для Modbus инверторов
                ]
                main_table_data.append(row)
                modbus_adr += 1

    # Добавление раздела "Sensors"
    sensor_count = ceil(power_wp / 5) * 2  # Два датчика температуры на каждые 5 Wp
    pyranometer_count = ceil(power_wp / 5)  # Один пирометр на каждые 5 Wp

    if sensor_count > 0 or pyranometer_count > 0:
        main_table_data.append(['Sensors', '', '', '', '', '', '', '', ''])

        for i in range(sensor_count):
            device_name = 'Temperature Sensor'
            connection_type = 'Modbus'
            device_id = i + 1
            ip_address = ''
            subnet_mask = ''
            type_mark = 'Tm-RS485-MB'
            main_table_data.append([device_name, connection_type, device_id, ip_address, subnet_mask, '', sensor_modbus_adr, type_mark, cable_type])
            sensor_modbus_adr += 1

        for i in range(pyranometer_count):
            device_name = 'Pyranometer'
            connection_type = 'Modbus'
            device_id = i + 1
            ip_address = ''
            subnet_mask = ''
            type_mark = 'SMP10'
            main_table_data.append([device_name, connection_type, device_id, ip_address, subnet_mask, '', sensor_modbus_adr, type_mark, cable_type])
            sensor_modbus_adr += 1

    # Создание Excel отчета
    file_name = create_excel(author_data, main_table_data)

    # Кнопка для скачивания Excel
    st.download_button(
        label="Скачать Excel",
        data=open(file_name, "rb").read(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
