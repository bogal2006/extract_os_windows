from winreg import *
import time
from datetime import datetime as dt, timedelta
from struct import unpack
import subprocess


# функция извлекающая данные о пользователях

def user_data():
    user_name = subprocess.check_output('wmic useraccount get name', shell=True, encoding='CP866').split()
    return user_name


# функция извлекающая время выключения хоста
def winreg_os():
    win_info = dict()
    if comp_shutdown := OpenKeyEx(HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Windows"):
        shutdown_time_bin = QueryValueEx(comp_shutdown, 'ShutdownTime')[0]
        shutdown_time = (dt(1601, 1, 1) + timedelta(microseconds=float(unpack("<Q", shutdown_time_bin)[0]) / 10)). \
            strftime('%Y-%m-%d %H:%M:%S')
        win_info.update({'shutdown': shutdown_time})
    return win_info


# функция извлекающая сетевой стек
def stack_network():
    network_info = dict()
    net_info = OpenKeyEx(HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Services\Tcpip\Parameters")
    network_info.update({'Current DHCP Server': QueryValueEx(net_info, 'DhcpNameServer')[0]})
    network_info.update({'Domain name': QueryValueEx(net_info, 'ICSDomain')[0]})
    return network_info


# функция извлекающая сетевой стек
def stack_network_inside1():
    network_info = dict()
    net_info = OpenKeyEx(HKEY_LOCAL_MACHINE, r'SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\{'
                                             r'e783fbfe-2a7e-417a-be12-a6d7bc94d6c2}')
    network_info.update({'DhcpDefaultGateway': QueryValueEx(net_info, 'DhcpDefaultGateway')[0]})
    network_info.update({'DhcpDomain': QueryValueEx(net_info, 'DhcpDomain')[0]})
    network_info.update({'DhcpIPAddress': QueryValueEx(net_info, 'DhcpIPAddress')[0]})
    network_info.update({'DhcpSubnetMask': QueryValueEx(net_info, 'DhcpSubnetMask')[0]})
    lease_obtained_time_bin = QueryValueEx(net_info, 'LeaseObtainedTime')[0]
    network_info.update(
        {'leaseObtainedTime': time.strftime('%Y-%m-%d %H:%M:%S', time.gmtime(float(lease_obtained_time_bin)))})

    return network_info


# функция извлекающая сетевой стек
def stack_network_inside2():
    network_info = dict()
    net_info = OpenKeyEx(HKEY_LOCAL_MACHINE, r'SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\{'
                                             r'e783fbfe-2a7e-417a-be12-a6d7bc94d6c2}\14C6E4163747')
    network_info.update({'DhcpDefaultGateway': QueryValueEx(net_info, 'DhcpDefaultGateway')[0]})
    network_info.update({'DhcpIPAddress': QueryValueEx(net_info, 'DhcpIPAddress')[0]})
    network_info.update({'DhcpSubnetMask': QueryValueEx(net_info, 'DhcpSubnetMask')[0]})
    lease_obtained_time_bin = QueryValueEx(net_info, 'LeaseObtainedTime')[0]
    network_info.update(
        {'leaseObtainedTime': time.strftime('%Y-%m-%d %H:%M:%S', time.gmtime(float(lease_obtained_time_bin)))})

    return network_info


# функция извлекающая сетевой стек
def stack_network_inside3():
    network_info = dict()
    net_info = OpenKeyEx(HKEY_LOCAL_MACHINE, r'SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\{'
                                             r'f40071ba-fc7f-4fe8-8881-6d83a938dfbf}')
    network_info.update({'DhcpDefaultGateway': QueryValueEx(net_info, 'DhcpDefaultGateway')[0]})
    network_info.update({'DhcpIPAddress': QueryValueEx(net_info, 'DhcpIPAddress')[0]})
    network_info.update({'DhcpSubnetMask': QueryValueEx(net_info, 'DhcpSubnetMask')[0]})
    lease_obtained_time_bin = QueryValueEx(net_info, 'LeaseObtainedTime')[0]
    network_info.update(
        {'leaseObtainedTime': time.strftime('%Y-%m-%d %H:%M:%S', time.gmtime(float(lease_obtained_time_bin)))})

    return network_info


# функция извлекающая сетевой стек
def stack_network_inside4():
    network_info = dict()
    net_info = OpenKeyEx(HKEY_LOCAL_MACHINE, r'SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\{'
                                             r'f9cb1bff-ccbf-4194-b99e-dc54bb512c68}')
    network_info.update({'DhcpDefaultGateway': QueryValueEx(net_info, 'DhcpDefaultGateway')[0]})
    network_info.update({'DhcpDomain': QueryValueEx(net_info, 'DhcpDomain')[0]})
    network_info.update({'DhcpIPAddress': QueryValueEx(net_info, 'DhcpIPAddress')[0]})
    network_info.update({'DhcpSubnetMask': QueryValueEx(net_info, 'DhcpSubnetMask')[0]})
    lease_obtained_time_bin = QueryValueEx(net_info, 'LeaseObtainedTime')[0]
    network_info.update(
        {'leaseObtainedTime': time.strftime('%Y-%m-%d %H:%M:%S', time.gmtime(float(lease_obtained_time_bin)))})

    return network_info


# функция извлекающая версию виндовс
def current_version():
    cur_ver_info = dict()
    net_info = OpenKeyEx(HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion')
    cur_ver_info.update({'ProductName': QueryValueEx(net_info, 'ProductName')[0]})
    cur_ver_info.update({'BuildLab': QueryValueEx(net_info, 'BuildLab')[0]})
    cur_ver_info.update({'EditionID': QueryValueEx(net_info, 'EditionID')[0]})
    cur_ver_info.update({'ProductId': QueryValueEx(net_info, 'ProductId')[0]})
    install_date = QueryValueEx(net_info, 'InstallDate')[0]
    cur_ver_info.update(
        {'InstallDate': time.strftime('%Y-%m-%d %H:%M:%S', time.gmtime(float(install_date)))})

    return cur_ver_info


# функция извлекающая временную зону
def time_zone():
    timezone_info = dict()
    net_info = OpenKeyEx(HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\TimeZoneInformation")
    timezone_info.update({'TimeZoneKeyName': QueryValueEx(net_info, 'TimeZoneKeyName')[0]})
    return timezone_info


# функция извлекающая имя компьютера
def comp_name():
    comp_name_info = dict()
    net_info = OpenKeyEx(HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName")
    comp_name_info.update({'ComputerName': QueryValueEx(net_info, 'ComputerName')[0]})
    return comp_name_info


# функция извлекающая установленные программы
def soft_list():
    data_soft_subkey = dict()
    path_soft_unistal_key = OpenKeyEx(HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
    # получаем количество подключей от основного ключа info = QueryInfoKey(path_soft_unistal_key)[0]
    for i in range(QueryInfoKey(path_soft_unistal_key)[0]):
        try:
            keynames = EnumKey(path_soft_unistal_key, i)
            data_soft_subkey.update({f'{i + 1}': f' {keynames}'})
        except FileNotFoundError:
            pass
    return data_soft_subkey


# extract processor data
def central_processor():
    processor_info = dict()
    net_info = OpenKeyEx(HKEY_LOCAL_MACHINE, r'HARDWARE\DESCRIPTION\System\CentralProcessor\0')
    processor_info.update({'ProcessorNameString': QueryValueEx(net_info, 'ProcessorNameString')[0]})
    return processor_info


# функция создания имени файла из имени хоста и даты
def create_name_file(name_comp: str) -> str:
    now_date = dt.now().strftime('%y_%m_%d')
    name_file = f'{now_date}_{name_comp}.txt'
    return name_file


# функция записи данных в файл
def print_in_file(name_file_txt):
    with open(name_file_txt, 'w', encoding='utf-8') as file:
        file.write(f'************************************************\n')
        file.write(f'имена зарегистрированных в системе пользователей:\n')
        length_list = len(user_data())
        for numer, name in enumerate(user_data(), start=1):
            if numer == length_list:
                file.write(f'{name}.\n')
            else:
                file.write(f'{name}, ')
        file.write(f'************************************************\n')
        file.write(f'************************************************\n')
        file.write(f'имя компьютера:\n')
        for key, value in comp_name().items():
            file.write(f'{key}:{value}\n')
        file.write(f'************************************************\n')
        file.write(f'************************************************\n')
        file.write(f'время последнего выключения:\n')
        for key, value in winreg_os().items():
            file.write(f'{key}:{value}\n')
        file.write(f'************************************************\n')
        file.write(f'************************************************\n')
        file.write(f'сведения о процессоре:\n')
        for key, value in central_processor().items():
            file.write(f'{key}:{value}\n')
        file.write(f'************************************************\n')
        file.write(f'************************************************\n')
        file.write(f'сведения о сетевых подключениях:\n')
        for key, value in stack_network().items():
            file.write(f'{key}:{value}\n')
        for key, value in stack_network_inside1().items():
            file.write(f'{key}:{value}\n')
        for key, value in stack_network_inside2().items():
            file.write(f'{key}:{value}\n')
        for key, value in stack_network_inside3().items():
            file.write(f'{key}:{value}\n')
        for key, value in stack_network_inside4().items():
            file.write(f'{key}:{value}\n')
        file.write(f'************************************************\n')
        file.write(f'************************************************\n')
        file.write(f'версия windows:\n')
        for key, value in current_version().items():
            file.write(f'{key}:{value}\n')
        file.write(f'************************************************\n')
        file.write(f'************************************************\n')
        file.write(f'сведения о установенной временной зоне:\n')
        for key, value in time_zone().items():
            file.write(f'{key}:{value}\n')
        file.write(f'************************************************\n')
        file.write(f'************************************************\n')
        file.write(f'установленное программное обеспечение:\n')
        for key, value in soft_list().items():
            file.write(f'{key}:{value}\n')
        file.write(f'************************************************\n')


def main():
    print_in_file(create_name_file(comp_name().get('ComputerName')))
    a1 = user_data()
    a = winreg_os()
    b = stack_network()
    c = stack_network_inside1()
    e = stack_network_inside2()
    d = stack_network_inside3()
    f = stack_network_inside4()
    g = current_version()
    h = time_zone()
    i1 = comp_name()
    i2 = soft_list()
    i3 = central_processor()
    print('************************************************')
    print('имена зарегистрированных пользователей:')
    print(a1)
    # print('************************************************')
    print('************************************************')
    print('имя компьютера:')
    print(i1)
    print('************************************************')
    print('************************************************')
    print('время последнего выключения')
    print(a)
    print('************************************************')
    print('************************************************')
    print('сведения о процессоре')
    print(i3)
    print('************************************************')
    print('************************************************')
    print('сведения о сетевых подключениях')
    print(b, sep='\n')
    print(c, sep='\n')
    print(e, sep='\n')
    print(d, sep='\n')
    print(f, sep='\n')
    print('************************************************')
    print('************************************************')
    print('версия windows')
    print(g, sep='\n')
    print('************************************************')
    print('************************************************')
    print('сведения о установенной временной зоне')
    print(h, sep='\n')
    print('************************************************')
    print('************************************************')
    print('установленное программное обеспечение')
    print(i2, sep='\n')
    print('************************************************')


if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
