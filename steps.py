from liteflow.core import *
from actions import *
from datetime import datetime as date, timedelta
from pywinauto.timings import wait_until_passes
from pywinauto import keyboard
from pywinauto import Application
from uiarecorder.play import find_element_by_uia
import pathlib
from actions import rpa_logging as log
from pywinauto.timings import wait_until, TimeoutError
import pandas as pd

recorder_elems_dir = str(pathlib.Path(__file__).parent.resolve().joinpath('recorder_elems'))
    
    
class SetDay(StepBody):
    '''
    Установка времени: начало текущего месяца и T-1 день
    и создание папки для выгрузки и сохранения отчетов
    '''
    
    def run(self, context: StepExecutionContext) -> ExecutionResult:
        log.info('Создание папки для сохранения отчетов')
        yesterday = datetime.now() - timedelta(days=1)
        save_date = yesterday.strftime('%d%m%Y')
        main_path = r'C:\Users\RPA_024\Desktop'
        if not os.path.exists(os.path.join(main_path, save_date)):
            os.makedirs(os.path.join(main_path, save_date))

        return ExecutionResult.next()
    
    
class StartFlow(StepBody):
    '''
    Признак начала CollectReports Workflow
    
    '''
    def run(self, context: StepExecutionContext) -> ExecutionResult:
        log.info("Started")
        return ExecutionResult.next()
    
    
class SendNotification(StepBody):
    '''
    Уведомление о начале роботы
    
    '''
    
    def run(self, context: StepExecutionContext) -> ExecutionResult:
        mail_data = {
            'sender': 'RPA_024@halykbank.nb',
            'subject': 'start',
            'body': 'Робот начал работу',
            'receiver': ['BauyrzhanAl@halykbank.kz']
        }
# , 'KAMILAK2@halykbank.kz'
        endpoint = "http://alav741/api/send/"

        requests.post(endpoint, data=mail_data)
        
        return ExecutionResult.next()
    
    
class GetReport20(StepBody):
    '''
    Выгрузка отчета 20 из SAP BW
    
    '''
    
    def run(self, context: StepExecutionContext) -> ExecutionResult:
#         report = "Приложение 20(н) ver.2.6_AS"
        report = "XA_DBUIO_PRIL20H_VAL_POZ"
        path = r"\\ala300n02\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\settings.xlsx"
        settings_data = pd.read_excel(path, engine='openpyxl', sheet_name='Учетки')
        sap_login = '000' + str(settings_data['Учетка SAP'][0])[:-2]
        log.info(f'Ввел логин {sap_login}' )
        sap_password = settings_data['Пароль SAP'][0][:-1] + '{+}'
        log.info(f'Ввел пароль {sap_password}')
        try:
            wait_until_passes(350, 5, lambda: start_uia_app(r'C:\Program Files\SAP BusinessObjects\Office AddIn\BiOfficeLauncher.exe'))
        except Exception as err:
            log.error(f'Ошибка в запуске Analysis For Office {err}')
            
        try:
            excel_app = Application(backend='uia').connect(path='excel.exe')
            excel_app.top_window().set_focus()

            try:
                wait_until_passes(10, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').set_focus())
                wait_until_passes(10, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            except TimeoutError as err:
                log.error(f'Ошибка нажатия на файл - {err}')
                keyboard.send_keys('^o')

            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_parametrs').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_nadstroiki').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_control').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_com').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_go').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_switch_analysis').click_input())
            keyboard.send_keys('{VK_ADD}')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_switch_ok').click_input())

        except Exception as err:
            log.error(f'Ошибка добавления Analysis в Excel {err}')
            
        try:
            excel = Application(backend='uia').connect(path='excel.exe')
            excel.top_window().maximize()

            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_analysis').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/analysis_open').click_input())
            keyboard.send_keys('{TAB 3}')
            keyboard.send_keys('{ENTER}')

            if find_element_by_uia(f'{recorder_elems_dir}/sap_analysis_for_office_window'):
                find_element_by_uia(f'{recorder_elems_dir}/sap_analysis_office_ok_btn').click_input()

            keyboard.send_keys('{ENTER}', pause=1)
            keyboard.send_keys('{TAB 5}')

            keyboard.send_keys(sap_login)
            keyboard.send_keys('{TAB}')

            keyboard.send_keys(sap_password)
            keyboard.send_keys('{ENTER}')

        except Exception as err:
            log.error(f'Ошибка в авторизации Analysis {err}')
            
        try:
            keyboard.send_keys('{TAB}')
            keyboard.send_keys('{TAB}')
            keyboard.send_keys(report, with_spaces=True)
            keyboard.send_keys('{ENTER 2}')

        except Exception as err:
            log.error(f'Ошибка ввода отчета {err}')
            
            
        try:
            my_date = date.today().strftime("%A")
            if my_date == "Monday":
                prev_date = (date.now() - timedelta(3)).strftime('%d.%m.%Y')
            else:
                prev_date = (date.now() - timedelta(1)).strftime('%d.%m.%Y')

            enter_date = prev_date
            period_date = '0' + prev_date[3:]

            time.sleep(120)  # Нужно время для того чтобы робот увидел окно
            
            keyboard.send_keys('{TAB 6}')
            log.info(f'Ввожу начало периода {enter_date}')
            keyboard.send_keys(enter_date)

            keyboard.send_keys('{TAB 2}', pause=1)
            log.info(f'Ввожу конец периода {enter_date}')
            keyboard.send_keys(enter_date)
            
            keyboard.send_keys('{TAB 2}', pause=1)
            log.info(f'Ввожу дату провизий {enter_date}')
            keyboard.send_keys(enter_date)
            
            keyboard.send_keys('{TAB 2}', pause=1)
            log.info(f'Ввожу Период/финансовый год {period_date}')
            keyboard.send_keys(period_date)
            
            keyboard.send_keys('{TAB 2}', pause=1)
            log.info(f'Ввожу дату собственного капитала {enter_date}')
            keyboard.send_keys(enter_date)

            keyboard.send_keys('{TAB}')
            keyboard.send_keys('{TAB 2}')
    
            log.info('Нажимаю OK')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/btn_ok').click_input())

        except Exception as err:
            log.error(f'Ошибка в заполнении периода {err}')
            
        try:
            save_date_month = enter_date.replace(".", "")[2:]
            save_date_day = enter_date.replace(".", "")
            path = rf"\\ala300n02\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\{save_date_month}\{save_date_day}\20.xls"
            log.info(f'Начинаю сохранение отчета')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_save_as').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/etot_computer_excel').double_click_input())
            time.sleep(10)
            keyboard.send_keys(path, with_spaces=True)
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/save_excel_type').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/save_as_xls').click_input())
            log.info(f'Ввод отчета')

            keyboard.send_keys('{ENTER}')
            log.info(f'Сохранил отчет {bank_branch_name}')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/continue_save_excel').double_click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/btn_no_in_save_excel').double_click_input())
        except Exception as err:
            log.error(f'Ошибка сохранения отчета {err}')
        
        
        return ExecutionResult.next()
    
    
class GetReport20(StepBody):
    '''
    Выгрузка отчета 22 из SAP BW
    
    '''
    
    def run(self, context: StepExecutionContext) -> ExecutionResult:
        report = "XA_DBUIO_PRILL22"
        path = r"\\ala300n02\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\settings.xlsx"
        settings_data = pd.read_excel(path, engine='openpyxl', sheet_name='Учетки')
        sap_login = '000' + str(settings_data['Учетка SAP'][0])[:-2]
        log.info(f'Ввел логин {sap_login}' )
        sap_password = settings_data['Пароль SAP'][0][:-1] + '{+}'
        log.info(f'Ввел пароль {sap_password}')
        try:
            wait_until_passes(350, 5, lambda: start_uia_app(r'C:\Program Files\SAP BusinessObjects\Office AddIn\BiOfficeLauncher.exe'))
        except Exception as err:
            log.error(f'Ошибка в запуске Analysis For Office {err}')
            
        try:
            excel_app = Application(backend='uia').connect(path='excel.exe')
            excel_app.top_window().set_focus()

            try:
                wait_until_passes(10, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').set_focus())
                wait_until_passes(10, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            except TimeoutError as err:
                log.error(f'Ошибка нажатия на файл - {err}')
                keyboard.send_keys('^o')

            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_parametrs').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_nadstroiki').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_control').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_com').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_go').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_switch_analysis').click_input())
            keyboard.send_keys('{VK_ADD}')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_switch_ok').click_input())

        except Exception as err:
            log.error(f'Ошибка добавления Analysis в Excel {err}')
            
        try:
            excel = Application(backend='uia').connect(path='excel.exe')
            excel.top_window().maximize()

            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_analysis').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/analysis_open').click_input())
            keyboard.send_keys('{TAB 3}')
            keyboard.send_keys('{ENTER}')

            if find_element_by_uia(f'{recorder_elems_dir}/sap_analysis_for_office_window'):
                find_element_by_uia(f'{recorder_elems_dir}/sap_analysis_office_ok_btn').click_input()

            keyboard.send_keys('{ENTER}', pause=1)
            keyboard.send_keys('{TAB 5}')

            keyboard.send_keys(sap_login)
            keyboard.send_keys('{TAB}')

            keyboard.send_keys(sap_password)
            keyboard.send_keys('{ENTER}')

        except Exception as err:
            log.error(f'Ошибка в авторизации Analysis {err}')
            
        try:
            keyboard.send_keys('{TAB}')
            keyboard.send_keys('{TAB}')
            keyboard.send_keys(report, with_spaces=True)
            keyboard.send_keys('{ENTER 2}')

        except Exception as err:
            log.error(f'Ошибка ввода отчета {err}')
            
            
        try:
            my_date = date.today().strftime("%A")
            if my_date == "Monday":
                prev_date = (date.now() - timedelta(3)).strftime('%d.%m.%Y')
            else:
                prev_date = (date.now() - timedelta(1)).strftime('%d.%m.%Y')

            enter_date = prev_date
            period_date = '0' + prev_date[3:]

            time.sleep(120)  # Нужно время для того чтобы робот увидел окно
            
            keyboard.send_keys('{TAB 6}')
            log.info(f'Ввожу начало периода {enter_date}')
            keyboard.send_keys(enter_date)

            keyboard.send_keys('{TAB 2}', pause=1)
            log.info(f'Ввожу Период/финансовый год {period_date}')
            keyboard.send_keys(period_date)
            

            keyboard.send_keys('{TAB}')
            keyboard.send_keys('{TAB 2}')
    
            log.info('Нажимаю OK')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/btn_ok').click_input())

        except Exception as err:
            log.error(f'Ошибка в заполнении периода {err}')
            
        try:
            save_date_month = enter_date.replace(".", "")[2:]
            save_date_day = enter_date.replace(".", "")
            path = rf"\\ala300n02\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\{save_date_month}\{save_date_day}\22.xlsx"
            log.info(f'Начинаю сохранение отчета')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_save_as').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/etot_computer_excel').double_click_input())
            time.sleep(10)
            keyboard.send_keys(path, with_spaces=True)
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/save_excel_type').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/save_as_xls').click_input())
            log.info(f'Ввод отчета')

            keyboard.send_keys('{ENTER}')
            log.info(f'Сохранил отчет {bank_branch_name}')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/continue_save_excel').double_click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/btn_no_in_save_excel').double_click_input())
        except Exception as err:
            log.error(f'Ошибка сохранения отчета {err}')
        
        
        return ExecutionResult.next()
    

class GetReport3(StepBody):
    '''
    Выгрузка отчета 3 из SAP BW
    
    '''
    
    def run(self, context: StepExecutionContext) -> ExecutionResult:
        report = "XA_DBUIO_PRIL3_PRUDIKI"
        path = r"\\ala300n02\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\settings.xlsx"
        settings_data = pd.read_excel(path, engine='openpyxl', sheet_name='Учетки')
        sap_login = '000' + str(settings_data['Учетка SAP'][0])[:-2]
        log.info(f'Ввел логин {sap_login}' )
        sap_password = settings_data['Пароль SAP'][0][:-1] + '{+}'
        log.info(f'Ввел пароль {sap_password}')
        try:
            wait_until_passes(350, 5, lambda: start_uia_app(r'C:\Program Files\SAP BusinessObjects\Office AddIn\BiOfficeLauncher.exe'))
        except Exception as err:
            log.error(f'Ошибка в запуске Analysis For Office {err}')
            
        try:
            excel_app = Application(backend='uia').connect(path='excel.exe')
            excel_app.top_window().set_focus()

            try:
                wait_until_passes(10, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').set_focus())
                wait_until_passes(10, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            except TimeoutError as err:
                log.error(f'Ошибка нажатия на файл - {err}')
                keyboard.send_keys('^o')

            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_parametrs').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_nadstroiki').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_control').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_com').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_go').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_switch_analysis').click_input())
            keyboard.send_keys('{VK_ADD}')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_switch_ok').click_input())

        except Exception as err:
            log.error(f'Ошибка добавления Analysis в Excel {err}')
            
        try:
            excel = Application(backend='uia').connect(path='excel.exe')
            excel.top_window().maximize()

            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_analysis').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/analysis_open').click_input())
            keyboard.send_keys('{TAB 3}')
            keyboard.send_keys('{ENTER}')

            if find_element_by_uia(f'{recorder_elems_dir}/sap_analysis_for_office_window'):
                find_element_by_uia(f'{recorder_elems_dir}/sap_analysis_office_ok_btn').click_input()

            keyboard.send_keys('{ENTER}', pause=1)
            keyboard.send_keys('{TAB 5}')

            keyboard.send_keys(sap_login)
            keyboard.send_keys('{TAB}')

            keyboard.send_keys(sap_password)
            keyboard.send_keys('{ENTER}')

        except Exception as err:
            log.error(f'Ошибка в авторизации Analysis {err}')
            
        try:
            keyboard.send_keys('{TAB}')
            keyboard.send_keys('{TAB}')
            keyboard.send_keys(report, with_spaces=True)
            keyboard.send_keys('{ENTER 2}')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/continue_in_app_3').double_click_input())

        except Exception as err:
            log.error(f'Ошибка ввода отчета {err}')
            
            
        try:
            my_date = date.today().strftime("%A")
            if my_date == "Monday":
                prev_date = (date.now() - timedelta(3)).strftime('%d.%m.%Y')
            else:
                prev_date = (date.now() - timedelta(1)).strftime('%d.%m.%Y')

            enter_date = prev_date
            period_date = '0' + prev_date[3:]

            time.sleep(120)  # Нужно время для того чтобы робот увидел окно
            
            keyboard.send_keys('{TAB 6}')
            log.info(f'Ввожу начало периода {enter_date}')
            keyboard.send_keys(enter_date)

            keyboard.send_keys('{TAB 7}', pause=1)
    
            log.info('Нажимаю OK')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/btn_ok').click_input())

        except Exception as err:
            log.error(f'Ошибка в заполнении периода {err}')
            
        try:
            save_date_month = enter_date.replace(".", "")[2:]
            save_date_day = enter_date.replace(".", "")
            path = rf"\\ala300n02\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\{save_date_month}\{save_date_day}\3.xlsx"
            log.info(f'Начинаю сохранение отчета')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_save_as').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/etot_computer_excel').double_click_input())
            time.sleep(10)
            keyboard.send_keys(path, with_spaces=True)
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/save_excel_type').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/save_as_xls').click_input())
            log.info(f'Ввод отчета')

            keyboard.send_keys('{ENTER}')
            log.info(f'Сохранил отчет {bank_branch_name}')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/continue_save_excel').double_click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/btn_no_in_save_excel').double_click_input())
        except Exception as err:
            log.error(f'Ошибка сохранения отчета {err}')
        
        
        return ExecutionResult.next()
    
    
class GetReportCastody(StepBody):
    '''
    Выгрузка отчета кастоди из SAP BW
    
    '''
    
    def run(self, context: StepExecutionContext) -> ExecutionResult:
        report = "XA_DBUIO_EJED_OT_KASTOD_CL"
        path = r"\\ala300n02\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\settings.xlsx"
        settings_data = pd.read_excel(path, engine='openpyxl', sheet_name='Учетки')
        sap_login = '000' + str(settings_data['Учетка SAP'][0])[:-2]
        log.info(f'Ввел логин {sap_login}' )
        sap_password = settings_data['Пароль SAP'][0][:-1] + '{+}'
        log.info(f'Ввел пароль {sap_password}')
        try:
            wait_until_passes(350, 5, lambda: start_uia_app(r'C:\Program Files\SAP BusinessObjects\Office AddIn\BiOfficeLauncher.exe'))
        except Exception as err:
            log.error(f'Ошибка в запуске Analysis For Office {err}')
            
        try:
            excel_app = Application(backend='uia').connect(path='excel.exe')
            excel_app.top_window().set_focus()

            try:
                wait_until_passes(10, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').set_focus())
                wait_until_passes(10, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            except TimeoutError as err:
                log.error(f'Ошибка нажатия на файл - {err}')
                keyboard.send_keys('^o')

            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_parametrs').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_nadstroiki').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_control').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_com').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_go').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_switch_analysis').click_input())
            keyboard.send_keys('{VK_ADD}')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_switch_ok').click_input())

        except Exception as err:
            log.error(f'Ошибка добавления Analysis в Excel {err}')
            
        try:
            excel = Application(backend='uia').connect(path='excel.exe')
            excel.top_window().maximize()

            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_analysis').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/analysis_open').click_input())
            keyboard.send_keys('{TAB 3}')
            keyboard.send_keys('{ENTER}')

            if find_element_by_uia(f'{recorder_elems_dir}/sap_analysis_for_office_window'):
                find_element_by_uia(f'{recorder_elems_dir}/sap_analysis_office_ok_btn').click_input()

            keyboard.send_keys('{ENTER}', pause=1)
            keyboard.send_keys('{TAB 5}')

            keyboard.send_keys(sap_login)
            keyboard.send_keys('{TAB}')

            keyboard.send_keys(sap_password)
            keyboard.send_keys('{ENTER}')

        except Exception as err:
            log.error(f'Ошибка в авторизации Analysis {err}')
            
        try:
            keyboard.send_keys('{TAB}')
            keyboard.send_keys('{TAB}')
            keyboard.send_keys(report, with_spaces=True)
            keyboard.send_keys('{ENTER 2}')

        except Exception as err:
            log.error(f'Ошибка ввода отчета {err}')
            
            
        try:
            my_date = date.today().strftime("%A")
            if my_date == "Monday":
                prev_date = (date.now() - timedelta(3)).strftime('%d.%m.%Y')
            else:
                prev_date = (date.now() - timedelta(1)).strftime('%d.%m.%Y')

            enter_date = prev_date
            period_date = '0' + prev_date[3:]

            time.sleep(120)  # Нужно время для того чтобы робот увидел окно
            
            keyboard.send_keys('{TAB 6}')
            log.info(f'Ввожу начало периода {enter_date}')
            keyboard.send_keys(enter_date)

            keyboard.send_keys('{TAB 3}', pause=1)
    
            log.info('Нажимаю OK')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/btn_ok').click_input())

        except Exception as err:
            log.error(f'Ошибка в заполнении периода {err}')
            
        try:
            save_date_month = enter_date.replace(".", "")[2:]
            save_date_day = enter_date.replace(".", "")
            path = rf"\\ala300n02\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\{save_date_month}\{save_date_day}\кастоди.xlsm"
            log.info(f'Начинаю сохранение отчета')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_save_as').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/etot_computer_excel').double_click_input())
            time.sleep(10)
            keyboard.send_keys(path, with_spaces=True)
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/save_excel_type').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/save_as_xls').click_input())
            log.info(f'Ввод отчета')

            keyboard.send_keys('{ENTER}')
            log.info(f'Сохранил отчет {bank_branch_name}')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/continue_save_excel').double_click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/btn_no_in_save_excel').double_click_input())
        except Exception as err:
            log.error(f'Ошибка сохранения отчета {err}')
        
        
        return ExecutionResult.next()
    
    
class GetReport2(StepBody):
    '''
    Выгрузка отчета 2 из SAP BW
    
    '''
    
    def run(self, context: StepExecutionContext) -> ExecutionResult:
        report = "XA_DBUIO_PRIL2_GGK_2020_V2"
        path = r"\\ala300n02\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\settings.xlsx"
        settings_data = pd.read_excel(path, engine='openpyxl', sheet_name='Учетки')
        sap_login = '000' + str(settings_data['Учетка SAP'][0])[:-2]
        log.info(f'Ввел логин {sap_login}' )
        sap_password = settings_data['Пароль SAP'][0][:-1] + '{+}'
        log.info(f'Ввел пароль {sap_password}')
        try:
            wait_until_passes(350, 5, lambda: start_uia_app(r'C:\Program Files\SAP BusinessObjects\Office AddIn\BiOfficeLauncher.exe'))
        except Exception as err:
            log.error(f'Ошибка в запуске Analysis For Office {err}')
            
        try:
            excel_app = Application(backend='uia').connect(path='excel.exe')
            excel_app.top_window().set_focus()

            try:
                wait_until_passes(10, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').set_focus())
                wait_until_passes(10, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            except TimeoutError as err:
                log.error(f'Ошибка нажатия на файл - {err}')
                keyboard.send_keys('^o')

            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_parametrs').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_nadstroiki').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_control').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_com').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_go').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_switch_analysis').click_input())
            keyboard.send_keys('{VK_ADD}')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_switch_ok').click_input())

        except Exception as err:
            log.error(f'Ошибка добавления Analysis в Excel {err}')
            
        try:
            excel = Application(backend='uia').connect(path='excel.exe')
            excel.top_window().maximize()

            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_analysis').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/analysis_open').click_input())
            keyboard.send_keys('{TAB 3}')
            keyboard.send_keys('{ENTER}')

            if find_element_by_uia(f'{recorder_elems_dir}/sap_analysis_for_office_window'):
                find_element_by_uia(f'{recorder_elems_dir}/sap_analysis_office_ok_btn').click_input()

            keyboard.send_keys('{ENTER}', pause=1)
            keyboard.send_keys('{TAB 5}')

            keyboard.send_keys(sap_login)
            keyboard.send_keys('{TAB}')

            keyboard.send_keys(sap_password)
            keyboard.send_keys('{ENTER}')

        except Exception as err:
            log.error(f'Ошибка в авторизации Analysis {err}')
            
        try:
            keyboard.send_keys('{TAB}')
            keyboard.send_keys('{TAB}')
            keyboard.send_keys(report, with_spaces=True)
            keyboard.send_keys('{ENTER 2}')

        except Exception as err:
            log.error(f'Ошибка ввода отчета {err}')
            
            
        try:
            my_date = date.today().strftime("%A")
            if my_date == "Monday":
                prev_date = (date.now() - timedelta(3)).strftime('%d.%m.%Y')
            else:
                prev_date = (date.now() - timedelta(1)).strftime('%d.%m.%Y')

            enter_date = prev_date
            period_date = '0' + prev_date[3:]

            time.sleep(120)  # Нужно время для того чтобы робот увидел окно
            
            keyboard.send_keys('{TAB 6}')
            log.info(f'Ввожу начало периода {enter_date}')
            keyboard.send_keys(enter_date)

            keyboard.send_keys('{TAB 7}', pause=1)
    
            log.info('Нажимаю OK')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/btn_ok').click_input())

        except Exception as err:
            log.error(f'Ошибка в заполнении периода {err}')
            
        try:
            save_date_month = enter_date.replace(".", "")[2:]
            save_date_day = enter_date.replace(".", "")
            path = rf"\\ala300n02\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\{save_date_month}\{save_date_day}\2.xlsx"
            log.info(f'Начинаю сохранение отчета')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_save_as').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/etot_computer_excel').double_click_input())
            time.sleep(10)
            keyboard.send_keys(path, with_spaces=True)
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/save_excel_type').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/save_as_xls').click_input())
            log.info(f'Ввод отчета')

            keyboard.send_keys('{ENTER}')
            log.info(f'Сохранил отчет {bank_branch_name}')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/continue_save_excel').double_click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/btn_no_in_save_excel').double_click_input())
        except Exception as err:
            log.error(f'Ошибка сохранения отчета {err}')
        
        
        return ExecutionResult.next()
    
    
class GetReportSK(StepBody):
    '''
    Выгрузка отчета СК из SAP BW
    
    '''
    
    def run(self, context: StepExecutionContext) -> ExecutionResult:
        report = "XA_SOBS_KAPITAL_GGK_V3_AS"
        path = r"\\ala300n02\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\settings.xlsx"
        settings_data = pd.read_excel(path, engine='openpyxl', sheet_name='Учетки')
        sap_login = '000' + str(settings_data['Учетка SAP'][0])[:-2]
        log.info(f'Ввел логин {sap_login}' )
        sap_password = settings_data['Пароль SAP'][0][:-1] + '{+}'
        log.info(f'Ввел пароль {sap_password}')
        try:
            wait_until_passes(350, 5, lambda: start_uia_app(r'C:\Program Files\SAP BusinessObjects\Office AddIn\BiOfficeLauncher.exe'))
        except Exception as err:
            log.error(f'Ошибка в запуске Analysis For Office {err}')
            
        try:
            excel_app = Application(backend='uia').connect(path='excel.exe')
            excel_app.top_window().set_focus()

            try:
                wait_until_passes(10, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').set_focus())
                wait_until_passes(10, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            except TimeoutError as err:
                log.error(f'Ошибка нажатия на файл - {err}')
                keyboard.send_keys('^o')

            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_parametrs').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_nadstroiki').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_control').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_com').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_go').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_switch_analysis').click_input())
            keyboard.send_keys('{VK_ADD}')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_switch_ok').click_input())

        except Exception as err:
            log.error(f'Ошибка добавления Analysis в Excel {err}')
            
        try:
            excel = Application(backend='uia').connect(path='excel.exe')
            excel.top_window().maximize()

            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_analysis').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/analysis_open').click_input())
            keyboard.send_keys('{TAB 3}')
            keyboard.send_keys('{ENTER}')

            if find_element_by_uia(f'{recorder_elems_dir}/sap_analysis_for_office_window'):
                find_element_by_uia(f'{recorder_elems_dir}/sap_analysis_office_ok_btn').click_input()

            keyboard.send_keys('{ENTER}', pause=1)
            keyboard.send_keys('{TAB 5}')

            keyboard.send_keys(sap_login)
            keyboard.send_keys('{TAB}')

            keyboard.send_keys(sap_password)
            keyboard.send_keys('{ENTER}')

        except Exception as err:
            log.error(f'Ошибка в авторизации Analysis {err}')
            
        try:
            keyboard.send_keys('{TAB}')
            keyboard.send_keys('{TAB}')
            keyboard.send_keys(report, with_spaces=True)
            keyboard.send_keys('{ENTER 2}')

        except Exception as err:
            log.error(f'Ошибка ввода отчета {err}')
            
            
        try:
            my_date = date.today().strftime("%A")
            if my_date == "Monday":
                prev_date = (date.now() - timedelta(3)).strftime('%d.%m.%Y')
            else:
                prev_date = (date.now() - timedelta(1)).strftime('%d.%m.%Y')

            enter_date = prev_date
            period_date = '0' + prev_date[3:]

            time.sleep(120)  # Нужно время для того чтобы робот увидел окно
            
            keyboard.send_keys('{TAB 6}')
            log.info(f'Ввожу начало периода {enter_date}')
            keyboard.send_keys(enter_date)

            keyboard.send_keys('{TAB 2}', pause=1)
            log.info(f'Ввожу Период/финансовый год {period_date}')
            keyboard.send_keys(period_date)

            keyboard.send_keys('{TAB 2}', pause=1)
            log.info(f'Ввожу год')
            keyboard.send_keys('31.12.2023')

            keyboard.send_keys('{TAB}')
            keyboard.send_keys('{TAB 2}')
    
            log.info('Нажимаю OK')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/btn_ok').click_input())

        except Exception as err:
            log.error(f'Ошибка в заполнении периода {err}')
            
        try:
            save_date_month = enter_date.replace(".", "")[2:]
            save_date_day = enter_date.replace(".", "")
            path = rf"\\ala300n02\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\{save_date_month}\{save_date_day}\СК.xlsx"
            log.info(f'Начинаю сохранение отчета')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_file').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/excel_save_as').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/etot_computer_excel').double_click_input())
            time.sleep(10)
            keyboard.send_keys(path, with_spaces=True)
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/save_excel_type').click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/save_as_xls').click_input())
            log.info(f'Ввод отчета')

            keyboard.send_keys('{ENTER}')
            log.info(f'Сохранил отчет {bank_branch_name}')
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/continue_save_excel').double_click_input())
            wait_until_passes(350, 1, lambda: find_element_by_uia(f'{recorder_elems_dir}/btn_no_in_save_excel').double_click_input())
        except Exception as err:
            log.error(f'Ошибка сохранения отчета {err}')
        
        
        return ExecutionResult.next()
    

class RunMainPrudikiMacro(StepBody):
    '''
    Формирование и процесинг файла прудиков
    
    '''
    
    def run(self, context: StepExecutionContext) -> ExecutionResult:
        import os        
        import win32com.client as win32
        import comtypes, comtypes.client  
        from pywinauto import Application
        from uiarecorder.play import find_element_by_uia  
        from actions import rpa_log as log
        
        prudiki_macro = '''
        Sub prudiki()
        Application.AskToUpdateLinks = False
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False

        reportDate = InputBox("Введите отчетную дату в формате ДДММГГГГ")

        xx = Left(reportDate, 2) & "." & Mid(reportDate, 3, 2) & "." & Right(reportDate, 4)

        Sheets("Пруденциальные").Select
        Range("D2:D170").Select
        Selection.Copy
        Range("E2").Select
        Selection.PasteSpecial Paste:=xlPasteValues

        Sheets("Пруденциальные").Range("D2") = xx

        Sheets("2").Select
        Range("D3:F155").Select
        Selection.Copy
        Range("I3").Select
        Selection.PasteSpecial Paste:=xlPasteValues


        ' Курс USD Copy

        Sheets("провизии мсфо").Range("J1").Formula = "=SUBSTITUTE(TEXT(TODAY(),""ДД.ММ""),""."","""",1)"
        Sheets("провизии мсфо").Range("J2").FormulaR1C1 = "=CONCATENATE(MID(SUBSTITUTE(TEXT(TODAY(),""ДД.ММ""),""."","""",1),3,2),MID(SUBSTITUTE(TEXT(TODAY(),""ММ.ГГ""),""."","""",1),3,2))"

        Dim kk As String
        kk = Sheets("провизии мсфо").Range("J1").Value

        Dim mes As String
        mes = Sheets("провизии мсфо").Range("J2").Value


        'Workbooks.Open Filename:="W:\Общебанковский_сетевой_ресурс\KURS\KURS18\" & mes & "\kurs" & kk & ".XLS"

            'Sheets("Лист1").Select
            'Sheets("Лист1").Range("G15").Copy
            'Windows("prudnorm_" & xx & ".xlsm").Activate
            'Sheets("профизии мсфо").Select
            'Range("E2").Select
            'Selection.PasteSpecial Paste:=xlPasteValues


        'Windows("kurs" & kk & ".XLS").Close False



        ' 2 баланса Copy
        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\баланс.xls"
        Sheets("Баланс_форма").Select
        Range("A:G").Select
        Selection.Copy
        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets("Баланс").Select
        Range("A1").Select
        ActiveSheet.Paste

        Windows("баланс.xls").Activate
        Sheets("Требования-обязательства_форма").Select
        Range("F7").ClearContents
        Range("A:G").Select
        Selection.Copy
        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets("Баланс_условн").Select
        Range("A1").Select
        ActiveSheet.Paste
        Windows("баланс.xls").Activate

        ActiveWindow.Close SaveChanges:=False



        ' прил 700 Copy
        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\баланс2.xls"
        Sheets("Отчет_новая форма").Select


            Cells.Select
            With Selection
                .HorizontalAlignment = xlGeneral
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            With Selection
                .HorizontalAlignment = xlGeneral
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.UnMerge



        Columns("D:EQ").Select
        Selection.ClearContents
        Range("A:C").Select
        Selection.Copy
        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets("прил 700").Select
        Range("A1").Select
        ActiveSheet.Paste


        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\кастоди.xlsm"
        Sheets("форма").Select


        Range("E4").FormulaR1C1 = "=SUMIFS(C[1],C[2],""Резидент"")/1000"

            Windows("кастоди.xlsm").Activate

            Range("D4").Formula = _
             "=SUMIFS(C[2],C[3],""Резидент"")/1000+SUMIFS(C[2],C[3],""Нерезидент"")/1000"

            Range("D4").Select
            Selection.Copy
            Windows("prudnorm_" & xx & ".xlsm").Activate
            Sheets("2").Select
            Range("D185").Select
            Selection.PasteSpecial Paste:=xlPasteValues

            Windows("кастоди.xlsm").Activate

            ActiveWindow.Close SaveChanges:=True



        ' КВА
        Windows("баланс2.xls").Activate
        Sheets("Отчет_новая форма").Select
        Range("A:C").Select
        Selection.Copy


        KVA = "КВА " & Mid(reportDate, 3, 2) & " " & Right(reportDate, 4) & ".xlsx"
        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\" & KVA
        Sheets("Баланс 700Н").Select
        Range("A1").Select
        ActiveSheet.Paste


        Windows("баланс2.xls").Activate
        ActiveWindow.Close SaveChanges:=True



        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\кастоди.xlsm"
        Sheets("Форма").Select
        Range("D4").Select
        Selection.Copy
        Windows(KVA).Activate
        Sheets("Баланс 700Н").Select
        Range("L4").Select
        Selection.PasteSpecial Paste:=xlPasteValues

        Windows("кастоди.xlsm").Activate
        ActiveWindow.Close SaveChanges:=False

        Windows(KVA).Activate
        Sheets("КВА").Select
        Range("C2:C484").Select
        Selection.Copy
        Sheets("КВА свод").Select
        mmm = Cells(2, 1).Value
        Cells(2, mmm).Select
        Selection.PasteSpecial Paste:=xlPasteValues

        Cells(1, mmm).Select
        Selection.Copy
        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets("динамика КВА").Select
        Cells(3, mmm - 1).Select
        Selection.PasteSpecial Paste:=xlPasteValues

        Windows(KVA).Activate
        Cells(231, mmm).Select
        Selection.Copy
        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets("динамика КВА").Select
        Cells(4, mmm - 1).Select
        Selection.PasteSpecial Paste:=xlPasteValues

        Windows(KVA).Activate
        Cells(484, mmm).Select
        Selection.Copy
        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets("динамика КВА").Select
        Cells(5, mmm - 1).Select
        Selection.PasteSpecial Paste:=xlPasteValues

        Windows(KVA).Activate
        ActiveWindow.Close SaveChanges:=True






        ' copy 20

        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\20.xls"
        Sheets("Приложение 20 - 2").Select
        Range("B7:D175").Select
        Selection.Copy
        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets("20 ").Select
        Range("B6").Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Range("B159:D173").Select
        Selection.ClearContents

        Windows("20.xls").Activate
        ActiveWindow.Close SaveChanges:=False



        ' copy 2

        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\2.xlsx"
        Sheets("Приложение 2(10)").Select
        Range("C11:E138").Select
        Selection.Copy
        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets("2").Select
        Range("D5").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False


        Windows("2.xlsx").Activate
        Range("C143:E172").Select
        Selection.Copy
        Windows("prudnorm_" & xx & ".xlsm").Activate
        Range("D137").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False


        Windows("2.xlsx").Activate
        ActiveWindow.Close SaveChanges:=False


        ' copy 3

        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\3.xlsx"
        Sheets("Приложение 3").Select
        Range("C13:C117").Select
        Selection.Copy
        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets("3").Select
        Range("C3").Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Windows("3.xlsx").Activate
        ActiveWindow.Close SaveChanges:=False


        ' copy СК

        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\СК.xlsx"
        Sheets("Новая форма (2)").Select
        Range("C9:C53").Select
        Selection.Copy
        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets("СК").Select
        Range("C9").Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Windows("СК.xlsx").Activate
        ActiveWindow.Close SaveChanges:=False


        ' 22
        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\22.xls"
        Sheets("Приложение 22_9").Select
        Range("C12:C42").Select
        Selection.Copy
        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets(" 22").Select
        Range("C3").Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Windows("22.xls").Activate
        ActiveWindow.Close SaveChanges:=False
        '

        ' ПФИ Copy
        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\пфи.xlsx"
        Sheets("Отчет (лист 1)").Select
        Range("A:N").Select
        Selection.Copy
        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets("ОВП ПФИ").Select
        Range("A1").Select
        ActiveSheet.Paste

        'Columns("A:N").Replace What:=".", Replacement:=","
        ПФИ = Sheets("ОВП ПФИ").Range("B:B").Find("ПФИ").Row


            Range("G" & ПФИ - 1).FormulaR1C1 = "=RC[-4]"
            Range("H" & ПФИ - 1).FormulaR1C1 = "=RC[-4]"


            Range("G" & ПФИ).FormulaR1C1 = "=(RC[-4]+0)/1000"
            Range("H" & ПФИ).FormulaR1C1 = "=(RC[-4]+0)/1000"


            Range("I" & ПФИ).FormulaR1C1 = "=(RC[-2]-RC[-1])"
            Range("J" & ПФИ).FormulaR1C1 = "=RC[-1]/Пруденциальные!R12C6"

            Range("J" & ПФИ).NumberFormat = "0.00%"

            Range("G" & ПФИ & ":I" & ПФИ).NumberFormat = "#,##0"


        Windows("пфи.xlsx").Activate
        ActiveWindow.Close SaveChanges:=False
        '

        Sheets("20 ").Select
            Range("H174:J174").Select
            Selection.Copy
            Range("B174").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False






        ' LCR ВКЛА

        cb = "Портфель ЦБ за " & xx & ".xlsx"
        repo = "Портфель РЕПО за " & xx & ".xlsx"
        mbd = "Контроль исполнения лимитов " & xx & ".xlsm"
        vkla = "ВКЛА " & xx & ".xlsx"
        ottoki = "LCR оттоки " & xx & ".xlsx"

        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\баланс2.xls"
        Sheets("Отчет_новая форма").Select
        Range("A:C").Select
        Selection.Copy

        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\" & vkla
        Sheets("700").Select
        Range("A1").Select
        ActiveSheet.Paste

        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\" & ottoki
        Sheets("Оттоки ЮЛ ФЛ").Select
        Range("A1").Select
        ActiveSheet.Paste
        Windows("баланс2.xls").Activate
        ActiveWindow.Close SaveChanges:=False


        Windows(ottoki).Activate
        Range("I82:I91").Select
        Selection.Copy

        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets("LCR").Select
        Range("D46").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False

        Windows(ottoki).Activate
        ActiveWindow.Close SaveChanges:=True




        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\kurs.xls"
        Sheets("Лист1").Select
        Range("A:J").Select
        Selection.Copy
        Windows(vkla).Activate
        Sheets("курсы").Select
        Range("A1").Select
        ActiveSheet.Paste
        Windows("kurs.xls").Activate
        ActiveWindow.Close SaveChanges:=False


        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\" & cb
        Sheets("Отчет (лист 1)").Select
        Columns("A:AR").Select
        Selection.Copy
        Windows(vkla).Activate
        Sheets("ЦБ 1").Select
        Range("J1").Select
        ActiveSheet.Paste
        Windows(cb).Activate
        ActiveWindow.Close SaveChanges:=False


        Workbooks.Open Filename:="W:\CO\ДЕПАРТАМЕНТ ФИНАНСОВЫХ РИСКОВ И ПОРТФЕЛЬНОГО АНАЛИЗА\Управление финансовых рисков\PRUDIKI\" & Right(reportDate, 6) & "\" & reportDate & "\" & mbd
        Windows(mbd).Activate
        Sheets("ММ").Select
        Columns("A:O").Select
        Selection.Copy
        Windows(vkla).Activate
        Sheets("МБД 2").Select
        Range("J1").Select
        ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False

        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets("1254").Select
        Range("A1").Select
        ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False




        Windows(mbd).Activate
        ActiveWindow.Close SaveChanges:=False


        Windows(vkla).Activate
        Sheets("ВКЛА").Select
        Range("C3").Select
        Selection.Copy



        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets("Пруденциальные").Select
        Range("K115").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False



        Windows(vkla).Activate
        Sheets("ВКЛА").Select
        Range("K2").Select
        Selection.Copy

        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets("NSFR").Select
        Range("D23").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False


        Windows(vkla).Activate
        Sheets("ВКЛА").Select
        Range("K3").Select
        Selection.Copy

        Windows("prudnorm_" & xx & ".xlsm").Activate
        Sheets("NSFR").Select
        Range("D26").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False



        Windows(vkla).Activate
        ActiveWindow.Close SaveChanges:=True

        'msg = MsgBox("Запустить копирование приложения 2", vbYesNo)
        'If msg = vbYes Then
        'Call двойка
        'Else
        'x = 1 + 1
        'End If


        Application.DisplayAlerts = False
        Application.ScreenUpdating = True



        End Sub
        '''
        
        working_directory = os.getcwd().__str__().replace('\\', '/')
        file_path = f"{working_directory}/excel"        
        self.out_excel_name = f'{file_path}/{self.in_excel_name}' + '.xlsx' 
        self.out_excel_name = self.out_excel_name.replace('/', '\\')
                
        
        log.info('Полный путь до Excel: %s.', self.out_excel_name)
        
        if os.path.isfile(self.out_excel_name):    
            log.info("Удалил существующий Excel.")
            os.remove(self.out_excel_name)
        
        excel = win32.gencache.EnsureDispatch('Excel.Application') 
        workbook_original = excel.Workbooks.OpenXML(Filename = report_excel_path)
        xlmodule = workbook_original.VBProject.VBComponents.Add(1)
        xlmodule.CodeModule.AddFromString(prudiki_macro.format(path = self.out_excel_name))
        excel.Application.Run('SaveAs') 
        excel.Application.Quit()
        
        log.info("Ввел название файла.")
                
        
        return ExecutionResult.next()
    
    
class RunSravnMacro(StepBody):
    '''
    Формирование итогово файла
    
    '''
    
    def run(self, context: StepExecutionContext) -> ExecutionResult:
        import os        
        import win32com.client as win32
        import comtypes, comtypes.client  
        from pywinauto import Application
        from uiarecorder.play import find_element_by_uia  
        from actions import rpa_log as log
        
        sravn_macro = '''
        Sub СравнБаланс()


        With Sheets("Баланс_форма")

        .Rows("14:14").AutoFilter

        first = .Range("B:B").Find("1 КЛАСС - АКТИВЫ").Row
        last = .Range("B:B").Find("Итого АКТИВ").Row
        last_H = .Range("H:H").Find("Итого ПАССИВ").Row

        For i = first + 1 To last
            If .Cells(i, 1).HorizontalAlignment = xlCenter Then
                .Cells(i, 1).Interior.Color = 65535
            End If
        Next i

        For i = first + 1 To last_H
            If .Cells(i, 7).HorizontalAlignment = xlCenter Then
                .Cells(i, 7).Interior.Color = 65535
            End If
        Next i

        .Cells(last, 1).Interior.Color = 65535
        .Cells(last_H, 7).Interior.Color = 65535


        For i = first + 1 To last - 1

            If .Cells(i, 1).Interior.Color = 65535 Then
                If .Cells(i, 5) > 0 Then
                   Positive = Positive + .Cells(i, 5)
                Else
                   Negative = Negative + .Cells(i, 5)
                End If
            End If

        Next i

        For i = first + 1 To last - 1

            If .Cells(i, 7).Interior.Color = 65535 Then
                If .Cells(i, 11) > 0 Then
                   Positive1 = Positive1 + .Cells(i, 11)
                Else
                   Negative1 = Negative1 + .Cells(i, 11)
                End If
            End If

        Next i


        .Range("E10").FormulaR1C1 = "=IF(R[1]C=R[" & last - 10 & "]C,""ОК"",""Расхождение"")"

        .Range("K10").FormulaR1C1 = "=IF(R[1]C=R[" & last_H - 10 & "]C,""ОК"",""Расхождение"")"




        .Cells(10, 3) = "+"
        .Cells(10, 4) = "-"

        .Cells(11, 3) = Positive
        .Cells(11, 4) = Negative


        .Cells(10, 9) = "+"
        .Cells(10, 10) = "-"

        .Cells(11, 9) = Positive1
        .Cells(11, 10) = Negative1

        .Cells(11, 5) = "=RC[-2]+RC[-1]"
        .Cells(11, 11) = "=RC[-2]+RC[-1]"

            Rows("10:11").Select
            With Selection.Font
                .Size = 13
                .Underline = xlUnderlineStyleNone
                .Color = -16776961
                .TintAndShade = 0
                .ThemeFont = xlThemeFontNone
            End With

            Selection.NumberFormat = "#,##0"
            Selection.Font.Bold = True

            Range("C10:E11").Borders.LineStyle = True
            Range("I10:K11").Borders.LineStyle = True

            Rows("10:11").Select
            Rows("10:11").EntireRow.AutoFit



        End With

            Cells.Select
            With Selection
                .HorizontalAlignment = xlGeneral
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            With Selection
                .HorizontalAlignment = xlGeneral
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.UnMerge


            Columns("E:E").Select
            Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                Formula1:="=0"
            Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            With Selection.FormatConditions(1).Font
                .Color = -16752384
                .TintAndShade = 0
            End With
            With Selection.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 13561798
                .TintAndShade = 0
            End With
            Selection.FormatConditions(1).StopIfTrue = False


            Columns("e:e").Select
            Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                Formula1:="=0"
            Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            With Selection.FormatConditions(1).Font
                .Color = -16383844
                .TintAndShade = 0
            End With
            With Selection.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 13551615
                .TintAndShade = 0
            End With
            Selection.FormatConditions(1).StopIfTrue = False
            Range("B39").Select





            Columns("k:k").Select
            Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                Formula1:="=0"
            Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            With Selection.FormatConditions(1).Font
                .Color = -16752384
                .TintAndShade = 0
            End With
            With Selection.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 13561798
                .TintAndShade = 0
            End With
            Selection.FormatConditions(1).StopIfTrue = False


            Columns("k:k").Select
            Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                Formula1:="=0"
            Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            With Selection.FormatConditions(1).Font
                .Color = -16383844
                .TintAndShade = 0
            End With
            With Selection.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 13551615
                .TintAndShade = 0
            End With
            Selection.FormatConditions(1).StopIfTrue = False
            Range("B39").Select

        End Sub
        '''
        
        working_directory = os.getcwd().__str__().replace('\\', '/')
        file_path = f"{working_directory}/excel"        
        self.out_excel_name = f'{file_path}/{self.in_excel_name}' + '.xlsx' 
        self.out_excel_name = self.out_excel_name.replace('/', '\\')
                
        
        log.info('Полный путь до Excel: %s.', self.out_excel_name)
        
        if os.path.isfile(self.out_excel_name):    
            log.info("Удалил существующий Excel.")
            os.remove(self.out_excel_name)
        
        excel = win32.gencache.EnsureDispatch('Excel.Application') 
        workbook_original = excel.Workbooks.OpenXML(Filename = report_excel_path)
        xlmodule = workbook_original.VBProject.VBComponents.Add(1)
        xlmodule.CodeModule.AddFromString(RunSravnMacro.format(path = self.out_excel_name))
        excel.Application.Run('SaveAs') 
        excel.Application.Quit()
                
        
        return ExecutionResult.next()
    
class ParseAndSend(StepBody):
    '''
    Парсинг сравнительного файла и отправка финального письма
    
    '''
    
    def run(self, context: StepExecutionContext) -> ExecutionResult:
        import pandas as pd
        import numpy as np
        import xlrd
        import requests
        from os import listdir
        from os.path import isfile, join
        
        file_path = r'C:\Users\RPA_024\Desktop\письмо'
        
        onlyfiles = [f for f in listdir(file_path) if isfile(join(file_path, f))]
        my_file = os.path.join(file_path, onlyfiles[0])
        
        data = pd.read_excel(my_file, skiprows=14)

        code_dict_actives = [1000, 1010, 1050, 1100, 1111, 1150, 1200, 1250, 1300, 1350, 1400, 1450, 1460, 1470, 1480, 1490, 1550, 1600, 1610, 1650, 1700, 1790, 1810, 1830, 1850, 1880, 1890]
        code_dict_obyaz = [2010, 2020, 2030, 2040, 2050, 2110, 2120, 2150, 2200, 2255, 2300, 2400, 2550, 2700, 2770, 2790, 2810, 2830, 2850, 2880, 2890, 3000, 3100, 3200, 3400, 3500,]

        workbook = xlrd.open_workbook(my_file)

        sheet_names = workbook.sheet_names()

        sheet = workbook.sheet_by_name(sheet_names[2])
        balance = round(sheet.cell(284, 4).value / 1000000, 1)
        difference = round(sheet.cell(284, 3).value / 1000000)
        trend = 'увеличилась' if balance > 0 else 'уменьшилась'

        def get_cell_range(start_col, start_row, end_col, end_row):
            return [sheet.row_slice(row, start_colx=start_col, end_colx=end_col+1) for row in range(start_row, end_row+1)]

        actives_result = list()

        data_actives = get_cell_range(0, 14, 4, 270)

        for row in data_actives:
            if int(row[0].value) in code_dict_actives:
                actives_result.append(row)

        actives_result = [[c.value for c in actives_result[n]] for n in range(0, len(actives_result))]

        data_obyaz = get_cell_range(6, 14, 10, 280)
        obyaz_result = []

        for row in data_obyaz:
            if row[0].value and int(row[0].value) in code_dict_obyaz:
                obyaz_result.append(row)

        actives_result = [[c.value for c in obyaz_result[n]] for n in range(0, len(obyaz_result))]

        df = pd.DataFrame(np.row_stack(actives_result))

        df[0] = pd.to_numeric(df[0])
        df[2] = pd.to_numeric(df[2])
        df[3] = pd.to_numeric(df[3])
        df[4] = pd.to_numeric(df[4])

        df = df.astype({0: int, 2: int, 3: int, 4: int})

        df = df.drop_duplicates()

        df_positive = df[(df[4] > 0) & (df[4] > 1000000)]
        df_positive_other = df[(df[4] > 0) & (df[4] < 1000000)]
        positive_sum = df_positive_other[4].sum()

        df_negative = df[(df[4] < 0) & (df[4] < -1000000)]
        df_negative_other = df[(df[4] < 0) & (df[4] > -1000000)]
        negative_sum = df_negative_other[4].sum()
        negative_sum = abs(negative_sum)

        df_positive[4] = (df_positive[4] / 1000000).round(1)
        df_negative[4] = (df_negative[4] / 1000000).round(1)

        positive_sum = positive_sum / 1000000
        positive_sum_obyaz = round(positive_sum, 1)

        negative_sum = negative_sum / 1000000
        negative_sum_obyaz = round(negative_sum, 1)

        obyaz_rise = round(df_positive[4].sum() + positive_sum, 1)
        obyaz_fall = round(abs(df_negative[4].sum()) + negative_sum, 1)

        obyaz_rise_message = ''
        obyaz_fall_message = ''

        for index, row in df_positive.iterrows():
            obyaz_rise_message += f'{row[1]} - на {row[4]} млрд. тг.<br>'

        for index, row in df_negative.iterrows():
            obyaz_fall_message += f'{row[1]} - на {abs(row[4])} млрд. тг.<br>'

        import xlrd

        workbook = xlrd.open_workbook(my_file)

        sheet_names = workbook.sheet_names()

        sheet = workbook.sheet_by_name(sheet_names[2])


        def get_cell_range(start_col, start_row, end_col, end_row):
            return [sheet.row_slice(row, start_colx=start_col, end_colx=end_col+1) for row in range(start_row, end_row+1)]

        actives_result = list()

        data_actives = get_cell_range(0, 14, 4, 270)

        for row in data_actives:
            if int(row[0].value) in code_dict_actives:
                actives_result.append(row)
        actives_result = [[c.value for c in actives_result[n]] for n in range(0, len(actives_result))]

        df = pd.DataFrame(np.row_stack(actives_result))

        df[0] = pd.to_numeric(df[0])
        df[2] = pd.to_numeric(df[2])
        df[3] = pd.to_numeric(df[3])
        df[4] = pd.to_numeric(df[4])

        df = df.astype({0: int, 2: int, 3: int, 4: int})

        df = df.drop_duplicates()

        df_positive = df[(df[4] > 0) & (df[4] > 1000000)]
        df_positive_other = df[(df[4] > 0) & (df[4] < 1000000)]
        positive_sum = df_positive_other[4].sum()

        df_negative = df[(df[4] < 0) & (df[4] < -1000000)]
        df_negative_other = df[(df[4] < 0) & (df[4] > -1000000)]
        negative_sum = df_negative_other[4].sum()
        negative_sum = abs(negative_sum)

        df_positive[4] = (df_positive[4] / 1000000).round(1)
        df_negative[4] = (df_negative[4] / 1000000).round(1)

        positive_sum = positive_sum / 1000000
        positive_sum = round(positive_sum, 1)

        negative_sum = negative_sum / 1000000
        negative_sum = round(negative_sum, 1)

        active_rise = round(df_positive[4].sum() + positive_sum, 1)
        active_fall = round((abs(df_negative[4].sum()) + negative_sum), 1)

        actives_rise_message = ''
        actives_fall_message = ''

        for index, row in df_positive.iterrows():
            actives_rise_message += f'{row[1]} - на {row[4]} млрд. тг.<br>'

        for index, row in df_negative.iterrows():
            actives_fall_message += f'{row[1]} - на {abs(row[4])} млрд. тг.<br>'

        my_message = f"""Валюта баланса {trend} на {difference} млрд. тг. и составила {balance} млрд. тг.<br><br>
        По активам:<br>
        Увеличение - на {active_rise}<br>
        {actives_rise_message}
        Другие активы (изменения менее 1,0 млрд. тг.) – на {positive_sum} млрд. тг.<br><br>
        Уменьшение - на {active_fall}<br>
        {actives_fall_message}
        Другие активы (изменения менее 1,0 млрд. тг.) – на {negative_sum} млрд. тг.<br><br><br>
        По обязательствам:<br>
        Увеличение - на {obyaz_rise}<br>
        {obyaz_rise_message}
        Другие активы (изменения менее 1,0 млрд. тг.) – на {positive_sum_obyaz} млрд. тг.<br><br>
        Уменьшение - на {obyaz_fall}<br>
        {obyaz_fall_message}
        Другие активы (изменения менее 1,0 млрд. тг.) – на {negative_sum_obyaz} млрд. тг.<br><br><br>
        """

        mail_data = {
            'sender': 'RPA_024@halykbank.nb',
            'subject': 'final',
            'body': my_message,
            'receiver': ['KAMILAK2@halykbank.kz', 'AzamatTi3@halykbank.kz', 'AdiletAs@halykbank.kz']
        }

        endpoint = "http://alav741/api/send/"

        requests.post(endpoint, data=mail_data)
        os.remove(my_file)
        
        
        return ExecutionResult.next()
    
    
class FinishFlow(StepBody):
    '''
    Признак конца CollectReports Workflow
    
    '''
    def run(self, context: StepExecutionContext) -> ExecutionResult:
        log.info("Finished")
        return ExecutionResult.next()
    
