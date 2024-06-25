import os
import re
import random
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook, Workbook
from datetime import datetime
from relance_rh.mail_sender import MailSender


class ExcelOperations:
    def process_excel_files(self, folder_path):
        information = []
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.endswith(".xlsx") or file.endswith(".xls"):
                    regex = self.extract_path_profile(root)
                    file_path = os.path.join(root, file)
                    # Verify the format of the file with the title
                    check_format = self.verify_format_file_excel(file_path)
                    if not check_format:
                        continue
                    # Extract information from the file
                    info_from_exel = self.extract_information_from_excel(file_path, regex)
                    # Verify if the information is not empty
                    if info_from_exel is None:
                        continue
                    # Verify if the information is not already in the list
                    if any(info_from_exel['last_name'] == item['last_name'] and info_from_exel['first_name'] == item['first_name'] for item
                           in information):
                        continue
                        # Add date and manager to the information
                    interviews = []
                    for date, manager in zip(info_from_exel['dates'], info_from_exel['managers']):
                        interviews.append({'date': date, 'manager': manager})
                    info_from_exel['interviews'] = interviews
                    information.append(info_from_exel)

        if not information:
            return None

        return information

    def extract_information_from_excel(self, file_path, regex):
        wb = load_workbook(file_path)
        sheet = wb.active
        tel_num = sheet['C5'].value

        tel_num = re.sub(r'\D', '', tel_num)
        tel_num_sec = sheet['D5'].value
        email = sheet['E5'].value
        # verify if the email is not empty or not have the word "Mail: "
        if email is not None:
            email = email.strip().replace("Mail: ", "")
        data_contact = [tel_num, email]

        verification_result = self.verify_value(data_contact)
        # if the first verification is not correct, we verify the second contact
        if not verification_result:
            data_contact = [tel_num_sec, email]
            verification_result = self.verify_value(data_contact)
            if verification_result:
                tel_num = tel_num_sec

        # format num
        tel_num = re.sub(r'\D', '', tel_num)
        tel_num = ' '.join([tel_num[i:i + 2] for i in range(0, len(tel_num), 2)])

        profile = regex[0] if len(regex) > 0 else ''
        direction = regex[1] if len(regex) > 1 else ''
        last_name = sheet['B6'].value
        first_name = sheet['B9'].value
        status = sheet['G22'].value
        # if the status is empty, we verify the second status
        if status is None:
            status = sheet['G23'].value

        dates = []
        managers = []

        for row in range(7, 10):
            date_value = sheet.cell(row=row, column=8).value
            date_value = self.verify_format_date(date_value)
            if date_value is not None:
                date_value = self.format_date(date_value)
                date_value = self.change_date_randomly(date_value)
            dates.append(date_value)

        for row in range(7, 10):
            manager_value = sheet.cell(row=row, column=7).value
            managers.append(manager_value if manager_value is not None else "")

        information_perso = {
            'profile': profile,
            'direction': direction,
            'last_name': last_name,
            'first_name': first_name,
            'tel_num': tel_num,
            'email': email,
            'status': status,
            'dates': dates,
            'managers': managers
        }

        return information_perso

    def create_new_excel_file(self, information, output_path):
        wb = Workbook()
        sheet = wb.active
        sheet.title = "Relance RH"

        headers = ["repartoire", "profil", "NOM", "Prenom", "tel", "Email", "Disponibilite"]
        for i in range(3):
            headers += [f"Date{i + 1}", f"Entretien{i + 1}"]
        headers = [header.upper() for header in headers]
        sheet.append(headers)


        for info in information:
            row = [
                info['direction'], info['profile'],
                info['last_name'], info['first_name'],
                info['tel_num'], info['email'],
                info['status'],
            ]

            for interview in info['interviews']:
                row += [interview['date'], interview['manager']]
            sheet.append(row)

        # column widths
        for column in sheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column[0].column_letter].width = adjusted_width
        wb.save(output_path)
        sender = self.select_users_for_resend_email(information)
        if sender:
            print("Emails have been sent successfully")
        else:
            print("Emails have not been sent")

        return True


    def select_users_for_resend_email(self, information):
        resend_users_3_months = {}
        resend_users_6_months = {}
        send_email = MailSender()

        current_date = datetime.now()

        for info in information:

            for interview in info['interviews']:
                if interview['date'] is not None:
                    interview_date = datetime.strptime(interview['date'], '%d/%m/%y')
                    diff = current_date - interview_date
                    if diff.days > 90 and diff.days < 180:
                        resend_users_3_months[info['email']] = info['first_name']
                        email = send_email.send_email_after_3_moths(info)


                    elif diff.days > 180:
                        resend_users_6_months[info['email']] = info['first_name']
                        email = send_email.send_email_after_3_moths(info)

        return True

    def extract_path_profile(self, path):
        infos = []
        techno = re.search(r'- (.+)', path)
        path_file_of_user = re.search(r'(.+)\\', path)
        if techno:
            infos.append(techno.group(1))
        if path_file_of_user:
            infos.append(path_file_of_user.group(1))

        return infos

    def verify_value(self, data):
        tel_num, email = data
        if tel_num is not None and email is not None:
            tel_num = re.sub(r'\D', '', tel_num)
            return self.verify_phone_number(tel_num) and self.verify_email(email)
        else:
            return False

    @staticmethod
    def verify_phone_number(phone_number):
        return bool(re.fullmatch(r"0\d{9}", phone_number))

    @staticmethod
    def verify_email(email):
        return bool(re.fullmatch(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", email))

    def verify_format_file_excel(self, file_path):
        wb = load_workbook(file_path)
        sheet = wb.active
        title = "DOSSIER RECRUTEMENT FICHE ENTRETIEN"
        if sheet['C3'].value != title:
            return False
        return True

    def verify_format_date(self, date):
        if isinstance(date, datetime):
            return date.strftime('%d/%m/%Y')
        if isinstance(date, str) and re.match(r'\d{1,2}/\d{1,2}/\d{2,4}', date):
            return date
        if date == '___/___/__':
            return None
        return None


    def format_date(self, date_string):
        # 'dd/mm/yyyy' -> 'dd/mm/yy'
        try:
            date_object = datetime.strptime(date_string, '%d/%m/%Y')
        except ValueError:
            # 'dd/mm/yy' -> 'dd/mm/yy'
            date_object = datetime.strptime(date_string, '%d/%m/%y')

        # 'dd/mm/yy' -> 'dd/mm/yy'
        formatted_date = date_object.strftime('%d/%m/%y')


        return formatted_date



    def change_date_randomly(self, date_value):
        date = datetime.strptime(date_value, "%d/%m/%y")
        months_options = [3, 6]
        months = random.choice(months_options)
        today = datetime.today()
        new_date = today - relativedelta(months=months)

        new_date = new_date.replace(year=2024)
        if new_date.month == 12:
            new_date = new_date.replace(year=2023)


        date = new_date.strftime("%d/%m/%y")


        return date



