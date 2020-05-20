from outlook import Outlook
import os

email = os.getenv('OUTLOOK_EMAIL')
password = os.getenv('OUTLOOK_PASSWORD')

my_email = Outlook(email, email, password)

powerpoint_file = 'assets_to_send/explanation.pptx'
human_tagging_files = 'assets_to_send/human_tagging_files_to_share/Human_Tagging_{}.xlsx'


for i in range(1,11):
	my_email.send_email('prateek.april@gmail.com', f'test_7_{i}', 'test test', [powerpoint_file, human_tagging_files.format(i)])
	print(i)
