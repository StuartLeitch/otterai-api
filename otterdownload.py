import json
import requests
import datetime
from otterai import OtterAI
import os
from docx import Document
from dotenv import load_dotenv


load_dotenv()
windows_path = os.getenv('WINDOWS_EXPORT_PATHNAME')

def docx_to_text(docx_path, txt_path):
    doc = Document(docx_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(full_text))


def get_latest_load(folder_path):
    # List all files in the folder
    files = os.listdir(folder_path)
    
    # Filter out directories
    files = [f for f in files if os.path.isfile(os.path.join(folder_path, f))]
    
    # Sort files alphabetically
    files.sort(reverse=True)
    
    # Return the latest filename
    return datetime.datetime.strptime(files[0][:8], "%Y%m%d") if files else None

while True:

    latest_load_date = get_latest_load(windows_path)
    print(f'Top file: {latest_load_date}')

    if latest_load_date is None:
        # set to 5 years ago
        latest_load_date = datetime.datetime.now() - datetime.timedelta(days=365*5)

    otter = OtterAI()
    otter.login(os.getenv('OTTER_USER'), os.getenv('OTTER_PASS'))

    def calc_dt_offset(top_file):
        offset = top_file + datetime.timedelta(days=10)
        return str(int(offset.timestamp()))

    offset_date = calc_dt_offset(latest_load_date)
    speechesObject = otter.get_speeches(page_size=50, last_load_ts=offset_date if latest_load_date else None)

    speeches = speechesObject['data']['speeches']

    print(f'Speeches found: {len(speeches)}')

    # print(otter.get_user())

    files = os.listdir(windows_path)

    found_file_to_download = False
    skipped_downloads = False

    for speech in reversed(speeches):
        # print(speech)
        if speech['summary'] is not None:
            speech_id = speech['otid']

            formatted_date = datetime.datetime.fromtimestamp(speech['start_time']).strftime("%Y%m%d %H:%M:%S")
            save_name = formatted_date + " " + (speech['title'] if not speech['title']==None else speech_id).replace('/', '-')
    
            if save_name + '.txt' in files:
                print('File already exists, skipping ' + save_name + '.txt')
                skipped_downloads = True
                continue

            speechObject = otter.get_speech(speech_id=speech_id)
            # if empty meeting, skip
            if speechObject['data']['speech']['summary'] == '':
                print('Empty content, skipping ' + save_name + '.txt')
                skipped_downloads = True
                continue
            # print(speechObject)



            # print(otter.get_speakers())

            # print(otter.query_speech(speech_id=speech_id, query='What was said in this meeting?'))
            print('Downloading ' + save_name)
            found_file_to_download = True
            downloadedspeech = otter.download_speech(speech_id=speech_id, name=save_name, pathname=windows_path, fileformat='docx')

            docx_to_text(windows_path + save_name + '.docx', windows_path + save_name + '.txt')
            # Delete the docx file
            os.remove(windows_path + save_name + '.docx')


            # with open('speech.txt', 'w') as file:
            #     file.write(downloadedspeech)

    if not skipped_downloads:
        print('WARNING. No files skipped. INVESTIGATE!!! Exiting....')
        break

    if not found_file_to_download:
        print('No new files to download. Exiting....')
        break

