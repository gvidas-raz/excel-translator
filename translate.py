import openpyxl
import os
import re
import requests

from dotenv import load_dotenv
from absl import app
from absl import flags

load_dotenv()

API_KEY = os.getenv("API_KEY")

flags.DEFINE_string('file', None, 'Excel file name to translate', short_name='f', required=True)
flags.DEFINE_string('sheet', None, 'Excel sheet to translate', short_name='sh')
flags.DEFINE_string('source', None, 'The first cell in the column to translate', short_name='s', required=True)
flags.DEFINE_string('dest', None, 'The first cell in the column to write the translation to', short_name='d', required=True)
flags.DEFINE_boolean('overwrite', False, 'Enables the overwriting of destination cells. USE WITH CAUTION', short_name='o')

FLAGS = flags.FLAGS

DEEPL_URL = 'https://api-free.deepl.com/v2/translate'

progress_count = 0

def translate(text, source_lang, target_lang):
    query = {
        'text': text,
        'source_lang': source_lang,
        'target_lang': target_lang
    }
    headers = {'Authorization': 'DeepL-Auth-Key {key}'.format(key=API_KEY)}
    response = requests.post(DEEPL_URL, params=query, headers=headers)
    response.raise_for_status()

    content = response.json()

    return content['translations'][0]['text']

def progress():
    global progress_count
    progress_count += 1

def move_cells_column(source, dest):
    '''
    Moves the source and destination cells down by 1 in their respective columns.
    '''
    source = list(filter(None, re.split(r'(\d+)', source)))
    dest = list(filter(None, re.split(r'(\d+)', dest)))

    source[1] = str(int(source[1])+1)
    dest[1] = str(int(dest[1])+1)

    return source[0]+source[1], dest[0]+dest[1]

def main(argv):
    global progress_count
    source_cell = FLAGS.source
    dest_cell = FLAGS.dest

    workbook = openpyxl.load_workbook(filename=FLAGS.file)

    filename_bits = FLAGS.file.split('.')
    # make a backup of the excel file in case we mess something up
    workbook.save(filename_bits[0]+'_backup.'+filename_bits[1])
    sheet = workbook.active
    
    if FLAGS.sheet is not None:
        sheet = workbook[FLAGS.sheet]

    while sheet[source_cell].value is not None:
        if sheet[dest_cell].value is not None and FLAGS.overwrite == False:
            print(
'''
Translation destination cell {dest} is not empty. To overwrite destination cells use the flag --overwrite
Skipping translation of cell: {source}.
'''.format(dest=dest_cell, source=source_cell))
            source_cell, dest_cell = move_cells_column(source_cell, dest_cell)
            continue

        print('Translating cell: '+ source_cell +' into -> '+dest_cell)

        targetText = sheet[source_cell].value
        try:
            translation = translate(targetText, 'ES', 'EN')
        except requests.exceptions.HTTPError as error:
            print(error)
            break
        except ValueError as error:
            print(error)
            break

        sheet[dest_cell] = translation
        progress()
        
        if progress_count >= 100:
            workbook.save(FLAGS.file)
            progress_count = 0

        source_cell, dest_cell = move_cells_column(source_cell, dest_cell)
    
    workbook.save(FLAGS.file)
    workbook.close()
        
if __name__ == "__main__":
    app.run(main)
