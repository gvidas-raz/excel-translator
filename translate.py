import openpyxl
import os
import sys
import re
import requests

from dotenv import load_dotenv
from absl import app
from absl import flags

load_dotenv()

API_KEY = os.getenv("API_KEY")

flags.DEFINE_string('file', None, 'Excel file name to translate', short_name='f')
flags.DEFINE_string('source', None, 'The first cell in the column to translate', short_name='s')
flags.DEFINE_string('dest', None, 'The first cell in the column to write the translation to', short_name='d')

flags.mark_flags_as_required('file', 'source', 'dest')

FLAGS = flags.FLAGS

DEEPL_URL = 'https://api-free.deepl.com/v2/translate'

def main(argv):
    source = re.split(r"([a-z]+)([0-9]+)", FLAGS.source)
    dest = re.split(r"([a-z]+)([0-9]+)", FLAGS.dest)

    workbook = openpyxl.load_workbook(filename=FLAGS.file)
    sheet = workbook.active

    while sheet[source[0]+source[1]].value is not None:
        targetText = sheet[source[0]+source[1]].value
        query = {
            'text': targetText,
            'source_lang': 'ES',
            'target_lang': 'EN'
        }
        headers = {'Authorization': 'DeepL-Auth-Key {API_KEY}'}
        try:
            response = requests.post(DEEPL_URL, query=query, headers=headers)
            response.raise_for_status()
        except requests.exceptions.HTTPError as error:
            print(error)
        
        try:
            content = response.json()
        except ValueError as error:
            print(error)

        translation = content['translation'][0]['text']

        sheet[dest[0]+dest[1]] = translation
        
        source[1] += 1
        dest[1] += 1
    
    workbook.save()
    workbook.close()
        
if __name__ == "__main__":
    app.run(main)
