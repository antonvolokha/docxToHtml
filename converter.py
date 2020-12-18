import zipfile
import math
import os
import shutil
import mammoth
import argparse
from tempfile import gettempdir

from PIL import Image
from os import listdir
from os.path import isfile, join
from queue import Queue
from threading import Thread
from uuid import uuid4

class Converter:

  imageQuality = 75

  def __init__(self):
    self.prefix = str(uuid4())
    self.tmpPath = gettempdir() + '/tmp' + str(uuid4())
    self.imgPath = '%s/word/media' % (self.tmpPath)

  def run(self, filepath: str):
    # Not workign on windows
    compress = os.name != 'nt'
    if compress:
      filepath = self.compressFile(filepath)

    self.convertToHtml(filepath, remove=compress)



  def zipdir(self, path, ziph):
    for root, dirs, files in os.walk(path):
      for file in files:
        src = os.path.join(root, file)
        dst = src.replace(self.tmpPath + '/', '')

        ziph.write(src, dst)

  def compressTask(self, q: Queue):
    while not q.empty():
      path = q.get()
      img = Image.open(path)
      x, y = img.size
      if x > 2000:
        x2, y2 = math.floor(x/2), math.floor(y/2)
        img = img.resize((x2,y2),Image.ANTIALIAS)

      img.save(path, quality=self.imageQuality)
      q.task_done()

  def compressThreds(self, q: Queue, threadCount=5):
    for i in range(threadCount):
      thread = Thread(target=self.compressTask, name=str(i), args=(q,))
      thread.start()

  def compressFile(self, filepath: str) -> str:
    try:
      with zipfile.ZipFile(filepath, 'r') as zip_ref:
        zip_ref.extractall(self.tmpPath)

      filepath = filepath.split('/')
      filename = self.prefix + filepath.pop()
      filepath.append(filename)

      filepath = '/'.join(filepath)

      if os.path.exists(self.imgPath):
        q = Queue()
        for image in [f for f in listdir(self.imgPath) if isfile(join(self.imgPath, f))]:
          q.put(join(self.imgPath, image))

        self.compressThreds(q)
        q.join()

      zipf = zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED)
      self.zipdir(self.tmpPath, zipf)
      zipf.close()

      shutil.rmtree(self.tmpPath)

      return filepath
    except Exception as ex:
      print('Exception %s' % (ex))
      return None

  def convertToHtml(self, filepath: str, remove=True):
    html = ''
    with open(filepath, "rb") as docx_file:
      result = mammoth.convert_to_html(docx_file)
      html = result.value

    oldfilepath = filepath
    filepath = filepath.split('/')
    filename = filepath.pop().replace('.docx', '.html').replace(self.prefix, '')
    filepath.append(filename)

    filepath = '/'.join(filepath)

    html = html.replace('<img src', '<img class="center" src')
    html = html.replace('<em>Рис', '<em class="center">Рис')
    html = html.replace('<p>Рис', '<p class="center">Рис')

    html = """
      <html><head><meta charset="utf-8"/><style type="text/css">
        @page { size: 21cm 29.7cm; margin: 2cm; }
        p {
          background: transparent;
          text-indent: 1.2cm;
          margin-bottom: 0cm;
          line-height: 100%;
          font-family: "Times New Roman", Times, serif;
          font-size: 14pt;
        }
        p img {
          display: block;
          margin-left: auto;
          margin-right: auto;
          width: 50%;
          text-align: center;
        }
        .center {
          display: block;
          margin-left: auto;
          margin-right: auto;
          width: 50%;
          text-align: center;
        }
        a:link {
          color: #000080;
          so-language: zxx;
          text-decoration: underline;
        }
        a:visited {
          color: #800000;
          so-language: zxx;
          text-decoration: underline;
        }
      </style></head><body>
    """  + html + "</body></html>"

    with open(filepath, 'w+', encoding='utf-8') as html_file:
      html_file.write(html)

    if remove:
      os.remove(oldfilepath)
