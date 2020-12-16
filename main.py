
import os
import sys
import argparse
import math

from queue import Queue
from os import listdir
from os.path import isfile, join
from threading import Thread
from converter import Converter

from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

class MainWidget(QMainWindow):

    width = 720
    height = 480

    def __init__(self):
      QMainWindow.__init__(self)

      self.size = QSize(self.width, self.height)

      self.setWindowTitle("Docx to HTML converter")
      self.setMinimumSize(self.size)
      self.setMaximumSize(self.size)
      self.setAcceptDrops(True)

      self.label = QLabel("Drag and Drop .docx file or folder", self)
      self.label.resize(self.width, 50)
      self.label.setAlignment(Qt.AlignCenter)
      self.label.setStyleSheet("""
        QLabel {
          font-size: 15pt;
        }
      """)

      self.label.move(0, math.floor(self.height/2))

    def dragEnterEvent(self, event):
      if event.mimeData().hasUrls():
        event.accept()
      else:
        event.ignore()

    def dropEvent(self, event):
      files = [u.toLocalFile() for u in event.mimeData().urls()]
      for f in files:
        runCLI(dest=f)

def run(q: Queue):
  while not q.empty():
    file = q.get()
    converter = Converter()
    converter.processFile(file)

    q.task_done()

def runCLI(dest=None):
  q = Queue()

  if os.path.isdir(dest):
    for file in [f for f in listdir(dest) if isfile(join(dest, f)) and join(dest, f).endswith('.docx')]:
      q.put(join(dest, file))
  elif isfile(dest) and dest.endswith('.docx'):
    q.put(dest)

  if q.empty():
    print('No args')
    return

  for i in range(5):
    thread = Thread(target=run, args=(q,))
    thread.start()

  q.join()

def runUI():
  app = QApplication(sys.argv)
  ui = MainWidget()
  ui.show()
  sys.exit(app.exec_())


if __name__ == '__main__':
  parser = argparse.ArgumentParser(description='Convert script')

  parser.add_argument('-d', action="store", dest="dest")
  parser.add_argument('-u', action="store", dest="ui", default=True, type=bool)


  args = parser.parse_args()

  if not args.ui or (args.dest):
    runCLI(dest=args.dest)
  else:
    runUI()
