import xml.etree.ElementTree as ET
import zipfile
import shutil
import os
import tkinter
from tkinter import filedialog
from tkinter import constants
from pptx import Presentation

class Application:
    save_path = ''
    ppt_path = ''

    def __init__(self):
        pass

    def set_ppt_path(self):
        pass

    def set_save_path(self):
        pass

class Powerpoint:
    embed = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'
    namespaces = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                  'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                  'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}

    def __init__(self, ppt_path):
        self.powerpoint = zipfile.ZipFile(ppt_path, 'r')
        self.load_files()
        #self.slide_names = []
        #self.slide_xmls = []
        #self.rel_names = []
        #self.rel_xmls = []
        self.slides = []
        self.load_xml()

    def load_files(self):
        """Read filenames from zip file"""
        self.files = self.powerpoint.namelist()

# FIXME Instead of storing in an array, create slides and store array of slides
    def load_xml(self):
        """Load slide and relationship XML from zip file"""

        slide_names = []
        rel_names = []

        for f in self.files:
            if f.startswith('ppt/slides') and f.endswith('.xml'):
                slide_names.append(f)

            if f.startswith('ppt/slides/_rels') and f.endswith('.rels'):
                rel_names.append(f)

        slide_names.sort()
        rel_names.sort()

        for i in range(len(slide_names)):
            self.slides.append(self.make_slide(slide_names[i],rel_names[i]))


    def make_slide(self, slide_name, rel_name):
        """Create a slide ready for conversion"""
        slide_xml = (ET.fromstring(self.powerpoint.read(slide_name)))
        rel_xml = (ET.fromstring(self.powerpoint.read(rel_name)))
        slide = Slide(slide_name, slide_xml, rel_xml)
        return slide



class Converter:

    def __init__(self, slide, page):
        self.slide = slide
        self.page = page

class Slide:

    def __init__(self, name, xml, rel_xml):
        self.name = name
        self.xml = xml
        self.rel_xml = rel_xml

        self.template = ''

        self.title = ''
        self.content = ''
        self.image_src = ''

        self.dir_name = ''
        self.dir_path = ''
        self.media_path = ''
        self.file_name = ''
        self.newfile = ''

    def extract_title(self, title_name):
        title_block = self.slide.find('.//*[@name="%s"].....' % title_name)
        title = ''
        title_text = title_block.findall('.//a:t', namespaces)
        for item in title_text:
            title += item.text
        return title

    def extract_content(self):
        pass

    def extract_image_id(self):
        pass

    def get_image_path(self):
        pass

class Page:

    def __init__(self):
        self.html = ''
        self.template = ''

    def save_to_disk(self):
        pass

# Move this into the converter?
class HTMLFormatter:
    def make_tags(self):
        pass

    def format_content(self):
        pass

