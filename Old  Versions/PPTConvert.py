###CURRENT ISSUES####
# FIXME Something with Unicode? Apostrophes render as question diamonds
# TODO Very Fragile: Only handles 1 slide type when explicitily named - make more robust - Elegantly handle multiple slide types
# Slide names incorrect - FIXED
# Titles break on mispelled words - FIXED
# File locations hardcoded need to accept input - FIXED
# TODO Will not overwrite files


##NEW SPECS##
# 1) move each page to individual folder - FIXED
# 2) each folder should have its own media folder for its local images to keep self contained - FIXED
# 3) update javascript file to include pages - HAXED
##END SPECS##

####IMPROT DEPENDANT PACKAGES####
import xml.etree.ElementTree as ET
import zipfile
import shutil
import os
import tkinter
from tkinter import filedialog
from tkinter import constants

####OS FILE MANIPULATION/UI####
# open powerpoint
# ppt_path = "C:/Users/eric_/Desktop/GreenMockups/xmltest/SAFE/BigTest.pptx"
# prompt for save location
# save_path = "C:/Users/eric_/Desktop/GreenMockups/xmltest/SAVETEST/"
# template location, need to auto decide, hardcoded for now
template_path = 'C:/Users/eric_/Desktop/GreenMockups/eta-sample/eta-sample/pages/_template/index.html'

####BASIC GUI####
# make this fancier later
root = tkinter.Tk()

# cheap js hack for demo
pages = []


def set_ppt_path():
    opts = {}
    opts['parent'] = root
    opts['title'] = 'Select Powerpoint to Convert'
    opts['filetypes'] = [('Powerpoint Files', '.pptx')]
    # hard code this for now for convienence
    opts['initialdir'] = 'C:/Users/eric_/Desktop/GreenMockups/xmltest/SAFE/'
    global ppt_path
    ppt_path = filedialog.askopenfilename(**opts)


def set_save_path():
    opts = {}
    opts['parent'] = root
    opts['title'] = 'Select Save Location'
    # hard code this for now for convienence
    opts['initialdir'] = 'C:/Users/eric_/Desktop/GreenMockups/xmltest/SAVETEST/'
    global save_path
    save_path = filedialog.askdirectory(**opts)
    print(save_path)


def start_conversion():
    # make media path from save location if not present
    # removed this to comply with new standards
    # media_path = save_path + 'media/'
    # if not os.path.exists(media_path):
    #    os.makedirs(media_path)

    #### FILE INITIALIZATION####
    # convert powerpoint to zip ## powerpoints can open like zip files happy day
    # open powerpoint file
    powerpoint = zipfile.ZipFile(ppt_path, 'r')

    # get slides
    files = powerpoint.namelist()
    slide_names = []
    slides = []
    for f in files:
        if f.startswith('ppt/slides') and f.endswith('.xml'):
            slide_names.append(f)
    # sort names from random order import
    slide_names.sort()

    # read in xml from slides
    for slide in slide_names:
        slides.append(powerpoint.read(slide))

    # get relationships
    rel_names = []
    for f in files:
        if f.startswith('ppt/slides/_rels') and f.endswith('.rels'):
            rel_names.append(f)
    # sort lists so they corespond with their names
    rel_names.sort()

    # read in xml from rels
    rels = []
    for r in rel_names:
        rels.append(powerpoint.read(r))

    # build relationship map ## sorted lists should solve this problem???

    ####XML RETRIVAL####
    # build slide root ## opening as zip lets us get files as strings
    slide_xmls = []
    rel_xmls = []

    for s in slides:
        slide_xmls.append(ET.fromstring(s))

    for r in rels:
        rel_xmls.append(ET.fromstring(r))

    # build namespaces
    namespaces = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                  'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                  'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
    # embed shortcut
    embed = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'

    # helper functions

    def extract_title(xml):
        title_name = 'Title 1'
        title_block = xml.find('.//*[@name="%s"].....' % title_name)
        title = ''
        title_text = title_block.findall('.//a:t', namespaces)
        for item in title_text:
            title += item.text
        return title

    def extract_content(xml):
        body_name = 'Text Placeholder 2'
        body_block = xml.find('.//*[@name="%s"].....' % body_name)
        body_ps = body_block.findall('.//a:p', namespaces)
        body_text = []
        for item in body_ps:
            paragraph = ''
            body_text_elements = body_block.findall('.//a:t', namespaces)
            for element in body_text_elements:
                paragraph += element.text
            body_text.append(paragraph)
        return body_text

    def extract_image_id(xml):
        image_name = 'Picture Placeholder 4'
        image_block = xml.find('.//*[@name="%s"].....' % image_name)
        image = xml.find('.//a:blip', namespaces)
        image_id = image.attrib[embed]
        return image_id

    def get_image_path(rid, rel_xml):
        return rel_xml.find('.//*[@Id="%s"]' % rid).attrib['Target'].split('..')[1][1:]

    # HTML helper functions
    def make_tags(text, tag):
        return '<%s>' % tag + text + '</%s>' % tag

    def format_content(text_list):
        formated = ''
        for item in text_list:
            formated += make_tags(item, 'p')
            formated += '\n'
        return formated

    # loop through all slides to build content
    for i in range(len(slides)):
        print('starting slide: %s' % i)
        # retrieve text from slide

        # cache slide/relationship xml
        slide = slide_xmls[i]
        rel = rel_xmls[i]

        title = extract_title(slide)
        print('title: %s' % title)
        content = extract_content(slide)

        # retrieve image relationship from slide
        rid = extract_image_id(slide)

        # retrieve images from rel file
        img_src = get_image_path(rid, rel)

        ####STRING MANIPULATION####
        # open template, could also generate from scratch
        with open(template_path, 'r') as temp:
            template = temp.read()
            # format text as HTML
            # {{CONTENT}} needs <p></p> wrappers on text
            # {{IMAGE}} needs "s to make links work properly
            # {{TITLE}} comes with <h1> tags #maybe we should get rid of that for consistency?
        formated_content = format_content(content)
        # create img links in template
        formatted_image = '"<img src="%s"></img>"' % img_src

        # replace text in template
        template = template.replace('{{CONTENT}}', formated_content)
        template = template.replace('{{TITLE}}', title)
        template = template.replace('{{IMAGE}}', img_src)

        ####OS FILE MANIPULATION/SAVING####


        # save template
        dir_name = os.path.splitext(os.path.basename(slide_names[i]))[0]
        dir_path = save_path + '/' + dir_name
        media_path = dir_path + '/media'
        file_name = dir_path + '/index.html'

        os.makedirs(dir_path)
        os.makedirs(media_path)

        newfile = open(file_name, 'w')
        newfile.write(template)
        newfile.close()

        # move images to new folder
        with open(media_path + '/' + img_src.split('/')[-1], 'wb') as image:
            image.write(powerpoint.read('ppt/' + img_src))

        # cheap js hack for demo
        pages.append('"pages/' + dir_name + '/index.html",\n')
    # loop until all slides are done
    pages.sort()
    for p in pages:
        print(p)


#### button creation. this is messy. fix later
button_opts = {'fill': constants.BOTH, 'padx': 5, 'pady': 5}
ppt_button = tkinter.Button(root, text='Select Powerpoint', command=set_ppt_path).pack(**button_opts)
save_button = tkinter.Button(root, text='Set Save Location', command=set_save_path).pack(**button_opts)
start_button = tkinter.Button(root, text='Convert', command=start_conversion).pack(**button_opts)
