import os
import tkinter
from tkinter import filedialog
from tkinter import constants
from pptx import Presentation
from mappings import templates, mappings, SlideType

class Application:
    save_path = ''
    ppt_path = ''
    root = tkinter.Tk()

    def __init__(self):
        pass

    def set_ppt_path(self):
        opts = {}
        opts['parent'] = Application.root
        opts['title'] = 'Select Powerpoint to Convert'
        opts['filetypes'] = [('Powerpoint Files', '.pptx')]
        # hard code this for now for convenience
        opts['initialdir'] = 'C:/Users/eric_/Desktop/GreenMockups/xmltest/SAFE/'
        Application.ppt_path = filedialog.askopenfilename(**opts)

    def set_save_path(self):
        opts = {}
        opts['parent'] = Application.root
        opts['title'] = 'Select Save Location'
        # hard code this for now for convenience
        opts['initialdir'] = 'C:/Users/eric_/Desktop/GreenMockups/xmltest/SAVETEST/'
        Application.save_path = filedialog.askdirectory(**opts)

    def run(self):
        self.powerpoint = Powerpoint(Application.ppt_path)
        for slide in self.powerpoint.slides:
            page = Converter.convert(slide)
            page.save_to_disk()



class Powerpoint:
    def __init__(self, ppt_path):
        self.powerpoint = Presentation(ppt_path)
        self.make_layout_list()

    def make_layout_list(self):
        self.layouts = []
        for i in range(len(self.powerpoint.slide_layouts)):
            self.layouts.append(self.powerpoint.slide_layouts[i])

    def name_slides(self):
        for slide in self.powerpoint.slides:
            slide_number = self.powerpoint.slides.index(slide) + 1
            slide.name = "Slide%02d" % slide_number

    def get_slide_type(self, slide):
        return SlideType(self.layouts.index(slide.slide_layout))


class Converter:
    def __init__(self, slide):
        pass

    def convert(self, slide):
        slide_type = Powerpoint.get_slide_type(slide)
        html = templates[slide_type]
        mapping = mappings[slide_type]
        page = Page(slide.name)

        for item in mapping:
            element = slide.placeholders[item.idx]
            if element.has_text_frame:
                formatted_text = self.make_tags(element.text, 'p') + '\n'
                html = html.replace(item.template_element, formatted_text)
            else : # for the time being can reasonably assume this means an image
                   # but should probably figure out a better way to handle this
                img = element.image
                page.add_image(img.filename, img.blob)
                img_src = '<img src="media/%s"></img>' % img.filename
                html = html.replace(item.template_element, img_src)

        page.build_html(html)
        return page

    def make_tags(self, text, tag):
        return '<%s>' % tag + text + '</%s>' % tag
    # A conversion involves extracting information from a slide object and formatting that information
    # for presentation in an HTML document
    # In order to make this conversion you must know both the type of slide you are extracting from
    # as well as the structure of the template you are formatting for
    # there should be a one to one mapping between information in the slide and in the template
    # or in the case of lists (ie, questions, paragpraphs) the ability to format a list and place in
    # a singular location
    # a one to one mapping implies concurrent lists or key value pairs would be appropriate to store
    # the information between the file types
    # the converter could look in this mapping to find what information to take from the slide type
    # and where to put it in the HTML
    #
    # For instance: A 'Title Only Slide' has only a title named 'Title 1'
    # The  HTML would have a corresponding '{{TITLE}}' in the template
    # A mechanism is neccesary to inform the converter to extract 'Title 1' and place it in '{{TITLE}}'
    # this would need to be dependant on the fact that it was a Title Only slide
    # Converter would detect type 'Title Only Slide' look up its pattern and recieve a {'Title 1', '{{TITLE}}'}
    # More complex structures would be returned as {'Title 1' : '{{TITLE}}',
    # Possible formats, two lists, 2d list, list of tuples, dictionary, objects

    # Working Example: Title Slide has 'Title 1' and 'Subtitle 2'
    # mappings = {}
    # mappings[SlideType.TitleSlide] = {'Title 1': '((TITLE}}', 'Subtitle 2' : '{{SUBTITLE}}'}
    # Conversion object needs to take in a slide and a template and a mapping
    # Set up instances in own module?
    # we're going to use namedtuples!
    # Map = namedtuple('Mapping',  ['slide_element', 'template_element'])
    # The dict will now return a list of Map objects on SlideType lookup

# This is slated for deletion since all of this can be done with pptx
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

# TODO Start here tomorrow. Work on saving. Then do GUI
class Page:
    template_head = '''<!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="utf-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta http-equiv="Cache-Control" Content="no-cache" />
        <meta http-equiv="Pragma" Content="no-cache" />
        <meta http-equiv="Expires" Content="0" />
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <meta name="apple-mobile-web-app-capable" content="yes">
        <meta name="apple-mobile-web-app-status-bar-style" content="black">
        <meta name="format-detection" content="telephone=no">

        <title></title>

        <link href="https://fonts.googleapis.com/css?family=Titillium+Web:300,400,400i,700" rel="stylesheet">
        <link href="https://fonts.googleapis.com/css?family=Permanent+Marker" rel="stylesheet">
        <link href="https://fonts.googleapis.com/css?family=Caveat:400,700" rel="stylesheet">
        <link href="https://fonts.googleapis.com/css?family=Open+Sans:400,400i,700,700i" rel="stylesheet">
        <link href="https://maxcdn.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css" rel="stylesheet">
        <link href="../../css/animate.min.css" rel="stylesheet" >
        <link href="../../css/sandstone.css" rel="stylesheet">
        <link href="../../css/main.css" rel="stylesheet">
        <link rel="apple-touch-icon-precomposed" href="../../images/_icons/touch-icon-iphone-60.png">
        <link rel="apple-touch-icon-precomposed" sizes="76x76" href="../../images/_icons/touch-icon-ipad-76.png">
        <link rel="apple-touch-icon-precomposed" sizes="120x120" href="../../images/_icons/touch-icon-iphone-retina-120.png">
        <link rel="apple-touch-icon-precomposed" sizes="152x152" href="../../images/_icons/touch-icon-ipad-retina-152.png">
        <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
        <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
        <!--[if lt IE 9]>
          <script src="https://oss.maxcdn.com/html5shiv/3.7.3/html5shiv.min.js"></script>
          <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
        <![endif]-->
    </head>
    <body id="main">

        <!--==================================================
        ======================================================
        START CONTENT - EDIT BELOW ONLY
        ======================================================
        ===================================================-->


        <!--===============================================================
        Section
        ================================================================-->
        '''
    template_tail = '''
    <!--==================================================
        ======================================================
        END CONTENT - EDIT ABOVE ONLY
        ======================================================
        ===================================================-->



        <script src="../../js/jquery-3.1.1.min.js"></script>
        <script src="../../js/bootstrap.min.js"></script>
        <script src="../../js/iframeResizer.contentWindow.min.js" defer></script>
        <script src="../../js/page.js"></script>


    </body>
    </html>'''
    def __init__(self, name):
        self.html = ''
        self.images = {}
        self.slide_name = name

    def build_html(self, html):
        self.html = Page.template_head + html + Page.template_tail

    def add_image(self, name, image):
        self.images[name] = image

    def save_to_disk(self):
        base_path = os.path.join(Application.save_path,  self.slide_name)
        media_path = os.path.join(base_path, 'media')
        page_name = os.path.join(base_path,  'index.html')

        os.makedirs(base_path, exist_ok=True)
        os.makedirs(media_path, exist_ok=True)

        with open(page_name, 'w') as new_file
            new_file.write(self.html)

        for image_name in self.images:
            image_path = os.path.join(media_path, image_name)
            with open (image_path, 'wb') as image:
                image.write(self.images[image_name])



# Move this into the converter? yup delete it now
class HTMLFormatter:
    def make_tags(self):
        pass

    def format_content(self):
        pass
