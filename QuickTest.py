from collections import namedtuple
from enum import IntEnum
from pptx import Presentation
import os

Map = namedtuple('Mapping', ['slide_element', 'idx', 'template_element'])

class SlideType(IntEnum):
    TitleSlide = 0
    Blank = 1
    Table = 2
    Question = 3
    Text = 4
    ImageText = 5
    TabsAccordianPills = 6
    ImageCarousel = 7
    Interactive = 8
    Video = 9
    VideoText = 10
    CanvasAnimation = 11
    CanvasAnimationText = 12
    Modal = 13
    FieldNote = 14
    SafetyTip = 15

p = Presentation("C:/Users/eric_/Desktop/GreenMockups/xmltest/SAFE/BigTest.pptx")
mappings = {}
templates = {}
layouts = []

mappings[SlideType.ImageText] = [Map('Title 1', 0, '{{TITLE}}'),
                                 Map('Text Placeholder 2', 2, '{{TEXT}}'),
                                 Map('Picture Placeholder 3', 10, '{{IMAGE}}')]

templates[SlideType.ImageText] = '''
<section>
        <div class="row">
            <div class="col-xs-10 col-xs-offset-1">
                <div class="row">
                    <div class="col-xs-12 col-sm-12">
                        <h1> {{TITLE}} </h1>
                    </div>
                </div>
            </div>
        </div>

        <div class="row">
            <div class="col-xs-10 col-xs-offset-1">
                <div class="row">
                    <div class="col-xs-12 col-sm-12 col-md-3">
                        <img class="img-responsive" src={{IMAGE}}>
                    </div>
                    <div class="col-xs-12 col-sm-12 col-md-9">
                         {{TEXT}}
                    </div>
                </div>
            </div>
        </div>
    </section>
    '''

for i in range(len(p.slide_layouts)):
            layouts.append(p.slide_layouts[i])

def get_slide_type(slide):
    return SlideType(layouts.index(slide.slide_layout))
