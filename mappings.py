from collections import namedtuple
from enum import IntEnum


# Used to access slide_layout types. New templates should be added here in the same order
# that they appear in the powerpoint dropdown
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


# TODO figure out a way to deal with variable num datatypes
    # possible solutions: first element contains instructions, custom functions/converter objects,

Map = namedtuple('Mapping', ['slide_element', 'idx', 'template_element'])
# Right now I'm using the named elements but it might make sense to use the idx since those remain consistent
mappings = {}
mappings[SlideType.TitleSlide] = [Map('Title 1', 0, '{{TITLE}}'),
                                  Map('Subtitle 2', 1, '{{SUBTITLE}}'),
                                  Map('Slide Number Placeholder 5', 12, '{{SLIDENUM}}')]
mappings[SlideType.Blank] = []
mappings[SlideType.Table] = [Map('Title 1', 0, '{{TITLE}}'),
                             Map('Table Placeholder 2', 10, '{{TABLE}}')]
# Questions can have variable # of answers -  need a solution to handle this and place all in template
mappings[SlideType.Question] = [Map('Text Placeholder 1',  2, '{{QUESTION}}'),
                                Map('Text Placeholder 2', 13, '{{OPTION1}}'),
                                Map('Text Placeholder 3', 14, '{{OPTION2}}'),
                                Map('Text Placeholder 4', 15, '{{OPTION3}}'),
                                Map('Text Placeholder 5', 16, '{{OPTION4}}'),
                                Map('Text Placeholder 6', 17, '{{REMEDIATION}}')]
mappings[SlideType.Text] = [Map('Title 1', 0, '{{TITLE}}'),
                            Map('Text Placeholder 2', 2, '{{TEXT}}')]
mappings[SlideType.ImageText] = [Map('Title 1', 0, '{{TITLE}}'),
                                 Map('Text Placeholder 2', 2, '{{TEXT}}'),
                                 Map('Picture Placeholder 3', 10, '{{IMAGE}}')]
# This one could also potentially have multiples
mappings[SlideType.TabsAccordianPills] = [Map('Text Placeholder 1',  2, '{{TEXT1}}'),
                                          Map('Text Placeholder 2', 10, '{{TEXT2}}'),
                                          Map('Text Placeholder 3', 11, '{{HEADING1}}'),
                                          Map('Text Placeholder 4', 12, '{{HEADING2}}')]
# Another multiple
mappings[SlideType.ImageCarousel] = [Map('Title 1', '{{TITLE}}'),
                                     Map('Picture Placeholder 2',  0, '{{IMAGE1}}'),
                                     Map('Picture Placeholder 3', 10, '{{IMAGE2}}'),
                                     Map('Picture Placeholder 4', 11, '{{IMAGE3}}'),
                                     Map('Picture Placeholder 5', 12, '{{IMAGE4}}'),
                                     Map('Picture Placeholder 6', 13, '{{IMAGE5}}'),
                                     Map('Picture Placeholder 7', 14, '{{IMAGE6}}'),
                                     Map('Picture Placeholder 8', 15, '{{IMAGE7}}'),
                                     Map('Picture Placeholder 9', 16, '{{IMAGE8}}')]
mappings[SlideType.Interactive] = []
# This one has a video which I don't think I can extract at the moment - Maybe flag as needing manual edits
mappings[SlideType.Video] = [Map('Title 1', 0, '{{TITLE}}')]
mappings[SlideType.VideoText] = [Map('Title 1', 0, '{{TITLE}}'), Map('Text Placeholder 2', 2, '{{TEXT}}')]
# Same problem as video
mappings[SlideType.CanvasAnimation] = [Map('Title 1', 0, '{{TITLE}}')]
# Same as above
mappings[SlideType.CanvasAnimationText] = [Map('Title 1', 0, '{{TITLE}}'),
                                           Map('Text Placeholder 2', 2, '{{TEXT}}')]
mappings[SlideType.Modal] = [Map('Subtitle 1', '{{SUBTITLE}}')]
mappings[SlideType.FieldNote] = [Map('Subtitle 1', '{{SUBTITLE}}')]
mappings[SlideType.SafetyTip] = [Map('Subtitle 1', '{{SUBTITLE}}')]

# Add additional mappings below in the format
# mappings[SlideType enum] = [Map(slide_element, template_element),...]

templates = {}
templates[SlideType.TitleSlide] = ''' '''
templates[SlideType.Blank] = ''' '''
templates[SlideType.Table] = ''' '''
templates[SlideType.Question] = ''' '''
templates[SlideType.Text] = ''' '''
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
templates[SlideType.TabsAccordianPills] = ''' '''
templates[SlideType.ImageCarousel] = ''' '''
templates[SlideType.Interactive] = ''' '''
templates[SlideType.Video] = ''' '''
templates[SlideType.VideoText] = ''' '''
templates[SlideType.CanvasAnimation] = ''' '''
templates[SlideType.CanvasAnimationText] = ''' '''
templates[SlideType.Modal] = ''' '''
templates[SlideType.FieldNote] = ''' '''
templates[SlideType.SafetyTip] = ''' '''
