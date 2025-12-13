class Slide():
    '''
    Base Class for all Slides.
    TODO : Maybe use this to define style?
    '''
    def make_boxes(self):
        pass

    def make_layout(self):
        pass

    def make_content(self):
        pass

    def make_titles(self):
        pass

    def make_slide(self):
        self.make_boxes()
        self.make_layout()
        self.make_content()
        self.make_titles()
