import collections 
import collections.abc
from pptx import Presentation
from PIL import Image
import os


def get_placeholder_info(prs):
    slide_layouts = prs.slide_master.slide_layouts
    for num in range(len(slide_layouts)):
        for placeholder in prs.slide_layouts[num].placeholders:
            print('%d %s' % (placeholder.placeholder_format.idx, placeholder.name))


def insert_img(prs, placeholder, img):
    picture = placeholder.insert_picture(img)
    return picture


def add_slide(prs, layout):
    new_slide = prs.slides.add_slide(layout)
    return new_slide



prs = Presentation('testproj2_pres.pptx')

get_placeholder_info(prs) #Helper for now

all_slides = prs.slides

main_slide = all_slides[0]
main_title_text = "This is the Main Pres Title"
main_slide_title_placeholder = main_slide.shapes.title
main_slide_title_placeholder.text = main_title_text

new_slide_layout = prs.slide_layouts[2]
new_slide = add_slide(prs, new_slide_layout)

prs.save('testproj2_pres.pptx')



