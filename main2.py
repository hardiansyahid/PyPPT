from datetime import date
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE_TYPE

text_top = [
    "Cepat",
    "Stabil",
    "Scalable"
]

text_bottom = [
    "Support Python",
    "Auto Generate PPT",
    "Ready for Report"
]

today = date.today()

def set_bullet_list(shape, items, font_size=14):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    for i, text in enumerate(items):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = text
        p.level = 0
        p.font.size = Pt(font_size)

def read_and_write_ppt():
    try:
        prs = Presentation("pyppt.pptx")
        slide = prs.slides[1]

        # title
        slide.shapes.title.text = f"Update Kondisi Siklon Tropis. Selasa, {today}"

        # === WAJIB pastikan index placeholder ===
        # debug dulu:
        for p in slide.placeholders:
            print(p.placeholder_format.idx, p.name)

        text_top_shape = slide.placeholders[0]
        text_bottom_shape = slide.placeholders[1]

        set_bullet_list(text_top_shape, text_top)
        set_bullet_list(text_bottom_shape, text_bottom)

        # replace image
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                slide.shapes._spTree.remove(shape._element)
                slide.shapes.add_picture("new_image.png", left, top, width=width, height=height)
                break

        prs.save("output.pptx")
    except Exception as e:
        print(e)

if __name__ == "__main__":
    read_and_write_ppt()
