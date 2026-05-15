from pptx import Presentation

def create_qbr(customer):

    prs = Presentation()

    # Slide 1
    slide_layout = prs.slide_layouts[0]

    slide = prs.slides.add_slide(slide_layout)

    title = slide.shapes.title

    subtitle = slide.placeholders[1]

    title.text = f"{customer} QBR"

    subtitle.text = "Quarterly Business Review"

    # Slide 2
    bullet_slide_layout = prs.slide_layouts[1]

    slide2 = prs.slides.add_slide(bullet_slide_layout)

    title2 = slide2.shapes.title

    title2.text = "Customer Highlights"

    body = slide2.placeholders[1]

    tf = body.text_frame

    tf.text = "Usage increased 25%"

    p = tf.add_paragraph()
    p.text = "Strong product adoption"

    p = tf.add_paragraph()
    p.text = "Expansion opportunity identified"

    filename = f"{customer}_QBR.pptx"

    prs.save(filename)

    return filename

