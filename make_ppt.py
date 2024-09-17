from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.slide import Slide
from pptx.shapes.autoshape import Shape
from pptx.enum.text import PP_ALIGN

def make_player_box(shape, name = "AA", score = "69420"):

    text_frame = shape.text_frame
    text_frame.clear()
    p = text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = f'{name}: {score}'

def make_player_layout(
        slide: Slide,
        num_players: int = 7,
        available_h: float = 6/25
        ) -> list[Shape]:
    # Adds player shapes at correct positions.

    k = num_players // 2

    # if numplayers is odd
    if num_players % 2:
        divs = 3 * k + 4
        top_ls = list(range(1, divs, 3))
        bottom_ls = [i + 0.5 for i in range(2, 3 * k + 2, 3)]
    else:
        divs = 3 * k + 1
        top_ls = [i for i in range(1, divs, 3)]
        bottom_ls = top_ls

    delta_h = available_h / 2

    top_h = (1 - available_h) * prs.slide_height
    bottom_h = top_h + delta_h * prs.slide_height

    shape_tree = slide.shapes

    print(k, divs)
    width = 2 / divs * prs.slide_width
    height = delta_h * prs.slide_height

    player_shapes = []

    for i in top_ls:
        left = i / divs * prs.slide_width
        player_shape = shape_tree.add_shape(MSO_SHAPE.RECTANGLE, left, top_h, width, height)
        make_player_box(player_shape)
        player_shapes.append(player_shape)

    for i in bottom_ls:
        left = i / divs * prs.slide_width
        player_shape = shape_tree.add_shape(MSO_SHAPE.RECTANGLE, left, bottom_h, width, height)
        make_player_box(player_shape)
        player_shapes.append(player_shape)

    return player_shapes

def set_player_detail(players: list[tuple[str, int]], player_boxes: list[Shape]):
    for i in range(len(players)):

        player_name, player_score = players[i]
        player_box = player_boxes[i]

        text_frame = player_box.text_frame
        text_frame.clear()
        p = text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = f'{player_name}: {player_score}'

prs = Presentation()
# print(prs.slide_width, prs.slide_height)
prs.slide_width = 12192 * 1000
blank_slide_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_slide_layout)

shapes = slide.shapes

# left = top = width = height = Inches(1.0)

left = prs.slide_width / 13
top = 19 / 25 * prs.slide_height
width = 2 * prs.slide_width / 13
height = 3 / 25 * prs.slide_height

player_boxes = make_player_layout(slide, 7)

players = [
    ("Be", 1),
    ("Ca", 2),
    ("Da", 3),
    ("Ja", 4),
    ("Jo", 5),
    ("Ka", 6),
    ("Su", 7),
]

set_player_detail(players, player_boxes)

prs.save("text.pptx")