r"""4살을 위한 사칙연산 숙제.

Todo:
    None

References:
    .. [] 책: 저자명. (년). 챕터명. In 편집자명 (역할), 책명 (쪽). 발행지 : 발행사
    .. [] 학위 논문: 학위자명, "논문제목", 대학원 이름 석사 학위논문, 1990
    .. [] 저널 논문: 저자. "논문제목". 저널명, . pp.

:auther: ok97465
:Date created: 22.08.15 16:32:24
"""
# %% Import
# Standard library imports
import datetime
from random import randint

# Local imports
from helper_pptx import add_hquiz, add_slide_title, add_vquiz, get_new_ppt
from misc import open_file_externally

# %% Parameter
date = datetime.datetime.now().strftime("%y년%m월%d일")
# date = "22년8월7일"
filename = f"{date}_4years.pptx"

slide_width = 21
slide_height = 29.7

# %% Generate slide
prs = get_new_ppt(slide_width, slide_height)
blank_slide_layout = prs.slide_layouts[6]
slide_quiz = prs.slides.add_slide(blank_slide_layout)
slide_solu = prs.slides.add_slide(blank_slide_layout)

add_slide_title(slide_quiz, f"{date} 예서 문제", slide_width)
add_slide_title(slide_solu, f"{date} 예서 정답", slide_width)

# %% Add Quiz
# vertical quiz
y_offset = 3.0
y_diff = 5.0
x_offset = 3.0
x_diff = 12.0

# add
for iy in range(3):
    for ix in range(2):
        pos_x = x_offset + ix * x_diff
        pos_y = y_offset + iy * y_diff
        v1 = randint(2, 15)
        v2 = randint(2, 5)
        add_vquiz(slide_quiz, v1, v2, False, pos_x, pos_y, "+")
        add_vquiz(slide_solu, v1, v2, True, pos_x, pos_y, "+")

# horizontal quiz
y_offset = 20.0
y_diff = 3.0
x_offset = 0.8
x_diff = 10.0

# add
for iy in range(3):
    for ix in range(2):
        pos_x = x_offset + ix * x_diff
        pos_y = y_offset + iy * y_diff
        v1 = randint(2, 15)
        v2 = randint(2, 5)
        add_hquiz(slide_quiz, v1, v2, False, pos_x, pos_y, "+")
        add_hquiz(slide_solu, v1, v2, True, pos_x, pos_y, "+")

# %% Save pptx
prs.save(filename)
open_file_externally(filename)
