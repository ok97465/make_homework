r"""Misc.

Todo:
    None

References:
    .. [] 책: 저자명. (년). 챕터명. In 편집자명 (역할), 책명 (쪽). 발행지 : 발행사
    .. [] 학위 논문: 학위자명, "논문제목", 대학원 이름 석사 학위논문, 1990
    .. [] 저널 논문: 저자. "논문제목". 저널명, . pp.

:auther: ok97465
:Date created: 22.08.15 16:36:23
"""
# %% Import
import os
import subprocess
import platform


def open_file_externally(path):
    """Open file externally."""
    current_platform = platform.system()
    if current_platform == "Linux":
        subprocess.call(["xdg-open", path], start_new_session=True)
    elif current_platform == "Windows":
        os.system("start " + path)
    elif current_platform == "Darwin":
        subprocess.call(["open", path])
