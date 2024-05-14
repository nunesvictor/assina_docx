import os
from math import floor
from typing import TypeAlias

import imgkit
from docx import Document
from docx.shared import Length, Mm

_PageSize: TypeAlias = dict[str, Length]


A4: _PageSize = {
    "page_height": Mm(297),
    "page_width": Mm(210),
    "left_margin": Mm(25.4),
    "right_margin": Mm(25.4),
    "top_margin": Mm(25.4),
    "bottom_margin": Mm(25.4),
    "header_distance": Mm(12.7),
    "footer_distance": Mm(12.7),
}

OPEN_FILE_CMD = "open"


class DocumentSigner:
    def __init__(self, filename: str, page_size: _PageSize = A4) -> None:
        self.document = Document(filename)
        self.page_size = page_size

        for section in self.document.sections:
            section.page_height = self.page_size["page_height"]
            section.page_width = self.page_size["page_width"]
            section.left_margin = self.page_size["left_margin"]
            section.right_margin = self.page_size["right_margin"]
            section.top_margin = self.page_size["top_margin"]
            section.bottom_margin = self.page_size["bottom_margin"]
            section.header_distance = self.page_size["header_distance"]
            section.footer_distance = self.page_size["footer_distance"]

    @staticmethod
    def __mm2px(mm: float, dpi: int = 96) -> int:
        return floor(mm * dpi / 25.4)

    @staticmethod
    def __get_imgkit_config() -> dict[str, str]:
        return {
            "format": "png",
            "enable-local-file-access": "",
            # 210mm - 25.4mm (left) - 25.4mm (right) + 15px (padding)
            "crop-w": DocumentSigner.__mm2px(210 - (25.4 * 2)) + 15,
            "transparent": "",
        }

    def __populate_footer(self, link: str, uuid: str) -> None:
        for section in self.document.sections:
            footer = section.footer

            # limpa todos os parágrafos do footer
            for paragraph in footer.paragraphs:
                paragraph.clear()

            paragraph = footer.paragraphs[0]
            # 159.2mm = 210mm - 25.4mm (left) - 25.4mm (right)
            paragraph.add_run().add_picture("build/banner.png", width=Mm(159.2))

    def sign(self, link: str, uuid: str) -> None:
        self.__populate_footer(link, uuid)
        self.document.save("build/file_out.docx")


if __name__ == "__main__":
    # cria o banner (pode ser feito via from_url, from_string ou from_file).
    # aqui estamos usando from_file
    imgkit.from_file(
        "./sign_banner.html",
        "build/banner.png",
        options=DocumentSigner.__get_imgkit_config(),
    )

    signer = DocumentSigner("assets/file_in.docx")
    signer.sign(
        link="https://solar.defensoria.to.def.br/docs/d/validar/",
        uuid="A6B56B39D2-195AD85977-740AB544E6-CB1D22BD03",
    )

    os.system(f"{OPEN_FILE_CMD} build/file_out.docx")
    exit(0)
