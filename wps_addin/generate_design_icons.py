from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).parent
CANVAS = 128
OUTPUT = 32
BLUE = (22, 119, 255, 255)
NAVY = (31, 41, 55, 255)
ORANGE = (245, 158, 11, 255)
PURPLE = (139, 92, 246, 255)
GREEN = (15, 157, 88, 255)
PINK = (225, 29, 72, 255)
TEAL = (8, 145, 178, 255)
TRANSPARENT = (0, 0, 0, 0)


def new_icon():
    image = Image.new("RGBA", (CANVAS, CANVAS), TRANSPARENT)
    return image, ImageDraw.Draw(image)


def rounded(draw, box, color, width=9):
    draw.rounded_rectangle(box, radius=18, outline=color, width=width)


def save(image, name):
    image.resize((OUTPUT, OUTPUT), Image.Resampling.LANCZOS).save(
        ROOT / name, optimize=True
    )


def make_style_brush():
    image, draw = new_icon()
    draw.line((34, 94, 72, 56), fill=BLUE, width=12)
    draw.line((76, 52, 90, 38), fill=NAVY, width=12)
    draw.polygon(((96, 26), (112, 42), (88, 66), (70, 48)), fill=ORANGE)
    save(image, "icon_design_style.png")


def make_text():
    image, draw = new_icon()
    rounded(draw, (22, 18, 106, 110), BLUE)
    draw.line((44, 46, 84, 46), fill=PURPLE, width=9)
    draw.line((44, 68, 84, 68), fill=NAVY, width=8)
    draw.line((44, 90, 70, 90), fill=NAVY, width=8)
    save(image, "icon_design_text.png")


def make_layout():
    image, draw = new_icon()
    rounded(draw, (20, 22, 108, 106), BLUE)
    draw.rectangle((68, 30, 99, 59), fill=ORANGE)
    draw.line((64, 22, 64, 106), fill=BLUE, width=7)
    draw.line((20, 64, 108, 64), fill=BLUE, width=7)
    save(image, "icon_design_layout.png")


def make_cleanup():
    image, draw = new_icon()
    draw.line((26, 34, 102, 34), fill=TEAL, width=10)
    draw.line((52, 22, 76, 22), fill=TEAL, width=9)
    rounded(draw, (36, 46, 92, 106), TEAL, width=9)
    draw.line((54, 64, 54, 88), fill=TEAL, width=8)
    draw.line((74, 64, 74, 88), fill=TEAL, width=8)
    save(image, "icon_design_cleanup.png")


def make_export():
    image, draw = new_icon()
    rounded(draw, (18, 38, 92, 106), BLUE)
    draw.line((34, 88, 54, 66), fill=NAVY, width=8)
    draw.line((50, 70, 66, 88), fill=NAVY, width=8)
    draw.line((78, 50, 108, 20), fill=ORANGE, width=12)
    draw.polygon(((88, 18), (112, 18), (112, 42)), fill=ORANGE)
    save(image, "icon_design_export.png")


def make_color():
    image, draw = new_icon()
    draw.polygon(((64, 14), (98, 62), (30, 62)), fill=BLUE)
    draw.ellipse((26, 54, 102, 112), fill=BLUE)
    draw.ellipse((42, 68, 60, 86), fill=ORANGE)
    draw.ellipse((68, 68, 86, 86), fill=PINK)
    draw.ellipse((55, 90, 73, 108), fill=GREEN)
    save(image, "icon_design_color.png")


def make_photoshop():
    image, draw = new_icon()
    rounded(draw, (16, 16, 112, 112), BLUE, width=10)
    try:
        font = ImageFont.truetype("DejaVuSans-Bold.ttf", 52)
    except OSError:
        try:
            font = ImageFont.load_default(size=52)
        except TypeError:
            font = ImageFont.load_default()
    text = "Ps"
    box = draw.textbbox((0, 0), text, font=font)
    draw.text((64 - (box[2] - box[0]) / 2, 64 - (box[3] - box[1]) / 2 - box[1]),
              text, font=font, fill=BLUE)
    save(image, "icon_design_photoshop.png")


def main():
    make_style_brush()
    make_text()
    make_layout()
    make_cleanup()
    make_export()
    make_color()
    make_photoshop()


if __name__ == "__main__":
    main()
