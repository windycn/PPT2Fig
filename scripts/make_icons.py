"""Draw the small, self-contained Ribbon icon set."""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


DEST = Path(__file__).resolve().parents[1] / "office/icons"
SCALE = 4
SIZE = 96
FONT = Path("/System/Library/Fonts/Supplemental/Arial Bold.ttf")


def icon(name: str, color: str, label: str, kind: str) -> None:
    image = Image.new("RGBA", (SIZE * SCALE, SIZE * SCALE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    def box(coords, fill, radius=0, outline=None, width=1):
        coords = tuple(round(v * SCALE) for v in coords)
        if radius:
            draw.rounded_rectangle(coords, radius * SCALE, fill=fill, outline=outline,
                                   width=width * SCALE)
        else:
            draw.rectangle(coords, fill=fill, outline=outline, width=width * SCALE)

    def line(points, fill, width=3, joint="curve"):
        draw.line([(round(x * SCALE), round(y * SCALE)) for x, y in points],
                  fill=fill, width=width * SCALE, joint=joint)

    ink = "#253447"
    pale = "#F3F6FA"
    if kind == "scope":
        box((14, 22, 58, 76), "#DCE7F2", 5, "#8396AB", 2)
        box((29, 12, 79, 71), "#FFFFFF", 5, ink, 3)
        box((37, 24, 69, 29), color, 2)
        box((37, 36, 65, 40), "#B7C7D8", 2)
        box((37, 46, 59, 50), "#B7C7D8", 2)
        line([(60, 62), (81, 62)], color, 5)
        line([(73, 53), (82, 62), (73, 71)], color, 5)
    elif kind == "selected":
        box((17, 18, 66, 73), "#DCE7F2", 5, "#8396AB", 2)
        box((30, 10, 78, 70), "#FFFFFF", 5, ink, 3)
        for x, y in ((37, 28), (52, 28), (37, 44), (52, 44)):
            box((x, y, x + 9, y + 9), color, 2)
        line([(60, 73), (68, 81), (83, 64)], color, 5)
    elif kind == "shapes":
        box((20, 22, 75, 72), "#FFFFFF", 4, color, 3)
        draw.ellipse((36*SCALE, 34*SCALE, 59*SCALE, 57*SCALE), fill=color)
        for x, y in ((16, 18), (71, 18), (16, 68), (71, 68)):
            box((x, y, x + 8, y + 8), "#FFFFFF", 1, ink, 2)
    elif kind == "all":
        box((12, 25, 57, 79), "#E1EAF3", 4, "#8DA0B5", 2)
        box((24, 17, 69, 71), "#F1F5F9", 4, "#687E93", 2)
        box((36, 9, 81, 63), "#FFFFFF", 4, ink, 3)
        for y in (23, 34, 45):
            box((45, y, 69, y + 4), color, 2)
    elif kind == "settings":
        for y, knob in ((24, 38), (46, 64), (68, 47)):
            line([(18, y), (80, y)], ink, 4)
            box((knob - 7, y - 7, knob + 7, y + 7), "#FFFFFF", 4, color, 3)
    elif kind == "raster":
        box((12, 19, 69, 71), "#FFFFFF", 5, ink, 3)
        box((23, 29, 59, 53), pale, 2)
        draw.ellipse((49 * SCALE, 32 * SCALE, 56 * SCALE, 39 * SCALE), fill=color)
        line([(24, 52), (36, 39), (44, 46), (51, 40), (60, 52)], color, 3)
        box((32, 67, 85, 83), color, 3)
    elif kind == "eps":
        box((21, 10, 73, 80), "#FFFFFF", 5, ink, 3)
        line([(30, 55), (46, 30), (66, 55)], color, 4)
        for x, y in [(30, 55), (46, 30), (66, 55)]:
            box((x-4, y-4, x+4, y+4), "#FFFFFF", 1, color, 3)
        box((31, 64, 63, 69), color, 2)
    else:
        box((21, 10, 73, 80), "#FFFFFF", 5, ink, 3)
        box((31, 23, 63, 27), color, 2)
        box((31, 34, 57, 38), "#B7C7D8", 2)
        box((31, 45, 62, 49), "#B7C7D8", 2)
        box((29, 58, 65, 74), color, 3)

    if label:
        font = (ImageFont.truetype(str(FONT), 11 * SCALE) if FONT.is_file()
                else ImageFont.load_default(size=11 * SCALE))
        bounds = draw.textbbox((0, 0), label, font=font)
        x = round((SIZE * SCALE - (bounds[2] - bounds[0])) / 2)
        if kind == "raster":
            y = 69 * SCALE
        elif kind == "eps":
            y = 64 * SCALE
            box((29, 62, 65, 78), color, 3)
        else:
            y = 59 * SCALE
        draw.text((x, y), label, font=font, fill="#FFFFFF", stroke_width=0)

    image.resize((SIZE, SIZE), Image.Resampling.LANCZOS).save(DEST / f"{name}.png")


def main() -> None:
    DEST.mkdir(parents=True, exist_ok=True)
    icon("scope", "#0879A9", "", "scope")
    icon("selected", "#4B719D", "", "selected")
    icon("shapes", "#725AA8", "", "shapes")
    icon("all", "#3C7E9D", "", "all")
    icon("png", "#16887C", "PNG", "file")
    icon("jpg", "#AD7840", "JPG", "file")
    icon("tiff", "#506DAD", "TIF", "file")
    icon("pdf", "#BB4745", "PDF", "file")
    icon("eps", "#AD6D2C", "EPS", "eps")
    icon("settings", "#435E79", "", "settings")


if __name__ == "__main__":
    main()
