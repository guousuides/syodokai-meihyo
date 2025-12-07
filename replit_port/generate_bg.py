from PIL import Image, ImageDraw

def generate_background():
    width = 595
    height = 842
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)

    # Coordinate conversion (same as applied.py)
    def to_pil(x, y):
        return x, height - y

    def draw_rect(x, y, w, h_rect):
        x0, y1 = to_pil(x, y)
        x1, y0 = to_pil(x + w, y + h_rect)
        draw.rectangle([x0, y0, x1, y1], outline="black", width=1)

    def draw_line(x1, y1, x2, y2):
        px1, py1 = to_pil(x1, y1)
        px2, py2 = to_pil(x2, y2)
        draw.line([px1, py1, px2, py2], fill="black", width=1)

    # Draw frames (same coordinates as applied.py)
    draw_rect(20, 20, 260, 190)
    draw_rect(20, 230, 260, 580)
    draw_line(120, 230, 120, 810)
    draw_line(192, 230, 192, 810)
    draw_line(236, 230, 236, 810)
    draw_line(236, 520, 280, 520)

    image.save("static/background.png")
    print("background.png generated.")

if __name__ == "__main__":
    import os
    os.makedirs("static", exist_ok=True)
    generate_background()
