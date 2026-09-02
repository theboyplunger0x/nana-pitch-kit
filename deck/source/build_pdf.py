from pathlib import Path

from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parents[2]
SOURCE_DIR = ROOT / "deck/current/preview"
OUTPUT = ROOT / "deck/current/Nana-Pitch-Deck-v2.pdf"
PAGE_SIZE = (960, 540)


def main() -> None:
    slides = sorted(SOURCE_DIR.glob("slide-*.png"))
    if len(slides) != 16:
        raise RuntimeError(f"Expected 16 rendered slides, found {len(slides)}")

    pdf = canvas.Canvas(str(OUTPUT), pagesize=PAGE_SIZE, pageCompression=1)
    pdf.setTitle("Nana Pitch Deck v2")
    pdf.setAuthor("Nana")
    for slide in slides:
        pdf.drawImage(
            str(slide),
            0,
            0,
            width=PAGE_SIZE[0],
            height=PAGE_SIZE[1],
            preserveAspectRatio=True,
            mask="auto",
        )
        pdf.showPage()
    pdf.save()


if __name__ == "__main__":
    main()
