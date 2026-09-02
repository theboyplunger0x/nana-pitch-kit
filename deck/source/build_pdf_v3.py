from pathlib import Path

from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parents[2]
PAGES = ROOT / "tmp/deck-v3-aesthetic/pdf-pages-hires"
OUT = ROOT / "deck/current/Nana-Pitch-Deck-v3.pdf"

width, height = 1280, 720
page_images = sorted(PAGES.glob("slide-*.png"), key=lambda p: int(p.stem.split("-")[-1]))
if len(page_images) != 19:
    raise RuntimeError(f"Expected 19 rendered slides, found {len(page_images)}")

OUT.parent.mkdir(parents=True, exist_ok=True)
pdf = canvas.Canvas(str(OUT), pagesize=(width, height), pageCompression=1)
pdf.setTitle("Nana Pitch Deck v3")
pdf.setAuthor("Nana Wallet")
pdf.setSubject("Agentic wallet for seniors, people with limited mobility, and their families")

for page_image in page_images:
    pdf.drawImage(str(page_image), 0, 0, width=width, height=height, preserveAspectRatio=False, mask="auto")
    pdf.showPage()

pdf.save()
print(OUT)
