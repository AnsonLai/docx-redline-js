import os
import fitz
from PIL import Image

out_dir = r'tmp/multilevel-bullet-visual'
os.makedirs(out_dir, exist_ok=True)

cases = [
    {
        'id': 'synthetic-nested-child',
        'dir': r'tmp/word-visual-review/rendered',
        'prefix': 'administrative-list-change-nested-child',
        'page': 0
    },
    {
        'id': 'superdoc-board-agenda-multiple-children',
        'dir': r'tmp/superdoc-word-visual-review/rendered',
        'prefix': '38-administrative-administrative-list-change-board-agenda-multiple-children',
        'page': 0
    },
    {
        'id': 'superdoc-bylaws-nested-list-batch',
        'dir': r'tmp/superdoc-word-visual-review/rendered',
        'prefix': '28-legal-legal-bylaws-nested-list-batch',
        'page': 0
    }
]

views = ['allMarkup', 'acceptAll', 'rejectAll']

for c in cases:
    case_images = []
    for v in views:
        pdf_path = os.path.join(c['dir'], f"{c['prefix']}-{v}.pdf")
        doc = fitz.open(pdf_path)
        page = doc.load_page(c['page'])
        pix = page.get_pixmap(dpi=150)
        img_name = f"{c['id']}--{v}.png"
        img_path = os.path.join(out_dir, img_name)
        pix.save(img_path)
        print(f"Saved {img_path} ({pix.width}x{pix.height})")
        case_images.append(img_path)
        doc.close()

    # Create 3-view contact sheet
    imgs = [Image.open(p) for p in case_images]
    total_w = sum(im.width for im in imgs) + 40
    max_h = max(im.height for im in imgs) + 20
    sheet = Image.new('RGB', (total_w, max_h), (240, 240, 240))
    x = 10
    for im in imgs:
        sheet.paste(im, (x, 10))
        x += im.width + 10
    sheet_path = os.path.join(out_dir, f"{c['id']}--contact-sheet.png")
    sheet.save(sheet_path)
    print(f"Saved contact sheet {sheet_path}")

print("ALL MULTILEVEL VISUAL CASES RENDERED.")
