import random
import pandas as pd
import os
import env
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
from io import BytesIO
from time import perf_counter
from collections import Counter
import os
import smtplib
from email.message import EmailMessage
import textwrap

# Start timer
start_time = perf_counter()

# Read configuration from environment variables (with defaults)
EMAIL_FUNCTIONALITY = os.getenv("EMAIL_FUNCTIONALITY", "False").strip().lower() == "true"
EMAIL_SENDER = os.getenv("EMAIL_SENDER", "")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT", "465"))
PRINT_CYCLES = os.getenv("PRINT_CYCLES", "False").strip().lower() == "true"
PRINT_NAMED_CYCLES = os.getenv("PRINT_NAMED_CYCLES", "False").strip().lower() == "true"

# Double check that EMAIL_FUNCTIONALITY is meant to be True
if EMAIL_FUNCTIONALITY:
    while True:
        confirmation = input(
            "\n⚠️  EMAIL FUNCTIONALITY IS ENABLED.\n"
            "Are you sure you want to send emails? (yes / no): "
            ).strip().lower()

        if confirmation in {"yes", "y"}:
            EMAIL_FUNCTIONALITY = True
            print("\n✅ Email sending confirmed.\n")
            break
        elif confirmation in {"no", "n"}:
            EMAIL_FUNCTIONALITY = False
            print("\n🚫 Email sending cancelled. No emails will be sent.\n")
            break
        else:
            print("Please type 'yes' or 'no'.\n")

def wrap_text(text, width):
    # Return text wrapped to a fixed width
    text = (text or "").strip()
    if not text:
        return ""
    return textwrap.fill(text, width=width)

# Define pairing function
def secret_santa(n):
    buyers = list(range(1, n + 1))
    receivers = buyers.copy()
    while True:
        random.shuffle(receivers)
        if all(buyer != receiver for buyer, receiver in zip(buyers, receivers)):
            return list(zip(buyers, receivers))

# Load participants from chosen xlsx file
df = pd.read_excel("participants_100.xlsx")
n = len(df)
buyer_receiver_pairs = secret_santa(n)

# Change empty requests to empty strings to avoid NaN
df['Request'] = df['Request'].fillna("").astype(str)

# Create dataframe
pairs_df = pd.DataFrame(buyer_receiver_pairs, columns=['Buyer', 'Receiver'])
pairs_df['Buyer_Name'] = pairs_df['Buyer'].apply(lambda x: df.loc[x - 1, 'Name'])
pairs_df['Receiver_Name'] = pairs_df['Receiver'].apply(lambda x: df.loc[x - 1, 'Name'])
pairs_df['Buyer_Email'] = pairs_df['Buyer'].apply(lambda x: df.loc[x - 1, 'Email'])
pairs_df['Receiver_Email'] = pairs_df['Receiver'].apply(lambda x: df.loc[x - 1, 'Email'])
pairs_df['Receiver_Extra'] = pairs_df['Receiver'].apply(lambda x: df.loc[x - 1, 'Request'])
pairs_df.to_excel("secret_santa_pairs.xlsx", index=False)

# PDF variables
template_pdf = "christmas_template.pdf"
output_folder = "secret_santa_notes"
os.makedirs(output_folder, exist_ok=True)

# Choose font and layout
font_name = "Helvetica-Bold"
italic_font_name = "Helvetica-Oblique" # for italic section
font_size = 32
max_text_margin = 150
line_spacing = font_size + 6
vertical_offset = -32

# Marker to indicate that a line should be rendered in italics
ITALIC_MARKER = "§§I§§"

# Message on PDFs
template = """★★★ Ho ho ho! ★★★

{buyer},

You have been chosen as Secret Santa for...

{receiver}!
{extra_section}
Remember to bring your present to the
❄ EVENT NAME AND DATE HERE ❄

The budget for presents is €10, so be creative!

Shhhh, it's a secret....
"""

def send_email(to_email, buyer_name, pdf_path):
    msg = EmailMessage()
    msg["Subject"] = "🎅 Your Secret Santa Assignment!"
    msg["From"] = EMAIL_SENDER
    msg["To"] = to_email

    body = f"""{buyer_name},

Ho ho ho! 🎁

Attached is your Secret Santa assignment for the EVENT NAME HERE.


Happy gifting!
"""
    msg.set_content(body)

    # Attach the PDF
    with open(pdf_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename=os.path.basename(pdf_path))

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)
        print(f"    ✅ Email sent to {buyer_name} ({to_email})\n")
    except Exception as e:
        print(f"    ❌ Failed to send to {buyer_name} ({to_email}): {e}\n")

# Put template text on PDF in centre
def create_centered_overlay(text, font_name, italic_font_name, font_size, page_width, page_height):
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))

    # Wrap lines within margins, tracking italic vs normal
    wrapped_lines = []  # list of (line_text, is_italic)

    for line in text.splitlines():
        is_italic = False
        if line.startswith(ITALIC_MARKER):
            is_italic = True
            line = line[len(ITALIC_MARKER):]

        if line.strip() == "":
            wrapped_lines.append(("", is_italic))
        else:
            # Use the correct font when computing wrapping
            base_font = italic_font_name if is_italic else font_name
            parts = simpleSplit(line, base_font, font_size, page_width - max_text_margin)
            for part in parts:
                wrapped_lines.append((part, is_italic))

    text_height = len(wrapped_lines) * line_spacing
    start_y = (page_height + text_height) / 2 + vertical_offset
    center_x = page_width / 2

    for line, is_italic in wrapped_lines:
        current_font = italic_font_name if is_italic else font_name
        c.setFont(current_font, font_size)
        line_width = c.stringWidth(line, current_font, font_size)
        c.drawString(center_x - line_width / 2, start_y, line)
        start_y -= line_spacing

    c.save()
    packet.seek(0)
    return PdfReader(packet)

pdf_paths = {}

print("\n  📝 Generating personalized PDFs...\n")

for _, row in pairs_df.iterrows():
    buyer = row["Buyer_Name"]
    receiver = row["Receiver_Name"]
    
    receiver_first_name = receiver.strip().split()[0]

    # Safely get and strip the extra info
    receiver_extra = (row.get("Receiver_Extra", "") or "").strip()

    # Build optional extra section
    if receiver_extra:
        wrapped_extra_section = wrap_text(receiver_extra, 60)

        # ONLY wrapped_extra_section is italic:
        italic_wrapped = "\n".join(
            ITALIC_MARKER + line for line in wrapped_extra_section.splitlines()
        )

        extra_section = f"""
{receiver_first_name} has provided the following information:
{italic_wrapped}
"""
    else:
        extra_section = ""

    text = template.format(
        buyer=buyer,
        receiver=receiver,
        extra_section=extra_section
    )

    background_pdf = PdfReader(template_pdf)
    page = background_pdf.pages[0]

    page_width = float(page.mediabox.width)
    page_height = float(page.mediabox.height)

    overlay_pdf = create_centered_overlay(
        text,
        font_name,
        italic_font_name,
        font_size,
        page_width,
        page_height
    )

    writer = PdfWriter()
    page.merge_page(overlay_pdf.pages[0])
    writer.add_page(page)

    out_path = f"{output_folder}/{buyer}.pdf"
    with open(out_path, "wb") as f:
        writer.write(f)

    pdf_paths[buyer] = out_path
    print(f"    ✅ PDF created for {buyer}")

print("\n  📄 All PDFs generated successfully!\n")

if EMAIL_FUNCTIONALITY:
    print("  📧 Sending Secret Santa emails...\n")

    for i, row in pairs_df.iterrows():
        buyer = row["Buyer_Name"]
        buyer_email = row["Buyer_Email"]
        pdf_path = pdf_paths.get(buyer)

        print(f"  ➡️  Sending email {i+1}/{len(pairs_df)}: {buyer} <{buyer_email}>")
        send_email(buyer_email, buyer, pdf_path)

    print("\n  🎉 All emails have been sent!\n")
else:
    print("  📭 Email functionality is disabled. No emails sent.\n")

# Set up for PRINT_CYCLES       
mapping = dict(buyer_receiver_pairs)
visited = set()
cycles = []          
cycle_lengths = [] 

if PRINT_CYCLES:
    if PRINT_NAMED_CYCLES:
        print("Pairing cycles:\n")
    for person in mapping:
        if person in visited:
            continue
        current = person
        cycle = []
        while current not in visited:
            visited.add(current)
            cycle.append(current)
            current = mapping[current]

        names = [df.loc[i - 1, 'Name'] for i in cycle]
        cycles.append(names)
        cycle_lengths.append(len(names))
        
        if PRINT_NAMED_CYCLES:
            print("\n".join(names))
            print("    Cycle Length :", len(names))
            print("\n")

    print("Summary:")
    print(f"  Total cycles: {len(cycle_lengths)}")
    print(f"  Cycle lengths: {sorted(cycle_lengths, reverse=True)}")
    print()

end_time = perf_counter()
time_taken = end_time - start_time

print("\n   " + "=" * 60)
print("   🎅  SECRET SANTA GENERATOR COMPLETE!  🎁")
print("   " + "=" * 60)
print(f"   {'📄  Total participants':<25}:   {n}")
print(f"   {'✅  Pairings saved to':<25}:   secret_santa_pairs.xlsx")
print(f"   {'📂  PDFs saved to':<25}:   {output_folder}/")
print(f"   {'⏱️   Time taken':<27}:   {time_taken:.4f} seconds")
print("   " + "=" * 60 + "\n")
print()
