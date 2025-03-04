from pathlib import Path
import time
import shutil
import os
import pandas as pd
import qrcode


# Function to create vCard 4.0 string from a panda row
def build_vcard(row):
    vcard = [
        "BEGIN:VCARD",
        "VERSION:4.0"
    ]

    # Add fields only if they exist and are not empty
    if pd.notna(row["Firstname"]) and row["Firstname"]:
        name = f"N:{row['Lastname']};{row['Firstname']};;;"
        vcard.append(name)
        vcard.append(f"FN:{row['Firstname']} {row['Lastname']}")
    if pd.notna(row["Company"]) and row["Company"]:
        vcard.append(f"ORG:{row['Company']}")
    if pd.notna(row["Title"]) and row["Title"]:
        vcard.append(f"TITLE:{row['Title']}")
    if pd.notna(row["Work Phone"]) and row["Work Phone"]:
        vcard.append(f"TEL;TYPE=WORK:{row['Work Phone']}")
    if pd.notna(row["Mobile Phone"]) and row["Mobile Phone"]:
        vcard.append(f"TEL;TYPE=CELL:{row['Mobile Phone']}")
    if pd.notna(row["Work Email"]) and row["Work Email"]:
        vcard.append(f"EMAIL;TYPE=WORK:{row['Work Email']}")
    if pd.notna(row["Personal Email"]) and row["Personal Email"]:
        vcard.append(f"EMAIL;TYPE=HOME:{row['Personal Email']}")
    if pd.notna(row["Website"]) and row["Website"]:
        vcard.append(f"URL:{row['Website']}")
    if any(pd.notna(row[field]) and row[field] for field in ["P.O. Box", "Street", "City", "State", "Zip Code", "Country"]):
        address = row["P.O. Box"] if pd.notna(row["P.O. Box"]) else ""
        street = row["Street"] if pd.notna(row["Street"]) else ""
        city = row["City"] if pd.notna(row["City"]) else ""
        region = row["State"] if pd.notna(row["State"]) else ""
        zipcode = row["Zip Code"] if pd.notna(row["Zip Code"]) else ""
        country = row["Country"] if pd.notna(row["Country"]) else ""
        vcard.append(
            f"ADR;TYPE=WORK:{address};;{street};{city};{region};{zipcode};{country}")

    vcard.append("END:VCARD")
    return "\n".join(vcard)


def generate_vcard_images(rows, output_dir: str, options: dict):

    # Process each row
    for index, row in rows:
        # Generate vCard string
        vcard_data = build_vcard(row)

        # Create QR code
        qr = qrcode.QRCode(
            version=options["version"],  # Auto-size based on data
            error_correction=options["error_correction"],
            box_size=options["pixel_box_size"],
            border=options["border"]
        )

        qr.add_data(vcard_data)
        qr.make(fit=True)

        # Generate and save QR code as image
        img = qr.make_image(
            fill_color=options["color"], back_color=options["background_color"])
        filename = output_dir / \
            f"{index+1}_{row['Firstname']}_{row['Lastname']}_vcard.png"
        img.save(filename)


# Generates  vCard QR codes optimized for business cards and zips them into a timestamped archive.
def main(upload_filepath: str, base_output_dir: str, qrcode_options: dict):

    # read the data from xlsx or csv
    if upload_filepath.endswith(".csv"):
        df = pd.read_csv(upload_filepath, dtype=str)
    else:
        df = pd.read_excel(upload_filepath, dtype=str)

    # Generate output directory
    timestamp = round(time.time() * 1000)
    output_dir = Path(f"{base_output_dir}/{timestamp}")
    output_dir.mkdir(exist_ok=True)

    # Generate the QR Codes
    generate_vcard_images(df.iterrows(), output_dir, qrcode_options)

    # zip the directory
    output = shutil.make_archive(output_dir, 'zip', output_dir)

    return f'{timestamp}.zip'
