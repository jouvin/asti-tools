"""
Script to download images about France regions
"""

import os
from argparse import ArgumentParser
from math import ceil

import requests
from pptx import Presentation
from pptx.util import Inches
from yaml import safe_load

CONFIG_FILE = "france_regions_images.yaml"
IMAGES_DIR = "images_regions"

IMAGES_PER_SLIDE = 5
IMAGE_HEIGHT_WIDTH_RATIO = 0.71
SLIDE_HEIGTH_INCHES = 7.2
SLIDE_WIDTH_INCHES = 10
IMAGE_LEFT_OFFSET_DEFAULT = 0.6
IMAGE_TOP_OFFSET = 1.7
LINE_INTERVAL_INCHES = 0.8

# Dossier pour stocker les images
os.makedirs(IMAGES_DIR, exist_ok=True)

# Images par région
with open(CONFIG_FILE, "r", encoding="utf-8") as f:
    config = safe_load(f.read())

# Données (régions + lieux)
if "regions" in config:
    regions = config["regions"]
else:
    raise Exception(f"No regions found in {CONFIG_FILE}")


# Fonction pour télécharger une image depuis Unsplash (source libre)
def download_image(url, filename):
    """
    Download an image from a url

    :param url: image url
    :param filename: image filename
    :return: status (0 for success, http status code otherwhise)
    """
    response = requests.get(url)
    if response.status_code == 200:
        with open(filename, "wb") as f:
            f.write(response.content)
            status = 0
    else:
        print(f"Error downloading {url} to {filename}: {response.reason}")
        status = response.status_code

    return status


def main():
    parser = ArgumentParser()
    parser.add_argument(
        "--image-per-page",
        type=int,
        default=None,
        help="Number of images per page",
    )
    parser.add_argument(
        "--no-download", action="store_true", default=False, help="Do not download images, use existing ones"
    )
    options = parser.parse_args()

    images_per_slide = IMAGES_PER_SLIDE
    if "layout" in config:
        if "slide" in config["layout"]:
            if "images" in config["layout"]["slide"]:
                images_per_slide = config["layout"]["slide"]["images"]

    # Command line options take precedence over config file
    if options.image_per_page:
        images_per_slide = options.image_per_page

    image_paths = {}

    for region, places in regions.items():
        image_paths[region] = []
        for place, url in places.items():
            filename = f"{IMAGES_DIR}/{region}_{place}.jpg".replace(" ", "_")
            if options.no_download:
                if os.path.exists(filename):
                    print(f"{place}: using existing {filename}")
                else:
                    print(f"{filename} does not exist and --no-download option specified")
                    continue
            else:
                print(f"Téléchargement : {place}")
                status = download_image(url, filename)
                if status != 0:
                    continue
            image_paths[region].append({"place": place, "file": filename})

    # Création du PowerPoint
    prs = Presentation()

    if images_per_slide <= 3:
        lines = 1
    elif images_per_slide <= 6:
        lines = 2
    else:
        raise Exception("More than 6 images per slide currently not supported")
    images_per_line = ceil(images_per_slide / lines)

    # Compute image width based on number of images per line
    # full width includes left margin
    image_left_offset = IMAGE_LEFT_OFFSET_DEFAULT
    image_full_width_inches = (SLIDE_WIDTH_INCHES - image_left_offset) / images_per_line
    image_width_inches = image_full_width_inches - image_left_offset
    image_heigth_inches = image_width_inches * IMAGE_HEIGHT_WIDTH_RATIO
    image_full_heigth_inches = image_heigth_inches + LINE_INTERVAL_INCHES

    # Adjust image width so that images fit in the slide heigth, based on heigth/width ratio
    # full heigth includes line interval
    total_heigth = image_full_heigth_inches * lines
    if total_heigth > SLIDE_HEIGTH_INCHES:
        image_full_heigth_inches = (SLIDE_HEIGTH_INCHES - IMAGE_TOP_OFFSET) / lines
        image_heigth_inches = image_full_heigth_inches - LINE_INTERVAL_INCHES
        image_width_inches = image_heigth_inches / IMAGE_HEIGHT_WIDTH_RATIO
        image_full_width_inches = image_width_inches / IMAGE_HEIGHT_WIDTH_RATIO
        image_left_offset = (SLIDE_WIDTH_INCHES - (image_width_inches * images_per_line)) / (images_per_line + 1)

    for region, images in image_paths.items():
        region_slide_num = 0

        for i, image_params in enumerate(images):
            if (i % images_per_slide) == 0:
                region_slide_num += 1
                # Create a new slile
                slide_layout = prs.slide_layouts[5]
                slide = prs.slides.add_slide(slide_layout)
                # Titre
                title = slide.shapes.title
                slide_num_text = "" if region_slide_num == 1 else region_slide_num
                if slide_num_text:
                    slide_num_text = f" ({slide_num_text})"
                title.text = f"{region} {slide_num_text}"
                line = 0
            else:
                line = (i % images_per_slide) // images_per_line

            # Add image
            place = image_params["place"]
            image = image_params["file"]
            left = image_left_offset + (i % images_per_line) * (image_width_inches + image_left_offset)
            top = IMAGE_TOP_OFFSET + (image_full_heigth_inches * line)
            slide.shapes.add_picture(image, Inches(left), Inches(top), width=Inches(image_width_inches))

            # Add caption
            caption = slide.shapes.add_textbox(
                Inches(left + 0.5),
                Inches(top + image_heigth_inches + 0.05),
                Inches(image_width_inches + image_left_offset),
                Inches(0.5),
            )
            caption.text_frame.text = place

    # Sauvegarde
    prs.save(f"{IMAGES_DIR}/diaporama_regions_images.pptx")

    print("\n✅ Terminé ! Fichier créé : diaporama_regions_images.pptx")


if __name__ == "__main__":
    exit(main())
