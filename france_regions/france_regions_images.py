"""
Script to download images about France regions
"""

import os
import re
import shutil
from argparse import ArgumentParser
from math import ceil

import imagesize
import requests
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from yaml import safe_load

CONFIG_FILE = "france_regions_images.yaml"
IMAGES_DIR = "images_regions"

IMAGES_PER_SLIDE = 5
IMAGE_HEIGHT_WIDTH_RATIO = 0.71
SLIDE_HEIGHT_INCHES = 7.2
SLIDE_WIDTH_INCHES = 10
IMAGE_LEFT_OFFSET_DEFAULT = 0.6
IMAGE_TOP_OFFSET_DEFAULT = 1.7
LINE_INTERVAL_INCHES_DEFAULT = 0.8

IMAGE_MAX_PIXEL_HEIGHT = 600

SLD_LAYOUT_TITLE_SLIDE = 0
SLD_LAYOUT_TITLE_ONLY = 5

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
    images_per_region = None
    presentation_title = None
    if "layout" in config:
        if "region" in config["layout"]:
            if "max_images" in config["layout"]["region"]:
                images_per_region = config["layout"]["region"]["max_images"]
        if "slide" in config["layout"]:
            if "images" in config["layout"]["slide"]:
                images_per_slide = config["layout"]["slide"]["images"]
        if "title" in config["layout"]:
            presentation_title = config["layout"]["title"]

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
                if re.match(r"https*:", url):
                    status = download_image(url, filename)
                    if status != 0:
                        continue
                else:
                    # Assume it is a local file and copy it
                    shutil.copyfile(url, filename)
            image_paths[region].append({"place": place, "file": filename})

    # Create PowerPoint file and title slide
    prs = Presentation()
    # Create a title slide if a title is defined in configuration
    if presentation_title:
        slide_layout = prs.slide_layouts[SLD_LAYOUT_TITLE_SLIDE]
        title_slide = prs.slides.add_slide(slide_layout)
        title = title_slide.shapes.title
        title.text = "Régions de France en image"

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
    line_interval_inches = LINE_INTERVAL_INCHES_DEFAULT
    image_top_offset = IMAGE_TOP_OFFSET_DEFAULT
    image_full_width_inches = (SLIDE_WIDTH_INCHES - image_left_offset) / images_per_line
    image_width_inches = image_full_width_inches - image_left_offset
    image_height_inches = image_width_inches * IMAGE_HEIGHT_WIDTH_RATIO
    image_full_height_inches = image_height_inches + line_interval_inches

    total_height = image_full_height_inches * lines
    if total_height > SLIDE_HEIGHT_INCHES:
        # Adjust image width so that images fit in the slide height, based on height/width ratio
        # full height includes line interval
        image_full_height_inches = (SLIDE_HEIGHT_INCHES - image_top_offset) / lines
        image_height_inches = image_full_height_inches - line_interval_inches
        image_width_inches = image_height_inches / IMAGE_HEIGHT_WIDTH_RATIO
        image_full_width_inches = image_width_inches / IMAGE_HEIGHT_WIDTH_RATIO
        image_left_offset = (SLIDE_WIDTH_INCHES - (image_width_inches * images_per_line)) / (images_per_line + 1)
    else:
        # Center vertically the images
        max_height = SLIDE_HEIGHT_INCHES - image_top_offset
        line_interval_inches = (max_height - (image_height_inches * lines)) / (lines + 0.5)
        image_full_height_inches = image_height_inches + line_interval_inches
        image_top_offset += line_interval_inches * 0.5

    for region, images in image_paths.items():
        region_slide_num = 0

        # Compute left offset for last line to balance free space around images if the number of
        # images on this line is less than the number of images on other lines. It can be different
        # for each region, depending on the actual number of images for the region
        region_image_num = len(images)
        if images_per_region and len(images) > images_per_region:
            region_image_num = images_per_region
        region_lines = ceil(region_image_num / images_per_line)
        region_current_line = 0
        last_line_images = region_image_num % images_per_line
        if last_line_images == 0:
            last_line_left_offset = image_left_offset
        else:
            last_line_left_offset = (SLIDE_WIDTH_INCHES - (image_width_inches * last_line_images)) / (
                last_line_images + 1
            )

        if "maps" in config and region in config["maps"]:
            width, height = imagesize.get(config["maps"][region])
            max_image_height_inches = SLIDE_HEIGHT_INCHES - IMAGE_TOP_OFFSET_DEFAULT - LINE_INTERVAL_INCHES_DEFAULT
            height_inches = height / IMAGE_MAX_PIXEL_HEIGHT * max_image_height_inches
            if height_inches > max_image_height_inches:
                slide_image_width = max_image_height_inches * width / height
            else:
                slide_image_width = SLIDE_WIDTH_INCHES - (2 * IMAGE_LEFT_OFFSET_DEFAULT)
            left_offset = (SLIDE_WIDTH_INCHES - slide_image_width) / 2
            slide_layout = prs.slide_layouts[SLD_LAYOUT_TITLE_ONLY]
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.add_picture(
                config["maps"][region],
                Inches(left_offset),
                Inches(IMAGE_TOP_OFFSET_DEFAULT),
                width=Inches(slide_image_width),
            )
            slide.shapes.title.text = region

        for i, image_params in enumerate(images):
            if images_per_region and i == images_per_region:
                break

            if (i % images_per_line) == 0:
                region_current_line += 1

            if (i % images_per_slide) == 0:
                region_slide_num += 1
                # Create a new slile
                slide_layout = prs.slide_layouts[SLD_LAYOUT_TITLE_ONLY]
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
            if region_current_line == region_lines:
                # Last line
                left_offset = last_line_left_offset
                if last_line_images == 0:
                    line_image_num = images_per_line
                else:
                    line_image_num = last_line_images
            else:
                left_offset = image_left_offset
                line_image_num = images_per_line
            left = left_offset + (i % line_image_num) * (image_width_inches + left_offset)
            top = image_top_offset + (image_full_height_inches * line)
            slide.shapes.add_picture(image, Inches(left), Inches(top), width=Inches(image_width_inches))

            # Add caption
            caption = slide.shapes.add_textbox(
                Inches(left),
                Inches(top + image_height_inches + 0.05),
                Inches(image_width_inches),
                Inches(0.5),
            )
            p = caption.text_frame.paragraphs[0]
            p.text = place
            p.alignment = PP_ALIGN.CENTER

    # Sauvegarde
    prs.save(f"{IMAGES_DIR}/diaporama_regions_images.pptx")

    print("\n✅ Terminé ! Fichier créé : diaporama_regions_images.pptx")


if __name__ == "__main__":
    exit(main())
