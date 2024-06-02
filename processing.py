import numpy as np
from PIL import Image
import fitz  # PyMuPDF
import io
import cv2
import matplotlib.pyplot as plt
import os

def pdf_to_images(pdf_path):
    doc = fitz.open(pdf_path)
    images = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap()
        img = Image.open(io.BytesIO(pix.tobytes()))
        images.append(img)
    return images

def crop_image(image, crop_box):
    return image.crop(crop_box)

def display_images(images, titles=None):
    for i, image in enumerate(images):
        # Convert the PIL Image to a NumPy array
        image_np = np.array(image)
        # Display the image using OpenCV
        if titles:
            cv2.imshow(titles[i], image_np)
        else:
            cv2.imshow(f"Image {i}", image_np)
        cv2.waitKey(0)
        cv2.destroyAllWindows()

def save_images_as_jpg(images, output_folder):
    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)
    
    # Save each image as a .jpg file in the output folder
    for i, img in enumerate(images):
        # Construct the file path for the image
        file_path = os.path.join(output_folder, f"image_{i}.jpg")
        # Convert PIL image to RGB mode if necessary
        if img.mode != 'RGB':
            img = img.convert('RGB')
        # Save the image as a .jpg file
        img.save(file_path, "JPEG")
        print(f"Image {i} saved as {file_path}")

def correct_skew(image):
    # Convert PIL Image to NumPy array
    image_np = np.array(image)

    # Convert to grayscale
    gray = cv2.cvtColor(image_np, cv2.COLOR_RGB2GRAY)
    # Apply edge detection
    edges = cv2.Canny(gray, 50, 150, apertureSize=3)
    # Use the Hough Line Transform to detect lines in the image
    lines = cv2.HoughLines(edges, 1, np.pi / 180, 200)
    
    # Calculate the angle of the lines
    angles = []
    if lines is not None:
        for rho, theta in lines[:, 0]:
            angle = (theta - np.pi / 2) * 180 / np.pi
            angles.append(angle)
    
    # Compute the median angle of the detected lines
    if len(angles) > 0:
        median_angle = np.median(angles)
    else:
        median_angle = 0  # If no lines are detected, assume no rotation is needed

    # Rotate the image to correct skew
    (h, w) = image_np.shape[:2]
    center = (w // 2, h // 2)
    M = cv2.getRotationMatrix2D(center, median_angle, 1.0)
    rotated = cv2.warpAffine(image_np, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    
    # Convert the corrected NumPy array back to a PIL Image
    corrected_image = Image.fromarray(cv2.cvtColor(rotated, cv2.COLOR_BGR2RGB))
    
    return corrected_image

# Read PDF file
pdf_path = "resources/klasa_1b.pdf"
images = pdf_to_images(pdf_path)

# Correct skew for each cropped image
corrected_images = [correct_skew(img) for img in images]

# Define crop box and crop images
crop_box = (0, 0, 418, 310)
cropped_images = [crop_image(img, crop_box) for img in corrected_images]

# Display images
display_images(cropped_images)

# # Display images
# display_images(enhanced_names)

# # Dump images to file
# save_images_as_jpg(enhanced_names, "processed_images")