def correct_deformity(image):
    # Convert PIL image to numpy array
    img_array = np.array(image)

    # Convert to grayscale
    gray = cv2.cvtColor(img_array, cv2.COLOR_BGR2GRAY)
    # Apply GaussianBlur to reduce noise and improve edge detection
    gray = cv2.GaussianBlur(gray, (5, 5), 0)
    # Edge detection
    edges = cv2.Canny(gray, 50, 150)

    # Find contours
    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # Assume the largest contour is the order form
    contours = sorted(contours, key=cv2.contourArea, reverse=True)
    form_contour = contours[0]

    # Approximate the contour to a polygon
    epsilon = 0.02 * cv2.arcLength(form_contour, True)
    approx = cv2.approxPolyDP(form_contour, epsilon, True)

    if len(approx) == 4:
        # Get the points in the form of (x, y)
        pts = approx.reshape(4, 2)
        
        # Order points: top-left, top-right, bottom-right, bottom-left
        rect = np.zeros((4, 2), dtype="float32")
        s = pts.sum(axis=1)
        rect[0] = pts[np.argmin(s)]
        rect[2] = pts[np.argmax(s)]

        diff = np.diff(pts, axis=1)
        rect[1] = pts[np.argmin(diff)]
        rect[3] = pts[np.argmax(diff)]

        # Destination points to map the form to
        (tl, tr, br, bl) = rect
        widthA = np.sqrt(((br[0] - bl[0]) ** 2) + ((br[1] - bl[1]) ** 2))
        widthB = np.sqrt(((tr[0] - tl[0]) ** 2) + ((tr[1] - tl[1]) ** 2))
        maxWidth = max(int(widthA), int(widthB))

        heightA = np.sqrt(((tr[0] - br[0]) ** 2) + ((tr[1] - br[1]) ** 2))
        heightB = np.sqrt(((tl[0] - bl[0]) ** 2) + ((tl[1] - bl[1]) ** 2))
        maxHeight = max(int(heightA), int(heightB))

        dst = np.array([
            [0, 0],
            [maxWidth - 1, 0],
            [maxWidth - 1, maxHeight - 1],
            [0, maxHeight - 1]], dtype="float32")

        # Apply the perspective transform
        M = cv2.getPerspectiveTransform(rect, dst)
        warped = cv2.warpPerspective(img_array, M, (maxWidth, maxHeight))

        # Convert back to PIL image
        corrected_image = Image.fromarray(warped)
        return corrected_image
    else:
        # If four corners are not detected, return the original image
        return image