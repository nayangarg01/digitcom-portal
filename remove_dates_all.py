import cv2
import numpy as np
import os
import glob

def remove_date(img_path):
    print(f"Processing {img_path}")
    img = cv2.imread(img_path)
    if img is None:
        print("Failed to read image")
        return
        
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    
    lower_orange = np.array([5, 150, 150])
    upper_orange = np.array([25, 255, 255])
    
    lower_orange2 = np.array([10, 100, 200])
    upper_orange2 = np.array([30, 255, 255])
    
    mask1 = cv2.inRange(hsv, lower_orange, upper_orange)
    mask2 = cv2.inRange(hsv, lower_orange2, upper_orange2)
    mask = cv2.bitwise_or(mask1, mask2)
    
    kernel = np.ones((5,5), np.uint8)
    mask = cv2.dilate(mask, kernel, iterations=1)
    
    # If there are no orange pixels (e.g. less than 50 pixels), don't process it to save time and prevent artifacting
    if cv2.countNonZero(mask) < 50:
        print(f"No date stamp detected in {img_path}")
        return

    result = cv2.inpaint(img, mask, 5, cv2.INPAINT_TELEA)
    cv2.imwrite(img_path, result)
    print(f"Cleaned {img_path}")

gallery_dir = "images/gallery/"
for ext in ('*.png', '*.jpg', '*.jpeg'):
    for path in glob.glob(os.path.join(gallery_dir, ext)):
        if "cleaned" not in path:
            remove_date(path)
