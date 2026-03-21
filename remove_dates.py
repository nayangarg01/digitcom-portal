import cv2
import numpy as np

def remove_date(img_path, out_path):
    print(f"Processing {img_path}")
    img = cv2.imread(img_path)
    if img is None:
        print("Failed to read image")
        return
        
    # Convert to HSV to easily select orange color
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    
    # Orange range
    lower_orange = np.array([5, 150, 150])
    upper_orange = np.array([25, 255, 255])
    
    # Yellow/Orange range (sometimes dates look a bit yellowish)
    lower_orange2 = np.array([10, 100, 200])
    upper_orange2 = np.array([30, 255, 255])
    
    mask1 = cv2.inRange(hsv, lower_orange, upper_orange)
    mask2 = cv2.inRange(hsv, lower_orange2, upper_orange2)
    mask = cv2.bitwise_or(mask1, mask2)
    
    # Dilate mask slightly to cover edges of text
    kernel = np.ones((5,5), np.uint8)
    mask = cv2.dilate(mask, kernel, iterations=1)
    
    # Inpaint
    result = cv2.inpaint(img, mask, 5, cv2.INPAINT_TELEA)
    cv2.imwrite(out_path, result)
    print(f"Saved to {out_path}")

remove_date("images/gallery/1.png", "images/gallery/1_cleaned.png")
remove_date("images/gallery/10.png", "images/gallery/10_cleaned.png")
