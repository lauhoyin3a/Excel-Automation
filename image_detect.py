import pyautogui

# Load the image of the "Replace All" button
replace_all_button_image = r"C:\Users\Tom\Downloads\replace_all_button.png"
r=None
while r is None:
    r=pyautogui.locateOnScreen(replace_all_button_image, grayscale=True, confidence=.3)
print(r)
# Locate the position of the "Replace All" button on the screen
button_position = r
print(button_position)
if button_position is not None:
    # Get the center coordinates of the button
    button_x, button_y, button_width, button_height = button_position
    button_center_x = button_x + button_width // 2
    button_center_y = button_y + button_height // 2

    # Move the mouse to the center of the button and click
    pyautogui.moveTo(button_center_x, button_center_y)
    pyautogui.click()
else:
    print("Unable to locate the 'Replace All' button.")