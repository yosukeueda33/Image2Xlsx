import cv2
import numpy as np
import xlsxwriter


###### Control area BEGIN

desired_scale   = 1.0
quantize_num    = 8    #from 0
inverse         = True
cell_size       = 10 

###### Control area END

img_bgr = cv2.imread("input.jpg")

img_gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
cv2.imwrite("gray.jpg", img_gray)

print(f"original size:{img_gray.shape}")

print(f"desired scale:{desired_scale}")
desired_size = (
    int(img_gray.shape[1] * desired_scale),
    int(img_gray.shape[0] * desired_scale),
)

img_resized = cv2.resize(img_gray, desired_size)
print(f"resized size:{img_resized.shape}")

bins = np.arange(
    start=0.0,
    stop=255,
    step=255/quantize_num
)
print(f"bins:{bins}")
img_mozaic = np.digitize(img_resized, bins)


with xlsxwriter.Workbook('result.xlsx') as workbook:
    worksheet = workbook.add_worksheet()
    # Write pixel intensity.
    for y in range(0, img_mozaic.shape[0]):
        for x in range(0, img_mozaic.shape[1]):
            e_num = img_mozaic.item(y, x)
            worksheet.write_number(
                row=y,
                col=x,
                number=e_num if not inverse else (quantize_num-e_num)
            )
    # Set width and height.
    for y in range(0, img_mozaic.shape[0]):
        worksheet.set_row(
            row=y,
            height=cell_size
        )
    worksheet.set_column(
        first_col=0,
        last_col=img_mozaic.shape[1],
        width=cell_size*0.1
    )
