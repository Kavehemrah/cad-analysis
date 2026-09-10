import win32com.client
import pythoncom
from win32com.client import VARIANT

acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument
ms = doc.ModelSpace

# نقطه اصلی
x = 563273.517648
y = 3594736.029122
z = 0.0

# مقدار جابه‌جایی
offset = 0.150  # 150 mm

# فرض: پایین صفحه = Y-
new_x = x
new_y = y - offset

def pt(x, y, z=0.0):
    return VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        (x, y, z)
    )

# نقطه اصلی
p1 = ms.AddPoint(pt(x, y, z))

# نقطه جابه‌جا شده
p2 = ms.AddPoint(pt(new_x, new_y, z))

# خط بین دو موقعیت
line = ms.AddLine(
    pt(x, y, z),
    pt(new_x, new_y, z)
)

doc.Regen(1)

print("Original:")
print(f"X = {x}")
print(f"Y = {y}")

print("\nOffset:")
print(f"X = {new_x}")
print(f"Y = {new_y}")

print("\nOffset = 150 mm toward Y-")