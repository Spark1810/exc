# PG 3 

Here's a step-by-step guide to run the VBA functions in Excel:

1. **Open Excel**: Open Microsoft Excel on your computer.

2. **Open or Create a Workbook**: You can either open an existing workbook where you want to use these functions or create a new workbook.

3. **Open the Visual Basic Editor**:
   - Press `ALT + F11` to open the Visual Basic for Applications (VBA) editor.
   - Alternatively, you can go to the "Developer" tab (if it's not visible, you need to enable it in Excel options), and click on "Visual Basic" to open the VBA editor.

4. **Insert a New Module**:
   - In the VBA editor, click on "Insert" in the menu bar.
   - Then select "Module" from the drop-down menu. This will insert a new module into the project.

5. **Copy and Paste the Functions**:
   - Copy the functions provided in the concise version (listed in the previous response).
   - Paste them into the module window in the VBA editor.

6. **Close the VBA Editor**: Close the VBA editor window.

7. **Use the Functions in Excel**:
   - Go back to your Excel workbook.
   - In any cell, you can now use these functions like any other Excel functions. For example:
     - To calculate the square of a number, you can type `=square(5)` in a cell and press Enter. This will return 25.
     - To calculate the cube of a number, you can type `=cube(3)` in a cell and press Enter. This will return 27.
     - To calculate the area of a rectangle, you can type `=Rectangle(4, 5)` in a cell and press Enter. This will return 20.
     - To calculate the area of a triangle, you can type `=Triangle(6, 8)` in a cell and press Enter. This will return 24.
     - To calculate the area of a circle, you can type `=hi(3)` in a cell and press Enter. This will return the area of a circle with radius 3.

That's it! You've successfully added and used the VBA functions in Excel. You can now use these functions in your Excel worksheets as needed.

-----------------------------------------------------------------------------------------------------------------------------------------------------

Function square(Side1 As Integer) As Integer
    square = Side1 ^ 2
End Function

Function cube(Side1 As Integer) As Integer
    cube = Side1 ^ 3
End Function

Function Rectangle(Side1 As Integer, Side2 As Integer) As Integer
    Rectangle = Side1 * Side2
End Function

Function Triangle(Side1 As Integer, Side2 As Integer) As Integer
    Triangle = (Side1 * Side2) / 2
End Function

Function hi(Side1 As Double) As Double
    hi = 3.14 * Side1 ^ 2
End Function
+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++




# Pg 4

