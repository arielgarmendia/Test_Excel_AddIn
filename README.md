## **About:**

Solution to develop an specifid add-in for BBVA Pricer. 

## **Important tips and notes:**

- Use Visual Studio 2022
- **Clone** this repository in your local repo folder then **Open** solution then **Rebuild** solution to restore **Nuget** packages.
- Install VS Office extensions if asked. 
- Install visual studio extension: Code Converter (VB - C#), using main menu Extensions/Manage Extentions - Search Online.
- Use provided VB.Net project Pricer.VBA.Conversion to copy from excel's VBA macros code, into VB classes and then converto to C#.
    - Conversion procedure is the following:
        - After creating a dummy VB class and pasting the VBA code.
        - Select in editor the desired code to convert, i.e. the whole class code.
        - Right click on selected code, menu: Convert to C$.
        - This will create a new .cs doc with a class and the content translated from VB to C#.
        - Then move code to the main add-in project, choose the desired location or container class.
        - Unfortunately we have to manually check the whole converion to have a working code.