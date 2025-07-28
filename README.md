# MEPCE Tools

PyRevit Extension with MEP and BIM tools created by Simeon Neese.


## Features

#### 🔧Mechanical Tools
- Clash Detection 💥
    - Checks for clashes between selected cateogries. Includes filters for system types. Draws boxes around clash geometry using detail lines.
- HVAC Zoning 🎨
    - Creates filled regions for each room in view.
    - Adds filled regions to color-coded zones.
#### 💧Plumbing Tools
- Pipe Volumes 📏
    - Exports a list of total pipe volumes for each system type.
- Water Demand 🚿
    - Enters water fixture type counts into water demand calculator excel sheet.
#### ⚡Electrical Tools
- Circuit Lighting 💡
    - Groups lights and adds them to a circuit based on filled region zoning.
#### 📃BIM Tools
- Rename 📝
    - Rename Sheets & Rename Views - use find / replace logic to rename sheets/views. Allows the addition of a prefix / suffix as well.
- Export Schedules 📩
    - Export room schedule with zones and loads (from filled regions).


## Installation

- Install [pyRevit](https://github.com/pyrevitlabs/pyRevit/releases)
- Download MEPCE_Tools_Extension.exe : 📦[MEPCE_Tools_Extension.exe](https://github.com/simneese/MEPCERevitTools/releases/download/v1.0.0/MEPCE_Tools_Extension.exe)
- Run MEPCE_Tools_Extension.exe
- Navigate to Extensions menu in the pyRevit toolbar:
  
  <img src="https://github.com/user-attachments/assets/cba76431-a299-4696-acf1-84b91ea5cee3" alt="Alt Text" width="500">

- Find MEPCE Tools in the menu and install (make sure you click the drop down with the install directory - it is easy to miss):
  
   <img src="https://github.com/user-attachments/assets/9ba0a7eb-7460-4b70-a18e-3fd87962022b" alt="Alt Text" width="500">

- Done!


## How to Update

- Update the toolbar using the "Update" button on the pyRevit tab. It is located right next to the "Extensions" button in the dropdown used to get to the pyRevit settings.


## Roadmap

- Remake Dynamo scripts in python

- Add tutorials


## Authors

- [@simneese](https://github.com/simneese)
