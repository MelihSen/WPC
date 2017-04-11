# WPC - Windows Power Control
Power Control .Net Library for Windows Operating Systems. Using this library (WPC.DLL), .Net developers can 
* Shut Down 
* Power Off
* Restart
* Log Off
* Sleep
* Hibernate
* Lock

the system. These functions can be called with no parameter or with advanced parameters (Major reasons, Minor reasons, predefined reasons, Force options, etc.)
All these functions call Windows APIs in a background

## Getting Started
Download "Release" folder of WPC project and copy folder to your .Net project folder (Visual Basic.Net, C#.Net, etc.) Release folder includes WPC.DLL and XML Documentation Comments file which provides explanation of functions and parameters for code editor. Name of "Release" folder can be changed in your project folder.

### Prerequisites
WPC needs Microsoft .Net Framework 4.0

### Referencing and using WPC.DLL
* Right click "References" item under your project in Solution Explorer
* Choose "Add Reference..." item in opened pop-up menu
* Click Browse..." button and choose WPC.DLL in your project directory (WPC Namespace will be added to your project)
* Create new object from WPC.WindowsPowerControl class
* All Power Control Functions will be listed under the created object from WindowsPowerControl class

All Power Control Functions and parameters have "XML Documentation Comments" that they are shown in code editor when you use them

## Version
Inital Version 1.0.0

## Author
* **Melih Senyurt** - *Initial work* - [WPC](https://github.com/melihsen/WPC)

## License
This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details
