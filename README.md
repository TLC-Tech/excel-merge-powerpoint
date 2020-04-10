# Excel Merge (to) PowerPoint
## Description
* This is a VBA Macro that should be run with an excel file open, which has at least 2 columns of data.
* When the Macro is run, it will open PowerPoint and construct a series of slides based on each row of data from the first 2 columns.
* The slide title is what's in column 1, and the slide text is what's in column 2.
* The text is centered in the slide to look more like flash cards
* The background of the slides is light grey to avoid eye strain from a white background.
* The PowerPoint project is NOT saved automatically, it's up to the User to save the ppt after its generated.
## How to Install
1. Add `excel-merge-powerpoint.xla` to `C:\Users\YOUR_USERNAME\AppData\Roaming\Microsoft\AddIns\`.
2. Open Excel and start a new workbook
3. File > Options > Add Ins > Go
4. Check `Excel Merge Powerpoint` and press OK
5. Alt + F11 to open VBA IDE
6. Tools > References
7. Check `Microsoft PowerPoint 16.0 Object Library` and press OK
8. Exit the VBA IDE
9. File > Options > Customize Ribbon
10. New Tab. Find the new tab, right click, rename "TLC"
11. Within the TLC tab, New Group.  Find the new group, right click, rename "Merges"
12. Left side, filter by Macros, click on `GeneratePowerPoint` and click Add >> to add it to the new ribbon tab group
13. Right click the GeneratePowerPoint, rename "PowerPoint" with a logo that looks like powerpoint slides
14. Press OK and get out of the menus
## How to Use
Open the excel sheet you want to merge into powerpoint, and press `PowerPoint` button from the `TLC` tab in the excel function ribbon
### Usage Notes
* The PowerPoint project is NOT saved automatically, it's up to the User to save the ppt after its generated.
* If you stop the process in the middle of its activity, you'll get a VBA error message. Press `End` and NOT `Debug`.
## How to Edit & Distribute
See [this webpage](https://www.ozgrid.com/VBA/excel-add-in-create.htm) for more info.  TLC has no affiliation with the site or its owner.
## License
MIT
