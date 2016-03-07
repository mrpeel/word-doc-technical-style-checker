# Word Technical Document Style Checker
The recently released Mapbox retext project (https://github.com/mapbox/retext-mapbox-standard) includes a series of useful resources to simplify the language in technical docuemntation.  I often need to write technical documents using MS Word and wanted similar funcionality for my Word documents.

This project has an MS Word VBA module to check the style of technical documents.  It checks for three categories of style improvement:
* words which generally add no meaning
* words which shouldn't start a sentence
* complex words which could be simplified

For each potential improvement found, a comment is added to the document.

#Installation
1. Enable the developer ribbon:
     Word Options -> Customize Ribbon -> Make sure "Developer" is checked
2. On the main screen go to the "Developer" ribbon and click "visual Basic".
3. Under Normal -> Modules, right-click and select "Import File"
4. A module called "techDocStyleChecker" will appear

#Running
Go to the "techDocStyleChecker" module and run "checkStyle()".  

###Assigning to a Hotkey
1. Word Options -> Customize Ribbon
2. Find "Keyboard shortcuts" and click the "Customize" button
3. In "Categories" select "Macros"
3. In "Macros" select "checkStyle"
4. Click into "Press new shortcut key"
5. Press key combination
6. Click "Assign"

###Adding to Quick Access Toolbar
1. Word Options -> Quick Access Toolbar
2. Set "Choose commands from" to "Macros"
3. Add "Normal.techDocStyleChcker.checkStyle()" to Quick access
"checkStyle " can be assigned to a hot-key, or a menu or 

###Adding to Ribbon
1. Word Options -> Customize Ribbon
2. Set "Choose commands from" to "Macros"
3. Choose tab, and click "New Group"
3. Add "Normal.techDocStyleChcker.checkStyle()" to new group

#Configuration and customisation
No configuration is necessary to run the style checker.  The terms used in each of the categories are set directly in the code, so they can be customised as required:
  - setupWordsToRemove()
  - setupSentenceStartWordsToRemove()
  - setupWordSimplifications()
