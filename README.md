A VBA Class for Excel that implements a wrapper for Microsoft's UI Automation 
(https://msdn.microsoft.com/en-us/library/ms747327(v=vs.110).aspx)

Programmatically identify and manipulate UI elements on the UI tree based on easily identifiable properties. To discover properties of UI element use Microsoft's Inspect.exe (https://msdn.microsoft.com/en-us/library/windows/desktop/dd318521(v=vs.85).aspx) or other utility (e.g. https://github.com/blackrosezy/gui-inspect-tool)

Supported Actions:
click, sendkeys, invoke, close, move, resize, minimize, maximize, normal, page down/up, scroll down/up, setvalue, getvalue

Supported Searches:
- Search with find first with multiple properties
- Search with Regular Expressions
- Get parent by defined levels
- Search parents by Regular Expressions

This class is based off of the UIAWrapper.au3 script created by junkew on AutoIT Forum
https://www.autoitscript.com/forum/topic/153520-iuiautomation-ms-framework-automate-chrome-ff-ie/

To use in Excel:
1) Download UIA_Wrapper.cls and HelperFunctions.bas files
2) Open Excel
3) Go to Visual Basic Editor
4) Select File -> Import File...
5) Select the UIA_Wrapper.cls and HelperFunctions.bas

To Do:
1) Create version without AutoItX - should work on base Excel install
2) Create helper function module for auxilary functions (drawLine, Sleep, MouseClick, etc.) - In Construction
3) Create Examples and demos
  - Notepad, Calculator, search Desktop,...
4) Add breadthFirst and depthFirst tree searches to avoid create of large UI arrays in RegEx search

<a rel="license" href="http://creativecommons.org/licenses/by-sa/4.0/"><img alt="Creative Commons License" style="border-width:0" src="https://i.creativecommons.org/l/by-sa/4.0/88x31.png" /></a><br />This work is licensed under a <a rel="license" href="http://creativecommons.org/licenses/by-sa/4.0/">Creative Commons Attribution-ShareAlike 4.0 International License</a>.
