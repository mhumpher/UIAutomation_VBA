VERSION 0.4a CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIA_Wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'INTRODUCTION#############################################################################################
'This class implements functions to help with the Microsoft UI Automation
'Author: Michael Humpherys
'Last Updated: 6/12/2017
'Version: 0.6a
'Requirements:
'   UIAuotmationClient (UIAutomationCore.dll) (this is the core UI dll)
'   Microsoft Scripting Runtime (scrrun.dll) (this is mainly for using dictionaries)
'   AutoItX3 1.0 Type Library (AutoItX3.dll) (additional automation capabilities)
'   Miscrosoft VBScript Regular Expression (vbscript.dll\3)
'Notes:
'   This class is based off of the UIAWrapper.au3 script created by junkew on AutoIT Forum
'   https://www.autoitscript.com/forum/topic/153520-iuiautomation-ms-framework-automate-chrome-ff-ie/
'TO DO:
'   1) Add documentation
'   2) Add error checking/handling to various functions
'   3) clear up references to dictionary in getFirstObjectOfElement and getObjectByFindFirst
'   4) Add normalizeExpression function for input string to get Object functions
'   5) Figure out how to get handle to use with AutoIt
'       -could get title with UIA, use title with AIt to get handle, etc.
'   6) Decide whether some attributes/methods need to be private or not
'   xxx7) Add a parent search for properties like find object (FindAll and FindFirst don't work on Ancestors)
'
'Contents:
'=======================================================================================================
'ATTRIBUTES
'Public UIA_oUIAutomation As New CUIAutomation
'Public UIA_oDesktop As IUIAutomationElement
'Public UIA_oTW As IUIAutomationTreeWalker
'Public UIA_oTRUECondition As IUIAutomationCondition
'
'Private UIA_oAutoIT As New AutoItX3
'Public UIA_dPropertiesSupportedArray As New Scripting.Dictionary
'Public UIA_dControlArray As New Scripting.Dictionary
'
'Public UIA_iDefaultWaitTime As Integer
'Private Const iUIA_MAXDEPTH As Integer = 25
'Private Const iUIA_tryMAX As Integer = 3
'=======================================================================================================
'METHODS
'Private Sub class_initialize()
'Function getPropertyValue(oUIElement As IUIAutomationElement, id As Long)
'Function getPropertyFromName(oUIElement As IUIAutomationElement, sPropName As String)
'Function getPropertyFromName(oUIElement As IUIAutomationElement, sPropName As String)
'Function getAllPropertyValues(oUIElement As IUIAutomationElement)
'Function getFirstObjectOfElement(oUIElement As IUIAutomationElement, sPropDesc As String, vTreeScope As Variant)
'Function getObjectByFindFirst(oUIElement As IUIAutomationElement, sPropDesc As String, vTreeScope As Variant)
'Function ActOnElemnt(oElement2ActOn As IUIAutomationElement, sAction As String, Optional p1 As Variant = "", Optional p2 As Variant = "")
'Function getParentElement(oUIElement As IUIAutomationElement, iParentLevel)
'Sub drawLine(x0 As Long, y0 As Long, x1 As Long, y1 As Long)
'Sub drawRect(x0 As Long, y0 As Long, x1 As Long, y1 As Long)
'Sub drawBoundingRect(oUIElement As IUIAutomationElement)
'########################################################################################################

Option Explicit

'DECLARE FUNCTIONS#######################################################################################
'========================================================================================================
'Sleep function is useful as part of automation
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If
'=========================================================================================================
'Additional Functions to help draw lines on screen to identify location of UI Elements
Private Declare Function LineTo _
Lib "gdi32.dll" _
(ByVal handle As Long, ByVal x As Long, ByVal y As Long) As Integer

Private Declare Function MoveToEx _
Lib "gdi32.dll" _
(ByVal handle As Long, ByVal x As Integer, ByVal y As Integer, ByVal lppoint As Long) As Integer

Private Declare Function GetDC _
Lib "user32" _
(ByVal hWnd As Long) As Long

Private Declare Function CreatePen _
Lib "gdi32" _
(ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Private Declare Function SelectObject _
Lib "gdi32" _
(ByVal hdc As Long, ByVal hObject As Long) As Long
'==========================================================================================================
'###########################################################################################################

'ATTRIBUTES###########################################################################################
'Attributes will begin with UIA_
'Implements the UIAutomation Framework
Public UIA_oUIAutomation As New CUIAutomation
Public UIA_oDesktop As IUIAutomationElement
'Public UIA_oUIElement As IUIAutomationElement 'Since this is specific element shouldn't be part of class
Public UIA_oTW As IUIAutomationTreeWalker
Public UIA_oTRUECondition As IUIAutomationCondition 'may need to set to object if not correct

'Include AutoIT Object to assit with automation
'Private because don't want user messing with it
Private UIA_oAutoIT As New AutoItX3

'Tables holding lookup values
Public UIA_dPropertiesSupportedArray As New Scripting.Dictionary
Public UIA_dControlArray As New Scripting.Dictionary

'Constants
Public UIA_iDefaultWaitTime As Integer
Private Const iUIA_MAXDEPTH As Integer = 25
Private Const iUIA_tryMAX As Integer = 3
'###########################################################################################################

'METHODS####################################################################################################
'===========================================================================================================
'Function: class_initialize
'Purpose: VBA internal subroutine called when class is initially declared. Initializes attributes of class.
'Parameters: None
'Returns: None
'Notes: VBA interal subroutine
'===========================================================================================================
Private Sub class_initialize()
    UIA_iDefaultWaitTime = 200
    Set UIA_oDesktop = UIA_oUIAutomation.GetRootElement 'Gets the desktop element
    Set UIA_oTW = UIA_oUIAutomation.RawViewWalker
    'may need to check that this worked
    Set UIA_oTRUECondition = UIA_oUIAutomation.CreateTrueCondition
    
    'Fill out property dictionary
    UIA_dPropertiesSupportedArray.CompareMode = vbTextCompare 'Set the key lookup to be case insensitive
    With UIA_dPropertiesSupportedArray
        '.Add "indexrelative", UIA_SpecialProperty                                                ' Special propertyname
        '.Add "index", UIA_SpecialProperty                                                            ' Special propertyname
        '.Add "instance", UIA_SpecialProperty                                                     ' Special propertyname
        .Add "title", UIA_NamePropertyId                                        ' Alternate propertyname
        .Add "text", UIA_NamePropertyId                                         ' Alternate propertyname
        .Add "regexptitle", UIA_NamePropertyId                                  ' Alternate propertyname
        .Add "class", UIA_ClassNamePropertyId                                    ' Alternate propertyname
        .Add "regexpclass", UIA_ClassNamePropertyId                          ' Alternate propertyname
        .Add "iaccessiblevalue", UIA_LegacyIAccessibleValuePropertyId            ' Alternate propertyname
        .Add "iaccessiblechildId", UIA_LegacyIAccessibleChildIdPropertyId        ' Alternate propertyname
        .Add "id", UIA_AutomationIdPropertyId                                   ' Alternate propertyname
        .Add "handle", UIA_NativeWindowHandlePropertyId                          ' Alternate propertyname
        .Add "RuntimeId", UIA_RuntimeIdPropertyId
        .Add "BoundingRectangle", UIA_BoundingRectanglePropertyId
        .Add "ProcessId", UIA_ProcessIdPropertyId
        .Add "ControlType", UIA_ControlTypePropertyId
        .Add "LocalizedControlType", UIA_LocalizedControlTypePropertyId
        .Add "Name", UIA_NamePropertyId
        .Add "AcceleratorKey", UIA_AcceleratorKeyPropertyId
        .Add "AccessKey", UIA_AccessKeyPropertyId
        .Add "HasKeyboardFocus", UIA_HasKeyboardFocusPropertyId
        .Add "IsKeyboardFocusable", UIA_IsKeyboardFocusablePropertyId
        .Add "IsEnabled", UIA_IsEnabledPropertyId
        .Add "AutomationId", UIA_AutomationIdPropertyId
        .Add "ClassName", UIA_ClassNamePropertyId
        .Add "HelpText", UIA_HelpTextPropertyId
        .Add "ClickablePoint", UIA_ClickablePointPropertyId
        .Add "Culture", UIA_CulturePropertyId
        .Add "IsControlElement", UIA_IsControlElementPropertyId
        .Add "IsContentElement", UIA_IsContentElementPropertyId
        .Add "LabeledBy", UIA_LabeledByPropertyId
        .Add "IsPassword", UIA_IsPasswordPropertyId
        .Add "NativeWindowHandle", UIA_NativeWindowHandlePropertyId
        .Add "ItemType", UIA_ItemTypePropertyId
        .Add "IsOffscreen", UIA_IsOffscreenPropertyId
        .Add "Orientation", UIA_OrientationPropertyId
        .Add "FrameworkId", UIA_FrameworkIdPropertyId
        .Add "IsRequiredForForm", UIA_IsRequiredForFormPropertyId
        .Add "ItemStatus", UIA_ItemStatusPropertyId
        .Add "IsDockPatternAvailable", UIA_IsDockPatternAvailablePropertyId
        .Add "IsExpandCollapsePatternAvailable", UIA_IsExpandCollapsePatternAvailablePropertyId
        .Add "IsGridItemPatternAvailable", UIA_IsGridItemPatternAvailablePropertyId
        .Add "IsGridPatternAvailable", UIA_IsGridPatternAvailablePropertyId
        .Add "IsInvokePatternAvailable", UIA_IsInvokePatternAvailablePropertyId
        .Add "IsMultipleViewPatternAvailable", UIA_IsMultipleViewPatternAvailablePropertyId
        .Add "IsRangeValuePatternAvailable", UIA_IsRangeValuePatternAvailablePropertyId
        .Add "IsScrollPatternAvailable", UIA_IsScrollPatternAvailablePropertyId
        .Add "IsScrollItemPatternAvailable", UIA_IsScrollItemPatternAvailablePropertyId
        .Add "IsSelectionItemPatternAvailable", UIA_IsSelectionItemPatternAvailablePropertyId
        .Add "IsSelectionPatternAvailable", UIA_IsSelectionPatternAvailablePropertyId
        .Add "IsTablePatternAvailable", UIA_IsTablePatternAvailablePropertyId
        .Add "IsTableItemPatternAvailable", UIA_IsTableItemPatternAvailablePropertyId
        .Add "IsTextPatternAvailable", UIA_IsTextPatternAvailablePropertyId
        .Add "IsTogglePatternAvailable", UIA_IsTogglePatternAvailablePropertyId
        .Add "IsTransformPatternAvailable", UIA_IsTransformPatternAvailablePropertyId
        .Add "IsValuePatternAvailable", UIA_IsValuePatternAvailablePropertyId
        .Add "IsWindowPatternAvailable", UIA_IsWindowPatternAvailablePropertyId
        .Add "ValueValue", UIA_ValueValuePropertyId
        .Add "ValueIsReadOnly", UIA_ValueIsReadOnlyPropertyId
        .Add "RangeValueValue", UIA_RangeValueValuePropertyId
        .Add "RangeValueIsReadOnly", UIA_RangeValueIsReadOnlyPropertyId
        .Add "RangeValueMinimum", UIA_RangeValueMinimumPropertyId
        .Add "RangeValueMaximum", UIA_RangeValueMaximumPropertyId
        .Add "RangeValueLargeChange", UIA_RangeValueLargeChangePropertyId
        .Add "RangeValueSmallChange", UIA_RangeValueSmallChangePropertyId
        .Add "ScrollHorizontalScrollPercent", UIA_ScrollHorizontalScrollPercentPropertyId
        .Add "ScrollHorizontalViewSize", UIA_ScrollHorizontalViewSizePropertyId
        .Add "ScrollVerticalScrollPercent", UIA_ScrollVerticalScrollPercentPropertyId
        .Add "ScrollVerticalViewSize", UIA_ScrollVerticalViewSizePropertyId
        .Add "ScrollHorizontallyScrollable", UIA_ScrollHorizontallyScrollablePropertyId
        .Add "ScrollVerticallyScrollable", UIA_ScrollVerticallyScrollablePropertyId
        .Add "SelectionSelection", UIA_SelectionSelectionPropertyId
        .Add "SelectionCanSelectMultiple", UIA_SelectionCanSelectMultiplePropertyId
        .Add "SelectionIsSelectionRequired", UIA_SelectionIsSelectionRequiredPropertyId
        .Add "GridRowCount", UIA_GridRowCountPropertyId
        .Add "GridColumnCount", UIA_GridColumnCountPropertyId
        .Add "GridItemRow", UIA_GridItemRowPropertyId
        .Add "GridItemColumn", UIA_GridItemColumnPropertyId
        .Add "GridItemRowSpan", UIA_GridItemRowSpanPropertyId
        .Add "GridItemColumnSpan", UIA_GridItemColumnSpanPropertyId
        .Add "GridItemContainingGrid", UIA_GridItemContainingGridPropertyId
        .Add "DockDockPosition", UIA_DockDockPositionPropertyId
        .Add "ExpandCollapseExpandCollapseState", UIA_ExpandCollapseExpandCollapseStatePropertyId
        .Add "MultipleViewCurrentView", UIA_MultipleViewCurrentViewPropertyId
        .Add "MultipleViewSupportedViews", UIA_MultipleViewSupportedViewsPropertyId
        .Add "WindowCanMaximize", UIA_WindowCanMaximizePropertyId
        .Add "WindowCanMinimize", UIA_WindowCanMinimizePropertyId
        .Add "WindowWindowVisualState", UIA_WindowWindowVisualStatePropertyId
        .Add "WindowWindowInteractionState", UIA_WindowWindowInteractionStatePropertyId
        .Add "WindowIsModal", UIA_WindowIsModalPropertyId
        .Add "WindowIsTopmost", UIA_WindowIsTopmostPropertyId
        .Add "SelectionItemIsSelected", UIA_SelectionItemIsSelectedPropertyId
        .Add "SelectionItemSelectionContainer", UIA_SelectionItemSelectionContainerPropertyId
        .Add "TableRowHeaders", UIA_TableRowHeadersPropertyId
        .Add "TableColumnHeaders", UIA_TableColumnHeadersPropertyId
        .Add "TableRowOrColumnMajor", UIA_TableRowOrColumnMajorPropertyId
        .Add "TableItemRowHeaderItems", UIA_TableItemRowHeaderItemsPropertyId
        .Add "TableItemColumnHeaderItems", UIA_TableItemColumnHeaderItemsPropertyId
        .Add "ToggleToggleState", UIA_ToggleToggleStatePropertyId
        .Add "TransformCanMove", UIA_TransformCanMovePropertyId
        .Add "TransformCanResize", UIA_TransformCanResizePropertyId
        .Add "TransformCanRotate", UIA_TransformCanRotatePropertyId
        .Add "IsLegacyIAccessiblePatternAvailable", UIA_IsLegacyIAccessiblePatternAvailablePropertyId
        .Add "LegacyIAccessibleChildId", UIA_LegacyIAccessibleChildIdPropertyId
        .Add "LegacyIAccessibleName", UIA_LegacyIAccessibleNamePropertyId
        .Add "LegacyIAccessibleValue", UIA_LegacyIAccessibleValuePropertyId
        .Add "LegacyIAccessibleDescription", UIA_LegacyIAccessibleDescriptionPropertyId
        .Add "LegacyIAccessibleRole", UIA_LegacyIAccessibleRolePropertyId
        .Add "LegacyIAccessibleState", UIA_LegacyIAccessibleStatePropertyId
        .Add "LegacyIAccessibleHelp", UIA_LegacyIAccessibleHelpPropertyId
        .Add "LegacyIAccessibleKeyboardShortcut", UIA_LegacyIAccessibleKeyboardShortcutPropertyId
        '.Add "LegacyIAccessibleSelection", UIA_LegacyIAccessibleSelectionPropertyId                        'Failed in testing - unknown
        .Add "LegacyIAccessibleDefaultAction", UIA_LegacyIAccessibleDefaultActionPropertyId
        .Add "AriaRole", UIA_AriaRolePropertyId
        .Add "AriaProperties", UIA_AriaPropertiesPropertyId
        .Add "IsDataValidForForm", UIA_IsDataValidForFormPropertyId
        .Add "ControllerFor", UIA_ControllerForPropertyId
        .Add "DescribedBy", UIA_DescribedByPropertyId
        .Add "FlowsTo", UIA_FlowsToPropertyId
        .Add "ProviderDescription", UIA_ProviderDescriptionPropertyId
        .Add "IsItemContainerPatternAvailable", UIA_IsItemContainerPatternAvailablePropertyId
        .Add "IsVirtualizedItemPatternAvailable", UIA_IsVirtualizedItemPatternAvailablePropertyId
        .Add "IsSynchronizedInputPatternAvailable", UIA_IsSynchronizedInputPatternAvailablePropertyId
    End With
End Sub

'===========================================================================================================
'Function: getPropertyValue
'Purpose: Get the property value from the UI Element based on the UIA ID Value
'Parameters:    oUIElement - UI Element from which to retrieve property value
'               id - Number from UIA Enumeration for property
'Returns:   Success: Propery value. If array, values separated by ";"
'           Failure: Empty string
'Notes:
'   If id is not in enumeration, will produce error. Need to add error handling.
'===========================================================================================================
Function getPropertyValue(oUIElement As IUIAutomationElement, id As Long)
    Dim vProp As Variant
    Dim sTemp As Variant
    Dim i As Long
    
    vProp = oUIElement.GetCurrentPropertyValue(id)
    If IsArray(vProp) Then
        sTemp = ""
        For i = 0 To UBound(vProp)
            sTemp = sTemp & Trim(vProp(i))
            If i <> UBound(vProp) Then
                'sTemp = sTemp & ";"
                sTemp = sTemp & "|"
            End If
        Next
    Else
        sTemp = vProp
    End If
    getPropertyValue = sTemp
End Function

'===========================================================================================================
'Function: getPropertyFromName
'Purpose: Use name of property rather than id number to get property value.
'Parameters:    oUIElement - UI Element from which to retrieve property value
'               sPropName - Name from property dictionary to get value for
'Returns:   Success: Propery value. If array, values separated by ";"
'           Failure: Empty string
'Notes: Add check to see if sPropName is in the dictionary
'===========================================================================================================
Function getPropertyFromName(oUIElement As IUIAutomationElement, sPropName As String)

    getPropertyFromName = getPropertyValue(oUIElement, UIA_dPropertiesSupportedArray(sPropName))

End Function

'===========================================================================================================
'Function: getAllPropertyValues
'Purpose: Get a dump of all of the property values in a string formatted as
'           Property Name1:= <Property Value1>
'           Property Name2:= <Property Value2>
'           ...
'Parameters:    oUIElement - UI Element from which to retrieve property values
'Returns:   Success: Formatted string of all properties values support by this class
'           Failure:
'Notes:
'===========================================================================================================
Function getAllPropertyValues(oUIElement As IUIAutomationElement)
    Dim vDictKey As Variant
    Dim sTemp As String
    
    sTemp = ""
    For Each vDictKey In UIA_dPropertiesSupportedArray
        sTemp = sTemp & vDictKey & ":= <" & getPropertyValue(oUIElement, UIA_dPropertiesSupportedArray(vDictKey)) & ">" & vbCrLf
    Next
    getAllPropertyValues = sTemp
End Function

'===========================================================================================================
'Function: getBasicPropertyValues
'Purpose: Get a dump of all of the property values in a string formatted as
'           Property Name1:= <Property Value1>;Property Name2:= <Property Value2>...
'Parameters:    oUIElement - UI Element from which to retrieve property values
'Returns:   Success: Formatted string of all properties values support by this class
'           Failure:
'Notes:
'===========================================================================================================
Function getBasicPropertyValues(oUIElement As IUIAutomationElement)
    Dim sBaseProps As String
    
    sBaseProps = "title:=" & getPropertyFromName(oUIElement, "title") & ";" & _
        "class:=" & getPropertyFromName(oUIElement, "class") & ";" & _
        "NativeWindowHandle:=" & Hex(getPropertyFromName(oUIElement, "NativeWindowHandle")) & ";" & _
        "BoundingRectangle:=" & getPropertyFromName(oUIElement, "BoundingRectangle") & ";"
    getBasicPropertyValues = sBaseProps

End Function
'This is a very simple way of retrieving an element. It only allows for one property
'to be searched for, but this still can be very powerful and helpful.
'Note: TreeScope_Ancestors is not a valid scope parameter - https://msdn.microsoft.com/en-us/library/windows/desktop/ee696029(v=vs.85).aspx
'===========================================================================================================
'Function: getFirstObjectOfElement
'Purpose: Get the first UI Element related to the passed element by the TreeScope with the passed property
'           description.
'Parameters:    oUIElement - The root UI Element from which to look
'               sPropDesc - A string with the property to look for in the UI element. Expected to be formatted
'                   as "property name:=property value." If no ":=" included than assume this is the title of
'                   UI element.
'Returns:   Success - First UI Element with the property value in sPropDesc
'           Failure - ""
'Notes:
'This is a very simple way of retrieving an element. It only allows for one property
'to be searched for, but this still can be very powerful and helpful.
'TreeScope_Ancestors is not a valid scope parameter - https://msdn.microsoft.com/en-us/library/windows/desktop/ee696029(v=vs.85).aspx
'===========================================================================================================
Function getFirstObjectOfElement(oUIElement As IUIAutomationElement, sPropDesc As String, vTreeScope As Variant)
    Dim aProps() As String
    Dim vPropID As Variant
    Dim vTempValue As Variant
    Dim vDictKey As Variant
    Dim oCondition As IUIAutomationCondition
    Dim iTry As Integer
    Dim oSearchUIElement As IUIAutomationElement
    
    aProps = Split(sPropDesc, ":=")
    
    'If property not identified then assume name/title
    If UBound(aProps) = 0 Then
        vPropID = UIA_NamePropertyId
        vTempValue = sPropDesc
    Else
        'Dictionary is set to be case insenstive so this can be rewritten - MGH
        For Each vDictKey In UIA_dPropertiesSupportedArray
            If LCase(CStr(vDictKey)) = LCase(aProps(0)) Then
                vPropID = UIA_dPropertiesSupportedArray(vDictKey)
                
                'Not sure this is correctly set up
                'aProps is a string and probably won't include a number
                'For my purposes, I don't generally look for an element based
                'on the control type id
                If vPropID = UIA_ControlTypePropertyId Then
                    vTempValue = CLng(aProps(1))
                Else
                    vTempValue = aProps(1)
                End If
            End If
        Next
    End If
    
    Set oCondition = UIA_oUIAutomation.CreatePropertyCondition(vPropID, vTempValue)
    iTry = 1
    
    While oSearchUIElement Is Nothing And iTry <= iUIA_tryMAX
        Set oSearchUIElement = oUIElement.FindFirst(vTreeScope, oCondition)
        If oSearchUIElement Is Nothing Then
            Sleep (100)
            iTry = iTry + 1
        End If
    Wend
    
    getFirstObjectOfElement = oSearchUIElement
End Function

'===========================================================================================================
'Function: getObjectByFindFirst
'Purpose: Get the first UI Element related to the passed element by the TreeScope with the passed property
'           description.
'Parameters:    oUIElement - The root UI Element from which to look
'               sPropDesc - A string with the properties to look for in the UI element. Expected to be formatted
'                   as "property1:=value1;property2:=value2" If no ":=" included than assume this is the title of
'                   UI element.
'Returns:   Success - First UI Element with the property value in sPropDesc
'           Failure - ""
'Notes: This is a more flexible version of the getFirstObjectOfElement
'This method uses the FindFirst method of UI Element, but instead of using one property to create
'the condition it creates an AND condition to search on multiple properties.
'If sPropDesc is "", then the function returns the first UI Element in the vTreeScope
'===========================================================================================================
Function getObjectByFindFirst(oUIElement As IUIAutomationElement, sPropDesc As String, vTreeScope As Variant)
    Dim aProps() As String
    Dim vPropID As Variant
    Dim vTempValue As Variant
    Dim vDictKey As Variant
    Dim oCondition As IUIAutomationCondition
    Dim oFinalCondition As IUIAutomationCondition
    Dim iTry As Integer
    Dim oSearchUIElement As IUIAutomationElement
    Dim i As Integer
    
    Set oFinalCondition = UIA_oTRUECondition
    aProps = Split(sPropDesc, ";")
    For i = 0 To UBound(aProps)
        Dim aProp() As String
        aProp = Split(aProps(i), ":=")
        'If property not identified then assume name/title
        If UBound(aProp) = 0 Then
            vPropID = UIA_NamePropertyId
            vTempValue = aProp(0)
        Else
            For Each vDictKey In UIA_dPropertiesSupportedArray
                If LCase(CStr(vDictKey)) = LCase(aProp(0)) Then
                    vPropID = UIA_dPropertiesSupportedArray(vDictKey)
                
                    'Not sure this is correctly set up
                    'aProps is a string and probably won't include a number
                    'For my purposes, I don't generally look for an element based
                    'on the control type id
                    If vPropID = UIA_ControlTypePropertyId Then
                        vTempValue = CLng(aProp(1))
                    Else
                        vTempValue = aProp(1)
                    End If
                End If
            Next
        End If
        Set oCondition = UIA_oUIAutomation.CreatePropertyCondition(vPropID, vTempValue)
        Set oFinalCondition = UIA_oUIAutomation.CreateAndCondition(oCondition, oFinalCondition)
    Next
        
    iTry = 1
    
    While oSearchUIElement Is Nothing And iTry <= iUIA_tryMAX
        Set oSearchUIElement = oUIElement.FindFirst(vTreeScope, oFinalCondition)
        If oSearchUIElement Is Nothing Then
            Sleep (100)
            iTry = iTry + 1
        End If
    Wend
    
    getObjectByFindFirst = oSearchUIElement
End Function

Function getObjectByRegEx(oUIElement As IUIAutomationElement, sPropDesc As String, vTreeScope As Variant)
    Dim aProps() As String
    Dim vPropID As Long
    Dim vTempValue As Variant
    Dim vDictKey As Variant
    Dim oUIElemArr As IUIAutomationElementArray
    Dim iTry As Integer
    Dim oTempUIElement As IUIAutomationElement
    Dim oFoundUIElement As IUIAutomationElement
    Dim iNumElems As Long
    Dim i As Integer
    Dim j As Long
    Dim RE As New RegExp
    Dim sActProp As String
    Dim bFoundElem As Boolean
    
    Set oUIElemArr = oUIElement.FindAll(vTreeScope, UIA_oTRUECondition)
    Set oFoundUIElement = Nothing
    
    iNumElems = oUIElemArr.Length
    
    aProps = Split(sPropDesc, ";")
    For j = 0 To iNumElems - 1 'Array is zero-indexed
        Set oTempUIElement = oUIElemArr.GetElement(j)
        Debug.Print getBasicPropertyValues(oTempUIElement)
        For i = 0 To UBound(aProps)
            Dim aProp() As String
            aProp = Split(aProps(i), ":=")
            'If property not identified then assume name/title
            If UBound(aProp) = 0 Then
                vPropID = UIA_NamePropertyId
                vTempValue = aProp(0)
            Else
                vPropID = UIA_dPropertiesSupportedArray(aProp(0))
                vTempValue = aProp(1)
            End If
            
            RE.Pattern = vTempValue
            RE.IgnoreCase = True
            
            sActProp = getPropertyValue(oTempUIElement, vPropID)
            If Not RE.test(sActProp) Then
                'Set oTempUIElement = Nothing
                bFoundElem = False
                GoTo NextElem:
            ElseIf i = UBound(aProps) Then
                bFoundElem = True
            End If
        Next
NextElem:
        If bFoundElem Then
            Set oFoundUIElement = oTempUIElement
            Exit For
        End If
    Next
    getObjectByRegEx = oFoundUIElement
    
    
End Function

'===========================================================================================================
'Function: ActOnElement
'Purpose: Perform some action on the passed UI Element
'Parameters:    oElement2ActOn - UI Element to perform action on
'               sAction - string name of the action to perform
'               p1,p2 - Auxiliary variables to define action when needed (Optional)
'Returns:
'Notes: Supported Actions
'...
'Control Pattern Ids: https://msdn.microsoft.com/en-us/library/windows/desktop/ee671195(v=vs.85).aspx
'===========================================================================================================

Function ActOnElement(oElement2ActOn As IUIAutomationElement, sAction As String, Optional p1 As Variant = "", Optional p2 As Variant = "")
    Dim controlType As Variant
    Dim hWnd As Variant
    Dim vRetValue As Variant
    
    Dim tInvokePat As IUIAutomationInvokePattern
    Dim tWinPat As IUIAutomationWindowPattern
    Dim tTransPat As IUIAutomationTransformPattern
    Dim tValuePat As IUIAutomationValuePattern
    Dim tScrollPat As IUIAutomationScrollPattern

    controlType = getPropertyValue(oElement2ActOn, UIA_ControlTypePropertyId)
    
    'May need to define the appropriate pattern type for each case within that case
    'Variant doesn't seem to work
    Select Case sAction
        Case Is = "click" 'simple for now, include more latter
            Dim aTemp As Variant
            Dim x As Integer
            Dim y As Integer
            aTemp = Split(getPropertyValue(oElement2ActOn, UIA_BoundingRectanglePropertyId), "|")
            x = CInt(aTemp(0) + (aTemp(2) / 2))
            y = CInt(aTemp(1) + (aTemp(3) / 2))
            
            'UIA Wrapper uses AutoIt to implement the mouse click
            UIA_oAutoIT.MouseMove x, y, 0
            UIA_oAutoIT.MouseClick "LEFT", x, y, 1, 0 'always assuming one left click
            
        Case Is = "setValue"
            'Dim tPattern As IUIAutomationValuePattern
            If controlType = UIA_WindowControlTypeId Then
                'hwnd = oElement2ActOn.CurrentNativeWindowHandle
                'uses AutoIt to set Window Title
                'UIA_oAutoIT.WinSetTitle hwnd, "", p1
            Else
                oElement2ActOn.SetFocus
                Sleep (UIA_iDefaultWaitTime)
                
                'Don't deep with the UIA_LegacyIAccessiblePatternId
                Set tValuePat = oElement2ActOn.GetCurrentPattern(UIA_ValuePatternId)
                tValuePat.SetValue (p1)
                
            End If
        
        'Case Is = "setValue using keys"
        'Case Is = "setValue using clipboard"
        Case Is = "getValue"
            oElement2ActOn.SetFocus
            UIA_oAutoIT.Send ("^a")
            UIA_oAutoIT.Send ("^c")
            vRetValue = UIA_oAutoIT.ClipGet
        
        Case Is = "sendkeys"
            oElement2ActOn.SetFocus
            UIA_oAutoIT.Send (p1)
            
        Case Is = "invoke"
            'Dim tPattern As IUIAutomationInvokePattern
            oElement2ActOn.SetFocus
            Set tInvokePat = oElement2ActOn.GetCurrentPattern(UIA_InvokePatternId)
            tInvokePat.Invoke
        Case Is = "close"
            'Dim tPattern As IUIAutomationWindowPattern
            Set tWinPat = oElement2ActOn.GetCurrentPattern(UIA_WindowPatternId)
            tWinPat.Close
        Case Is = "move"
            'Dim tPattern As IUIAutomationTransformPattern
            Set tTransPat = oElement2ActOn.GetCurrentPattern(UIA_TransformPatternId)
            tTransPat.Move p1, p2
        Case Is = "resize"
            'Dim tPattern As IUIAutomationWindowPattern
            'Dim tPattern2 As IUIAutomationTransformPattern
            
            Set tWinPat = oElement2ActOn.GetCurrentPattern(UIA_WindowPatternId)
            tWinPat.SetWindowVisualState (WindowVisualState_Normal)
            
            Set tTransPat = oElement2ActOn.GetCurrentPattern(UIA_TransformPatternId)
            tTransPat.Resize p1, p2
        Case Is = "minimize"
            'Dim tPattern As IUIAutomationWindowPattern
            Set tWinPat = oElement2ActOn.GetCurrentPattern(UIA_WindowPatternId)
            tWinPat.SetWindowVisualState (WindowVisualState_Minimized)
        Case Is = "maximize"
            'Dim tPattern As IUIAutomationWindowPattern
            Set tWinPat = oElement2ActOn.GetCurrentPattern(UIA_WindowPatternId)
            tWinPat.SetWindowVisualState (WindowVisualState_Maximized)
        Case Is = "normal"
            'Dim tPattern As IUIAutomationWindowPattern
            Set tWinPat = oElement2ActOn.GetCurrentPattern(UIA_WindowPatternId)
            tWinPat.SetWindowVisualState (WindowVisualState_Normal)
        
        'Scroll Options - https://msdn.microsoft.com/en-us/library/system.windows.automation.scrollamount(v=vs.110).aspx
        'Moves scroll as if page down or up
        Case Is = "page down"
            Set tScrollPat = oElement2ActOn.GetCurrentPattern(UIA_ScrollPatternId)
            tScrollPat.Scroll ScrollAmount_NoAmount, ScrollAmount_LargeIncrement
        Case Is = "page up"
            Set tScrollPat = oElement2ActOn.GetCurrentPattern(UIA_ScrollPatternId)
            tScrollPat.Scroll ScrollAmount_NoAmount, ScrollAmount_LargeDecrement
            
        'Moves scroll as if pushing arrow key or click error button on scrollbar
        Case Is = "scroll down"
            Set tScrollPat = oElement2ActOn.GetCurrentPattern(UIA_ScrollPatternId)
            tScrollPat.Scroll ScrollAmount_NoAmount, ScrollAmount_SmallIncrement
        Case Is = "scroll up"
            Set tScrollPat = oElement2ActOn.GetCurrentPattern(UIA_ScrollPatternId)
            tScrollPat.Scroll ScrollAmount_NoAmount, ScrollAmount_SmallDecrement
        Case Else
        
        
    End Select
    
    ActOnElement = vRetValue
End Function

'===========================================================================================================
'Function: getParentElement
'Purpose: Get the ancestor fo the passed UI Element
'Parameters:    oUIElement - UI Element from which to find parent UI Element
'               iParent Level - 1 = Parent, 2 = grandparent, 3 = ggrandparent, etc.
'Returns:   Sucess: UI Element parent of oUIElement
'           Failure: Nothing
'Notes:
'If iParent is < 1, then returns immediate parent
'add handling for this case to return Nothing!
'===========================================================================================================
Function getParentElement(oUIElement As IUIAutomationElement, iParentLevel)
    Dim i As Integer
    Dim oUIParent As IUIAutomationElement
    
    Set oUIParent = UIA_oTW.getParentElement(oUIElement)
    For i = 2 To iParentLevel
        Set oUIParent = UIA_oTW.getParentElement(oUIParent)
    Next

    getParentElement = oUIParent

End Function

'===========================================================================================================
'Function: getParentElement
'Purpose: Get the ancestor fo the passed UI Element
'Parameters:    oUIElement - UI Element from which to find parent UI Element
'               sProps - String of parent properties to search for
'Returns:   Sucess: UI Element parent of oUIElement
'           Failure: Nothing
'Notes:
'
'===========================================================================================================
Function getParentByRegEx(oUIElement As IUIAutomationElement, sProps As String)
    Dim i As Integer
    Dim oUIParent As IUIAutomationElement
    Dim oFoundElement As IUIAutomationElement
    
    Set oUIParent = UIA_oTW.getParentElement(oUIElement)
    While oFoundElement Is Nothing And getPropertyFromName(oUIParent, "title") <> "Desktop"
        Set oFoundElement = getObjectByRegEx(oUIParent, sProps, TreeScope_Element)
        Set oUIParent = UIA_oTW.getParentElement(oUIParent)
    Wend
    getParentByRegEx = oFoundElement

End Function

'===========================================================================================================
'Function: drawLine
'Purpose: Draw a red line on desktop
'Parameters:    x0, y0 - coordinates for start of line
'               x1, y1 - coordinates for end of line
'Returns: N/A
'Notes:
'Const PS_SOLID = 0
'Const PS_DASH = 1
'Const PS_DOT = 2
'Const PS_DASHDOT = 3
'Const PS_DASHDOTDOT = 4
'Const PS_NULL = 5
'Const PS_INSIDEFRAME = 6
'http://www.jasinskionline.com/windowsapi/ref/c/createpen.html
'===========================================================================================================
Sub drawLine(x0 As Long, y0 As Long, x1 As Long, y1 As Long)
    Dim hdc As Variant
    Dim pen As Variant
    Dim res As Variant

    hdc = GetDC(0)
    'pen = CreatePen(PS_SOLID, 4, RGB(255, 0, 0))
    pen = CreatePen(0, 4, RGB(255, 0, 0))
    res = SelectObject(hdc, pen)
    res = MoveToEx(hdc, x0, y0, 0)
    res = LineTo(hdc, x1, y1)

End Sub

'===========================================================================================================
'Function: drawRect
'Purpose: Draw a red rectangle
'Parameters:    x0, y0 - coordinates for top left of rectangle
'               x1, y1 - coordinates for bottom right of rectangle
'Returns: N/A
'Notes:
'===========================================================================================================
Sub drawRect(x0 As Long, y0 As Long, x1 As Long, y1 As Long)
    drawLine x0, y0, x1, y0 'top
    drawLine x0, y0, x0, y1 'left
    drawLine x1, y0, x1, y1 'right
    drawLine x0, y1, x1, y1 'bottom
End Sub

'===========================================================================================================
'Function: drawBoundingRect
'Purpose: Draw a rectangle around the current UI Element
'Parameters:    oUIElement - UI Element around which to draw rectangle
'Returns: N/A
'Notes:
'===========================================================================================================
Sub drawBoundingRect(oUIElement As IUIAutomationElement)
    Dim sProp As String
    Dim aBoundRect() As String
    Dim x0 As Long
    Dim y0 As Long
    Dim x1 As Long
    Dim y1 As Long
    
    sProp = getPropertyFromName(oUIElement, "BoundingRectangle")
    'aBoundRect() = Split(sProp, ";")
    aBoundRect() = Split(sProp, "|")
    x0 = CLng(aBoundRect(0)) - 1
    y0 = CLng(aBoundRect(1)) - 1
    x1 = CLng(aBoundRect(2)) + CLng(aBoundRect(0)) + 1
    y1 = CLng(aBoundRect(3)) + CLng(aBoundRect(1)) + 1
    drawRect x0, y0, x1, y1

End Sub
