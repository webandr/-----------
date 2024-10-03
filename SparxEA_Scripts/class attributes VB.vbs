option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Example illustrating how to add and delete Attributes and Methods.
' 
' NOTE: Requires an element to be selected in the Project Browser
' 
' Related APIs
' =================================================================================
' Element API - http://www.sparxsystems.com/enterprise_architect_user_guide/12.1/automation_and_scripting/element2.html
' Attribute API - http://www.sparxsystems.com/enterprise_architect_user_guide/12.1/automation_and_scripting/attribute.html
' Method API - http://www.sparxsystems.com/enterprise_architect_user_guide/12.1/automation_and_scripting/method.html
'
sub ManageAttributesMethodsExample()

	' Show the script output window
	Repository.EnsureOutputVisible "Script"
	
	' Get the currently selected element in the tree to work on
	dim theElement as EA.Element
	set theElement = Repository.GetTreeSelectedObject()
	
	if not theElement is nothing and theElement.ObjectType = otElement then
	
		dim i
	
		Session.Output( "VBScript MANAGE ATTRIBUTES/METHODS EXAMPLE" )
		Session.Output( "=======================================" )
		Session.Output( "Working on element '" & theElement.Name & "' (Type=" & theElement.Type & _
			", ID=" & theElement.ElementID & ")" )
			
		' ==================================================
		' MANAGE ATTRIBUTES
		' ==================================================
		' Add an attribute
		dim attributes as EA.Collection
		set attributes = theElement.Attributes
		
		dim newAttribute as EA.Attribute
		set newAttribute = attributes.AddNew( "m_newAttribute", "string" )
		newAttribute.Update()
		attributes.Refresh()
		
		Session.Output( "Added attribute: " & newAttribute.Name )
		
		set newAttribute = nothing
		
		' Search the attribute collection for the added attribute and delete it
		for i = 0 to attributes.Count - 1
			dim currentAttribute as EA.Attribute
			set currentAttribute = attributes.GetAt( i )
			
			Session.Output( "Attribute: " & currentAttribute.Name )
			
			' Delete the attribute we just added
			if currentAttribute.Name = "m_newAttribute" then
				attributes.DeleteAt i, false
				Session.Output( "Deleted Attribute: " & currentAttribute.Name )
			end if
		next
		
		set attributes = nothing

		' ==================================================
		' MANAGE METHODS
		' ==================================================
		' Add a method
		dim methods as EA.Collection
		set methods = theElement.Methods
		
		dim newMethod as EA.Method
		set newMethod = methods.AddNew( "NewMethod", "int" )
		newMethod.Update()
		methods.Refresh()
		
		Session.Output( "Added method: " & newMethod.Name )
		
		set newMethod = nothing
		
		' Search the method collection for the added method and delete it
		for i = 0 to methods.Count - 1
			dim currentMethod as EA.Method
			set currentMethod = methods.GetAt( i )
			
			Session.Output( "Method: " & currentMethod.Name )
			
			' Delete the method we just added
			if currentMethod.Name = "NewMethod" then
				methods.DeleteAt i, false
				Session.Output( "Deleted Method: " & currentMethod.Name )
			end if
		next
		
		set methods = nothing

		Session.Output( "Done!" )
		
	else
		' No item selected in the tree, or the item selected was not an element
		MsgBox( "This script requires an element be selected in the Project Browser." & vbCrLf & _
			"Please select an element in the Project Browser and try again." )
	end if


end sub

ManageAttributesMethodsExample
