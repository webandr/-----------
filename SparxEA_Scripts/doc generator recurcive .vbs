option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Demonstrates how the DocumentGenerator API can be used to programmatically generate RTF or HTML
' documentation for model elements
'
' This example generates documentation for all actors under the package currently selected in
' the project browser, and all use cases they are linked to
'
' Related APIs
' =================================================================================
'
dim MAX_UC_DEEP_LEVEL
MAX_UC_DEEP_LEVEL=2

dim ACTOR_TEMPLATE, USECASE_TEMPLATE
ACTOR_TEMPLATE = "Model Report"
' USECASE_TEMPLATE = "My Use Case Details"
USECASE_TEMPLATE = "Use Case Details SimpleConnectors"

dim DOCUMENTATION_TYPE, OUTPUT_FILE
DOCUMENTATION_TYPE = dtRTF
OUTPUT_FILE = "c:\\temp\\DocumentationExample.rtf"

sub DocumentationExample()

	dim i, j

	' Show the script output window
	Repository.EnsureOutputVisible "Script"

	Session.Output "VBScript DOCUMENTATION EXAMPLE"
	Session.Output "======================================="
	
	' Get the currently selected package in the Project Browser
	dim currentPackage as EA.Package
	set currentPackage = Repository.GetTreeSelectedPackage()
	
	if not currentPackage is nothing then
	
		' Create a document generator object
		dim docGenerator as EA.DocumentGenerator
		set docGenerator = Repository.CreateDocumentGenerator()
		
		' Create a new document
		if docGenerator.NewDocument("") = true then
		
			dim generationSuccess
			generationSuccess = false
			
			' Insert table of contents
			docGenerator.InsertText "Table of Contents", alignLeft
			generationSuccess = docGenerator.InsertTableOfContents()
			if generationSuccess = false then
				ReportWarning "Error inserting Table of Contents: " + docGenerator.GetLastError()
			end if
			
			' Insert page break
			docGenerator.InsertBreak( breakPage )
			
			' Iterate over all actors under the currently selected package
			dim packageElements as EA.Collection
			set packageElements = currentPackage.Elements

			dim reportedUCs
			reportedUCs = "!"
			' set reportedUCs = CreateObject( "System.Collections" )
						
			for i = 0 to packageElements.Count - 1
			
				' Get the current element
				dim currentElement as EA.Element;
				set currentElement = packageElements.GetAt( i )
				
				if currentElement.Type = "Actor" then
				
					' Generate Actor documentation
					ReportInfo "Generating documentation for actor: " + currentElement.Name
					generationSuccess = docGenerator.DocumentElement( currentElement.ElementID, 1, ACTOR_TEMPLATE )
					if generationSuccess = false then
						ReportWarning "Error generating Actor documentation: " + docGenerator.GetLastError()
					end if
					
					' Generate documentation for all Use Cases connected to the current actor
					dim elementConnectors as EA.Collection
					set elementConnectors = currentElement.Connectors
					
					for j = 0 to elementConnectors.Count - 1
					
						' Get the current connector and the element that it connects to
						dim currentConnector as EA.Connector
						set currentConnector = elementConnectors.GetAt( j )
						dim connectedElement as EA.Element
						set connectedElement = Repository.GetElementByID( currentConnector.SupplierID )
						set connectedElementCli = Repository.GetElementByID( currentConnector.ClientID )
						

						if not (connectedElement.Type = "UseCase" OR connectedElement.Type = "Collaboration") then
							ReportInfo currentElement.Name + " <====== " + connectedElementCli.Type + " " + connectedElementCli.Name
							set connectedElement = connectedElementCli
						else
							ReportInfo currentElement.Name + " ======> " + connectedElement.Type + " " + connectedElement.Name 
						end if
						
						if (connectedElement.Type = "UseCase" OR connectedElement.Type = "Collaboration") then
						
							' Generate Use Case documentation
							ReportInfo "Generating documentation for connected UseCase: " + connectedElement.Name
							
							
							generationSuccess = docGenerator.DocumentElement( connectedElement.ElementID, 2, USECASE_TEMPLATE )
							if generationSuccess = false then
								ReportWarning "Error generating UseCase documentation: " + docGenerator.GetLastError()
							end if
							
							
							dim ConnectedElementConnectors as EA.Collection
							set ConnectedElementConnectors = connectedElement.Connectors
							dim CEconnectedElement as EA.Element
							dim sstereotype 
							
							
							if ConnectedElementConnectors.Count > 1 then
								Session.Output "The " + connectedElement.name + " has " + CStr(ConnectedElementConnectors.Count) + " connections ."
								for x = 0 to ConnectedElementConnectors.Count - 1
									generateUC reportedUCs, docGenerator, connectedElement, ConnectedElementConnectors, x, 1, MAX_UC_DEEP_LEVEL
								next
								for x = 0 to ConnectedElementConnectors.Count - 1
									generateSubUC reportedUCs, docGenerator, connectedElement, ConnectedElementConnectors, x, 1, MAX_UC_DEEP_LEVEL
								next
							end if 'if ConnectedElementConnectors.Count > 1 then
						
						end if 'if connectedElement.Type = "UseCase" then
						
					next 'packageElements.GetAt( i ).Connectors
				else
					ReportInfo "Skipping element " + currentElement.Name + " - not an actor"
				end if
				
			next
			
			' Save the document
			dim saveSuccess
			saveSuccess = docGenerator.SaveDocument( OUTPUT_FILE, DOCUMENTATION_TYPE )
			
			if saveSuccess = true then
				ReportInfo "Documentation complete!"
			else
				ReportWarning "Error saving file: " + docGenerator.GetLastError()
			end if
		
		else
			ReportFatal "Could not create new document: " + docGenerator.GetLastError()
		end if
	
	else
	
		ReportFatal "This script requires a package to be selected in the Project Browser.\n" &_
			"Please select a package in the Project Browser and try again."
	
	end if
	
	Session.Output "Done!"
	
end sub


function generateUC(ByRef reportedUCs, ByRef docGenerator, ByVal connectedElement, ByVal ConnectedElementConnectors , ByVal x, ByVal level, ByVal levelMax)  

	dim subConnectedElementConnectors as EA.Collection
	dim subCEconnectedElement as EA.Element
	dim sstereotype 

	' Get the current connector and the element that it connects to
	dim currentCEConnector as EA.Connector
	set currentCEConnector = ConnectedElementConnectors.GetAt( x )
	set currentCECElement = Repository.GetElementByID( currentCEConnector.SupplierID )
	set currentCECElementClient = Repository.GetElementByID( currentCEConnector.ClientID )
						
		' Get the current connector and the element that it connects to
		dim sub_currentConnector as EA.Connector
		dim sub_connectedElement as EA.Element
						
	if (currentCECElement.Name = connectedElement.Name) then 'and (colContains(reportedUCs, connectedElement.ElementID) > 0) then
		set CCconnectedElement = Repository.GetElementByID( currentCEConnector.ClientID )
		'Session.Output CCconnectedElement.Name + "<" + CCconnectedElement.Type  + ">" 
		if (CCconnectedElement.Type = "UseCase" OR CCconnectedElement.Type = "Collaboration") then
			Session.Output "{" + currentCEConnector.Stereotype + "}" + CCconnectedElement.Name + " <S-Dest> element = " + connectedElement.Name
			if currentCEConnector.Stereotype = "extend" then 
				sstereotype = "extending" 
			else 
				sstereotype = "Including" 
			end if
			if colContains(reportedUCs, CCconnectedElement.ElementID) = 0 then
				if (currentCEConnector.Direction = "Source -> Destination") then
					' level+2 так как когда level=1 то это значит что в отчет надо вставить Header3
					generationSuccess = docGenerator.DocumentElement( CCconnectedElement.ElementID, level+2, "My " + sstereotype + " UC details" )
				else
					generationSuccess = docGenerator.DocumentElement( CCconnectedElement.ElementID, level+2, "My " + sstereotype + " UC details" )
				end if
				if generationSuccess = false then
					ReportWarning "Error generating UseCase documentation: " + docGenerator.GetLastError()
				else
					'reportedUCs.AddNew CCconnectedElement.Name, "String"
					' reportedUCs = reportedUCs & CCconnectedElement.ElementID & "!"
				end if	
			end if
		end if
	else
		set CCconnectedElement = Repository.GetElementByID( currentCEConnector.SupplierID )
		if (CCconnectedElement.Type = "UseCase" OR CCconnectedElement.Type = "Collaboration") then
			Session.Output "{" + currentCEConnector.Stereotype + "}" + connectedElement.Name + " <D-Source> element = " + currentCECElement.Name
			if currentCEConnector.Stereotype = "extend" then 
				sstereotype = "extended" 
			else 
				sstereotype = "Included" 
			end if
			if colContains(reportedUCs, CCconnectedElement.Name) = 0 then
				if (currentCEConnector.Direction = "Source -> Destination") then
					generationSuccess = docGenerator.DocumentElement( CCconnectedElement.ElementID, level+2, "My " + sstereotype + " UC details" )
				else
					generationSuccess = docGenerator.DocumentElement( CCconnectedElement.ElementID, level+2, "My " + sstereotype + " UC details" )
				end if
				if generationSuccess = false then
					ReportWarning "Error generating UseCase documentation: " + docGenerator.GetLastError()
				else
					'reportedUCs.AddNew CCconnectedElement.Name, "String"
				end if	
			end if
		end if
	end if
	reportedUCs = reportedUCs & connectedElement.ElementID & "!"
	level = level + 1
	generateSubUC reportedUCs, docGenerator, connectedElement, ConnectedElementConnectors, x, level, MAX_UC_DEEP_LEVEL
end function

function generateSubUC(ByRef reportedUCs, ByRef docGenerator, ByVal connectedElement, ByVal ConnectedElementConnectors , ByVal x, ByVal level, ByVal levelMax)  
	
	if (level > levelMax) then 
		exit function 
	end if
	level = level + 1
	
	
	dim subConnectedElementConnectors as EA.Collection
	dim subCEconnectedElement as EA.Element
	dim sstereotype 

	' Get the current connector and the element that it connects to
	dim currentCEConnector as EA.Connector
	set currentCEConnector = ConnectedElementConnectors.GetAt( x )
	set currentCECElement = Repository.GetElementByID( currentCEConnector.SupplierID )
	set currentCECElementClient = Repository.GetElementByID( currentCEConnector.ClientID )

						
		' Get the current connector and the element that it connects to
		dim sub_currentConnector as EA.Connector
		dim sub_connectedElement as EA.Element
						
	if (currentCECElement.Name = connectedElement.Name) then 'and (colContains(reportedUCs, connectedElement.ElementID) > 0) then
		set CCconnectedElement = Repository.GetElementByID( currentCEConnector.ClientID )
		'Session.Output CCconnectedElement.Name + "<" + CCconnectedElement.Type  + ">" 
		if (CCconnectedElement.Type = "UseCase" OR CCconnectedElement.Type = "Collaboration") then
			' recurcive call for sub cases 
			set subConnectedElementConnectors = CCconnectedElement.Connectors	
			if subConnectedElementConnectors.Count > 1 then
				if colContains(reportedUCs, CCconnectedElement.ElementID) = 0 then
					Session.Output "The sub " + CCconnectedElement.name + " has " + CStr(subConnectedElementConnectors.Count) + " connections ."
						sident = ""
						for i = 2 to level
							sident = sident & ">>>>> "
						next
						docGenerator.InsertText sident & " Составляющие ВИ: " & CCconnectedElement.Name, alignLeft
						generationSuccess = docGenerator.InsertBreak(1)
					for z = 0 to subConnectedElementConnectors.Count - 1
						generateUC reportedUCs, docGenerator, CCconnectedElement, subConnectedElementConnectors, z, level, levelMax			
					next
				end if
			end if
			reportedUCs = reportedUCs & connectedElement.ElementID & "!"
		end if
	else
		set CCconnectedElement = Repository.GetElementByID( currentCEConnector.SupplierID )
		if (CCconnectedElement.Type = "UseCase" OR CCconnectedElement.Type = "Collaboration") then
			' recurcive call for sub cases 
			set subConnectedElementConnectors = CCconnectedElement.Connectors	
			if subConnectedElementConnectors.Count > 1 then
				if colContains(reportedUCs, CCconnectedElement.ElementID) = 0 then
					Session.Output "The sub " + CCconnectedElement.Name + " has " + CStr(subConnectedElementConnectors.Count) + " connections ."
						sident = ""
						for i = 2 to level
							sident = sident & ">>>>> "
						next
						docGenerator.InsertText sident & " Составляющие ВИ: " & CCconnectedElement.Name, alignLeft					
						generationSuccess = docGenerator.InsertBreak(1)
					for z = 0 to subConnectedElementConnectors.Count - 1
						generateUC reportedUCs, docGenerator, CCconnectedElement, subConnectedElementConnectors, z, level, levelMax			
					next
				end if
			end if
			reportedUCs = reportedUCs & connectedElement.ElementID & "!"
		end if
	end if
end function

function colContains(col, val) 
	'for each c in col
	'	if c=val then 
	'		return true
	'	end if
	'next 
	if InStr(col,"!" & val & "!") >0 then
		ReportFLK val & ". " & Repository.GetElementByID( val ).name & " уже есть в отчете"
		colContains = 1
		exit function
	end if
	colContains = 0
end function 

sub ReportFLK( message )
	Session.Output "[CONTROL] " + message
end sub

sub ReportInfo( message )
	Session.Output "[INFO] " + message
end sub

sub ReportWarning( message )
	Session.Output "[WARNING] " + message
end sub

sub ReportFatal( message )
	Session.Output "[FATAL] " + message
	Session.Prompt message, promptOK
end sub

DocumentationExample