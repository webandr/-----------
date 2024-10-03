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
			for i = 0 to packageElements.Count - 1
			
				' Get the current element
				dim currentElement as EA.Element;
				set currentElement = packageElements.GetAt( i )
				
				if currentElement.Type = "Actor" then
				
					' Generate Actor documentation
					ReportInfo "Generating documentation for actor: " + currentElement.Name
					generationSuccess = docGenerator.DocumentElement( currentElement.ElementID, 0, ACTOR_TEMPLATE )
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
							
							
							generationSuccess = docGenerator.DocumentElement( connectedElement.ElementID, 1, USECASE_TEMPLATE )
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
															
									' Get the current connector and the element that it connects to
									dim currentCEConnector as EA.Connector
									set currentCEConnector = ConnectedElementConnectors.GetAt( x )
									set currentCECElement = Repository.GetElementByID( currentCEConnector.SupplierID )
									set currentCECElementClient = Repository.GetElementByID( currentCEConnector.ClientID )
									'Session.Output "1st level connectedElement.Name = " & connectedElement.Name								
									'Session.Output "2nd level SupplyID Name = " & currentCECElement.Name
									'Session.Output "2nd level ClientID Name = " & currentCECElementClient.Name
									
									if (currentCECElement.Name = connectedElement.Name) then
										set CCconnectedElement = Repository.GetElementByID( currentCEConnector.ClientID )
										'Session.Output CCconnectedElement.Name + "<" + CCconnectedElement.Type  + ">" 
										if (CCconnectedElement.Type = "UseCase" OR CCconnectedElement.Type = "Collaboration") then
											Session.Output "{" + currentCEConnector.Stereotype + "}" + CCconnectedElement.Name + " <S-Dest> element = " + connectedElement.Name
											if currentCEConnector.Stereotype = "extend" then 
												sstereotype = "extending" 
											else 
												sstereotype = "Including" 
											end if
											if (currentCEConnector.Direction = "Source -> Destination") then
												generationSuccess = docGenerator.DocumentElement( CCconnectedElement.ElementID, 2, "My " + sstereotype + " UC details" )
											else
												generationSuccess = docGenerator.DocumentElement( CCconnectedElement.ElementID, 2, "My " + sstereotype + " UC details" )
											end if
											if generationSuccess = false then
												ReportWarning "Error generating UseCase documentation: " + docGenerator.GetLastError()
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
											if (currentCEConnector.Direction = "Source -> Destination") then
												generationSuccess = docGenerator.DocumentElement( CCconnectedElement.ElementID, 2, "My " + sstereotype + " UC details" )
											else
												generationSuccess = docGenerator.DocumentElement( CCconnectedElement.ElementID, 2, "My " + sstereotype + " UC details" )
											end if
											if generationSuccess = false then
												ReportWarning "Error generating UseCase documentation: " + docGenerator.GetLastError()
											end if											
										end if
									end if
	
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