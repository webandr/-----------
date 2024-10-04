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

dim STATUS_FILTER
STATUS_FILTER = ""
'STATUS_FILTER = "Approved"

dim STATUS_VERSION
STATUS_VERSION = ""
'STATUS_VERSION = "1.1" 'подумать как быть 5 ВИ всего 3 ВИ версии 1.0 и не менялись 2 версии и 1.0 и 1.1 как тут фильтровать? без сложного запроса.
					   ' вариант 1 - 3 ВИ 1.0 остаются там где были и копи паст линки на диаграммы. Версия := 1.1. Фильтр по 1.1.
					   ' вариант 2 - 3 ВИ 1.0 копируются как новые ВИ (неудобно если там много связей) и версия := 1.1. ФИльтр по 1.1.

dim STATUS_PATH
STATUS_PATH = "Поведение системы (ПС) в.1.1 "

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
						
						if (connectedElement.Type = "UseCase" OR connectedElement.Type = "Collaboration") and filterElement(Repository.GetElementByID(connectedElement.ElementID)) then
						
							' Generate Use Case documentation
							ReportInfo "Generating documentation for " & directionArrow(currentConnector.Direction,"") & " UseCase: " + connectedElement.Name
							
							generationSuccess = docGenerator.InsertBreak(0)
							docGenerator.InsertText directionArrow(currentConnector.Direction,""), alignLeft
							'generationSuccess = docGenerator.InsertBreak(1)
							
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
		if (CCconnectedElement.Type = "UseCase" OR CCconnectedElement.Type = "Collaboration") and filterElement(Repository.GetElementByID(CCconnectedElement.ElementID)) then
			Session.Output "(" & CCconnectedElement.Name & ")" & + " === {" + currentCEConnector.Stereotype + "} =====> " + connectedElement.Name
			if currentCEConnector.Stereotype = "extend" then 
				sstereotype = "extending" 
			else 
				sstereotype = "Including" 
			end if
			if colContains(reportedUCs, CCconnectedElement.ElementID) = 0  then
				generationSuccess = docGenerator.InsertBreak(0)
				docGenerator.InsertText directionArrow("Destination -> Source", currentCEConnector.Stereotype), alignLeft
				generationSuccess = docGenerator.DocumentElement( CCconnectedElement.ElementID, level+2, "My UC details" )
				'if (currentCEConnector.Direction = "Source -> Destination") then
					' level+2 так как когда level=1 то это значит что в отчет надо вставить Header3
				'	generationSuccess = docGenerator.DocumentElement( CCconnectedElement.ElementID, level+2, "My " + sstereotype + " UC details" )
				'else
				'	generationSuccess = docGenerator.DocumentElement( CCconnectedElement.ElementID, level+2, "My " + sstereotype + " UC details" )
				'end if
				'if generationSuccess = false then
				'	ReportWarning "Error generating UseCase documentation: " + docGenerator.GetLastError()
				'end if	
			else
				'generationSuccess = docGenerator.InsertBreak(1)
				docGenerator.InsertText directionArrow("Destination -> Source", currentCEConnector.Stereotype), alignLeft
				generationSuccess = docGenerator.DocumentElement( CCconnectedElement.ElementID, level+2, "My UC name only" )
				docGenerator.InsertText "... ВИ уже присутствует в документе выше. ", alignCenter
				'generationSuccess = docGenerator.InsertBreak(1)
			end if
		end if
	else
		set CCconnectedElement = Repository.GetElementByID( currentCEConnector.SupplierID )
		if (CCconnectedElement.Type = "UseCase" OR CCconnectedElement.Type = "Collaboration") and filterElement(Repository.GetElementByID(CCconnectedElement.ElementID))  then
			Session.Output connectedElement.Name + " <===== {" + currentCEConnector.Stereotype + "} ===" + "(" + currentCECElement.Name + ")"
			if currentCEConnector.Stereotype = "extend" then 
				sstereotype = "extended" 
			else 
				sstereotype = "Included" 
			end if
			if colContains(reportedUCs, CCconnectedElement.Name) = 0 then
				generationSuccess = docGenerator.InsertBreak(0)
				docGenerator.InsertText directionArrow("Source -> Destination", currentCEConnector.Stereotype), alignLeft
				generationSuccess = docGenerator.DocumentElement( CCconnectedElement.ElementID, level+2, "My UC details" )

				'if (currentCEConnector.Direction = "Source -> Destination") then
				'	generationSuccess = docGenerator.DocumentElement( CCconnectedElement.ElementID, level+2, "My " + sstereotype + " UC details" )
				'else
				'	generationSuccess = docGenerator.DocumentElement( CCconnectedElement.ElementID, level+2, "My " + sstereotype + " UC details" )
				'end if
				'if generationSuccess = false then
				'	ReportWarning "Error generating UseCase documentation: " + docGenerator.GetLastError()
				'end if	
			else
				'generationSuccess = docGenerator.InsertBreak(1)
				docGenerator.InsertText directionArrow("Source -> Destination", currentCEConnector.Stereotype), alignLeft
				generationSuccess = docGenerator.DocumentElement( CCconnectedElement.ElementID, level+2, "My UC name only" )
				docGenerator.InsertText "... ВИ уже присутствует в документе выше. ", alignCenter
				'generationSuccess = docGenerator.InsertBreak(1)
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

	sident = ""
	for i = 2 to level
		sident = sident & ">>>>> "
	next
					
	if (currentCECElement.Name = connectedElement.Name)  then 'and (colContains(reportedUCs, connectedElement.ElementID) > 0) then
		set CCconnectedElement = Repository.GetElementByID( currentCEConnector.ClientID )
		'Session.Output CCconnectedElement.Name + "<" + CCconnectedElement.Type  + ">" 
		if (CCconnectedElement.Type = "UseCase" OR CCconnectedElement.Type = "Collaboration") and filterElement(Repository.GetElementByID(CCconnectedElement.ElementID)) then
			' recurcive call for sub cases 
			set subConnectedElementConnectors = CCconnectedElement.Connectors	
			if subConnectedElementConnectors.Count > 1 then
				if colContains(reportedUCs, CCconnectedElement.ElementID) = 0 then
					Session.Output "The sub " + CCconnectedElement.name + " has " + CStr(subConnectedElementConnectors.Count) + " connections ."
						docGenerator.InsertText sident & " Составляющие ВИ: " & CCconnectedElement.Name, alignLeft
						'generationSuccess = docGenerator.InsertBreak(1)
					for z = 0 to subConnectedElementConnectors.Count - 1
						generateUC reportedUCs, docGenerator, CCconnectedElement, subConnectedElementConnectors, z, level, levelMax			
					next
				else
					generationSuccess = docGenerator.InsertBreak(1)
					docGenerator.InsertText sident & " Составляющие ВИ: " & CCconnectedElement.Name & " уже приведены выше.", "Header4" 'alignLeft
					generationSuccess = docGenerator.InsertBreak(1)			
				end if
			end if
			reportedUCs = reportedUCs & connectedElement.ElementID & "!"
		end if
	else
		set CCconnectedElement = Repository.GetElementByID( currentCEConnector.SupplierID )
		if (CCconnectedElement.Type = "UseCase" OR CCconnectedElement.Type = "Collaboration") and filterElement(Repository.GetElementByID(CCconnectedElement.ElementID)) then
			' recurcive call for sub cases 
			set subConnectedElementConnectors = CCconnectedElement.Connectors	
			if subConnectedElementConnectors.Count > 1 then
				if colContains(reportedUCs, CCconnectedElement.ElementID) = 0 then
					Session.Output "The sub " + CCconnectedElement.Name + " has " + CStr(subConnectedElementConnectors.Count) + " connections ."
						docGenerator.InsertText sident & " Составляющие ВИ: " & CCconnectedElement.Name, alignLeft					
						generationSuccess = docGenerator.InsertBreak(1)
					for z = 0 to subConnectedElementConnectors.Count - 1
						generateUC reportedUCs, docGenerator, CCconnectedElement, subConnectedElementConnectors, z, level, levelMax			
					next
				else
					generationSuccess = docGenerator.InsertBreak(1)
					docGenerator.InsertText sident & " Составляющие ВИ: " & CCconnectedElement.Name & " уже приведены выше.", "Header4" 'alignLeft
					generationSuccess = docGenerator.InsertBreak(1)						
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


function directionArrow(ByVal direction, ByVal stereotype)
	if (direction="Source -> Destination") then
		if (stereotype = "extend") then
			directionArrow = " ==(" & stereotype &")===> "
		else 
			if(stereotype <> "" ) then
				directionArrow = " ==(" & stereotype &")===> "
			else
				directionArrow = " =====> "
			end if
		end if
	else
		if (stereotype = "extend") then
			directionArrow = " <===(" & stereotype &")== "
		else 
			if(stereotype <> "" ) then
				directionArrow = " <===(" & stereotype &")== "
			else
				directionArrow = " <===== "
			end if
		end if
	end if
end function 


function filterElement(ByRef element)
	filterElement = (true and filterByStatus(element) and filterByVersion(element) and filterByPath(element))
end function

function filterByStatus(ByRef element)
	if STATUS_FILTER = "" then 
		filterByStatus = true
		exit function 
	end if
	
	if (element.Status = STATUS_FILTER) then 
		filterByStatus = true
	else
		filterByStatus = false
	end if
	
end function

function filterByVersion(ByRef element)
	if STATUS_VERSION = "" then 
		filterByVersion = true
		exit function 
	end if
	
	if (element.Version = STATUS_VERSION) then 
		filterByVersion = true
	else
		filterByVersion = false
	end if

end function

function filterByPath(ByRef element)
	if STATUS_PATH = "" then 
		filterByPath = true
		exit function 
	end if
	
	if (InStr(element.FQName, STATUS_PATH)>0) then 
		filterByPath = true
	else
		filterByPath = false
	end if

end function
		
	
	
DocumentationExample