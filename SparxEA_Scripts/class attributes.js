function moveNext()
	{
		if(this.iElem > -1)
		{
			this.iElem++;
			if(this.iElem < this.Package.Count)
			{
				return true;
			}
			this.iElem = this.Package.Count;
		}
		return false;
	}
	function item()
	{
		if( this.iElem > -1 && this.iElem < this.Package.Count)
		{
			return this.Package.GetAt(this.iElem);
		}
		return null;
	}

	function atEnd()
	{
		if((this.iElem > -1) && (this.iElem < this.Package.Count))
		{
			return false;
		}
		// Session.Output("at end!");
		return true;
	}

	function Check( obj)
	{
		if(obj == undefined)
		{
			Session.Output("Undefined object");
			return false;
		}
		return true;
	}	


function Enumerator( object )
{
	this.iElem = 0;
	this.Package = object;
	this.atEnd = atEnd;
	this.moveNext = moveNext;
	this.item = item;
	this.Check = Check;
	if(!Check(object))
	{
		this.iElem = -1;
	}
}
//
// Iterates through an EAP file using recursion.
// 
// Related APIs
// =================================================================================
// Repository API -  http://www.sparxsystems.com/uml_tool_guide/sdk_for_enterprise_architect/repository3.html
//
function RecursiveModelDumpExample(packageID, packageName, diagramID)
{
	Session.Output( "JScript RECURSIVE MODEL DUMP EXAMPLE" );
	Session.Output( "=======================================" );
	
	if (packageID==0)
	{
		// Iterate through all models in the project
		var modelEnumerator = new Enumerator( Models );
		while ( !modelEnumerator.atEnd() )
		{
			var currentModel as EA.Package;
			
			currentModel =  modelEnumerator.item();
		
			// Recursively process this package
			DumpPackage( "", currentModel, packageName, diagramID );
			
			modelEnumerator.moveNext();
		}
	}
	else
	{
		currentModel =  Repository.GetPackageByID(packageID); 
		DumpPackage( "", currentModel, packageName, diagramID );		
	}
	Session.Output( "Done!" );
}

//
// Outputs the packages name and elements, and then recursively processes any child 
// packages
//
// Parameters:
// - indent A string representing the current level of indentation
// - thePackage The package object to be processed
//
function DumpPackage( indent, thePackage , packageName, diagramID)
{
	// Cast thePackage to EA.Package so we get intellisense
	// var currentPackage as EA.Package.FindObject("_Атрибуты_ПФ");
	var currentPackage as EA.Package;
	currentPackage = thePackage;

	
	// Add the current package's name to the list
	Session.Output( indent + currentPackage.Name + " (PackageID=" + currentPackage.PackageID + ")" );
	
	// Dump the elements this package contains
	if (currentPackage.Name==packageName){
		DumpElements( indent + "    ", currentPackage,  diagramID);
	}
	
	// Recursively process any child packages
	var childPackageEnumerator = new Enumerator( currentPackage.Packages );
	while ( !childPackageEnumerator.atEnd() )
	{
		var childPackage as EA.Package;
		childPackage = childPackageEnumerator.item();
		
		DumpPackage( indent + "    ", childPackage );
		
		childPackageEnumerator.moveNext();
	}
}

//
// Outputs the elements of the provided package to the Script output window
//
// Parameters:
// - indent A string representing the current level of indentation
// - thePackage The package object to be processed
//
function DumpElements( indent, thePackage, diagramID )
{
	// Cast thePackage to EA.Package so we get intellisense
	var currentPackage as EA.Package;
	currentPackage = thePackage;
	
	var diagr as EA.Diagram;
	diagr = Repository.GetDiagramByGuid(diagramID);
	
	// Iterate through all elements and add them to the list
	var elementEnumerator = new Enumerator( currentPackage.Elements );
	while ( !elementEnumerator.atEnd() )
	{
		var sElemType = "";
		var currentElement as EA.Element;
		var attrRel as EA.Element;
		
		currentElement = elementEnumerator.item();
		
		Session.Output( indent + "::" + currentElement.Name +
			" (" + currentElement.Type +
			", ID=" + currentElement.ElementID + ")" );
		
		
		var childAttr as EA.Attribute;
		sElemType  = currentElement.Type
		var attrEnumerator = new Enumerator( currentElement.Attributes );		
		if (sElemType == "Class"){
			while ( !attrEnumerator.atEnd() ){
				var aElem as EA.Element;
				childAttr = attrEnumerator.item();
				Session.Output( indent + "    :: attribute : " + childAttr.AttributeID + "." + childAttr.Name);
				attrEnumerator.moveNext();
			}
		}
		attrEnumerator = new Enumerator( currentElement.Attributes );
		
		Session.Output( "Connecttors " + currentElement.Connectors.Count );
			
		elementEnumerator.moveNext();
	}
	
	
	// getAllDiagramList();
		//
		// смотри D:\ЯДиск\YandexDisk\#GD\АНОЦТ\Проекты\АИС ГИН на ПКНД\Инструменты\Eye_2_SparxEA.xlsx
		// для понимания связности сущностей в SparxEA
		//

		// GetDiagramByGuid("{E4FFF907-6F13-46d8-A4E1-511BE264E835}");
		//"{828158C8-1681-4644-8DA1-C6CA22D616D7}");
		conectsOfClass = new Enumerator( diagr.DiagramLinks );
		while ( !conectsOfClass.atEnd() ){
			Session.Output("----------------------------------------------------------------------------------------------");
			var conOfClass as EA.Connector;
			conOfClass = conectsOfClass.item();
			
			//Session.Output("Connection ID = " + conOfClass.ConnectorID);
			//Session.Output("Connection SourceInstanceUID  = " + conOfClass.SourceInstanceUID );
			
			var tCon as EA.Connector;
			tCon = Repository.GetConnectorByID(conOfClass.ConnectorID);
			
			//Session.Output("Connection SourceInstanceUID атрибут GUID  = " + tCon.StyleEx);
			
			var s ="";
			s = tCon.StyleEx.substring(6);
			s = s.substring(0, s.indexOf("}"));
			s = "{"+s+"}";
			//Session.Output("Будем искать атрибут с GUID = " + s);
			
			var tAttr as EA.Attribute;
			tAttr = Repository.GetAttributeByGuid(s);
			//Session.Output("Найденый атрибут = " + tAttr.Name);
			
		
			objectsOfDiagram = new Enumerator( diagr.DiagramObjects );
			while ( !objectsOfDiagram.atEnd() )
			{
				//var objOfDiagr as EA.Element;
				objOfDiagr = objectsOfDiagram.item();
				if (0 || conOfClass.SourceInstanceUID ==  objOfDiagr.InstanceGUID) {
					
					//Session.Output(" InstanceGUID = " + objOfDiagr.InstanceGUID + " >> ElementID = " + objOfDiagr.ElementID + " InstanceID = " + objOfDiagr.InstanceID );
					
					var zz as EA.Element;
					zz = Repository.GetElementByID(objOfDiagr.ElementID);
					if (typeof tAttr == "undefined")
					{
						Session.Output("Связь от " + zz.Name );
					} else 
					{
						Session.Output("Связь от " + zz.Name + "." + tAttr.Name);				
					}
				}
				if (0 || conOfClass.TargetInstanceUID ==  objOfDiagr.InstanceGUID) {
					
					//Session.Output(" InstanceGUID = " + objOfDiagr.InstanceGUID + " >> ElementID = " + objOfDiagr.ElementID + " InstanceID = " + objOfDiagr.InstanceID );
					
					var zz as EA.Element;
					zz = Repository.GetElementByID(objOfDiagr.ElementID);

					Session.Output("Связь к " + zz.Name);				
				}
				objectsOfDiagram.moveNext(); 
			}				
			conectsOfClass.moveNext(); 
		}
		
		Session.Output("----------------------------------------------------------------------------------------------");
	
}

function findAttr( enumA, id ) 
{
	
	while ( !enumA.atEnd() ){
					
		var aElem as EA.Element;
		aElem = enumA.item();
		if (aElem.AttributeID==id){
			return aElem;
		}
		enumA.moveNext();
	}
}

function getAllDiagramList()
{
	
	Session.Output( "JScript RECURSIVE MODEL DUMP EXAMPLE" );
	Session.Output( "=======================================" );
	
	// Iterate through all models in the project
	var modelEnumerator = new Enumerator( Models );
	while ( !modelEnumerator.atEnd() )
	{
		var currentModel as EA.Package;
		currentModel = modelEnumerator.item();
		//Repository.GetPackageByID(632); 
		//modelEnumerator.item();
				
		// Recursively process this package
		DumpPackageDiagrams( "", currentModel );
		
		modelEnumerator.moveNext();
	}
	
	Session.Output( "Done!" );
	

}


function DumpPackageDiagrams( indent, thePackage )
{
	// Cast thePackage to EA.Package so we get intellisense
	// var currentPackage as EA.Package.FindObject("_Атрибуты_ПФ");
	var currentPackage as EA.Package;
	currentPackage = thePackage;
	
	// Add the current package's name to the list
	Session.Output( indent + currentPackage.Name + " (PackageID=" + currentPackage.PackageID + ")" );
	
	// Dump the elements this package contains
	//f (currentPackage.Name=="_Атрибуты_ПФ"){
		getPackageDiagramList( indent + "    ", currentPackage );
	//}
	
	// Recursively process any child packages
	var childPackageEnumerator = new Enumerator( currentPackage.Packages );
	while ( !childPackageEnumerator.atEnd() )
	{
		var childPackage as EA.Package;
		childPackage = childPackageEnumerator.item();
		
		DumpPackageDiagrams( indent + "    ", childPackage );
		
		childPackageEnumerator.moveNext();
	}
}

function getPackageDiagramList(indent, thePackage)
{
	var currentPackage as EA.Package;
	currentPackage = thePackage;
	
	var diagramsEnumerator = new Enumerator( currentPackage.Diagrams );
	
	while ( !diagramsEnumerator.atEnd() )
	{
		var d as EA.Diagram;
		d = diagramsEnumerator.item();
		Session.Output(indent + "    " + d.DiagramID + " :: " + d.Name + "(DiagramID =" + d.DiagramID + ")");
		diagramsEnumerator.moveNext();
	}
}


// getAllDiagramList();
// RecursiveModelDumpExample(819, "Scripts",  "{102F62A6-4DE8-486d-96DB-3FE6EC388DD8}");
RecursiveModelDumpExample(1781 , "Scripts",  "{102F62A6-4DE8-486d-96DB-3FE6EC388DD8}");
