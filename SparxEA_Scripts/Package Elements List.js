//
// Iterates through an EAP file using recursion.
// 
// Related APIs
// =================================================================================
// Repository API - http://www.sparxsystems.com/enterprise_architect_user_guide/12.1/automation_and_scripting/repository3.html
//
function RecursiveModelDumpExample(PackID)
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "JavaScript RECURSIVE MODEL DUMP EXAMPLE" );
	Session.Output( "=======================================" );
	
	var currentModel as EA.Package;
	currentModel = Repository.GetPackageByGuid(PackID);
	
	// // Iterate through all models in the project
	// for (var i=0; i < Repository.Models.Count; i++)
	// {
	// 	var currentModel as EA.Package;
	//	currentModel = Repository.Models.GetAt(i);
	//	Session.Output(currentModel.Name + " == " + PackID);
	//	if (currentModel.Name == PackID) {
	//		// Recursively process this package
		DumpPackage( "", currentModel , PackID);
	//	}
	//}
	
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
function DumpPackage( indent, thePackage, PackID)
{
			// Cast thePackage to EA.Package so we get intellisense
		var currentPackage as EA.Package;
		currentPackage = thePackage;
	
	//Session.Output( PackID + " " + currentPackage.Name);
	
	if (currentPackage.PackageGUID == PackID) {

		
		// Add the current package's name to the list
		Session.Output( indent + currentPackage.Name + " (PackageID=" + 
			currentPackage.PackageID + ")" );
		
		// Dump the elements this package contains
		DumpElements( indent + "    ", currentPackage );
		
		// Recursively process any child packages
		for (var i=0; i < currentPackage.Packages.Count; i++)
		{
			var childPackage as EA.Package;
			childPackage = currentPackage.Packages.GetAt(i);
			
			DumpPackage( indent + "    ", childPackage , currentPackage.Packages.GetAt(i).PackageGUID);
		}

	}
	else
		// Recursively process any child packages
		for (var i=0; i < currentPackage.Packages.Count; i++)
		{
			var childPackage as EA.Package;
			childPackage = currentPackage.Packages.GetAt(i);
			
			DumpPackage( indent + "    ", childPackage , PackID);
		}
}

//
// Outputs the elements of the provided package to the Script output window
//
// Parameters:
// - indent A string representing the current level of indentation
// - thePackage The package object to be processed
//
function DumpElements( indent, thePackage )
{
	// Cast thePackage to EA.Package so we get intellisense
	var currentPackage as EA.Package;
	currentPackage = thePackage;
	
	// Iterate through all elements and add them to the list
	for (var i=0; i < currentPackage.Elements.Count; i++)
	{
		var currentElement as EA.Element;
		currentElement = currentPackage.Elements.GetAt(i);

		Session.Output( indent + currentElement.Type + "::" + currentElement.Name +
			" (" + 
			"ID=" + currentElement.ElementID + ")" );
	}
}

//RecursiveModelDumpExample("051 :: ВИ системы");
 RecursiveModelDumpExample(Repository.GetTreeSelectedPackage().PackageGUID);