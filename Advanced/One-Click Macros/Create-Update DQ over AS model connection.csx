#r "Microsoft.VisualBasic"
using Microsoft.VisualBasic;
using MSTabular = Microsoft.AnalysisServices.Tabular;
using System.Windows.Forms;


/*
 * Title: Create/Update DQ over AS Model
 * 
 * Author: rem-bou
 * 
 * This script is to duplicate PowerBI behaviour on creation or update of a Module 
 * If Shared Expression selected containing a DQ over AS expression, it will update tables and measures
 * If any other object selected when script is run, it will check for shared expression and offer to update them or create a new one
 * 
 * Limitation: The measure type isn't 100% accurate and can default to Variant type
 *
 * Requirements: Null
 */


string directQueryName = "";
string workspaceName = "Workspace";
string workspaceConnectionString;
string databaseName = "Example Model";
string connectString;
bool askServerDatabaseUserInput = false;
bool UpdateCurrentCompositeModel = false;
bool selectedItemIsExpression = false;
bool dev_mode = false;
string ExpressionName;
string API_PBI_https = "powerbi://api.powerbi.com/v1.0/myorg/";


// This function will extract all quoted values from the M expression, returning a list of strings
// with the values extracted (in order), but ignoring any quoted values where a hashtag (#) precedes
// the quotation mark:
var split = new Func<string, List<string>>(m => { 
    var result = new List<string>();
    var i = 0;
    foreach(var s in m.Split('"')) {
        if(s.EndsWith("#") && i % 2 == 0) i = -2;
        if(i >= 0 && i % 2 == 1) result.Add(s);
        i++;
    }
    return result;
});
var GetServer = new Func<string, string>(m => split(m)[0]);    // Server name is usually the 1st encountered string
var GetDatabase = new Func<string, string>(m => split(m)[1]);  // Database name is usually the 2nd encountered string

MSTabular.Model modelRemote = new MSTabular.Model{};

//Delete all tables in DQ over AS - for dev purpose only
//foreach(var t in Model.Tables.ToList()){
//    if(t.Partitions.Any( p=> p.SourceType.ToString() ==  "Entity")){
//        t.Delete();
//    }
//}




// -------------- User Interaction -----------------
do{
    // Check that shared expression selected is a Composite Model string
    // used to define if using existing connection name or not
    ScriptHelper.WaitFormVisible = false;
    askServerDatabaseUserInput = false;
    UpdateCurrentCompositeModel = false;
    directQueryName = "";
    

    if (dev_mode){
        workspaceName = "";
        databaseName = "";
    }
    else{

        if(!selectedItemIsExpression){
            foreach( var exp in Selected){
                if( exp.ObjectType.ToString() == "Expression" ){  
                    if( Model.Expressions.Any( e => e.Name == exp.Name && e.Expression.Contains("AnalysisServices.Database(\"" + API_PBI_https ))){
                        selectedItemIsExpression = true;
                        directQueryName = exp.Name;
                        
                        DialogResult dialogResult = MessageBox.Show( "Do you want to change workspace or database for Composite Model " +directQueryName + "?", "Changing workspace or database connection", MessageBoxButtons.YesNo);
                        UpdateCurrentCompositeModel = (dialogResult == DialogResult.Yes);

                    }
                }
            }
            if(!selectedItemIsExpression && Model.Expressions.Count( e => e.Expression.Contains("AnalysisServices.Database(\"" + API_PBI_https) )>0 ){
                DialogResult dialogResult = MessageBox.Show( "Do you want to update tables and measures from an existing model in Direct Query Mode?", "Updating model in DQ", MessageBoxButtons.YesNo);
                bool UpdateACompositeModel = (dialogResult == DialogResult.Yes);
                if(UpdateACompositeModel){
                    var compositeModelExpression = Model.Expressions.Where(e=> e.Expression.Contains("AnalysisServices.Database(\"" + API_PBI_https )); // get expression(s) used for DQ over AS model
                    var SeletSharedExpression = SelectObject( compositeModelExpression, label:"Select Composite Model to update." ); // Box to select DQ over AS model to update locally
                    if( SeletSharedExpression != null){
                        if(SeletSharedExpression.Expression.Contains("AnalysisServices.Database(\"" + API_PBI_https ) ){
                            directQueryName = SeletSharedExpression.Name;
                        }
                    }
                }
            }
            // if retrieving Shared Expression - get current workspace and database
            if( directQueryName != ""){
                workspaceName = GetServer(Model.Expressions[directQueryName].Expression).Replace(API_PBI_https,"");
                databaseName = GetDatabase(Model.Expressions[directQueryName].Expression);
            }
        }

        if(UpdateCurrentCompositeModel ||directQueryName == "" ){
            workspaceName = Interaction.InputBox("What is the name of the workspace where the model is?\nProvide the exact name (case sensitive).", "Workspace Connection String", workspaceName );
            databaseName = Interaction.InputBox("What is the name of the model in " +workspaceName + "?\nProvide the exact name (case sensitive).", "Model Name", databaseName );
        }
    }
    
    workspaceConnectionString = "powerbi://api.powerbi.com/v1.0/myorg/" + workspaceName.Replace(" ","%20");
    connectString = $"DataSource={workspaceConnectionString};";

    if(directQueryName == ""){
        directQueryName = "DirectQuery to AS - " + databaseName;
    }

    // connect to the Power BI workspace referenced in connect string
    MSTabular.Server serverRemote = new MSTabular.Server();
    try{
        serverRemote.Connect(connectString);
    }
    catch{
        Warning("Enter valid workspace name.");
        askServerDatabaseUserInput = true;
    }

    string DatabaseID = "";
    // enumerate through datasets in workspace to get DatabaseID
    foreach (MSTabular.Database databaseRemote in serverRemote.Databases)
    {
        if ( databaseRemote.Name == databaseName ) {
        DatabaseID = databaseRemote.ID;
        break;
        }
    }
    
    if(!askServerDatabaseUserInput){
        try{
            modelRemote = serverRemote.Databases[DatabaseID].Model;
        }
        catch{
            Warning("Enter valid model name.");
            askServerDatabaseUserInput = true;
        }
    }
    if(askServerDatabaseUserInput){
        DialogResult dialogResult = MessageBox.Show("Do you want to correct your inputs??", "Workspace and Database Input", MessageBoxButtons.YesNo);
        askServerDatabaseUserInput = (dialogResult == DialogResult.Yes);
        if(!askServerDatabaseUserInput){return;}
    }
    
} while (askServerDatabaseUserInput);




// -------------- End User Interaction -----------------
string sharedExpressionMCode = String.Join(
    Environment.NewLine,
    "let",
    "    Source = AnalysisServices.Database(\"" + workspaceConnectionString + "\", \"" + databaseName + "\"),",
    "    Cubes = Table.Combine(Source[Data]),",
    "    Cube = Cubes{[Id=\"Model\", Kind=\"Cube\"]}[Data]",
    "in",
    "    Cube"
);
if ( !Model.Expressions.Any( e => e.Name == directQueryName)){
    Model.AddSharedExpression(directQueryName, expression:  sharedExpressionMCode);
    Model.Expressions[directQueryName].Kind = ExpressionKind.M;
    Model.Expressions[directQueryName].SetAnnotation(name:"PBI_IncludeFutureArtifacts", value:"True");
    Model.Expressions[directQueryName].SetAnnotation(name:"TabularEditor_CompositeModel", value:directQueryName);
}

// Delete all tables in DQ over AS if they do not contains local measures or local columns and they are for the model we are interested in.
foreach(var t in Model.Tables.ToList()){
    if( 
        (
            t.Partitions.Any( p=> p.SourceType.ToString() ==  "Entity") && (t.Measures.Count(m => m.Expression.Contains("EXTERNALMEASURE(")) == t.Measures.Count() ) )
            && ( t.Partitions.Any( p=> p.SourceType.ToString() ==  "Entity") && ( t.Columns.Count( c=> c.ObjectTypeName != "Calculated Column") == t.Columns.Count() ) )
            && (t.Partitions[t.Name] as EntityPartition).ExpressionSource.Name == directQueryName
        ){
        t.Delete();
    }
    else{
        if(t.Partitions.Any( p=> p.SourceType.ToString() ==  "Entity") && (t.Partitions[t.Name] as EntityPartition).ExpressionSource.Name == directQueryName ){
            if(t.GetAnnotation("TabularEditor_CompositeModel") == directQueryName){
                ("Table \"" + t.Name + "\" was not deleted because it has local object(s) created on it.").Output();
                foreach( var m in t.Measures.ToList()){
                    if( m.Expression.Contains("EXTERNALMEASURE(")){ 
                        m.Delete();
                    }
                }
                foreach( var c in t.Columns.ToList()){
                    if( c.ObjectTypeName != "Calculated Column" ){ 
                        c.Delete();
                    }
                }
            }
        }
    }
}



//if you don't specify a database, it will only grab models from the first database in the list
string tableName;
foreach (MSTabular.Table tableRemote in modelRemote.Tables)
{
    if(tableRemote.Name.StartsWith("Measures")){
        tableName =  tableRemote.Name;
    }
    else{
        tableName =  tableRemote.Name + " - " + databaseName.Replace(" Model","");
    }

    bool IsTableAccepted = true;
    if( IsTableAccepted ){
        
        if( !Model.Tables.Any( t => t.Name == tableName) ){
            Model.AddTable(tableName).AddEntityPartition("ToBeDefined",tableRemote.Name);
            foreach( var par in  Model.Tables[tableName].Partitions.ToList()){
                if( par.SourceType.ToString() !=  "Entity"){
                    par.Delete();
                }
            }
            Model.Tables[tableName].Partitions["ToBeDefined"].Name = tableName;
            Model.Tables[tableName].Partitions[tableName].Mode = ModeType.DirectQuery;
            (Model.Tables[tableName].Partitions[tableName] as EntityPartition).ExpressionSource = Model.Expressions[directQueryName];
            Model.Tables[tableName].SourceLineageTag = tableRemote.SourceLineageTag;
            Model.Tables[tableName].SetAnnotation(name:"TabularEditor_CompositeModel", value:directQueryName);
        }
        //Add column to table
        foreach(MSTabular.Column columnRemote in tableRemote.Columns){
            if(! columnRemote.Name.StartsWith("RowNumber-")){
                var externalColumn = Model.Tables[tableName].AddDataColumn(columnRemote.Name);
                externalColumn.Description = columnRemote.Description;
                externalColumn.DisplayFolder = columnRemote.DisplayFolder;
                externalColumn.IsHidden = columnRemote.IsHidden;
                externalColumn.FormatString = columnRemote.FormatString;
                externalColumn.SourceLineageTag = columnRemote.LineageTag;
                externalColumn.SourceProviderType = columnRemote.SourceProviderType;
                (externalColumn as DataColumn).SourceColumn = columnRemote.Name;
                // Define DataType of the column before attribuing it
                var columnDataType = DataType.String;
                switch(columnRemote.DataType.ToString()) 
                {
                    case "Binary":
                        columnDataType = DataType.Binary;
                        break;
                    case "Boolean":
                        columnDataType = DataType.Boolean;
                        break;
                    case "DateTime":
                        columnDataType = DataType.DateTime;
                        break;
                    case "Decimal":
                        columnDataType = DataType.Decimal;
                        break;
                    case "Double":
                        columnDataType = DataType.Double;
                        break;
                    case "Int64":
                        columnDataType = DataType.Int64;
                        break;
                    case "String":
                        columnDataType = DataType.String;
                        break;
                    case "Unknown":
                        columnDataType = DataType.Unknown;
                        break;
                    case "Variant":
                        columnDataType = DataType.Variant;
                        break;
                    default:
                        columnDataType = DataType.String;
                        break;
                }
                (externalColumn as DataColumn).DataType = columnDataType;
                externalColumn.SetAnnotation(name:"TabularEditor_CompositeModel", value:directQueryName);
            }
        }
        // sort by columns need to be done after all columns were added
        foreach(MSTabular.Column columnRemote in tableRemote.Columns){
            if(! columnRemote.Name.StartsWith("RowNumber-")){
                try{
                    Model.Tables[tableName].Columns[columnRemote.Name].SortByColumn = (Model.Tables[tableName].Columns[columnRemote.SortByColumn.Name] as DataColumn);
                }
                catch{ continue;}
            }
        }

        // Create measure to reference composite model's one
        // measureType is not correct, unable to get the right measure types - no impact found
        foreach(MSTabular.Measure measureRemote in tableRemote.Measures){
            string measureName = measureRemote.Name;
            string measureType = measureRemote.DataType.ToString();
            switch(measureType){
                case "Binary":
                        measureType = "BOOLEAN";
                        break;
                    case "Boolean":
                        measureType = "BOOLEAN";
                        break;
                    case "DateTime":
                        measureType = "DATETIME";
                        break;
                    case "Decimal":
                        measureType = "CURRENCY";
                        break;
                    case "Double":
                        measureType = "DOUBLE";
                        break;
                    case "Int64":
                        measureType = "INTEGER";
                        break;
                    case "String":
                        measureType = "STRING";
                        break;
                    default:
                        measureType = "VARIANT";
                        break;
            }
            string measureExpression = "EXTERNALMEASURE(\"" + measureName + "\"," + measureType + ", \"" + directQueryName + "\")";
            var externalMeasure = Model.Tables[tableName].AddMeasure(measureRemote.Name, measureExpression);
            externalMeasure.Description = measureRemote.Description;
            externalMeasure.DisplayFolder = measureRemote.DisplayFolder;
            externalMeasure.IsHidden = measureRemote.IsHidden;
            externalMeasure.FormatString = measureRemote.FormatString;
            externalMeasure.SourceLineageTag = measureRemote.LineageTag;
            externalMeasure.SetAnnotation(name:"TabularEditor_CompositeModel", value:directQueryName);
        }
    }
}

foreach (  MSTabular.SingleColumnRelationship relationshipRemote in modelRemote.Relationships)
{
    string fromTable = relationshipRemote.FromTable.Name;
    string toTable = relationshipRemote.ToTable.Name;
    string fromColumn = relationshipRemote.FromColumn.Name ;
    string toColumn = relationshipRemote.ToColumn.Name;
    var relType = relationshipRemote.Type;
    
    var relCrossFilteringBehavior = CrossFilteringBehavior.OneDirection;
    switch (relationshipRemote.CrossFilteringBehavior.ToString() ){
        case "Automatic":
            relCrossFilteringBehavior = CrossFilteringBehavior.Automatic;
            break;
        case "BothDirections":
            relCrossFilteringBehavior = CrossFilteringBehavior.BothDirections;
            break;
        case "OneDirection":
            relCrossFilteringBehavior = CrossFilteringBehavior.OneDirection;
            break;
    }
    
    var fromCardinality = RelationshipEndCardinality.None;
    switch(relationshipRemote.FromCardinality.ToString()){
        case "None":
            fromCardinality = RelationshipEndCardinality.None ;
            break;
        case "One":
            fromCardinality = RelationshipEndCardinality.One;
            break;
        case "Many":
            fromCardinality = RelationshipEndCardinality.Many;
            break;
    }
    
    var toCardinality  =  RelationshipEndCardinality.None;
    switch(relationshipRemote.ToCardinality.ToString()){
        case "None":
            toCardinality = RelationshipEndCardinality.None ;
            break;
        case "One":
            toCardinality = RelationshipEndCardinality.One;
            break;
        case "Many":
            toCardinality = RelationshipEndCardinality.Many;
            break;
    }
    
    var relSecurityFilteringBehavior = SecurityFilteringBehavior.None;
    switch (relationshipRemote.SecurityFilteringBehavior.ToString() ){
        case "None":
            relSecurityFilteringBehavior = SecurityFilteringBehavior.None;
            break;
        case "BothDirections":
            relSecurityFilteringBehavior = SecurityFilteringBehavior.BothDirections;
            break;
        case "OneDirection":
            relSecurityFilteringBehavior = SecurityFilteringBehavior.OneDirection;
            break;
    }
    
    bool relRelyOnReferentialIntegrity = relationshipRemote.RelyOnReferentialIntegrity;
    bool relActive = relationshipRemote.IsActive;
    
    if(fromTable.StartsWith("Measures")){
        fromTable =  fromTable;
    }
    else{
        fromTable =  fromTable + " - " + databaseName.Replace(" Model","");
    }
    if(toTable.StartsWith("Measures")){
        toTable =  toTable;
    }
    else{
        toTable =  toTable + " - " + databaseName.Replace(" Model","");
    }

    var rel = Model.AddRelationship();
    rel.FromColumn = Model.Tables[fromTable].Columns[fromColumn];
    rel.ToColumn = Model.Tables[toTable].Columns[toColumn];
    rel.CrossFilteringBehavior = relCrossFilteringBehavior;
    rel.SecurityFilteringBehavior =relSecurityFilteringBehavior;
    rel.FromCardinality = fromCardinality;
    rel.ToCardinality = toCardinality;
    rel.RelyOnReferentialIntegrity = relRelyOnReferentialIntegrity;
    rel.IsActive = relActive;
    rel.SetAnnotation(name:"TabularEditor_CompositeModel", value:directQueryName);
}


Info("Model " + directQueryName + " has been created/updated.");