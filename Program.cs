﻿using System.Data;
using System.Diagnostics;
using Microsoft.AnalysisServices.AdomdClient;
using ClosedXML.Excel;
using System.IO.Compression;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using System.Collections;


DataTable QueryResult = new DataTable();

var filePath = "C:\\Users\\RahulGaurMAQSoftware\\Downloads\\UnzipTest\\Refresh Tracker.pbix";
var extractPath = Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath));

ZipFile.ExtractToDirectory(filePath, extractPath, true);
Console.WriteLine("Extraction complete");

var layoutFilePath = Path.Combine(extractPath, "Report", "Layout");
var layoutContent = Regex.Replace(File.ReadAllText(layoutFilePath), @"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "");

dynamic fileContents = JsonConvert.DeserializeObject(layoutContent);
var parserArray = new List<Dictionary<string, string>>();

DataTable dataTable = new DataTable();

// Add columns to the DataTable
dataTable.Columns.Add("PageName", typeof(string));
dataTable.Columns.Add("VisualName", typeof(string));
dataTable.Columns.Add("MeasureName", typeof(string));
dataTable.Columns.Add("ColumnName", typeof(string));
dataTable.Columns.Add("DimensionName", typeof(string));
dataTable.Columns.Add("VisualTitle", typeof(string));
dataTable.Columns.Add("ReportName", typeof(string)).DefaultValue = Path.GetFileNameWithoutExtension(filePath);

// Process the layout content
foreach (var section in fileContents.sections)
{
    string pageName = section.displayName;

    foreach (var visualContainer in section.visualContainers)
    {
        try
        {
            dynamic configData = visualContainer.config;
            configData = JObject.Parse(configData.ToString());
            dynamic data = configData.singleVisual;

            if (data.visualType == "textbox")
                continue;

            string capturedVisualName = data.visualType;
            string capturedVisualTitle = "";

            try
            {
                capturedVisualTitle = data.vcObjects.title[0].properties.text.expr.Literal.Value;
                capturedVisualTitle = capturedVisualTitle.Replace("'", "").Replace(",", "");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            dynamic selectData = data.prototypeQuery.Select;

            List<string> columnList = new List<string>();
            List<string> measureList = new List<string>();
            List<string> dimensionList = new List<string>();

            foreach (var item in selectData)
            {
                if (item.Column != null)
                {
                    columnList.Add(item.Column.Property.ToString());
                    string dimension = item.Name;
                    dimension = dimension.Split('.')[0].Replace("Min(", "");
                    dimensionList.Add(dimension);
                }

                if (item.Measure != null)
                    measureList.Add(item.Measure.Property.ToString());
            }

            foreach (var measure in measureList)
            {
                if (columnList.Count == 0)
                {
                    DataRow row = dataTable.NewRow();
                    row["PageName"] = pageName;
                    row["VisualName"] = capturedVisualName;
                    row["MeasureName"] = measure;
                    row["ColumnName"] = "";
                    row["DimensionName"] = "";
                    row["VisualTitle"] = capturedVisualTitle;
                    dataTable.Rows.Add(row);
                }
                else
                {
                    for (int e = 0; e < columnList.Count; e++)
                    {
                        DataRow row = dataTable.NewRow();
                        row["PageName"] = pageName;
                        row["VisualName"] = capturedVisualName;
                        row["MeasureName"] = measure;
                        row["ColumnName"] = columnList[e];
                        row["DimensionName"] = dimensionList[e];
                        row["VisualTitle"] = capturedVisualTitle;
                        dataTable.Rows.Add(row);
                    }
                }
            }
        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
        }
    }
}

Console.WriteLine(dataTable);
Console.WriteLine("Created Array");

//Next Program
















DataTable cvcq = new DataTable();
var modelName = "Refresh Tracker";
var endpoint = "powerbi://api.powerbi.com/v1.0/myorg/Sprouts%20EDW%20-%20PPE;";
var checkforlocal = "Y";
var connectionstring = "Provider=MSOLAP.8;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=f7aa610f-23d0-48db-9e31-4322e8c92a31;Data Source=localhost:60735;MDX Compatibility=1;Safety Options=2;MDX Missing Member Mode=Error;Update Isolation Level=2";
var thresholdValue = "0.000001";
var parsedDataFrame = dataTable;
var runningFirstTime = "1";
var con = new AdomdConnection();

if (checkforlocal == "Y" || checkforlocal == "y")
{
    connectionstring = connectionstring;
}
else
{
    connectionstring = "Provider=MSOLAP.8;Data Source=" + endpoint + ";initial catalog=" + modelName + ";UID=;PWD=";
}

con.ConnectionString = connectionstring;



void CreateExcelSheet(DataTable dataTable, string ExcelName)
{
    var excelFileName = Path.Combine(extractPath, ExcelName);
    using (var workbook = new XLWorkbook())
    {
        var worksheet = workbook.Worksheets.Add(dataTable, "Sheet1");
        workbook.SaveAs(excelFileName);
    }

    Console.WriteLine("Saved DataTable to Excel file: " + excelFileName);
}

DataTable ConvertToDataTable(IEnumerable dataRows)
{
    DataTable dataTable = new DataTable();
    List<string> columnList = new List<string>();

    Type itemType = null;
    foreach (var item in dataRows)
    {
        itemType = item.GetType();
        break;
    }

    foreach (var column in itemType.GetProperties().Select(prop => prop.Name))
    {
        dataTable.Columns.Add(column, typeof(string));
        columnList.Add(column);
    }

    // Add rows from the LINQ query result to the DataTable
    foreach (var row in dataRows)
    {
        DataRow newRow = dataTable.NewRow();
        foreach (var column in columnList)
        {
            // Access property dynamically using reflection
            object columnValue = row.GetType().GetProperty(column)?.GetValue(row);
            newRow[column] = columnValue != null ? columnValue.ToString() : string.Empty;
        }
        dataTable.Rows.Add(newRow);
    }

    return dataTable;
}

DataTable ExecuteDataTable(string query, List<string> columnsList)
{
    try
    {
        con.Open();
    }
    catch
    {
        Console.WriteLine("Already Open...");
    }

    var command = new AdomdCommand(query, con);
    var reader = command.ExecuteReader();


    DataTable dataTable = new DataTable();
    for (int i = 0; i < reader.FieldCount; i++)
    {
        dataTable.Columns.Add(columnsList[i], typeof(string));
    }

    while (reader.Read())
    {
        DataRow row = dataTable.NewRow();
        for (int i = 0; i < reader.FieldCount; i++)
        {
            var columnData = reader[i] != null ? reader[i] : "NULL";
            row[$"{columnsList[i]}"] = columnData.ToString();
        }
        dataTable.Rows.Add(row);
    }

    con.Close();

    return dataTable;
}

DataTable MeasureListSQLQuery()
{

    string query = "SELECT [MEASURE_NAME],[MEASUREGROUP_NAME],[EXPRESSION],[CUBE_NAME] FROM $SYSTEM.MDSCHEMA_MEASURES WHERE MEASURE_IS_VISIBLE AND MEASUREGROUP_NAME <> 'Reporting Filters' ORDER BY [MEASUREGROUP_NAME]";
    List<string> columnsList = new List<string> { "Measure", "MeasureGroup", "EXPRESSION", "CubeName" };
    var result = ExecuteDataTable(query, columnsList);
    CreateExcelSheet(result, "MeasureListSQLQuery.xlsx");
    return result;
}

DataTable MeasureReferenceQuery()
{

    string query = "SELECT DISTINCT [Object] ,[Referenced_Table] FROM $SYSTEM.DISCOVER_CALC_DEPENDENCY WHERE [Object_Type] = 'MEASURE'";
    List<string> columnsList = new List<string> { "Measure", "Referenced_Table" };
    var result = ExecuteDataTable(query, columnsList);
    CreateExcelSheet(result, "MeasureReferenceQuery.xlsx");
    return result;
}

DataTable RelationshipQuery()
{

    string query = "SELECT DISTINCT [FromTableID],[FromColumnID],[ToTableID],[ToColumnID]FROM $SYSTEM.TMSCHEMA_RELATIONSHIPS WHERE [IsActive]";
    List<string> columnsList = new List<string> { "FromTableID", "FromColumnID", "ToTableID", "ToColumnID" };
    var result = ExecuteDataTable(query, columnsList);
    CreateExcelSheet(result, "RelationshipQuery.xlsx");
    return result;
}

DataTable TableQuery()
{

    string query = "SELECT DISTINCT [Name],[ID] FROM $SYSTEM.TMSCHEMA_TABLES";
    List<string> columnsList = new List<string> { "TableName", "TableID" };
    var result = ExecuteDataTable(query, columnsList);
    CreateExcelSheet(result, "TableQuery.xlsx");
    return result;
}

DataTable ColumnsQuery()
{

    string query = "SELECT DISTINCT [TableID],[ExplicitName],[ID] FROM $SYSTEM.TMSCHEMA_COLUMNS WHERE [Type] <> 3 AND NOT [IsDefaultImage] AND [State] = 1";
    List<string> columnsList = new List<string> { "TableID", "ColumnName", "ColumnID" };
    var result = ExecuteDataTable(query, columnsList);
    CreateExcelSheet(result, "ColumnsQuery.xlsx");
    return result;
}



int ColumnValuesCountQueryforprogress()
{
    DataTable df = TableQuery();
    DataTable df1 = ColumnsQuery();

    // Convert DataTables to IEnumerable<DataRow> for LINQ
    var dfRows = df.AsEnumerable();
    var df1Rows = df1.AsEnumerable();

    // Perform inner join using LINQ
    var query = from row1 in dfRows
                join row2 in df1Rows
                on row1.Field<string>("TableID") equals row2.Field<string>("TableID")
                select new
                {
                    TableName = row1.Field<string>("TableName"),
                    TableID = row1.Field<string>("TableID"),
                    ColumnName = row2.Field<string>("ColumnName"),
                    ColumnID = row2.Field<string>("ColumnID")
                };

    df = ConvertToDataTable(query);

    df.Columns.Add("ValuesQuery", typeof(string));
    df.Columns.Add("ID", typeof(string));

    for (int i = 0; i < df.Rows.Count; i++)
    {
        DataRow row = df.Rows[i];
        row["ValuesQuery"] = "WITH MEMBER [Measures].[Count] AS [" + row["TableName"] + "].[" + row["ColumnName"] + "].[" + row["ColumnName"] + "].Count SELECT {[Measures].[Count]} ON COLUMNS  FROM [Model]";
        row["ID"] = i + 1;
    }

    cvcq = df;
    CreateExcelSheet(df, "columnvaluescountquery.xlsx");
    return df.Rows.Count;
}


DataTable ColumnValuesCountQuery()
{
    DataTable df = cvcq;
    df.Columns.Add("Count", typeof(int));
    for (int i = 0; i < df.Rows.Count; i++)
    {
        DataRow row = df.Rows[i];

        string query = row["ValuesQuery"].ToString();
        try
        {
            List<string> columnsList = new List<string> { "Count" };
            DataTable tempDF = ExecuteDataTable(query, columnsList);
            df.Rows[i]["Count"] = tempDF.Rows[0]["Count"];
            Console.WriteLine(tempDF.Rows[0]["Count"]);
            Console.WriteLine("Column Values Count queries are running....");

        }
        catch
        {
            df.Rows[i]["Count"] = int.MaxValue;
            Console.WriteLine("Failed Column Values Count queries are running....");
            Debug.WriteLine(query);

        }
    }
    CreateExcelSheet(df, "df.xlsx");
    return df;
}

DataTable FinalColumnsFromTablesQuery()
{
    DataTable ColumnValuesCount = ColumnValuesCountQuery();
    DataTable RowNumberPerDimension = ColumnValuesCount;
    RowNumberPerDimension.DefaultView.Sort = "TableName ASC, Count ASC";
    RowNumberPerDimension = RowNumberPerDimension.DefaultView.ToTable();
    RowNumberPerDimension.Columns.Add("RowNumber", typeof(string));
    var groupedRows = RowNumberPerDimension.AsEnumerable()
    .GroupBy(row => row.Field<string>("TableName"));

    int rowCount = 0;
    foreach (var group in groupedRows)
    {
        int rowNumber = 1;
        foreach (DataRow row in group)
        {
            RowNumberPerDimension.Rows[rowCount]["RowNumber"] = rowNumber++;
            rowCount++;
        }
    }

    Dictionary<string, int> MeanRowNumber = new Dictionary<string, int>();

    foreach (DataRow row in RowNumberPerDimension.Rows)
    {
        string key = row["TableName"].ToString();
        int val = Convert.ToInt32(row["RowNumber"]);

        if (!MeanRowNumber.ContainsKey(key))
        {
            MeanRowNumber.Add(key, val);
        }
        else
        {
            if (val > MeanRowNumber[key])
            {
                MeanRowNumber[key] = val;
            }
        }
    }


    DataTable MeanRowNumberdf = new DataTable();
    MeanRowNumberdf.Columns.Add("TableName", typeof(string));
    MeanRowNumberdf.Columns.Add("MeanRowNumber", typeof(string));

    foreach (KeyValuePair<string, int> kvp in MeanRowNumber)
    {
        DataRow newRow = MeanRowNumberdf.NewRow();
        newRow["TableName"] = kvp.Key;
        newRow["MeanRowNumber"] = Convert.ToInt32(Math.Ceiling(kvp.Value / 2.0));
        MeanRowNumberdf.Rows.Add(newRow);
    }

    DataTable finalColumns = new DataTable();
    finalColumns.Columns.Add("TableName", typeof(string));
    finalColumns.Columns.Add("ColumnName", typeof(string));
    finalColumns.Columns.Add("RowNumber", typeof(string));

    foreach (DataRow row in RowNumberPerDimension.Rows)
    {
        string tableName = row["TableName"].ToString();
        string columnName = row["ColumnName"].ToString();
        int rowNumber = Convert.ToInt32(row["RowNumber"]);

        foreach (DataRow meanRow in MeanRowNumberdf.Rows)
        {
            if (tableName == meanRow["TableName"].ToString() && rowNumber == Convert.ToInt32(meanRow["MeanRowNumber"]))
            {
                finalColumns.Rows.Add(tableName, columnName, rowNumber);
                break;
            }
        }
    }
    CreateExcelSheet(finalColumns, "resultDataFrame.xlsx");
    return finalColumns;
}

DataTable MeasureWithDimensionsQuery()
{
    DataTable TempMeasureCalculationQuery = MeasureListSQLQuery();
    DataTable MeasureReferences = MeasureReferenceQuery();
    DataTable Relationships = RelationshipQuery();
    DataTable Tables = TableQuery();
    DataTable FinalColumnsFromTables = FinalColumnsFromTablesQuery();
    DataTable Columns = ColumnsQuery();

    var tempMeasureCalculationQueryRows = TempMeasureCalculationQuery.AsEnumerable();
    var measureReferencesRows = MeasureReferences.AsEnumerable();
    var relationshipsRows = Relationships.AsEnumerable();
    var tablesRows = Tables.AsEnumerable();
    var finalColumnsFromTablesRows = FinalColumnsFromTables.AsEnumerable();
    var columnsRows = Columns.AsEnumerable();

    var query = from relationship in relationshipsRows
                join fromTable in tablesRows
                    on relationship.Field<string>("FromTableID") equals fromTable.Field<string>("TableID")
                join toTable in tablesRows
                    on relationship.Field<string>("ToTableID") equals toTable.Field<string>("TableID")
                join fromColumn in columnsRows
                    on relationship.Field<string>("FromColumnID") equals fromColumn.Field<string>("ColumnID")
                join toColumn in columnsRows
                    on relationship.Field<string>("ToColumnID") equals toColumn.Field<string>("ColumnID")
                join measureReference in measureReferencesRows
                    on fromTable.Field<string>("TableName") equals measureReference.Field<string>("Referenced_Table")
                join tempMeasureCalculation in tempMeasureCalculationQueryRows
                    on measureReference.Field<string>("Measure") equals tempMeasureCalculation.Field<string>("Measure")
                join finalColumnFromTable in finalColumnsFromTablesRows
                    on toTable.Field<string>("TableName") equals finalColumnFromTable.Field<string>("TableName")
                select new
                {
                    FromTableID = fromTable.Field<string>("TableID"),
                    FromColumnID = fromColumn.Field<string>("ColumnID"),
                    ToTableID = toTable.Field<string>("TableID"),
                    ToColumnID = toColumn.Field<string>("ColumnID"),
                    FromTableName = fromTable.Field<string>("TableName"),
                    ToTableName = toTable.Field<string>("TableName"),
                    FromColumnName = fromColumn.Field<string>("ColumnName"),
                    ToColumnName = toColumn.Field<string>("ColumnName"),
                    Measure = tempMeasureCalculation.Field<string>("Measure"),
                    Referenced_Table = measureReference.Field<string>("Referenced_Table"),
                    MeasureGroup = tempMeasureCalculation.Field<string>("MeasureGroup"),
                    EXPRESSION = tempMeasureCalculation.Field<string>("EXPRESSION"),
                    CubeName = tempMeasureCalculation.Field<string>("CubeName"),
                    TableName = finalColumnFromTable.Field<string>("TableName"),
                    ColumnName = finalColumnFromTable.Field<string>("ColumnName"),
                    RowNumber = finalColumnFromTable.Field<string>("RowNumber"),
                };

    DataTable MeasuresWithDimensions = ConvertToDataTable(query);

    CreateExcelSheet(MeasuresWithDimensions, "MeasuresWithDimensions.xlsx");
    return MeasuresWithDimensions;

}

DataTable MeasureTimeWithoutDimensionsQuery()
{
    DataTable MeasureTimeWithoutDimensions = new DataTable();
    DataTable TempMeasureCalculation = MeasureListSQLQuery();
    MeasureTimeWithoutDimensions.Columns.Add("Measure", typeof(string));
    MeasureTimeWithoutDimensions.Columns.Add("MeasureGroup", typeof(string));
    MeasureTimeWithoutDimensions.Columns.Add("EXPRESSION", typeof(string));
    MeasureTimeWithoutDimensions.Columns.Add("CubeName", typeof(string));
    MeasureTimeWithoutDimensions.Columns.Add("Query", typeof(string));
    MeasureTimeWithoutDimensions.Columns.Add("WithDimension", typeof(string));
    MeasureTimeWithoutDimensions.Columns.Add("DimensionName", typeof(string));
    MeasureTimeWithoutDimensions.Columns.Add("ColumnName", typeof(string));

    foreach (DataRow row in TempMeasureCalculation.Rows)
    {
        MeasureTimeWithoutDimensions.Rows.Add(
        row["Measure"],
        row["MeasureGroup"],
        row["EXPRESSION"],
        row["CubeName"],
        $"SELECT [Measures].[{ row["Measure"]}] ON 0 FROM [{ row["CubeName"]}]",
        0,
        DBNull.Value,
        DBNull.Value
        );

    }
    CreateExcelSheet(MeasureTimeWithoutDimensions, "MeasureTimeWithoutDimensions.xlsx");
    return MeasureTimeWithoutDimensions;
}

DataTable MeasureTimeWithDimensionsQuery()
{
    DataTable MeasureTimeWithDimensions = new DataTable();
    DataTable MeasuresWithDimensions = MeasureWithDimensionsQuery();
    MeasureTimeWithDimensions.Columns.Add("Measure", typeof(string));
    MeasureTimeWithDimensions.Columns.Add("MeasureGroup", typeof(string));
    MeasureTimeWithDimensions.Columns.Add("EXPRESSION", typeof(string));
    MeasureTimeWithDimensions.Columns.Add("CubeName", typeof(string));
    MeasureTimeWithDimensions.Columns.Add("ColumnName", typeof(string));
    MeasureTimeWithDimensions.Columns.Add("DimensionName", typeof(string));
    MeasureTimeWithDimensions.Columns.Add("Query", typeof(string));
    MeasureTimeWithDimensions.Columns.Add("WithDimension", typeof(string));

    foreach (DataRow row in MeasuresWithDimensions.Rows)
    {
        MeasureTimeWithDimensions.Rows.Add(
            row["Measure"],
            row["MeasureGroup"],
            row["EXPRESSION"],
            row["CubeName"],
            row["ColumnName"],
            row["ToTableName"],
            string.Equals(row["ColumnName"].ToString(), "NULL", StringComparison.OrdinalIgnoreCase) ? "" : $"SELECT [{row["Measure"]}] ON 0, NON EMPTY {{[{row["ToTableName"]}].[{row["ColumnName"]}].children}} ON 1 FROM [{row["CubeName"]}]",
            1
        );
    }
    CreateExcelSheet(MeasureTimeWithDimensions, "MeasureTimeWithDimensions.xlsx");
    return MeasureTimeWithDimensions;   
}

DataTable GetLoadTime()
{
    DataTable MeasuresWithDimensions = MeasureTimeWithDimensionsQuery();
    DataTable MeasuresWithoutDimensions = MeasureTimeWithoutDimensionsQuery();

    // Add columns with default values directly
    MeasuresWithDimensions.Columns.Add("LoadTime", typeof(string)).DefaultValue = "x";
    MeasuresWithDimensions.Columns.Add("isMeasureUsedInVisual", typeof(string)).DefaultValue = "0";
    MeasuresWithDimensions.Columns.Add("PageName", typeof(string)).DefaultValue = "-";
    MeasuresWithDimensions.Columns.Add("VisualName", typeof(string)).DefaultValue = "-";
    MeasuresWithDimensions.Columns.Add("VisualTitle", typeof(string)).DefaultValue = "-";
    MeasuresWithDimensions.Columns.Add("ReportName", typeof(string)).DefaultValue = "-";
    MeasuresWithDimensions.Columns.Add("hasDimension", typeof(string)).DefaultValue = "1";

    foreach (DataRow row in MeasuresWithDimensions.Rows)
    {
        row["LoadTime"] = "x";
        row["isMeasureUsedInVisual"] = "0";
        row["PageName"] = "-";
        row["VisualName"] = "-";
        row["VisualTitle"] = "-";
        row["ReportName"] = "-";
        row["hasDimension"] = "1";
    }

    MeasuresWithoutDimensions.Columns.Add("LoadTime", typeof(string)).DefaultValue = "x";
    MeasuresWithoutDimensions.Columns.Add("isMeasureUsedInVisual", typeof(string)).DefaultValue = "0";
    MeasuresWithoutDimensions.Columns.Add("PageName", typeof(string)).DefaultValue = "-";
    MeasuresWithoutDimensions.Columns.Add("VisualName", typeof(string)).DefaultValue = "-";
    MeasuresWithoutDimensions.Columns.Add("VisualTitle", typeof(string)).DefaultValue = "-";
    MeasuresWithoutDimensions.Columns.Add("ReportName", typeof(string)).DefaultValue = "-";
    MeasuresWithoutDimensions.Columns.Add("hasDimension", typeof(string)).DefaultValue = "0";

    foreach (DataRow row in MeasuresWithoutDimensions.Rows)
    {
        row["LoadTime"] = "x";
        row["isMeasureUsedInVisual"] = "0";
        row["PageName"] = "-";
        row["VisualName"] = "-";
        row["VisualTitle"] = "-";
        row["ReportName"] = "-";
        row["hasDimension"] = "0";
    }


    // Set default values for parsedDataFrame
    parsedDataFrame.Columns.Add("LoadTime", typeof(string));
    parsedDataFrame.Columns.Add("isMeasureUsedInVisual", typeof(string));
    parsedDataFrame.Columns.Add("hasDimension", typeof(string));
    foreach (DataRow row in parsedDataFrame.Rows)
    {
        row["LoadTime"] = "x";
        row["isMeasureUsedInVisual"] = "1";
        row["hasDimension"] = "0";
    }

    // Rename columns and add Query column
    parsedDataFrame.Columns["MeasureName"].ColumnName = "Measure";
    parsedDataFrame.Columns.Add("Query", typeof(string)).DefaultValue = ""; ;

    // Populate Query column based on column values
    foreach (DataRow row in parsedDataFrame.Rows)
    {
        if (row["ColumnName"].ToString() == "")
        {
            row["Query"] = $"SELECT [Measures].[{row["Measure"]}] ON 0 FROM [{MeasuresWithDimensions.Rows[0]["CubeName"]}]";
        }
        else
        {
            row["Query"] = $"SELECT [Measures].[{row["Measure"]}] ON 0, NON EMPTY [{row["DimensionName"]}].[{row["ColumnName"]}].Children ON 1 FROM [{MeasuresWithDimensions.Rows[0]["CubeName"]}]";
        }
    }

    // Group parsedDataFrame by Measure and select the first row of each group
    DataTable tempDF = parsedDataFrame.AsEnumerable()
        .GroupBy(r => r.Field<string>("Measure"))
        .Select(g => g.First())
        .CopyToDataTable();
    CreateExcelSheet(tempDF, "tempDF.xlsx");


    Console.WriteLine("Done till here.");



    var leftQuery = (
    from tempRow in tempDF.AsEnumerable()
    join measuresRow in MeasuresWithoutDimensions.AsEnumerable()
    on tempRow.Field<string>("Measure") equals measuresRow.Field<string>("Measure") into temp
    from measuresRow in temp.DefaultIfEmpty()
    where measuresRow == null
    select new
    {
        Measure = tempRow?["Measure"],
        ColumnName_x = tempRow?["ColumnName"],
        DimensionName_x = tempRow?["DimensionName"],
        ReportName_x = tempRow?["ReportName"],
        hasDimension_x = tempRow?["hasDimension"],
        MeasureGroup = measuresRow?["MeasureGroup"], // Add null-conditional operator here
        EXPRESSION = measuresRow?["EXPRESSION"],     
        Query_y = measuresRow?["Query"],             // Add null-conditional operator here
        WithDimension = measuresRow?["WithDimension"],// Add null-conditional operator here
        DimensionName_y = measuresRow?["DimensionName"], // Add null-conditional operator here
        ColumnName_y = measuresRow?["ColumnName"],   // Add null-conditional operator here
        LoadTime_y = measuresRow?["LoadTime"],       // Add null-conditional operator here
        ReportName_y = measuresRow?["ReportName"],   // Add null-conditional operator here
        hasDimension_y = measuresRow?["hasDimension"] // Add null-conditional operator here
    }
);
    // Merge tempDF with MeasuresWithoutDimensions to find missing measures
    var rightQuery = (
    from measuresRow in MeasuresWithoutDimensions.AsEnumerable()
    join tempRow in tempDF.AsEnumerable()
    on measuresRow.Field<string>("Measure") equals tempRow.Field<string>("Measure") into temp
    from tempRow in temp.DefaultIfEmpty()
    where tempRow == null
    select new
    {
        Measure = measuresRow?["Measure"],
        ColumnName_x = tempRow?["ColumnName"],
        DimensionName_x = tempRow?["DimensionName"],
        ReportName_x = tempRow?["ReportName"],
        hasDimension_x = tempRow?["hasDimension"],
        MeasureGroup = measuresRow?["MeasureGroup"], // Add null-conditional operator here
        EXPRESSION = measuresRow?["EXPRESSION"],
        Query_y = measuresRow?["Query"],             // Add null-conditional operator here
        WithDimension = measuresRow?["WithDimension"],// Add null-conditional operator here
        DimensionName_y = measuresRow?["DimensionName"], // Add null-conditional operator here
        ColumnName_y = measuresRow?["ColumnName"],   // Add null-conditional operator here
        LoadTime_y = measuresRow?["LoadTime"],       // Add null-conditional operator here
        ReportName_y = measuresRow?["ReportName"],   // Add null-conditional operator here
        hasDimension_y = measuresRow?["hasDimension"]
    }
);

    var query = leftQuery.Union(rightQuery);
    var mergedDF =  ConvertToDataTable(query);

    CreateExcelSheet(mergedDF, "mergedDF.xlsx");

    mergedDF.Columns["LoadTime_y"].ColumnName = "LoadTime";
    mergedDF.Columns["Query_y"].ColumnName = "Query";

    // Add additional columns to the mergedDF DataTable
    mergedDF.Columns.Add("isMeasureUsedInVisual", typeof(string));
    mergedDF.Columns.Add("PageName", typeof(string));
    mergedDF.Columns.Add("VisualName", typeof(string));
    mergedDF.Columns.Add("VisualTitle", typeof(string));
    mergedDF.Columns.Add("ColumnName", typeof(string));
    mergedDF.Columns.Add("DimensionName", typeof(string));
    mergedDF.Columns.Add("hasDimension", typeof(string));


    // Assign default values to the additional columns in each row
    foreach (DataRow row in mergedDF.Rows)
    {
        row["isMeasureUsedInVisual"] = "0";
        row["PageName"] = "-";
        row["VisualName"] = "-";
        row["VisualTitle"] = "-";
        row["ColumnName"] = "-";
        row["DimensionName"] = "-";
        row["hasDimension"] = "0";
    }


    // Concatenate parsedDataFrame, mergedDF, and MeasuresWithDimensions
    var possibleCombinations = new DataTable();
    possibleCombinations.Merge(parsedDataFrame);
    possibleCombinations.Merge(mergedDF);
    possibleCombinations.Merge(MeasuresWithDimensions);

    // Select required columns from possibleCombinations
    possibleCombinations = possibleCombinations.DefaultView.ToTable(false,
        "Measure", "DimensionName", "ColumnName", "LoadTime", "isMeasureUsedInVisual",
        "ReportName", "PageName", "VisualName", "VisualTitle", "Query", "hasDimension");

    QueryResult = possibleCombinations;

    CreateExcelSheet(possibleCombinations, "possibleCombinations.xlsx");
    return possibleCombinations;

}

var a = ColumnValuesCountQueryforprogress();
GetLoadTime();

    void GetQueryExecutionTime(string query, double thresholdTime, int rowIndex, DataTable allQueries)
    {

            //var con = new AdomdConnection("Provider=MSOLAP.8;Data Source=powerbi://api.powerbi.com/v1.0/myorg/Sprouts EDW;Initial Catalog=UPC Shrink & DOS");
            con.Open();
            var command = new AdomdCommand(query, con);
            int CommandTimeout = (int)(thresholdTime + 1);
            int thresholdTimeMS = (int)(thresholdTime * 1000);
            command.CommandTimeout = CommandTimeout;
            double queryExecutionTime = 0;
            Console.WriteLine($"Currently running {query}");
            try
            {
                DateTime startTime = DateTime.Now;
                var cancellationTokenSource = new CancellationTokenSource();
                var executionTask = Task.Run(() =>
                {
                    using (cancellationTokenSource.Token.Register(() => command.Cancel()))
                    {
                        return command.ExecuteReader();
                    }
                });
        if (!executionTask.Wait(CommandTimeout))
        {
            cancellationTokenSource.Cancel();
            Console.WriteLine($"Query took too long to execute. Aborting query...");

            queryExecutionTime = thresholdTime;
        }
        else
        {
            DateTime endTime = DateTime.Now;
            queryExecutionTime = (endTime - startTime).TotalSeconds;
        }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error executing query: {ex.Message}");
                queryExecutionTime = - 1; // Error occurred
            }
    con.Close();
    allQueries.Rows[rowIndex]["LoadTime"] = queryExecutionTime;

    }


void ExecuteQuery(DataTable allQueries)
{
    ThreadPool.SetMinThreads(5, 5);
    ThreadPool.SetMaxThreads(5, 5);
    double thresholdValue = 0.01;
    List<Task> tasks = new List<Task>();

    for (int i = 0; i < allQueries.Rows.Count; i++)
    {
        string query = allQueries.Rows[i]["Query"].ToString(); // Assuming "Query" is the column name
        int rowIndex = i;
        GetQueryExecutionTime(query, thresholdValue, rowIndex, allQueries);
        //tasks.Add(Task.Run(() => GetQueryExecutionTime(query, thresholdValue, rowIndex, allQueries)));
    }
    //Task.WaitAll(tasks.ToArray());

    CreateExcelSheet(allQueries, "RES.xlsx");
}


ExecuteQuery(QueryResult);