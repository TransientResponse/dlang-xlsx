module xlsx;

import std.file, std.format, std.regex, std.conv, std.algorithm, std.range, std.utf;
import archive.core, archive.zip;

import dxml.parser;

private struct Coord {
    int row;
    int column;
}

/// Aliases a Sheet as a two-dimensional array of strings.
alias Sheet = string[][];

private enum configSplitYes = makeConfig(SplitEmpty.yes);

/++
Reads a sheet from an XLSX file. 

Params:
fileName = The path of the XLSX file to read
sheetNum = The sheet number to read (one-indexed)

Returns: The contents of the sheet, as a two-dimensional array of strings.
+/
Sheet readSheet(string fileName, int sheetNum) {
    assert(sheetNum > 0);
    
    auto zip = new ZipArchive(read(fileName));
    auto sheet = zip.getFile(format!"xl/worksheets/sheet%d.xml"(sheetNum));
    if(sheet is null) throw new Exception("Invalid sheet number");
    //std.utf.validate!(ubyte[])(sheet.data);
    string xml = cast(string) sheet.data;
    
    validate(xml);

    auto sstFile = zip.getFile("xl/sharedStrings.xml");
    string sstXML = cast(string) sstFile.data;
    string[] sst = parseStringTable(sstXML);
    
    return parseSheetXML(xml, sst);
}

/++
Reads a sheet from an XLSX file by its name.

Params:
fileName = the path of the XLSX file to read
sheetNum = The name of the sheet to read

Returns: the contents of the sheet, as a two-dimensional array of strings.
+/
Sheet readSheetByName(string fileName, string sheetName) {
    auto zip = new ZipArchive(read(fileName));
    auto workbook = zip.getFile("xl/workbook.xml");
    if(workbook is null) throw new Exception("Invalid XLSX file");
    //std.utf.validate!(ubyte[])(sheet.data);
    string xml = cast(string) workbook.data;

    const int id = getSheetId(xml, sheetName);

    return readSheet(fileName, id);
}

/++
Parses an XLSX sheet from XML.

Params:
xmlString = The XML to parse.

Returns: the equivalent sheet, as a two dimensional array of strings.
+/
Sheet parseSheetXML(string xmlString, string[] sst) {
    Sheet temp;
    
    int cols = 0;

    string[] theRow;

    auto range = parseXML!configSplitYes(xmlString);
    while(!range.empty) {
        if(range.front.type == EntityType.elementStart) {
            if(range.front.name == "dimension") {
                auto attr = range.front.attributes.front;
                assert(attr.name == "ref");
                auto dims = parseDimensions(attr.value);
                cols = dims.column;
                assert(cols > 0);
            }
            else if(range.front.name == "row") {
                assert(cols > 0);
                theRow = new string[cols];
            }
            else if(range.front.name == "c") {
                auto attrs = range.front.attributes;
                Coord loc;
                bool isref;
                foreach(attr; attrs) {
                    if(attr.name == "r") {
                        loc = parseLocation(attr.value);
                    }
                    else if(attr.name == "t") {
                        if(attr.value == "s") isref=true;
                        else isref = false;
                    }
                }
                assert(loc.row >= 0 && loc.column >= 0);
                range.popFront;
                if(range.front.type == EntityType.elementStart) {
                    if(range.front.name == "f") {
                        range.popFront; 
                        if(range.front.type == EntityType.text) range.popFront; 
                        range.popFront;
                    }
                    range.popFront;
                    assert(range.front.type == EntityType.text);
                    string text = range.front.text;
                    assert(theRow.length > 0);
                    assert(loc.column < theRow.length);
                    if(isref) theRow[loc.column] = sst[parse!int(text)];
                    else theRow[loc.column] = text;
                }
                else {
                    theRow[loc.column] = "";
                }
            }
        }
        else if(range.front.type == EntityType.elementEnd) {
            if(range.front.name == "row") {
                temp ~= theRow;
            }
        }
        range.popFront;
    }
    
    
    return temp;
}
unittest {
    const string test = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><dimension ref="A1:C3"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A4" sqref="A4"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/><sheetData><row r="1" spans="1:3" x14ac:dyDescent="0.25"><c r="A1"><v>1</v></c><c r="B1"><v>5</v></c><c r="C1"><v>7</v></c></row><row r="2" spans="1:3" x14ac:dyDescent="0.25"><c r="A2"><v>2</v></c><c r="B2"><v>4</v></c><c r="C2"><f t="shared" si="0" /><v>3</v></c></row><row r="3" spans="1:3" x14ac:dyDescent="0.25"><c r="A3"><v>7</v></c><c r="B3"><v>82</v></c><c r="C3"><v>1</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
    const Sheet testSheet = [["1","5","7"],["2","4","3"],["7","82","1"]];
    const Sheet result = parseSheetXML(test, null);
    assert(result == testSheet);
}
unittest {
    const string test = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><dimension ref="A1:C6"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A7" sqref="A7"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/><sheetData><row r="1" spans="1:3" x14ac:dyDescent="0.25"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row><row r="2" spans="1:3" x14ac:dyDescent="0.25"><c r="A2"><v>1</v></c><c r="B2" s="1"><f>ABS(-6.2)</f><v>6.2</v></c><c r="C2" s="1"/></row><row r="3" spans="1:3" x14ac:dyDescent="0.25"><c r="A3"><v>2</v></c><c r="B3" s="1"><v>3.4</v></c><c r="C3" s="1"/></row><row r="4" spans="1:3" x14ac:dyDescent="0.25"><c r="A4"><v>3</v></c><c r="B4" s="1"><v>87.1</v></c><c r="C4" s="1"/></row><row r="5" spans="1:3" x14ac:dyDescent="0.25"><c r="A5"><v>4</v></c><c r="B5" s="2"><v>83.2</v></c><c r="C5" s="2"/></row><row r="6" spans="1:3" x14ac:dyDescent="0.25"><c r="A6"><v>5</v></c><c r="B6" s="1"><v>3</v></c><c r="C6" s="1"/></row></sheetData><mergeCells count="1"><mergeCell ref="B5:C5"/></mergeCells><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
    const Sheet expected = [["Param", "Value", ""], ["1", "6.2", ""], ["2", "3.4", ""], ["3", "87.1", ""], ["4", "83.2", ""], ["5", "3", ""]];
    const Sheet result = parseSheetXML(test, ["Param", "Value"]);
    assert(result == expected);
}

/++
Gets the numeric (> 1) id of a sheet with a given name.

Params: 
wbXml = The XML content of the workbook.xml file in the main XLSX zip archive.
sheetName = The name of the sheet to find.

Returns: the id of the given sheet

Exceptions:
Exception if the given sheet name is not in the workbook.
+/
private int getSheetId(string wbXml, string sheetName) {
    auto range = parseXML!configSplitYes(wbXml);
    while(!range.empty) {
        if(range.front.type == EntityType.elementStart && range.front.name == "sheet") {
            auto attrs = range.front.attributes;
            bool nameFound;
            foreach(attr; attrs) {
                if(attr.name == "name" && attr.value == sheetName) nameFound = true;
                else if(nameFound && attr.name == "sheetId") {
                    return parse!int(attr.value);
                }
            }
        } 
        range.popFront;
    }
    
    throw new Exception("No sheet with that name!");
}
unittest {
    const string test = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"><fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="18625"/><workbookPr defaultThemeVersion="166925"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x15"><x15ac:absPath url="C:\Users\rraab.ADVILL\Documents\Code\D\dlang-xlsx\" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice></mc:AlternateContent><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12210" activeTab="1" xr2:uid="{045295F2-59E2-495E-BB00-149A1C289780}"/></bookViews><sheets><sheet name="Test1" sheetId="1" r:id="rId1"/><sheet name="Test2" sheetId="2" r:id="rId2"/></sheets><calcPr calcId="171027"/><extLst><ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:workbookPr chartTrackingRefBase="1"/></ext></extLst></workbook>`;
    assert(getSheetId(test, "Test1") == 1);
    assert(getSheetId(test, "Test2") == 2);
}

/++
Parses a Shared String Table XML string into an array of strings. 

Params:
sst = The Shared String Table XML string to parse
Returns: 
A simple string array that can be indexed into from an "s" type cell's value (which is a zero-based index, thankfully)
+/
private string[] parseStringTable(string sst) {
    //auto doc = new DocumentParser(sst);
    string[] table;

    auto range = parseXML!configSplitYes(sst);

    while(!range.empty) {
        if(range.front.type == EntityType.elementStart && range.front.name == "t") {
            range.popFront;
            table ~= range.front.text;
        }
        range.popFront;
    }
    return table;
}

/++
Gets a column from a sheet, with each element parsed as the templated type.

Params:
sheet = The Sheet from which to extract the column.
col = The (zero-based) index of the column to extract.

Returns: the given column, as a newly-allocated array.
+/
T[] getColumn(T)(Sheet sheet, size_t col) {
    return sheet.map!(x => parse!T(x[col])).array;
}
unittest {
    assert(getColumn!int([["1","5","7"],["2","4","3"],["7","82","1"]], 1) == [5, 4, 82]);
}

///Parses an Excel letter-number column-major one-based location string to a pair of row-major zero-based indeces
private Coord parseLocation(string location) {
    auto pat = ctRegex!"([A-Z]+)([0-9]+)";
    auto m = location.matchFirst(pat);
    Coord temp;
    string rrow = m[2];
    temp.row = parse!int(rrow)-1;
    temp.column = m[1].columnNameToNumber-1;
    return temp;
}
unittest {
    const Coord C3 = {2,2};
    const Coord B15 = {14,1};
    assert(parseLocation("C3") == C3);
    assert(parseLocation("B15") == B15);
}

private Coord parseDimensions(string dims) {
    auto pat = ctRegex!"([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)";
    auto m = dims.matchFirst(pat);
    Coord temp;
    string m4 = m[4];
    string m2 = m[2];
    temp.row = parse!int(m4) - parse!int(m2) + 1;
    temp.column = m[3].columnNameToNumber - m[1].columnNameToNumber + 1;
    return temp;
}

private int columnNameToNumber(string col) {
    reverse(col.dup);
    int num;
    foreach(i, c; col) {
        num += (c - 'A' + 1)*26^^i;
    }
    return num;
}

unittest {
    assert(columnNameToNumber("AA") == 27);
}
