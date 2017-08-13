module xlsx;

import std.file, std.xml, std.format, std.regex, std.conv, std.algorithm, std.range, std.utf;
import archive.core, archive.zip;

private struct Coord {
    int row;
    int column;
}

/// Aliases a Sheet as a two-dimensional array of strings.
alias Sheet = string[][];

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
    
    return parseSheetXML(xml);
}

/++
Parses an XLSX sheet from XML.

Params:
xmlString = The XML to parse.

Returns: the equivalent sheet, as a two dimensional array of strings.
+/
Sheet parseSheetXML(string xmlString) {
    Sheet temp;
    
    int cols = 0;
    auto doc = new DocumentParser(xmlString);
    doc.onEndTag["dimension"] = (in Element dim) {
        auto dims = parseDimensions(dim.tag.attr["ref"]);
        //temp ~= new string[dims.row];
        cols = dims.column;
    };
    
    doc.onStartTag["row"] = (ElementParser rowTag) {
        //int r = parse!int(rowTag.tag.attr["r"])-1;
        
        auto theRow = new string[cols];
        rowTag.onStartTag["c"] = (ElementParser cTag) {
            Coord loc = parseLocation(cTag.tag.attr["r"]);
            string val;
            cTag.onEndTag["v"] = (in Element v) { val = v.text; };
            cTag.parse();

            theRow[loc.column] = val;
        };
        rowTag.parse();
        temp ~= theRow;
    };
    doc.parse();
    return temp;
}
unittest {
    const string test = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><dimension ref="A1:C3"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A4" sqref="A4"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/><sheetData><row r="1" spans="1:3" x14ac:dyDescent="0.25"><c r="A1"><v>1</v></c><c r="B1"><v>5</v></c><c r="C1"><v>7</v></c></row><row r="2" spans="1:3" x14ac:dyDescent="0.25"><c r="A2"><v>2</v></c><c r="B2"><v>4</v></c><c r="C2"><v>3</v></c></row><row r="3" spans="1:3" x14ac:dyDescent="0.25"><c r="A3"><v>7</v></c><c r="B3"><v>82</v></c><c r="C3"><v>1</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
    Sheet testSheet = [["1","5","7"],["2","4","3"],["7","82","1"]];
    Sheet result = parseSheetXML(test);
    assert(result == testSheet);
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
    Coord C3 = {2,2};
    Coord B15 = {14,1};
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
    int num = 0;
    foreach(i, c; col) {
        num += (c - 'A' + 1)*26^^i;
    }
    return num;
}

unittest {
    assert(columnNameToNumber("AA") == 27);
}
