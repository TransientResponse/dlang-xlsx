# dlang-xlsx
An XLSX (Excel 2007+) sheet extractor/parser written in D

Really defines just one useful function: `readSheet(string fileName, int sheet)`. See the docs folder for specifics. 

Build with `dub build`. I would have written it with only the D standard library (Phobos), but std.zip has issues with `expandedData`
on members of XLSX files returning empty (zero-length) byte arrays. The Dub archive package, on the other hand, works perfectly.

Example usage:

```
import std.stdio;
import xlsx;

void main() {
    //Reads sheet 1 from file "test.xlsx"
    writeln(readSheet("test.xlsx", 1));
    
    //Read a named sheet
    writeln(readSheet("test.xlsx", "My Sheet"));
}
```

As of version 0.0.4 now properly reads from the Shared String Table for spreadsheets with many oft-repeating strings and/or whenever Excel decides to make use of it.

Tested on Windows, but written purely in D with no external dependencies, so it should run on all platforms D supports.
