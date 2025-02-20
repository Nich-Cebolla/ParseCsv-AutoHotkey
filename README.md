# ParseCsv-AutoHotkey
An all-purpose, easy-to-use CSV parser.

# Introduction
I haven't written the documentation for this yet, but it is very easy to use. Set the parameters by using the ParseCsvConfig.ahk file, Then call `result := ParseCsv()`, or passing an object with the params to the first parameter `result := ParseCsv(ParamsObject)`.  The records will be on the `result.Collection` object. `result.Collection` is a composition of `Array`, and the objects contained in the array are each compositions of `Map`, so you can access the objects and items in the usual ways:
```ahk
CsvPath := 'C:\Users\MyName\Downloads\MyCsv.csv'

; Input parameters are passed via an object (or the config file).
; Minimally you will also need to set `FieldDelimiter` and `RecordDelimiter`. "Field" would be
; synonymous with a cell in a table; the standard "Field" delimiter is the comma.
; "Record" would be synonymous with a row in a table, the standard delimiter being a newline.
; If your CSV has quoted strings, you must set the `QuoteChar` option to the quote character used.
; If your CSV does not have quoted strings, leave `QuoteChar` unset.
; Let's say this example CSV has quoted strings, is delimited by semicolons, and each record is
; separated by LFCR
Opts := {
    FieldDelimiter: ';',
    RecordDelimiter: '`r`n',
    PathIn: CsvPath,
    QuoteChar: '"'
}
; There's other customization options but for basic usage this is all that's needed.

; Now we parse the CSV
Result := ParseCsv(Opts)
; Assuming everything worked as intended, the `Result` object is an object you can use to handle the parsed content.

for Record in Result {
    for HeaderName, Value in Record {
        MsgBox(HeaderName '`r`n' Value)
    }
}

MsgBox(Result[2]['SomeHeader']) ; Line #2, arbitrary header name
```

The objects contained in the `Collection` are `ParseCsv.Record` objects, which are basically maps. They are case sensitive. Attempting to access a value using incorrect case, even using `object.path` notation, will result in an error. You can set all map objects to not be case sensitive by default by including the below code somewhere in your script before the script parses the CSV.

```
Map.Prototype.DefineProp('__New', {Call: NewConstructor})
NewConstructor(self, items*) {
  self.CaseSense := false
  if items.Length
    self.Set(items*)
}
```
