# ParseCsv-AutoHotkey
An all-purpose, easy-to-use CSV parser.

# Introduction
I haven't written the documentation for this yet, but it is very easy to use. Set the parameters by using the ParseCsvConfig.ahk file, or passing an object to `ParseCsv()`. Then call `result := ParseCsv()`, and the records will be on the `result.Collection` object. `result.Collection` is a composition of `Array`, and the objects contained in the array are each compositions of `Map`, so you can access the objects and items in the usual ways:
```ahk
ParseCsvConfig {
 ; params here
}
result := ParseCsv()
collection := result.Collection
for record in collection {
  for header, field in record {
    MsgBox('Item ' A_Index ' - ' header ': ' field)
  }
}

; individual items
MsgBox(collection[5]['One Of The Headers'])

; If you prefer `object.path` notation, you can access the values using that notation instead of `map['bracket']` notation. (If the header has spaces, you can replace those with underscores to use object notation, but if the header has any other invalid characters in it, you cannot use object notation for the associated value on any of the objects).
MsgBox(collection[11].One_Of_The_Headers == collection[11]['One Of The Headers']) ; 1
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
