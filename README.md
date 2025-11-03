# ParseCsv-AutoHotkey

An all-purpose, easy-to-use CSV parser.

# Introduction

See the documentation in the code file.

# v2.0.0

This rewrite mostly impacts internal logic. Most importantly, the performance when parsing csv
with quoted fields has been improved significantly. There are also changes to some options and
methods. I updated the documentation in src\ParseCsv.ahk to be accurate to v2.0.0.

# Tested methods

Items marked with X have been tested and verified to work as expected.

|  Name                     |  Is Tested  |
|  -------------------------|-----------  |
|  Call                     |      X      |
|  Find                     |             |
|  FindAll                  |             |
|  FindAllF                 |             |
|  FindAllR                 |             |
|  FindF                    |             |
|  FindR                    |             |
|  GetProgress              |      X      |
|  ToAhkCode                |      X      |

# Changelog

- **2025-11-03** - v2.0.0 **breaking**
  - Rewrote core logic.

- **2025-04-03**
  - Fixed bug: The function that is passed to the `Constructor` parameter was intended to receive the `ParseCsv` instance object as the second parameter, but it was not.

- **2025-03-08**
  - Changes to ParseCsv.Collection.Prototype:
    - Removed `__Add`, `__MakeRecord`, and `SetHeaders`, and removed associated code from `__New`.
  - Changes to ParseCsv.Prototype:
    - Rearranged order of methods in script.
    - Added `__Add`, `__MakeRecord`, and `__SetHeaders`. ParseCsv.Prototype now handles the creation of record objects.
    - Added `__ThrowInvalidInputError`. This provides information and context when a RegExMatch matches at an invalid position. Wherever RegExMatch is used to parse the input content, the function checks the position of the match to validate it. If the match is invalid, this method is called. It will check the line endings and compare them to the record delimiter, then throw an error with the details. Additional context is passed to `OutputDebug`.
    - Adjusted `__CheckLineEndings` to evaluate if the error described above is likely caused by an incorrect line ending value used as the record delimiter.
    - Added `Destroy`. This clears the collection array and deletes some properties.
    - Corrected the `LoopReadLine` method. Previously it did not handle breakpoints correctly.  Breakpoints now work as expected.
    - Adjusted the three `Find` methods to include a `OutRecord` var ref parameter, so the record object can be obtained during the function call.
    - Adjusted the three `FindAll` methods to include a `IncludeRecords` parameter, to indicate whether the records should be included in the result object. Also refactored the function code.
    - Adjusted `End`. When the content has been parsed, the property `Complete` is set to `1`.
    - Minor changes: Fixed a typo: "ReadStle" -> "ReadStyle", Removed some unneeded lines, corrected an invalid method call, made some minor optimizations.
  - Changes to ParseCsv:
    - Added `SetMapCaseSense`. This can be used to set the initial CaseSense value of all new map objects. This is intended to be used when one wants the record objects to not be case sensitive.
    - Added `__ReplaceNewLines`.
    - ParseCsv no longer has a `Headers` property assigned. Headers are accessible from the instance object.
  - Changes to ParseCsv.Params:
    - Removed `DisableLineEndingsCheck`. The line endings check always occurs when the error is encountered, but the information is conveyed via the error message and via OutputDebug.
