#Include ParseCsv.ahk
#Include GenerateCSV.ahk

GenerateCsvConfig.QuoteChar := ParseCsvConfig2.QuoteChar
GenerateCsvConfig.FieldDelimiter := ParseCsvConfig2.FieldDelimiter
GenerateCsvConfig.RecordDelimiter := ParseCsvConfig2.RecordDelimiter
GenerateCsvConfig.Headers := ParseCsvConfig2.Headers
GenerateCsvConfig.OtherChars := StrReplace(GenerateCsvConfig.OtherChars, ParseCsvConfig2.RecordDelimiter, '')


result := ParseCsv(ParseCsvConfig2, GenerateCsv())



class ParseCSVConfig2 {
    static Start := true
    static PathIn := ''
    static Encoding := ''
    static FieldDelimiter := ','
    static RecordDelimiter := '' ; if blank, general newline characters are used
    static Constructor := '' ; if blank, default constructor is used, else the `Record` array is passed to the constructor
    static Headers := '' ; if blank, the first record is used
    static MaxReadSizeBytes := 0 ; Only needed if there is a memory constraint. Is ignored when `QuoteChar` is false and `RecordDelimiter` is blank because the file is read by line anyway.
    static QuoteChar := '"' ; Set to the character used to quote strings, or an empty string if fields are not quoted.
    static Breakpoint := 0 ; Number of records parsed before the `BreakpointAction` is invoked.
    static BreakpointAction := '' ; Set to a function object to call the function when the breakpoint is reached. Else, set to any non-falsy value to return the `ParseCsv` instance.
    static CollectionArrayBuffer := 10000 ; A buffer so the `Record` array doesn't need to be constantly resized.
}

class GenerateCsvConfig {
    static PathOut := ''
    static QuoteChar := '"'
    static FieldDelimiter := ','
    static RecordDelimiter := '`n'
    static Columns := 10
    static Rows := 100
    static MinWordsPerField := 0
    static MaxWordsPerField := 10
    static Headers := ''
    static Encoding := ''
    static OtherChars := '1234567890+_)(*&^%$#@!~``;/\|><.?'
    static OtherCharsProbability := 0.3 ; The probability a group of other chars is used instead of a word
    static MinOtherCharsPerGroup := 1 ; The minimum number of other chars in a group
    static MaxOtherCharsPerGroup := 3 ; The maximum number of other chars in a group
    static OtherCharsOnlyInQuotedStrings := false
    static LineEndings := '`n'
    static RandomEscapedQuotes := true
    static RandomFieldDelimiters := true
    static RandomItemDelimiters := true
    static RandomLineEndings := true
}