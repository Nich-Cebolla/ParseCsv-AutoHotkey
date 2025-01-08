

class ParseCSVConfig {
    static Start := true
    static PathIn := ''
    static Encoding := ''
    static FieldDelimiter := ','
    static RecordDelimiter := '' ; if blank, general newline characters are used
    static Constructor := '' ; if blank, default constructor is used, else the `Record` array is passed to the constructor
    static Headers := '' ; if blank, the first record is used
    static MaxReadSizeBytes := 0 ; Only needed if there is a memory constraint. Is ignored when `QuoteChar` is false and `RecordDelimiter` is blank because the file is read by line anyway.
    static QuoteChar := '' ; Set to the character used to quote strings, or an empty string if fields are not quoted.
    static Breakpoint := 0 ; Number of records parsed before the `BreakpointAction` is invoked.
    static BreakpointAction := '' ; Set to a function object to call the function when the breakpoint is reached. Else, if Breakpoint has a value and this is not a Func, `ParseCsv` will return `this` (instance object).
    static CollectionArrayBuffer := 10000 ; A buffer so the `Record` array doesn't need to be constantly resized.
}