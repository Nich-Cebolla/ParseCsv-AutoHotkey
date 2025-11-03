
#include ..\src\ParseCsv.ahk
#include ..\src\GenerateCSV.ahk

if !A_IsCompiled && A_LineFile = A_ScriptFullPath {
    test_misc()
}

class test_misc {
    static __New() {
        this.DeleteProp('__New')
        this.parseCsvOpt := {
            Breakpoint: 9
          , BreakpointAction: test_BreakpointAction()
          , Constructor: ''
          , Encoding: 'cp1200'
          , FieldDelimiter: ','
          , FieldsContainRecordDelimiter: true
          , Headers: ''
          , PathIn: ParseCsv.GetPath()
          , QuoteChar: '"'
          , RecordDelimiter: '`n'
          , Start: true
        }
        this.generateCsvOpt := {
            Columns: 10
          , Encoding: 'cp1200'
          , FieldDelimiter: ','
          , Headers: ''
          , LineEnding: '`n'
          , MaxOtherCharsPerGroup: 3 ; The maximum number of other chars in a group
          , MaxWordsPerField: 10
          , MinOtherCharsPerGroup: 1 ; The minimum number of other chars in a group
          , MinWordsPerField: 0
          , NoHeaders: false
          , OtherChars: '1234567890+_)(*&^%$#@!~``|><?'
          , OtherCharsOnlyInQuotedStrings: false
          , OtherCharsProbability: 0.3 ; The probability a group of other chars is used instead of a word
          , OutputDisplayStr: false
          , Overwrite: true
          , PathOut: '..\.dev\out.csv'
          , ProbabilityQuotedString: 0.6
          , QuoteChar: '"'
          , RandomEscapedQuotes: true
          , RandomFieldDelimiters: true
          , RandomLineEnding: true
          , RandomRecordDelimiters: true
          , RecordDelimiter: '`n'
          , Rows: 100
          , ShowTooltip: false
        }
    }
    static Call(parseCsvOpt?, generateCsvOpt?) {
        if !IsSet(parseCsvOpt) {
            parseCsvOpt := this.parseCsvOpt
        }
        if !IsSet(generateCsvOpt) {
            generateCsvOpt := this.generateCsvOpt
        }
        this.Resume(parseCsvOpt, generateCsvOpt)
    }
    static Resume(parseCsvOpt?, generateCsvOpt?) {
        if !IsSet(parseCsvOpt) {
            parseCsvOpt := this.parseCsvOpt
        }
        if !IsSet(generateCsvOpt) {
            generateCsvOpt := this.generateCsvOpt
        }
        this.PrepareOptions(parseCsvOpt, generateCsvOpt)
        csv := GenerateCsv(generateCsvOpt)
        pcsv := ParseCsv(parseCsvOpt)
        while !pcsv.Complete {
            if A_Index > Ceil(generateCsvOpt.Rows / parseCsvOpt.Breakpoint) {
                throw Error('Too many loops.', , A_Index)
            }
            OutputDebug('Loop ' A_Index '; Index: ' pcsv.Index '; ParsedChars: ' pcsv.ParsedChars '`n')
            pcsv()
        }
    }
    static PrepareOptions(parseCsvOpt, generateCsvOpt) {
        for prop in [ 'RecordDelimiter', 'FieldDelimiter', 'QuoteChar', 'Encoding' ] {
            generateCsvOpt.%prop% := parseCsvOpt.%prop%
        }
        generateCsvOpt.RandomRecordDelimiters := generateCsvOpt.RandomLineEnding := parseCsvOpt.FieldsContainRecordDelimiter
        generateCsvOpt.PathOut := parseCsvOpt.PathIn
    }
}

class test_BreakpointAction {
    __New() {
        this.Index := 0
    }
    Call(pcsv) {
        OutputDebug('breakpoint ' (++this.Index) '; progress: ' pcsv.GetProgress() '`n')
        return 1
    }
}
