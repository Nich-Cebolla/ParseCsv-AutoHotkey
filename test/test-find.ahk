
#include ..\src\ParseCsv.ahk
#include ..\src\GenerateCSV.ahk

; Not finished

/*

if !A_IsCompiled && A_LineFile = A_ScriptFullPath {
    test_find()
}

class test_find {
    static Call() {

        parseCsvOpt := {
            Breakpoint: ''
          , BreakpointAction: ''
          , Constructor: ''
          , Encoding: 'cp1200'
          , FieldDelimiter: ','
          , FieldsContainRecordDelimiter: false
          , Headers: ''
          , PathIn: ParseCsv.GetPath()
          , QuoteChar: ''
          , RecordDelimiter: '`n'
          , Start: true
        }
        generateCsvOpt := {
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
        for prop in [ 'RecordDelimiter', 'FieldDelimiter', 'QuoteChar', 'Encoding' ] {
            generateCsvOpt.%prop% := parseCsvOpt.%prop%
        }
        generateCsvOpt.RandomRecordDelimiters := parseCsvOpt.FieldsContainRecordDelimiter
        generateCsvOpt.PathOut := parseCsvOpt.PathIn

        csv := GenerateCsv(generateCsvOpt)
        pcsv := ParseCsv(parseCsvOpt)
    }
}
