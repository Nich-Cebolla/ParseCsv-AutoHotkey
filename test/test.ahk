
#include ..\src\ParseCsv.ahk
#include ..\src\GenerateCSV.ahk

if !A_IsCompiled && A_LineFile = A_ScriptFullPath {
    test({
        quoteChar: [ '"', "'", '' ]
      , breakpoint: [ 0, 10 ]
      , row: [ 21 ]
      , breakpointAction: [ test_BreakpointAction() ]
      ; If you change the fieldDelimiter or recordDelimiter values, you might also need to change
      ; the value of generateCsvOpt.OtherChars (remove any characters that are used by fieldDelimiter
      ; or recordDelimiter).
      , fieldDelimiter: [ ',', ',,', ',./' ]
      , recordDelimiter: [ '`n', '`r`n', ';', ';;]', ';\-' ]
      , fieldsContainRecordDelimiter: [ false ]
    })
}

class test {
    static __New() {
        this.DeleteProp('__New')
    }
    /**
     * @param {Object} options
     * @param {String[]} options.quoteChar - An array of quote characters.
     * @param {Integer[]} options.breakpoint - An array of breakpoint values.
     * @param {Integer[]} options.row - An array of row (number of records) values.
     * @param {*[]} options.breakpointAction - An array of callable objects.
     * @param {String[]} options.fieldDelimiter - An array of field delimiters.
     * @param {String[]} options.recordDelimiter - An array of record delimiters.
     * @param {Boolean[]} options.fieldsContainRecordDelimiter - An array of true / false.
     */
    static Call(options) {

        parseCsvOpt := {
            Breakpoint: 0
          , BreakpointAction: ''
          , Constructor: ''
          , Encoding: 'cp1200'
          , FieldDelimiter: ','
          , FieldsContainRecordDelimiter: false
          , Headers: ''
          , PathIn: ''
          , QuoteChar: ''
          , RecordDelimiter: ''
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
          , Rows: 1000
          , ShowTooltip: false
        }
        path := ParseCsv.GetPath()
        this.OnExitFunc := _DeleteFile.Bind(path)
        OnExit(this.OnExitFunc, 1)
        generateCsvOpt.PathOut := parseCsvOpt.PathIn := path
        i := 0
        for quoteChar in options.quoteChar {
            generateCsvOpt.quoteChar := parseCsvOpt.quoteChar := quoteChar
            for breakpoint in options.breakpoint {
                parseCsvOpt.breakpoint := breakpoint
                for row in options.row {
                    generateCsvOpt.Rows := row
                    for breakpointAction in options.breakpointAction {
                        parseCsvOpt.breakpointAction := breakpointAction
                        for fieldDelimiter in options.fieldDelimiter {
                            generateCsvOpt.fieldDelimiter := parseCsvOpt.fieldDelimiter := fieldDelimiter
                            for recordDelimiter in options.recordDelimiter {
                                generateCsvOpt.recordDelimiter := parseCsvOpt.recordDelimiter := recordDelimiter
                                for fieldsContainRecordDelimiter in options.fieldsContainRecordDelimiter {
                                    ++i
                                    OutputDebug(i '`n')
                                    generateCsvOpt.RandomRecordDelimiters := generateCsvOpt.RandomLineEnding := parseCsvOpt.fieldsContainRecordDelimiter := fieldsContainRecordDelimiter
                                    csv := GenerateCsv(generateCsvOpt)
                                    pcsv := ParseCsv(parseCsvOpt)
                                    if quoteChar && pcsv.ParsedChars != csv.ContentLen {
                                        throw Error('Invalid content length.', , 'Generated: ' csv.ContentLen '; Parsed: ' pcsv.ParsedChars)
                                    }
                                    if csv.Headers.Length != pcsv.Headers.Length {
                                        throw Error('Invalid number of headers.', , 'Generated: ' csv.Headers.Length '; Parsed: ' pcsv.Headers.Length)
                                    }
                                    loop csv.Headers.Length {
                                        if csv.Headers[A_Index] != pcsv.Headers[A_Index] {
                                            throw Error('Mismatched headers.', , 'Generated: ' csv.Headers[A_Index] '; Parsed: ' pcsv.Headers[A_Index])
                                        }
                                    }
                                    if pcsv.Length != csv.Records.Length {
                                        throw Error('Invalid number of records.', , 'Generated: ' csv.Records.Length '; Parsed: ' pcsv.Length)
                                    }
                                    loop pcsv.Length {
                                        parsed := pcsv[A_Index]
                                        generated := csv.Records[A_Index]
                                        loop parsed.Length {
                                            if parsed[A_Index] != generated[A_Index] {
                                                throw Error('Mismatched fields.', , 'Generated: ' generated[A_Index] '; Parsed: ' parsed[A_Index])
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        return

        _DeleteFile(path, *) {
            if FileExist(path) {
                try {
                    FileDelete(path)
                }
            }
        }
    }
}

class test_BreakpointAction {
    __New() {
        this.Index := 0
    }
    Call(pcsv) {
        OutputDebug('breakpoint ' (++this.Index) '; progress: ' pcsv.GetProgress() '`n')
    }
}
