
#include ..\src\ParseCsv.ahk

if !A_IsCompiled && A_LineFile = A_ScriptFullPath {
    test(true)
}

class test {
    static __New() {
        this.DeleteProp('__New')
        this.Headers := [
            [ 'A', 'AAA', 'AAAAAAAAA', 'AAAAAAAAAAAAAAAAAAAAAAAAAAA', 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA' ]
          , [ '{1}A{1}', '{1}AAA{1}', '{1}AAAAAAAAA{1}', '{1}AAAAAAAAAAAAAAAAAAAAAAAAAAA{1}', '{1}AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA{1}' ]
          , [ '{1}A{1}', 'AAA', '{1}AAAAAAAAA{1}', 'AAAAAAAAAAAAAAAAAAAAAAAAAAA', '{1}AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA{1}' ]
          , [ 'A', '{1}AAA{1}', 'AAAAAAAAA', '{1}AAAAAAAAAAAAAAAAAAAAAAAAAAA{1}', 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA' ]
        ]
        this.Encodings := [ 'cp1200', 'utf-8' ]
    }
    static Call(showTooltip := false, fieldDelimiter := ',', recordDelimiter := '`n', quoteChar := '"') {
        path := this.GetPath()
        parseCsvOpt := {
            Breakpoint: 0
          , BreakpointAction: ''
          , Constructor: ''
          , Encoding: ''
          , FieldDelimiter: fieldDelimiter
          , FieldsContainRecordDelimiter: true
          , Headers: ''
          , PathIn: path
          , QuoteChar: quoteChar
          , RecordDelimiter: recordDelimiter
          , SuppressErrorWindow: false
          , Start: false
        }
        for fieldsContainRecordDelimiterValue in [ true, false ] {
            parseCsvOpt.FieldsContainRecordDelimiter := fieldsContainRecordDelimiterValue
            for encoding in this.Encodings {
                parseCsvOpt.Encoding := encoding
                OutputDebug('Encoding: ' encoding '`n')
                for headers in this.Headers {
                    OutputDebug('Headers array: ' A_Index '`n')
                    line := ''
                    loop headers.Length {
                        line .= headers[A_Index] fieldDelimiter
                    }
                    line := SubStr(Format(line, quoteChar), 1, -StrLen(fieldDelimiter)) recordDelimiter
                    f := FileOpen(path, 'w', encoding)
                    f.Write(line)
                    f.Close()
                    parseCsvObj := ParseCsv(parseCsvOpt)
                    parseCsvObj.ReadChars := 1
                    parseCsvObj.__GetPattern()
                    parseCsvObj.__PrepareContent()
                    _headers := parseCsvObj.Headers
                    loop headers.Length {
                        if Format(headers[A_Index], quoteChar) != _headers[A_Index] {
                            throw Error('Mismatched headers.', , 'Actual: ' headers[A_Index] '; Matched: ' _headers[A_Index])
                        }
                    }
                }
            }
            OutputDebug('Done`n')
        }
    }
    static GetPath() {
        s := ''
        loop 10 {
            s .= Chr(Random(65, 90))
        }
        path := A_Temp '\' s '.csv'
        while FileExist(path) {
            s .= Chr(Random(65, 90))
            path := A_Temp '\' s '.csv'
        }
        OnExit(_DeleteFile.Bind(path), 1)

        return path

        _DeleteFile(path, *) {
            if FileExist(path) {
                try {
                    FileDelete(path)
                }
            }
        }
    }
}
