
#SingleInstance force
#Include ..\src\GenerateCSV.ahk

; This creates a directory .dev in the parent directory.

if !A_IsCompiled && A_LineFile = A_ScriptFullPath {
    test(true)
}

class test {
    static Call(showTooltip := false) {
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
          , OtherChars: '1234567890+_)(*&^%$#@!~``;/\|><.?'
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
          , ShowTooltip: true
        }
        if !DirExist('..\.dev') {
            DirCreate('..\.dev')
        }

        p1 := this.Pattern1 := (
            'S)'
            '(?<={2}|^)'
            '(?<value>'
                '{3}{3}'
            '|'
                '{3}(*COMMIT)'
                '[^{3}]*+'
                '(?:'
                    '(?:{3}{3})*+'
                '|'
                    '(?:'
                        '(?:{3}{3})*+'
                        '[^{3}]*+'
                    ')++'
                ')'
                '{3}'
            '|'
                '[^{1}{2}]*+'
            ')'
            '{1}'
        )
        p2 := this.Pattern2 := (
            'S)'
            '(?<={1})'
            '(?<value>'
                '{3}{3}'
            '|'
                '{3}(*COMMIT)'
                '[^{3}]*+'
                '(?:'
                    '(?:{3}{3})*+'
                '|'
                    '(?:'
                        '(?:{3}{3})*+'
                        '[^{3}]*+'
                    ')++'
                ')'
                '{3}'
            '|'
                '[^{1}{2}]*+'
            ')'
            '{1}'
        )
        p3 := this.Pattern3 := (
            'S)'
            '(?<={1})'
            '(?<value>'
                '{3}{3}'
            '|'
                '{3}(*COMMIT)'
                '[^{3}]*+'
                '(?:'
                    '(?:{3}{3})*+'
                '|'
                    '(?:'
                        '(?:{3}{3})*+'
                        '[^{3}]*+'
                    ')++'
                ')'
                '{3}'
            '|'
                '[^{1}{2}]*+'
            ')'
            '(?:{2}|$)'
        )
        _p1 := this._p1 := Format(p1, GenerateCsvConfig.FieldDelimiter, GenerateCsvConfig.RecordDelimiter, GenerateCsvConfig.QuoteChar)
        _p2 := this._p2 := Format(p2, GenerateCsvConfig.FieldDelimiter, GenerateCsvConfig.RecordDelimiter, GenerateCsvConfig.QuoteChar)
        _p3 := this._p3 := Format(p3, GenerateCsvConfig.FieldDelimiter, GenerateCsvConfig.RecordDelimiter, GenerateCsvConfig.QuoteChar)
        csv := this.csv := GenerateCsv(generateCsvOpt)
        cols := csv.Columns
        pos := InStr(csv.Content, csv.options.RecordDelimiter)
        records := csv.records
        displayRecords := csv.displayRecords
        fd := csv.Options.FieldDelimiter
        rd := csv.Options.RecordDelimiter
        len := csv.MaxRecordLen
        f := FileOpen(csv.options.PathOut, 'r', generateCsvOpt.Encoding)
        startPos := f.Pos
        f.Read(1)
        ratio := f.Pos - startPos
        f.Pos := startPos
        str := f.Read(len)
        pos := InStr(str, rd)
        f.Pos := startPos + pos * ratio
        i := 0
        start := A_TickCount
        for record in records {
            ++i
            chars := Min(len, (f.Length - f.Pos) / ratio)
            str := f.Read(chars)
            pos := 1
            if RegExMatch(str, _p1, &m, pos) {
                if record[1] != m['value'] {
                    OutputDebug('record`n' record[1] '`nmatch`n' m['value'] '`n')
                    _Error('Matched string did not match the record.`r`nmatch:`r`n' StrReplace(StrReplace(m['value'], '`r', '``r'), '`n', '``n') '`r`nrecord:`r`n' displayRecords[i][1], _p1, displayRecords[i],  i > 1 ? displayRecords[i - 1] : '')
                    f.Close()
                    break
                }
                if m.Pos != pos {
                    _Error('The position was invalid.`r`nmatch:`r`n' StrReplace(StrReplace(m['value'], '`r', '``r'), '`n', '``n') '`r`nrecord:`r`n' displayRecords[i][1] '`r`npos: ' pos '; m.Pos: ' m.Pos, _p1, displayRecords[i], i > 1 ? displayRecords[i - 1] : '')
                    f.Close()
                    break
                }
                pos += m.Len
                _m := m
            } else {
                _Error('Pattern1 did not match.', _p1, displayRecords[i],  i > 1 ? displayRecords[i - 1] : '')
                f.Close()
                break
            }
            loop cols - 2 {
                if RegExMatch(str, _p2, &m, pos) {
                    if record[A_Index + 1] != m['value'] {
                        OutputDebug('record`n' record[A_Index + 1] '`nmatch`n' m['value'] '`n')
                        _Error('Matched string did not match the record.`r`nmatch:`r`n' StrReplace(StrReplace(m['value'], '`r', '``r'), '`n', '``n') '`r`nrecord:`r`n' displayRecords[i][A_Index + 1], _p2, displayRecords[i],  i > 1 ? displayRecords[i - 1] : '')
                        f.Close()
                        break 2
                    }
                    if m.Pos != pos {
                        _Error('The position was invalid.`r`nmatch:`r`n' StrReplace(StrReplace(m['value'], '`r', '``r'), '`n', '``n') '`r`nrecord:`r`n' displayRecords[i][A_Index + 1] '`r`npos: ' pos '; m.Pos: ' m.Pos, _p2, displayRecords[i],  i > 1 ? displayRecords[i - 1] : '')
                        f.Close()
                        break 2
                    }
                    pos := m.Pos + m.Len
                    _m := m
                } else {
                    _Error('Pattern2 did not match. Inner loop index + 1: ' (A_Index + 1), _p2, displayRecords[i],  i > 1 ? displayRecords[i - 1] : '')
                    f.Close()
                    break 2
                }
            }
            if RegExMatch(str, _p3, &m, pos) {
                if record[-1] != m['value'] {
                    OutputDebug('record`n' record[-1] '`nmatch`n' m['value'] '`n')
                    _Error('Matched string did not match the record.`r`nmatch:`r`n' StrReplace(StrReplace(m['value'], '`r', '``r'), '`n', '``n') '`r`nrecord:`r`n' displayRecords[i][-1], _p3, displayRecords[i],  i > 1 ? displayRecords[i - 1] : '')
                    f.Close()
                    break
                }
                if m.Pos != pos {
                    _Error('The position was invalid.`r`nmatch:`r`n' StrReplace(StrReplace(m['value'], '`r', '``r'), '`n', '``n') '`r`nrecord:`r`n' displayRecords[i][-1] '`r`npos: ' pos '; m.Pos: ' m.Pos, _p3, displayRecords[i],  i > 1 ? displayRecords[i - 1] : '')
                    f.Close()
                    break
                }
                pos := m.Pos + m.Len
                _m := m
            } else {
                _Error('Pattern3 did not match.', _p3, displayRecords[i],  i > 1 ? displayRecords[i - 1] : '')
                f.Close()
                break
            }
            f.Pos -= (chars - pos + 1) * ratio
        }
        dur := A_TickCount - start
        if showTooltip {
            MouseO := CoordMode('Mouse', 'Screen')
            TTO := CoordMode('Tooltip', 'Screen')
            MouseGetPos(&mx, &my)
            Tooltip('Done. Parsing ' records.Length ' records (' cols ' columns each) took ' Round(dur / 1000 / 60, 2) ' minutes.', mx, my)
            SetTimer(Tooltip, -3000)
            CoordMode('Mouse', MouseO)
            CoordMode('Tooltip', TTO)
        }

        return

        _Error(message, pattern, record, previousRecord) {
            g := Gui('+Resize')
            g.SetFont('s11 q5', 'Segoe Ui')
            g.Add('Text', 'w780 Section vTxtError', message '`r`nRecord ' i)
            s := ''
            w := 0
            context := SelectFontIntoDc(g['TxtError'].Hwnd)
            for h in csv.Headers {
                w := Max(w, GetTextExtentPoint32(context.hdc, h ':'))
            }
            context()
            edits := []
            if previousRecord {
                for field in previousRecord {
                    s .= field (A_Index < previousRecord.Length ? fd : '')
                }
                s .= '`r`n`r`n'
            }
            for h in csv.Headers {
                g.Add('Text', 'xs Right w' w ' Section', h ':')
                edits.Push(g.Add('Edit', 'ys -Wrap', record[A_Index]))
                s .= record[A_Index] (A_Index < record.Length ? fd : '')
            }
            context := SelectFontIntoDc(g['TxtError'].Hwnd)
            w2 := Min(w + Max(GetTextExtentPoint32(context.hdc, pattern), GetTextExtentPoint32(context.hdc, s)) + 10, 800)
            w3 := w2 - w - g.MarginX
            context()
            g.Add('Edit', 'xs w' w2 ' Section vPattern', StrReplace(StrReplace(pattern, '`r', '``r'), '`n', '``n'))
            g.Add('Edit', 'xs w' w2 ' vContext', s)
            for edt in edits {
                edt.Move(, , w3)
            }
            g.Add('Button', 'xs', 'Exit').OnEvent('Click', (*) => ExitApp())
            g.Show()
        }
    }
}

class GenerateCsvConfig {
    static Columns := 10
    static Encoding := ''
    static FieldDelimiter := ','
    static Headers := ''
    static LineEnding := '`n'
    static MaxOtherCharsPerGroup := 3 ; The maximum number of other chars in a group
    static MaxWordsPerField := 10
    static MinOtherCharsPerGroup := 1 ; The minimum number of other chars in a group
    static MinWordsPerField := 0
    static NoHeaders := false
    static OtherChars := '1234567890+_)(*&^%$#@!~``;/\|><.?'
    static OtherCharsOnlyInQuotedStrings := false
    static OtherCharsProbability := 0.3 ; The probability a group of other chars is used instead of a word
    static Overwrite := false
    static PathOut := ''
    static ProbabilityQuotedString := 0.6
    static QuoteChar := '"'
    static RandomEscapedQuotes := true
    static RandomFieldDelimiters := true
    static RandomLineEnding := true
    static RandomRecordDelimiters := true
    static RecordDelimiter := '`n'
    static Rows := 100
}


GetTextExtentPoint32(hdc, Str) {
    ; Measure the text
    if DllCall('C:\Windows\System32\Gdi32.dll\GetTextExtentPoint32'
        , 'Ptr', hdc
        , 'Ptr', StrPtr(Str)
        , 'Int', StrLen(Str)
        , 'Ptr', sz := Buffer(8)
        , 'Int'
    ) {
        return NumGet(sz, 0, 'int')
    } else {
        throw OSError()
    }
}

/**
 * @classdesc - Use this as a safe way to access a window's font object. This handles accessing and
 * releasing the device context and font object.
 */
class SelectFontIntoDc {

    __New(hWnd) {
        this.hWnd := hWnd
        if !(this.hdc := DllCall('GetDC', 'Ptr', hWnd, 'ptr')) {
            throw OSError()
        }
        OnError(this.Callback := ObjBindMethod(this, '__ReleaseOnError'), 1)
        if !(this.hFont := SendMessage(0x0031, 0, 0, , hWnd)) { ; WM_GETFONT
            throw OSError()
        }
        if !(this.oldFont := DllCall('SelectObject', 'ptr', this.hdc, 'ptr', this.hFont, 'ptr')) {
            throw OSError()
        }
    }

    /**
     * @description - Selects the old font back into the device context, then releases the
     * device context.
     */
    Call() {
        if err := this.__Release() {
            throw err
        }
    }

    __ReleaseOnError(thrown, mode) {
        if err := this.__Release() {
            thrown.Message .= '; ' err.Message
        }
        throw thrown
    }

    __Release() {
        if this.oldFont {
            if !DllCall('SelectObject', 'ptr', this.hdc, 'ptr', this.oldFont, 'int') {
                err := OSError()
            }
            this.DeleteProp('oldFont')
        }
        if this.hdc {
            if !DllCall('ReleaseDC', 'ptr', this.hWnd, 'ptr', this.hdc, 'int') {
                if IsSet(err) {
                    err.Message .= '; Another error occurred: ' OSError().Message
                }
            }
            this.DeleteProp('hdc')
        }
        OnError(this.Callback, 0)
        return err ?? ''
    }

    __Delete() => this()

    static __New() {
        if this.Prototype.__Class == 'SelectFontIntoDc' {
            Proto := this.Prototype
            Proto.DefineProp('hdc', { Value: '' })
            Proto.DefineProp('hFont', { Value: '' })
            Proto.DefineProp('oldFont', { Value: '' })
        }
    }
}
