
class ParseCsv {
    class Params {
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

        __New(params?) {
            for Name, Val in ParseCsv.Params.OwnProps() {
                if IsSet(params) && params.HasOwnProp(Name)
                   this.DefineProp(Name, {Value: params.%Name%})
                else if IsSet(ParseCsvConfig) && ParseCsvConfig.HasOwnProp(Name)
                   this.DefineProp(Name, {Value: ParseCsvConfig.%Name%})
                else
                   this.DefineProp(Name, {Value: Val})
            }
        }
    }

    Fields := [], Index := 0, CharPos := 0, LastPos := 0, Length := 0, Paused := 0, InputString := 0
    __New(params?, InputString?) {
        params := ParseCsv.Params(params??unset)
        this.params := params, this.Collection := ParseCsv.Collection(params.Constructor, params.CollectionArrayBuffer)
        if IsSet(InputString)
            this.Content := InputString, this.InputString := true
        if params.Breakpoint
            this.ReturnOnBreakpoint := SubStr(params.BreakpointAction, 1, 1) = 'R'
        if (params.RecordDelimiter && RegExMatch(params.RecordDelimiter, '[^\r\n]')) || params.QuoteChar || this.InputString {
            ; When there are quoted fields, the most straightforward parsing method is to loop with RegExMatch.
            if params.QuoteChar {
                this.ReadStyle := 'Quote'
            } else {
                ; When there are no quoted fields but the `RecordDelimiter` is not a newline, we can use `StrSplit`, either on the whole file or looping the content.
                this.ReadStyle := 'Split'
            }
        } else {
            ; When there are no quoted fields and the `RecordDelimiter` is a newline, we can loop by line, which is both fast and uses minimal memory.
            this.ReadStyle := 'Line'
        }
        if params.Start
            this()
    }

    static GetPattern(Quote, FieldDelimiter, RecordDelimiter?) {
        ; if IsSet(RecordDelimiter)
            Pattern := Format('JS)(?<=^|{2}|{3})(?:{1}(?<value>(?:[^{1}]*(?:{1}{1}){0,1})*){1}'
            '|(?<value>[^\r\n{1}{2}{4}]*?))(?={2}|{3}(*MARK:item)|$(*MARK:end))'
            , Quote, FieldDelimiter, RecordDelimiter, RegExReplace(RecordDelimiter, '\\[rnR]|``r|``n', ''))
        ; else
            ; Pattern := Format('JS)(?<=^|{2})(?:{1}(?<value>(?:[^{1}]*(?:{1}{1}){0,1})*){1}'
            ; '|(?<value>[^{1}{2}{3}{4}]*?))(?={2}|$(*MARK:item))', Quote, FieldDelimiter)
        ; The Pattern above is designed to be used in situations where the RecordDelimiter is a newline,
        ; and the fields may be quoted. I ended up not using it, but it does work if a use arises for it.
        try
            RegExMatch(' ', Pattern)
        catch Error as err {
            if err.message == 'Compile error 25 at offset 6: lookbehind assertion is not fixed length'
                throw Error('The procedure received "Compile error 25 at offset 6: lookbehind assertion'
                ' is not fixed length". To fix this, change the ``RecordDelimiter`` and/or ``FieldDelimiter``'
                ' to a value that is a fixed length.', -1)
            else
                throw err
        }
        return Pattern
    }

    Call() {
        if !this.Paused
            this.PrepareContent(), this.SetHeaders()
        this.Paused := false, this.Parse()
        if !this.Paused {
            if this.HasOwnProp('File')
                this.File.Close()
            this.Collection.__Item.Length := this.Collection.Count
        }
        return this
    }

    __Enum(VarCount) => this.Collection.__Enum(VarCount)

    GetField(match) {
        this.Fields[++this.Index] := match['value']
        if match.mark {
            if this.Index != this.RecordLength
                ParseCsv.__ThrowInvalidItemsQty(this.Fields, -2)
            this.Collection.Add(this.Fields)
            if match.mark == 'end'
                return 1
            else if match.mark == 'item'
                this.Fields := [], this.Fields.Length := this.RecordLength, this.Index := 0
            else
                ParseCsv.__ThrowUnexpectedMark(match.Mark, -2)
        }
    }

    GetHeaders(AhkStyle := false) {
        StrReplace(this.Content, '`r`n',,, &CRLFCount), StrReplace(this.Content, '`n',,, &LFCount)
        StrReplace(this.Content, '`r',,, &CRCount)
        if CRLFCount {
            if LFCount != CRCount
                _Throw()
            else
                return AhkStyle ? '`r`n' : '\r\n'
        } else {
            if LFCount && CRCount
                _Throw()
            else if LFCount
                return AhkStyle ? '`n' : '\n'
            else if CRCount
                return AhkStyle ? '`r' : '\r'
        }
        _Throw() {
            throw ValueError('The line endings are inconsistent, which will cause ``ParseCsv``'
            ' to fail to parse the input. Please correct the line endings before proceeding.', -2)
        }
    }

    GetProgress() {
        if this.ReadStyle == 'Line' {
            return this.File.Pos / this.File.Length
        } else if this.ReadStyle == 'Quote' {
            if this.HasOwnProp('File')
                return (this.LastPos + this.CharPos * this.bpc) / this.File.Length
            else
                return this.CharPos / this.Length
        } else if this.ReadStyle == 'Split' {
            len := 0
            Loop this.Content.Length
                len += StrLen(this.Content[A_Index])
            return (this.LastPos + (this.Length - len) * this.bpc) / this.File.Length
        } else
            throw ValueError('Unexpected read style: ' this.ReadStyle, -1)
    }

    LoopReadLine() {
        local params := this.params, BPA := params.BreakpointAction, f := this.File, Collection := this.Collection
        if params.Breakpoint {
            loop params.Breakpoint {
                Collection.Add(StrSplit(f.ReadLine(), params.FieldDelimiter))
                if f.AtEOF
                    break
            }
            this.Paused := true
            if not BPA is Func || BPA(this)
                return
        } else {
            while !f.AtEOF
                Collection.Add(StrSplit(f.ReadLine(), params.FieldDelimiter))
        }
    }

    LoopReadQuote() {
        local params := this.params, Pattern := this.Pattern, BPA := params.BreakpointAction, Collection := this.Collection
        if params.Breakpoint {
            next := Collection.Count + params.Breakpoint
            while Collection.Count < next {
                if RegExMatch(this.content, Pattern, &match, this.CharPos + 1) {
                    if match.Pos != this.CharPos + 1
                        _Throw()
                    this.CharPos := match.Pos + match.len
                    if this.GetField(match)
                        return
                } else {
                    if !this.HasOwnProp('File') || this.File.AtEOF
                        return
                    this.ReadNextQuote()
                }
            }
            this.Paused := true
            if not BPA is Func || BPA(this)
                return
        } else {
            Loop {
                while RegExMatch(this.content, Pattern, &match, this.CharPos + 1) {
                    if match.Pos != this.CharPos + 1
                        _Throw()
                    this.CharPos := match.Pos + match.len
                    if this.GetField(match)
                        return
                }
                if !this.HasOwnProp('File') || this.File.AtEOF
                    return
                this.ReadNextQuote()
            }
        }
        _Throw() {
            ParseCsv.__ThrowIncorrectMatchPos(match, this.CharPos, SubStr(this.Content, this.CharPos + 1, match.Pos - this.CharPos), -3)
        }
    }

    LoopReadSplit() {
        local params := this.params, BPA := params.BreakpointAction, Collection := this.Collection
        if params.Breakpoint {
            ; I wrote this block to use a method that, I believe, requires the fewest top-side calculations to accomplish the task.
            ; Not sure what the interpreter does so I cannot say if it truly requires the fewest calculations.
            start := this.Collection.Count + 1
            if (i := params.Breakpoint - this.Content.Length) > 0 {
                Loop {
                    if _LoopContentLength(&len)
                        return
                    if (i -= len) <= 0
                        break
                }
                Loop params.Breakpoint - this.Collection.Count + start
                    this.Collection.Add(StrSplit(this.Content.RemoveAt(1), params.FieldDelimiter))
            } else {
                Loop params.Breakpoint
                    this.Collection.Add(StrSplit(this.Content.RemoveAt(1), params.FieldDelimiter))
            }
            this.Paused := true
            if not BPA is Func || BPA(this)
                return
        } else
            while !_LoopContentLength()
                continue
        
        _LoopContentLength(&len?) {
            Loop this.Content.Length
                this.Collection.Add(StrSplit(this.Content[A_Index], params.FieldDelimiter))
            if !this.HasOwnProp('File') || this.File.AtEOF
                return 1
            this.ReadNextSplit()
            len := this.Content.Length
        }
    }

    Parse() {
        local params := this.params
        if this.ReadStyle == 'Quote'
            return this.LoopReadQuote()
        else if this.ReadStyle == 'Split'
            return this.LoopReadSplit()
        else if this.ReadStyle == 'Line'
            return this.LoopReadLine()
        else
            throw ValueError('Unexpected read style: ' this.ReadStyle, -1)
    }

    PrepareContent() {
        local params := this.params
        if this.InputString {
            if this.ReadStyle == 'Split'
                this.Content := StrSplit(this.Content, params.RecordDelimiter||this.GetHeaders(true))
            else
                this.Length := StrLen(this.Content)
            return
        }
        if this.ReadStyle == 'Line' || params.MaxReadSizeBytes {
            this.File := FileOpen(params.PathIn, 'r', params.Encoding||unset)
            this.File.Read(1), this.bpc := this.File.Pos, this.File.Pos := 0
            if this.ReadStyle == 'Line'
                return
            this.Content := this.File.Read(params.MaxReadSizeBytes)
        } else
            this.Content := FileRead(params.PathIn, params.Encoding||unset)
        if this.ReadStyle == 'Split'
            this.Content := StrSplit(this.Content, params.RecordDelimiter)
        else if !this.MaxReadSizeBytes
            this.Length := StrLen(this.Content)
    }

    ReadNext() {
        if this.ReadSyle == 'Quote'
            this.ReadNextQuote()
        else if this.ReadStyle == 'Split'
            this.ReadNextSplit()
        else if this.ReadStyle == 'Line'
            this.ReadNextLine()
        else
            throw ValueError('Unexpected read style: ' this.ReadStyle, -1)
    }

    /**
     * Constructs a record from a line of text in a file. Do not use this method to
     * loop a file, as it is less efficient compared to `LoopReadLine`.
     */
    ReadNextLine() {
        this.Collection.Add(StrSplit(this.File.ReadLine(), this.params.FieldDelimiter))
    }

    ReadNextQuote() {
        ; Chances are good that, when `RegExMatch` failed to return a match (which would then cause
        ; `ParseCsv` to call this method) there was some leftover text that was not utilized.
        ; This accounts for that condition.
        this.File.Pos -= (StrLen(this.Content) - this.CharPos) * this.bpc
        this.LastPos := this.File.Pos
        this.Content := this.File.Read(this.File.Pos + this.params.MaxReadSizeBytes > this.File.Length
        ? this.File.Length - this.File.Pos : this.params.MaxReadSizeBytes)
        this.CharPos := 0
    }

    ReadNextSplit() {
        this.LastPos := this.File.Pos
        this.Content := this.File.Read(this.File.Pos + this.params.MaxReadSizeBytes > this.File.Length
        ? this.File.Length - this.File.Pos : this.params.MaxReadSizeBytes)
        this.Length := StrLen(this.Content)
        this.Content := StrSplit(this.Content, this.params.RecordDelimiter||this.GetHeaders(true))
        ; The last item in the array is probably not a complete record, so it is removed.
        if !this.File.AtEOF
            this.File.Pos -= (len:=StrLen(this.Content.Pop())) * this.bpc
        this.Length -= len??0
    }

    SetHeaders() {
        local params := this.params
        if params.Headers {
            if params.Headers is Array
                this.Collection.SetHeaders(params.Headers)
            else if params.Headers is String
                this.Collection.SetHeaders(StrSplit(params.Headers, params.FieldDelimiter))
            else
                throw TypeError('The headers must be a string or an array.', -1)
        } else {
            if this.ReadStyle == 'Line'
                this.Collection.SetHeaders(StrSplit(this.File.ReadLine(), params.FieldDelimiter))
            else if this.ReadStyle == 'Quote'
                _SetHeadersLoopReadQuote()
            else if this.ReadStyle == 'Split'
                this.Collection.SetHeaders(StrSplit(this.Content.RemoveAt(1), params.FieldDelimiter))
            else
                throw ValueError('Unexpected read style: ' this.ReadStyle, -1)
        }
        
        this.RecordLength := this.Collection.Headers.Length
        if this.ReadStyle == 'Quote'
            this.Fields.Length := this.RecordLength

        _SetHeadersLoopReadQuote() {
            this.Pattern := ParseCsv.GetPattern(params.QuoteChar, params.FieldDelimiter, params.RecordDelimiter||this.GetHeaders())
            TempGetField := this.GetField
            this.DefineProp('GetField', {Call: _GetHeader})
            this.LoopReadQuote(), this.DefineProp('GetField', {Call: TempGetField})
            
            _GetHeader(self, match) {
                static Headers
                if !IsSet(Headers)
                    Headers := []
                Headers.Push(match['value'])
                if match.mark {
                    if match.mark == 'item' {
                        self.Collection.SetHeaders(Headers), Headers := unset
                        return 1
                    } else if match.mark == 'end'
                        throw Error('``ParseCsv`` encountered the end of the content while getting the headers.', -1)
                    else
                        ParseCsv.__ThrowUnexpectedMark(match.mark, -2)
                }
            }
        }
    }

    class Collection {
        Count := 0, BaseObj := {}
        __New(Constructor, BufferLength) {
            if Constructor
                this.DefineProp('MakeRecord', {Call: Constructor})
            this.BufferLength := this.__Item.Length := BufferLength
        }
        
        Add(Record) {
            this.__Item[++this.Count] := this.MakeRecord(Record)
            if this.Count == this.__Item.Length
                this.__Item.Length += this.BufferLength
        }
        
        InsertionSort(start, end?, CompareFn := (a, b) => a - b, arr?) {
            arr := arr??this.__Item
            if !IsSet(end) || end > this.Count
                end := this.Count
            i := start - 1
            while ++i <= end {
                current := arr[i]
                j := i - 1
                while (j >= start && CompareFn(arr[j], current) > 0) {
                    arr[j + 1] := arr[j]
                    j--
                }
                arr[j + 1] := current
            }
            return arr
        }
        
        QuickSort(CompareFn := (a, b) => a - b, ArrSizeThreshold := 3, PivotCandidates := 3) {
            arr := this.__Item
            if arr.length <= 1
                return arr
            if arr.Length <= ArrSizeThreshold
                return this.InsertionSort(1,, CompareFn)
            if PivotCandidates > 1 && arr.Length > PivotCandidates {
                Candidates := [], Candidates.Length := PivotCandidates
                Loop PivotCandidates
                    Candidates[A_Index] := arr[Random(1, arr.Length)]
                this.InsertionSort(1, , CompareFn, Candidates)
                pivot := Candidates[(PivotCandidates-Mod(PivotCandidates,2))/2]
            } else
                pivot := arr[arr.Length]
            left := [], right := [], left.Length := right.Length := arr.Length
            i := j := k := 0
            while ++i <= arr.Length {
                if (CompareFn(arr[i], pivot) < 0)
                    left[++k] := arr[i]
                else
                    right[++j] := arr[i]
            }
            left.Length := k, right.Length := j
            result := this.QuickSort(left, CompareFn), result.Push(this.QuickSort(right, CompareFn)*)
            return result
        }
        
        MakeRecord(RecordArray) {
            ObjSetBase(Rec := {}, this.BaseObj)
            Rec.__Item := Map()
            for Field in RecordArray
                Rec.Set(Rec.Headers[A_Index], Field)
            return Rec
        }

        SetHeaders(Headers) {
            ObjSetBase(this.BaseObj, ParseCsv.Record.Prototype)
            this.BaseObj.Headers := this.Headers := ParseCsv.Headers := Headers
        }
        
        __Item := Array()
        __Enum(VarCount) {
            i := 0
            if VarCount == 1
                return enum1
            else if VarCount == 2
                return enum2
            else
                throw TargetError('Unexpected number of parameters passed to the enumerator: ' VarCount, -1)

            enum1(&val) {
                if ++i > this.Count
                    return false
                val := this.__Item[i]
                return 1
            }
            enum2(&index, &val) {
                if ++i > this.Count
                    return false
                index := i, val := this.__Item[i]
                return 1
            }
        }
        Clone() => this.__Item.Clone()
        Delete(index) => this.__Item.Delete(index)
        Get(index?) {
            if IsSet(index) {
                if this.__Item.Has(index)
                    return this.__Item[index]
                else if this.__Item.HasOwnProp('Default')
                    return this.__Item.Default
                else if index > this.__Item.length
                    throw IndexError('The index ' index ' is out of range.', -1)
                else
                    throw UnsetItemError('The array does not have a value at index ' index, -1)
            } else
                return this.__Item
        }
        Has(index) => this.__Item.Has(index)
        InsertAt(index, val*) {
            this.Count += val.Length
            this.__Item.InsertAt(index, val*)
        }
        Pop() => this.__Item.RemoveAt(this.Count--)
        Push(val*) {
            this.__Item.InsertAt(this.Count, val*)
            this.Count += val.Length
        }
        RemoveAt(index, length?) {
            this.Count -= length??1
            return this.__Item.RemoveAt(index, length??unset)
        }
        Length => this.Count
        Capacity {
            Get => this.__Item.Capacity
            Set => this.__Item.Capacity := value
        }
        __Get(Name, *) {
            if this.__Item.HasOwnProp('Default')
                return this.__Item.Default
            else
                throw PropertyError('The property ' Name ' does not exist.', -1)
        }
    }

    class Record {
        __Enum(VarCount) => this.__Item.__Enum(VarCount)
        Delete(key) => this.__Item.Delete(key)
        Get(key?) {
            if IsSet(key) {
                if this.__Item.Has(key)
                    return this.__Item[key]
                else if this.__Item.HasOwnProp('Default')
                    return this.__Item.Default
                else
                    throw UnsetItemError('The key ' key ' does not exist.', -1)
            } else
                return this.__Item
        }
        Set(key, val) => this.__Item.Set(key, val)
        Has(key) => this.__Item.Has(key)
        Clone() => this.__Item.Clone()
        Clear() => this.__Item.Clear()
        Capacity {
            Get => this.__Item.Capacity
            Set => this.__Item.Capacity := value
        }
        CaseSense {
            Get => this.__Item.CaseSense
            Set => this.__Item.CaseSense := value
        }
        __Get(Name, *) {
            if this.__Item.Has(Name)
                return this.__Item[Name]
            if this.__Item.Has(StrReplace(Name, '_', ' '))
                return this.__Item[StrReplace(Name, '_', ' ')]
            if this.__Item.HasOwnProp('Default')
                return this.__Item.Default
            throw PropertyError('The property ' Name ' does not exist.', -1)
        }
    }

    static __ThrowIncorrectMatchPos(match, Pos, subContent, errorTarget) {
        throw ValueError(Format('Invalid CSV format. Note this can be caused by inconsistent line endings.'
        '`nThe current match should have began at position {1}, but the match occurred at position {2}.'
        '`nThe invalid matched content: {3}`nThe content before the match that should have been included'
        ' in the match:`n{4}', Pos + 1, match.Pos, match[0], subContent), errorTarget)
    }
    static __ThrowInvalidItemsQty(RecordArray, errorTarget) {
        Loop RecordArray.length {
            if RecordArray.Has(A_Index)
                errorstr .= RecordArray[A_Index] '`n'
        }
        throw ValueError('The number of values in the ``instance.Fields`` array does not match the'
        ' number of headers. This probably indicates an incorrectly formatted CSV.'
        ' Here are the fields for the current record:`n' Trim(errorstr, '`n'), errorTarget)
    }
    static __ThrowUnexpectedMark(mark, errorTarget) {
        throw ValueError('The RegExMatchInfo object contains an unexpect "MARK" value: ' mark, errorTarget)
    }
}

