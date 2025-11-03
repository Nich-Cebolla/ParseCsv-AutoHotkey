/*
    Github: https://github.com/Nich-Cebolla/ParseCsv-AutoHotkey/blob/main/ParseCsv.ahk
    Author: Nich-Cebolla
    License: MIT
*/

class ParseCsv extends Array {
    static __New() {
        this.DeleteProp('__New')
        proto := this.Prototype
        proto.InputString := false
        proto.File := proto.ReadStyle := proto.OnExitFunc := ''
        proto.FileStartByte := proto.Index := proto.ParsedChars := proto.RecordDelimiterLen := proto.ContentLen := 0
        proto.ReadChars := 250
        proto.BytesToCharRatio := 1
        proto.ErrorMessage := [
            'There is a syntax error in the csv content.'
          , 'The content contains only headers and no records.'
          , 'Unexpected read style.'
          , 'Unexpected value.'
          , 'There is a syntax error in the csv content.'
        ]
        proto.ErrorContext := [
            'The pattern failed to match.'
          , 'The match position was invalid.'
        ]
        proto.Options := ParseCsv.Options()
        proto.__LargeBreakpoint := 4294967295
        proto.__Pattern1 := (
            'S)'
            '%%'
            '(?<value>'
                '(?<empty_quotes>\Q{3}{3}\E)'
            '|'
                '\Q{3}\E(*COMMIT)'
                '(?<quoted_value>[^{6}]*+'
                    '(?:'
                        '(?:\Q{3}{3}\E)*+'
                        '[^{6}]*+'
                    ')++'
                ')'
                '\Q{3}\E'
            '|'
                '[^{4}{5}]*+'
            ')?'
            '\Q{1}\E'
        )
        proto.__PatternPart1 := '(?<=\Q{2}\E|^)'
        proto.__PatternPart2 := '^'
        proto.__Pattern2 := (
            'S)'
            '(?<=\Q{1}\E)'
            '(?<value>'
                '(?<empty_quotes>\Q{3}{3}\E)'
            '|'
                '\Q{3}\E(*COMMIT)'
                '(?<quoted_value>[^{6}]*+'
                    '(?:'
                        '(?:\Q{3}{3}\E)*+'
                        '[^{6}]*+'
                    ')++'
                ')'
                '\Q{3}\E'
            '|'
                '[^{4}{5}]*+'
            ')?'
            '\Q{1}\E'
        )
        proto.__Pattern3 := (
            'S)'
            '(?<=\Q{1}\E)'
            '(?<value>'
                '(?<empty_quotes>\Q{3}{3}\E)'
            '|'
                '\Q{3}\E(*COMMIT)'
                '(?<quoted_value>[^{6}]*+'
                    '(?:'
                        '(?:\Q{3}{3}\E)*+'
                        '[^{6}]*+'
                    ')++'
                ')'
                '\Q{3}\E'
            '|'
                '[^{4}{5}]*+'
            ')?'
            '%%'
        )
        proto.__PatternPart3 := '(?:\Q{2}\E|$(*MARK:end))'
        proto.__PatternPart4 := '$'
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
        return path
    }

    /**
     * @param {Object} [Options] - An object with options as property : value pairs. The following
     * are notes regarding the options your code will most likely need to set.
     * - `Options.RecordDelimiter` is required. Set this to the string that separates the records,
     *   usually "`n" or "`r`n".
     * - `Options.FieldDelimiter` has a default value of a comma ( , ); if your csv has fields that
     *   are separated by something other than a comma, set `Options.FieldDelimiter` with the string
     *   that separates the fields.
     * - Set `Options.PathIn` with the file path to the content, unless passing by string to
     *   `InputString`.
     * - `Options.Encoding` must be the correct encoding.
     *
     * @param {String} [InputString] - A string to parse. If not set, the file at `Options.PathIn` is
     * used. When the following conditions are true, {@link ParseCsv} creates a temporary file in
     * %TEMP% (A_Temp) and reads it back into memory incementally when parsing.
     * - `InputString` is set.
     * - `Options.QuoteChar` is set.
     * - `Options.FieldsContainRecordDelimiter` is nonzero.
     *
     * The file will be deleted when {@link ParseCsv} completed, and an `OnExit` callback is set
     * to delete the file in the event the script exits before {@link ParseCsv} completes. The
     * `OnExit` callback is disabled after the file is deleted.
     *
     * @param {Integer} [Options.Breakpoint] - Sets a threshold directing {@link ParseCsv} to pause
     * processing after processing `Options.Breakpoint` number of records. When using
     * `Options.Breakpoint` but not using `Options.BreakpointAction`, {@link ParseCsv} returns after
     * reaching the `Options.Breakpoint` threshold. To determine if {@link ParseCsv} returned because
     * it finished or because it is paused, check {@link ParseCsv#Paused} or {@link ParseCsv#Complete}.
     *
     * To start the parse process from where it left off, just call the {@link ParseCsv} object.
     *
     * @example
     * parseCsvOpt := { Breakpoint: 100, PathIn: "C:\users\shared\MyCsv.csv" }
     * pcsv := ParseCsv(parseCsvOpt)
     * while pcsv.Paused {
     *     OutputDebug(pcsv.GetProgress() "`n")
     *     pcsv()
     * }
     * @
     *
     * @param {*} [Options.BreakpointAction] - If set, a `Func` or callable object to call after
     * reaching the `Options.Breakpoint` threshold.
     *
     * Parameters:
     * 1. The {@link ParseCsv} object.
     *
     * Returns {Boolean}:
     * - Return a nonzero value to direct {@link ParseCsv} to stop processing. You can resume processing
     *   by calling the {@link ParseCsv} object, e.g. `ParseCsvObj()`.
     * - Return zero or an empty string to direct {@link ParseCsv} to resume processing.
     *
     * @example
     * MyCallback(pcsv) {
     *     OutputDebug(pcsv.GetProgress() "`n")
     *     if SomeCondition {
     *          return 1 ; Stop processing
     *     } else {
     *          return 0 ; Continue processing
     *     }
     * }
     * @
     *
     * @param {*} [Options.Constructor] - A `Func` or callable object that will be called every time
     * a complete record is parsed.
     *
     * Parameters:
     * 1. The array of field values
     *   - If `Options.QuoteChar` is set, the array contains RegExMatchInfo objects. The subcapture
     *     groups are:
     *     - empty_quotes - If the field contained only a pair of quote characters, this subcapture
     *       group is set with the quote characters.
     *     - quoted_value - If the field was a quoted field, the value between the quote characters.
     *     - value - The entire field value including quote characters if present.
     *   - If `Options.QuoteChar` is not set, the array contains the string values from each field
     *     of the record.
     * 2. The {@link ParseCsv} object. Your code can access the array of headers from
     *   {@link ParseCsv#Headers}.
     *
     * Returns: {*} The record object. If the function returns zero or an empty string, the record
     * is skipped.
     *
     * If unset, the default constructor {@link ParseCsv.RecordConstructor} is used, which returns
     * {@link ParseCsv.Record} objects.
     *
     * @param {string} [Options.Encoding = ""] - The encoding of the file.
     *
     * @param {string} [Options.FieldDelimiter = ","] - The string that separates fields. For example,
     * a comma.
     *
     * @param {Boolean} [Options.FieldsContainRecordDelimiter = false] - This option is only relevant
     * when `Options.QuoteChar` is set. Set `Options.FieldsContainRecordDelimiter` to a nonzero value
     * to direct {@link ParseCsv} to use an alternative parsing process that can handle fields that
     * might have a record delimiter character within them. For example, if `Options.RecordDelimiter = "`n"`,
     * and if the quoted fields in the csv might have a "`n" character, then this option must be
     * nonzero or else an error will occur. The inverse is not true; this can always be nonzero whether
     * or not there are any record delimiter characters in any fields.
     *
     * @param {string[]} [Options.Headers] - If set, the headers as an array of strings. If unset, the
     * first record in the CSV input is used.
     *
     * @param {Integer} [Options.InitialCapacity = 1000] - An integer to set the {@link ParseCsv}
     * object's capacity to before processing.
     *
     * @param {string} [Options.PathIn] - The path to the file to parse. If a relative path, it is
     * assumed to be relative to the working directory. Note that {@link ParseCsv} can usually handle
     * blank lines at the end of a csv file, except if `Options.QuoteChar` is set and if
     * `Options.FieldsContainRecordDelimiter` is nonzero. In this case your code would need to remove
     * any blank lines from the end of the file.
     *
     * @param {string} [Options.QuoteChar] - The character used to quote fields. This must be set when
     * some or all fields are quoted. If not set, the fields are assumed to be unquoted, which will
     * most likely result in an error or an incorrectly parsed CSV if the CSV actually does contain
     * quoted fields. If the fields are not quoted and this is set, {@link ParseCsv} will still parse
     * correctly but will be slower.
     *
     * If none of the fields are quoted, leaving this unset will result in significantly improved
     * performance.
     *
     * @param {string} Options.RecordDelimiter - The string that separates records. For example,
     * either "`n" or "`r`n". This option is required.
     *
     * @param {Boolean} [Options.Start = true] - If true, the parsing process is called immediately
     * upon instantiation. If false, your code must call {@link ParseCsv.Prototype.Call}.
     */
    __New(Options?, InputString?) {
        options := this.Options := ParseCsv.Options(Options ?? unset)
        this.Index := this.Paused := this.Complete := 0
        if options.QuoteChar {
            this.ReadStyle := 'Quote'
            if options.FieldsContainRecordDelimiter {
                if IsSet(InputString) {
                    this.ContentLen := StrLen(InputString)
                    options.PathIn := ParseCsv.GetPath()
                    f := FileOpen(options.PathIn, 'w', options.Encoding)
                    f.Write(InputString)
                    f.Close()
                    this.OnExitFunc := _DeleteFile.Bind(options.PathIn)
                    OnExit(this.OnExitFunc, 1)
                }
            } else if IsSet(InputString) {
                this.Content := StrSplit(InputString, options.RecordDelimiter)
                this.InputString := true
            }
        } else if IsSet(InputString) {
            this.ContentLen := StrLen(InputString)
            this.ReadStyle := 'Split'
            this.Content := StrSplit(InputString, options.RecordDelimiter)
            this.InputString := true
        } else if RegExMatch(options.RecordDelimiter, '[^\r\n]') {
            ; If `InputString` was set, or if the record delimiter contains characters other than line
            ; break characters, we can't use "Line" read style which is the fastest, but we can still
            ; loop using StrSplit which is still pretty fast.
            this.ReadStyle := 'Split'
        } else {
            ; When there are no quoted fields and the `RecordDelimiter` is a newline, we can loop by line, which is both fast and uses minimal memory.
            this.ReadStyle := 'Line'
        }
        this.Capacity := options.InitialCapacity
        if options.Start {
            this()
        }

        _DeleteFile(path, *) {
            if FileExist(path) {
                try {
                    FileDelete(path)
                }
            }
        }
    }

    /**
     * @description The general-purpose activation method.
     *
     * @returns {ParseCsv}
     */
    Call() {
        if this.Paused {
            this.Paused := false
        } else {
            if this.ReadStyle = 'Quote' {
                this.__GetPattern()
            }
            this.__PrepareContent()
            this.Constructor := this.Options.Constructor || ParseCsv.RecordConstructor(this.Headers, this.Options.QuoteChar || unset)
        }
        this.__Parse()
        return this
    }
    /**
     * @description - Searches the CSV for the first field which contains the input string
     * and returns the index number of the record which contains the field.
     *
     * @param {String} StrToFind - The string to find.
     *
     * @param {String|Integer|Array} [Headers] -
     * - If a string, the header name to search within.
     * - If an integer, the index value of the header to search within. If using an integer to refer
     *   to a header by index, its type must be `Number` or `Integer`, and not `String`. If it is a
     *   `String`, it will be considered a header name and not an index number.
     * - If an array, an array of integers or strings as described above. These will be searched
     *   in the order they are in the array.
     * - If unset, all headers will be searched.
     *
     * @param {Integer} [IndexStart = 1] - The record index number to start searching from.
     *
     * @param {Integer} [IndexEnd = this.Length] - The record index number to end searching at. If unset, the
     * all records after and including IndexStart are searched.
     *
     * @param {Boolean} [RequireFullMatch = true] - If true, the field must match the input string
     * exactly. If false, the field must contain the input string.
     *
     * @param {Boolean} [CaseSensitive = false] - If true, the search is case-sensitive. If false,
     * the search is case-insensitive.
     *
     * @param {Integer} [StartingPos] - This is only used when `RequireFullMatch` is false. The
     * position of the field's string (in number of characters) to search within. This is passed
     * to `InStr`.
     *
     * @param {VarRef} [OutField] - This variable will receive the string value of the field that
     * contains the input string.
     *
     * @param {VarRef} [OutHeader] - This variable will receive the string value of the header name
     * that contains `OutField`.
     *
     * @param {VarRef} [OutRecord] - This variable will receive the record object that contains
     * `OutField`.
     *
     * @returns {Integer} - The index number of the record which satisfies the conditions set by
     * the input parameters.
     */
    Find(StrToFind, Headers?, IndexStart := 1, IndexEnd := this.Length, RequireFullMatch := true, CaseSensitive := false, &OutField?, &OutHeader?, &OutRecord?) {
        Headers := this.__GetHeaders(Headers ?? unset)
        if RequireFullMatch {
            Process := CaseSensitive ? _Process_RFM_CS : _Process_RFM
        } else {
            Process := _Process
        }
        i := IndexStart - 1
        loop IndexEnd - i {
            ++i
            for header in Headers {
                if Process(&Header) {
                    OutHeader := header
                    OutRecord := this[i]
                    OutField := OutRecord[header]
                    return i
                }
            }
        }
        _Process(&Header) => InStr(this[i][Header], StrToFind, CaseSensitive)
        _Process_RFM(&Header) => this[i][Header] = StrToFind
        _Process_RFM_CS(&Header) => this[i][Header] == StrToFind
    }
    /**
     * @description - Loops the CSV between `IndexStart` and `IndexEnd`, adding the results from
     * {@link ParseCsv.Prototype.Find} to an array.
     *
     * The csv is iterated by searching all the records between IndexStart and IndexEnd for one header
     * before moving on to the next header. The arrays themselves are added to a `Map` object, where
     * the key is the header as it is passed to the `Headers` parameter (meaning if indices are used,
     * the keys are integers, else the keys are strings), and the value is the array.
     * If the {@link ParseCsv.Prototype.Find} method does not return a value for a given header,
     * the key is excluded from the resulting `Map` object.
     *
     * @param {String} StrToFind - The string to find.
     *
     * @param {String|Integer|Array} [Headers] -
     * - If a string, the header name to search within.
     * - If an integer, the index value of the header to search within. If using an integer to refer
     *   to a header by index, its type must be `Number` or `Integer`, and not `String`. If it is a
     *   `String`, it will be considered a header name and not an index number.
     * - If an array, an array of integers or strings as described above. These will be searched
     *   in the order they are in the array.
     * - If unset, all headers will be searched.
     *
     * @param {Integer} [IndexStart = 1] - The record index number to start searching from.
     *
     * @param {Integer} [IndexEnd = this.Length] - The record index number to end searching at. If unset, the
     * all records after and including IndexStart are searched.
     *
     * @param {Boolean} [RequireFullMatch = true] - If true, the field must match the input string
     * exactly. If false, the field must contain the input string.
     *
     * @param {Boolean} [CaseSensitive = false] - If true, the search is case-sensitive. If false,
     * the search is case-insensitive.
     *
     * @param {Boolean} [IncludeFields = true] - If true, the output includes the found field values.
     *
     * @param {Boolean} [IncludeRecords = true] - If true, the output includes the record object associated
     * with the found field values.
     *
     * @returns {Map} - If {@link ParseCsv.Prototype.Find} returns a value at least one time, this
     * returns a map object with the following characteristics:
     * - The map keys are any values passed to `Headers` which resulted in a match for at least
     *   one field. If any header did not have a field that matched the input string, that header
     *   is not represented in the result object.
     * - The items are objects with two to four properties:
     *   - {Integer} Index - The found index.
     *   - {String} Header- The header name associated with the found field value.
     *   - {String} Field - If `IncludeFields` is true, the found field value.
     *   - {Object} Record - If `IncludeRecords` is true, the record object associated with the found
     *     field value.
     *
     * If no values are returned by {@link ParseCsv.Prototype.Find}, this returns an empty string.
     */
    FindAll(StrToFind, Headers?, IndexStart := 1, IndexEnd := this.Length, RequireFullMatch := true, CaseSensitive := false, IncludeFields := true, IncludeRecords := true) {
        local i, field, result
        if IncludeFields {
            Add := IncludeRecords ? _Add4 : _Add2
        } else {
            Add := IncludeRecords ? _Add3 : _Add1
        }
        return this.__FindAll(_Process, Headers ?? unset)

        _Add1(&Header) => result.Push({ Header: Header, Index: i })
        _Add2(&Header) => result.Push({ Header: Header, Index: i, Field: Field })
        _Add3(&Header) => result.Push({ Header: Header, Index: i, Record: this[i] })
        _Add4(&Header) => result.Push({ Header: Header, Index: i, Field: Field, Record: this[i] })
        _Process(&Header) {
            Result := []
            i := IndexStart - 1
            while ++i <= IndexEnd {
                if i := this.Find(StrToFind, Header, i, IndexEnd, RequireFullMatch, CaseSensitive, &Field) {
                    Add(&Header)
                } else {
                    break
                }
            }
            return Result.Length ? Result : ''
        }
    }
    /**
     * @description - Loops the CSV between `IndexStart` and `IndexEnd`, adding the results from
     * {@link ParseCsv.Prototype.FindF} to an array.
     *
     * The csv is iterated by searching all the records between IndexStart and IndexEnd for one header
     * before moving on to the next header. The arrays themselves are added to a `Map` object, where
     * the key is the header as it is passed to the `Headers` parameter (meaning if indices are used,
     * the keys are integers, else the keys are strings), and the value is the array.
     * If the {@link ParseCsv.Prototype.FindF} method does not return a value for a given header,
     * the key is excluded from the resulting `Map` object.
     *
     * @param {*} Callback - A `Func` or callable object. When the function returns any nonzero value,
     * the result from {@link ParseCsv.Prototype.FindF} is added to an array.
     *
     * The function can accept up to three parameters:
     * 1. {Integer} The current record index number.
     * 2. {String} The current header name.
     * 3. {ParseCsv} The ParseCsv object.
     *
     * Returns: {Boolean} - Return nonzero to end the search.
     *
     * @param {String|Integer|Array} [Headers] -
     * - If a string, the header name to search within.
     * - If an integer, the index value of the header to search within. If using an integer to refer
     *   to a header by index, its type must be `Number` or `Integer`, and not `String`. If it is a
     *   `String`, it will be considered a header name and not an index number.
     * - If an array, an array of integers or strings as described above. These will be searched
     *   in the order they are in the array.
     * - If unset, all headers will be searched.
     *
     * @param {Integer} [IndexStart = 1] - The record index number to start searching from.
     *
     * @param {Integer} [IndexEnd = this.Length] - The record index number to end searching at. If unset, the
     * all records after and including IndexStart are searched.
     *
     * @param {Boolean} [IncludeFields = true] - If true, the output includes the found field values.
     *
     * @param {Boolean} [IncludeRecords = true] - If true, the output includes the record objects associated
     * with the found field values.
     *
     * @returns {Map} - If {@link ParseCsv.Prototype.FindF} returns a value at least one time, this
     * returns a map object with the following characteristics:
     * - The map keys are any values passed to `Headers` which resulted in a match for at least
     *   one field. If any header did not have a field cause the callback to return true, that
     *   header is not represented in the result object.
     * - The items are objects with two to four properties:
     *   - {Integer} Index - The found index.
     *   - {String} Header- The header name associated with the found field value.
     *   - {String} Field - If `IncludeFields` is true, the found field value.
     *   - {Object} Record - If `IncludeRecords` is true, the record object associated with the found
     *     field value.
     *
     * If no values are returned by {@link ParseCsv.Prototype.FindF}, this returns an empty string.
     */
    FindAllF(Callback, Headers?, IndexStart := 1, IndexEnd := this.Length, IncludeFields := true, IncludeRecords := true) {
        local Result, Field, i
        if IncludeFields {
            Add := IncludeRecords ? _Add4 : _Add2
        } else {
            Add := IncludeRecords ? _Add3 : _Add1
        }
        return this.__FindAll(_Process, Headers ?? unset)

        _Add1(&Header) => result.Push({ Header: Header, Index: i })
        _Add2(&Header) => result.Push({ Header: Header, Index: i, Field: Field })
        _Add3(&Header) => result.Push({ Header: Header, Index: i, Record: this[i] })
        _Add4(&Header) => result.Push({ Header: Header, Index: i, Field: Field, Record: this[i] })
        _Process(&Header) {
            Result := []
            i := IndexStart - 1
            while ++i <= IndexEnd {
                if i := this.FindF(Callback, Header, i, IndexEnd, &Field) {
                    Add(&Header)
                } else {
                    break
                }
            }
            return Result.Length ? Result : ''
        }
    }
    /**
     * @description - Loops the CSV between `IndexStart` and `IndexEnd`, adding the results from
     * {@link ParseCsv.Prototype.FindR} to an array.
     *
     * The csv is iterated by searching all the records between IndexStart and IndexEnd for one header
     * before moving on to the next header. The arrays themselves are added to a `Map` object, where
     * the key is the header as it is passed to the `Headers` parameter (meaning if indices are used,
     * the keys are integers, else the keys are strings), and the value is the array.
     * If the {@link ParseCsv.Prototype.FindR} method does not return a value for a given header,
     * the key is excluded from the resulting `Map` object.
     *
     * @param {String} Pattern - The Regular Expression pattern to match with.
     *
     * @param {String|Integer|Array} [Headers] -
     * - If a string, the header name to search within.
     * - If an integer, the index value of the header to search within. If using an integer to refer
     *   to a header by index, its type must be `Number` or `Integer`, and not `String`. If it is a
     *   `String`, it will be considered a header name and not an index number.
     * - If an array, an array of integers or strings as described above. These will be searched
     *   in the order they are in the array.
     * - If unset, all headers will be searched.
     *
     * @param {Integer} [IndexStart = 1] - The record index number to start searching from.
     *
     * @param {Integer} [IndexEnd = this.Length] - The record index number to end searching at. If unset, the
     * all records after and including IndexStart are searched.
     *
     * @param {Boolean} [IncludeMatch=true] - If true, the output includes the `RegExMatchInfo` objects.
     *
     * @param {Boolean} [IncludeFields = true] - If true, the output includes the found field values.
     *
     * @param {Boolean} [IncludeRecords = true] - If true, the output includes the record objects associated
     * with the found field values.
     *
     * @returns {Map} - If {@link ParseCsv.Prototype.FindR} returns a value at least one time, this
     * returns a map object with the following characteristics:
     * - The map keys are any values passed to `Headers` which resulted in a match for at least
     * one field. If any header did not have a field cause the callback to return true, that
     * header is not represented in the result object.
     * - The items are objects with two to five properties:
     *   - {Integer} Index - The found index.
     *   - {String} Header- The header name associated with the found field value.
     *   - {RegExMatchInfo} Match - If `IncludeMatch` is true, the match object.
     *   - {String} Field - If `IncludeFields` is true, the found field value.
     *   - {Object} Record - If `IncludeRecords` is true, the record object associated with the found
     *     field value.
     *
     * If no values are returned by {@link ParseCsv.Prototype.FindR}, this returns an empty string.
     */
    FindAllR(Pattern, Headers?, IndexStart := 1, IndexEnd := this.Length, IncludeMatch := true, IncludeFields := true, IncludeRecords := true) {
        local Result, Field, i, Add, Match, Record
        IncludeProps := []
        if IncludeMatch {
            IncludeProps.Push('Match')
        }
        if IncludeFields {
            IncludeProps.Push('Field')
        }
        if IncludeRecords {
            IncludeProps.Push('Record')
        }
        return this.__FindAll(_Process, Headers ?? unset)

        _Process(&Header) {
            Result := []
            i := IndexStart - 1
            while ++i <= IndexEnd {
                if i := this.FindR(Pattern, Header, i, IndexEnd, &Match, &Field, , &Record) {
                    ResultObj := { Header: Header, Index: i }
                    result.Push(ResultObj)
                    for Name in IncludeProps {
                        ResultObj.DefineProp(Name, { Value: %Name% })
                    }
                } else {
                    break
                }
            }
            return Result.Length ? Result : ''
        }
    }
    /**
     * @description - Iterates the fields in the CSV, passing the values to a callback function.
     * When the function returns true, this function returns the index number of the record.
     *
     * @param {*} Callback - A `Func` or callable object. When the function returns
     * any nonzero value, this function also returns.
     *
     * The function can accept up to three parameters:
     * 1. {Integer} The current record index number.
     * 2. {String} The current header name.
     * 3. {ParseCsv} The ParseCsv object.
     *
     * Returns: {Boolean} - Return nonzero to end the search.
     *
     * @param {String|Integer|Array} [Headers] -
     * - If a string, the header name to search within.
     * - If an integer, the index value of the header to search within. If using an integer to refer
     *   to a header by index, its type must be `Number` or `Integer`, and not `String`. If it is a
     *   `String`, it will be considered a header name and not an index number.
     * - If an array, an array of integers or strings as described above. These will be searched
     *   in the order they are in the array.
     * - If unset, all headers will be searched.
     *
     * @param {Integer} [IndexStart = 1] - The record index number to start searching from.
     *
     * @param {Integer} [IndexEnd = this.Length] - The record index number to end searching at. If unset, the
     * all records after and including IndexStart are searched.
     *
     * @param {VarRef} [OutField] - This variable will receive the string value of the field that
     * was passed to the callback function when the function returns true.
     *
     * @param {VarRef} [OutHeader] - This variable will receive the string value of the header name
     * that contains `OutField`.
     *
     * @param {VarRef} [OutRecord] - This variable will receive the record object that contains
     * `OutField`.
     *
     * @returns {Integer} - The index number of the record that contains the field that was passed
     * to the function when the function returned true.
     */
    FindF(Callback, Headers?, IndexStart := 1, IndexEnd := this.Length, &OutField?, &OutHeader?, &OutRecord?) {
        Headers := this.__GetHeaders(Headers ?? unset)
        i := IndexStart - 1
        loop IndexEnd - i {
            ++i
            for Header in Headers {
                if Callback(i, Header, this) {
                    OutRecord := this[i]
                    OutHeader := Header
                    OutField := OutRecord[Header]
                    return i
                }
            }
        }
    }
    /**
     * @description - Searches the CSV for a field that matches with the input pattern using
     * `RegExMatch`.
     *
     * @param {String} Pattern - The Regular Expression pattern to match with.
     *
     * @param {String|Integer|Array} [Headers] -
     * - If a string, the header name to search within.
     * - If an integer, the index value of the header to search within. If using an integer to refer
     *   to a header by index, its type must be `Number` or `Integer`, and not `String`. If it is a
     *   `String`, it will be considered a header name and not an index number.
     * - If an array, an array of integers or strings as described above. These will be searched
     *   in the order they are in the array.
     * - If unset, all headers will be searched.
     *
     * @param {Integer} [IndexStart = 1] - The record index number to start searching from.
     *
     * @param {Integer} [IndexEnd = this.Length] - The record index number to end searching at. If unset, the
     * all records after and including IndexStart are searched.
     *
     * @param {VarRef} [OutMatch] - This variable will receive the `RegExMatchInfo` object.
     *
     * @param {VarRef} [OutField] - This variable will receive the string value of the field that
     * matches, if any.
     *
     * @param {VarRef} [OutHeader] - This variable will receive the string value of the header name
     * that contains `OutField`.
     *
     * @param {VarRef} [OutRecord] - This variable will receive the record object that contains `OutField`.
     *
     * @returns {Integer} - The index number of the record that contains the field that matched
     * the pattern.
     */
    FindR(Pattern, Headers?, IndexStart := 1, IndexEnd := this.Length, &OutMatch?, &OutField?, &OutHeader?, &OutRecord?) {
        Headers := this.__GetHeaders(Headers ?? unset)
        i := IndexStart - 1
        loop IndexEnd - i {
            ++i
            for Header in Headers {
                if RegExMatch(this[i][Header], Pattern, &OutMatch) {
                    OutRecord := this[i]
                    OutHeader := Header
                    OutField := OutRecord[Header]
                    return i
                }
            }
        }
    }
    /**
     * @returns {Float} - Returns the progress of the parsing procedure as a float between 0 and 1.
     */
    GetProgress() {
        if this.Complete {
            return 1
        }
        switch this.ReadStyle, 0 {
            case 'Line': return this.File.Pos / this.File.Length
            case 'Quote':
                if this.parse_quote_incremental {
                    return this.File.Pos / this.File.Length
                } else if this.parse_quote_split {
                    return this.Index / this.Content.Length
                }
            case 'Split': return this.Index / this.Content.Length
            default: throw ValueError(this.ErrorMessage[3], , this.ReadStyle)
        }
    }
    Join(&OutStr, recordDelimiter?, fieldDelimiter?, includeHeaders := false) {
        if !IsSet(recordDelimiter) {
            recordDelimiter := this.Options.RecordDelimiter
        }
        if !IsSet(fieldDelimiter) {
            fieldDelimiter := this.Options.fieldDelimiter
        }
        headers := this.Headers
        loopLen := headers.Length - 1
        if includeHeaders {
            OutStr .= headers[1]
            i := 1
            loop loopLen {
                OutStr .= fieldDelimiter headers[++i]
            }
        }
        OutStr .= recordDelimiter
        for record in this {
            OutStr .= record[1]
            i := 1
            loop loopLen {
                OutStr .= fieldDelimiter record[++i]
            }
            OutStr .= recordDelimiter
        }
        OutStr := SubStr(OutStr, 1, -StrLen(recordDelimiter))
    }
    /**
     * Generates a string representing a function that returns the records as an array of objects with
     * a property for each header. The following data shows a comparison between the amount of time
     * it takes to load the data using the function created by this method, compared with the amount
     * of time it takes to load the data using {@link ParseCsv}. The data sample used was a 232965
     * record csv with 18 columns (61.7MB).
     *
     * It does not seem that the function generated by this method saves processing time.
     *
     * - Time to call the function generated by this method: 2.674 minutes
     * - Time to parse using {@link ParseCsv.Prototype.__ParseLine}: 0.018 minutes
     * - Time to parse using {@link ParseCsv.Prototype.__ParseQuoteIncremental}: 0.271 minutes
     * - Time to parse using {@link ParseCsv.Prototype.__ParseQuoteSplit}: 0.302 minutes
     * - Time to parse using {@link ParseCsv.Prototype.__ParseSplit}: 0.024 minutes
     */
    ToAhkCode(pathOut, funcName := 'GetRecords', arrayVarName := 'records', itemsPerArray := 500, quoteChar := "'") {
        headers := this.headers
        indent := '      , '
        endBytes := StrPut(', ', 'cp1200') - 2
        f := FileOpen(pathOut, 'w')
        f.Write(funcName '() {`n    ' arrayVarName '1 := [`n        { ')
        record := this[1]
        for h in headers {
            f.Write(h ': ' (IsNumber(record.%h%) && !InStr(record.%h%, '.') ? record.%h% : quoteChar StrReplace(StrReplace(StrReplace(record.%h%, '``', '````'), '`n', '``n'), '`r', '``r') quoteChar) ', ')
        }
        f.pos -= endBytes
        f.Write(' }`n')
        z := 0
        n := i := k := 1
        loop {
            if ++i > this.Length {
                break
            }
            if ++k > itemsPerArray {
                ++n
                k := 1
                f.Write('    ]`n    ' arrayVarName n ' := [`n        { ')
                record := this[i]
                for h in headers {
                    f.Write(h ': ' (IsNumber(record.%h%) && !InStr(record.%h%, '.') ? record.%h% : quoteChar StrReplace(StrReplace(StrReplace(record.%h%, '``', '````'), '`n', '``n'), '`r', '``r') quoteChar) ', ')
                }
                f.pos -= endBytes
                f.Write(' }`n')
                if ++i > this.Length {
                    break
                }
            }
            f.Write(indent '{ ')
            record := this[i]
            for h in headers {
                f.Write(h ': ' (IsNumber(record.%h%) && !InStr(record.%h%, '.') ? record.%h% : quoteChar StrReplace(StrReplace(StrReplace(record.%h%, '``', '````'), '`n', '``n'), '`r', '``r') quoteChar) ', ')
            }
            f.pos -= endBytes
            f.Write(' }`n')
        }
        f.Write('    ]`n    result := []`n    loop ' n ' {`n        result.Push(' arrayVarName '%A_Index%*)`n    }`n    return result`n}`n')
        f.Close()
    }
    __ConvertCharsToFilePos(CharCount) {
        return this.FileStartByte + CharCount * this.BytesToCharRatio
    }
    /**
     * @description - Handles end-of-procedure actions.
     */
    __End() {
        if this.HasOwnProp('File') {
            this.File.Close()
            if this.OnExitFunc {
                OnExit(this.OnExitFunc, 0)
                this.DeleteProp('OnExitFunc')
            }
            this.DeleteProp('File')
        }
        if this.HasOwnProp('Content') {
            this.DeleteProp('Content')
        }
        this.Capacity := this.Length
        this.Complete := 1
    }
    /**
     * @description - Used internally when any of `FindAll`, `FindAllF`, or `FindAllR` are called.
     */
    __FindAll(Callback, Headers?) {
        if IsSet(Headers) {
            if not Headers is Array {
                Headers := [Headers]
            }
        } else {
            Headers := this.Headers
        }
        Result := Map()
        for Header in Headers {
            if item := Callback(&Header) {
                Result.Set(Header, item)
            }
        }
        return Result.Count ? Result : ''
    }
    /**
     * @description - Used internally when any of {@link ParseCsv.Prototype.Find}, {@link ParseCsv.Prototype.FindF}, {@link ParseCsv.Prototype.FindR}, or `Join` are called
     */
    __GetHeaders(Headers?) {
        if IsSet(Headers) {
            if not Headers is Array {
                Headers := [Headers]
            }
            Result := []
            for Header in Headers {
                Result.Push(Header is Number ? this.Headers[Header] : Header)
            }
            return Result
        } else {
            return this.Headers
        }
    }
    __GetPattern() {
        options := this.Options
        if options.FieldsContainRecordDelimiter {
            this.Pattern1 := Format(
                StrReplace(this.__Pattern1, '%%', this.__PatternPart1)
              , options.FieldDelimiter
              , options.RecordDelimiter
              , options.QuoteChar
              , StrReplace(StrReplace(options.FieldDelimiter, '\', '\\'), ']', '\]')
              , StrReplace(StrReplace(options.RecordDelimiter, '\', '\\'), ']', '\]')
              , StrReplace(StrReplace(options.QuoteChar, '\', '\\'), ']', '\]')
            )
            this.Pattern3 := Format(
                StrReplace(this.__Pattern3, '%%', this.__PatternPart3)
              , options.FieldDelimiter
              , options.RecordDelimiter
              , options.QuoteChar
              , StrReplace(StrReplace(options.FieldDelimiter, '\', '\\'), ']', '\]')
              , StrReplace(StrReplace(options.RecordDelimiter, '\', '\\'), ']', '\]')
              , StrReplace(StrReplace(options.QuoteChar, '\', '\\'), ']', '\]')
            )
        } else {
            this.Pattern1 := Format(
                StrReplace(this.__Pattern1, '%%', this.__PatternPart2)
              , options.FieldDelimiter
              , options.RecordDelimiter
              , options.QuoteChar
              , StrReplace(StrReplace(options.FieldDelimiter, '\', '\\'), ']', '\]')
              , StrReplace(StrReplace(options.RecordDelimiter, '\', '\\'), ']', '\]')
              , StrReplace(StrReplace(options.QuoteChar, '\', '\\'), ']', '\]')
            )
            this.Pattern3 := Format(
                StrReplace(this.__Pattern3, '%%', this.__PatternPart4)
              , options.FieldDelimiter
              , options.RecordDelimiter
              , options.QuoteChar
              , StrReplace(StrReplace(options.FieldDelimiter, '\', '\\'), ']', '\]')
              , StrReplace(StrReplace(options.RecordDelimiter, '\', '\\'), ']', '\]')
              , StrReplace(StrReplace(options.QuoteChar, '\', '\\'), ']', '\]')
            )
        }
        this.Pattern2 := Format(
            this.__Pattern2
          , options.FieldDelimiter
          , options.RecordDelimiter
          , options.QuoteChar
          , StrReplace(StrReplace(options.FieldDelimiter, '\', '\\'), ']', '\]')
          , StrReplace(StrReplace(options.RecordDelimiter, '\', '\\'), ']', '\]')
          , StrReplace(StrReplace(options.QuoteChar, '\', '\\'), ']', '\]')
        )
    }
    /**
     * @description Redirects the function to the correct LoopRead method.
     */
    __Parse() {
        switch this.ReadStyle {
            case 'Line': return this.__ParseLine()
            case 'Quote':
                if this.parse_quote_incremental {
                    this.__ParseQuoteIncremental()
                } else if this.parse_quote_split {
                    this.__ParseQuoteSplit()
                }
            case 'Split': return this.__ParseSplit()
            default: throw ValueError(this.ErrorMessage[3], , this.ReadStyle)
        }
    }
    /**
     * @description This loops the input file one line at a time. It is used when the following are true:
     * - The `RecordDelimiter` is a newline.
     * - The fields are not quoted.
     * - `PathIn` is set.
     */
    __ParseLine() {
        options := this.Options
        constructor := this.Constructor
        f := this.File
        fd := options.FieldDelimiter
        bp := options.Breakpoint || this.__LargeBreakpoint
        if bpa := options.BreakpointAction {
            breakCondition := !IsObject(bpa)
        } else {
            breakCondition := true
        }
        loop {
            loop bp {
                if f.AtEoF {
                    this.__End()
                    return
                }
                fields := StrSplit(f.ReadLine(), fd)
                if fields.Length {
                    if record := constructor(fields, this) {
                        this.Push(record)
                    }
                }
            }
            if breakCondition || bpa(this) {
                this.Paused := true
                return
            }
        }
    }
    __ParseQuoteIncremental() {
        options := this.Options
        chars := this.ReadChars
        constructor := this.Constructor
        f := this.File
        pattern1 := this.Pattern1
        pattern2 := this.Pattern2
        pattern3 := this.Pattern3
        ratio := this.BytesToCharRatio
        parsedChars := this.ParsedChars
        fieldDelimiterLen := StrLen(options.FieldDelimiter)
        recordDelimiterLen := StrLen(options.RecordDelimiter)
        matchPos1 := recordDelimiterLen + 1
        matchPos2 := matchPos3 := fieldDelimiterLen + 1
        bp := options.Breakpoint || this.__LargeBreakpoint
        if bpa := options.BreakpointAction {
            breakCondition := !IsObject(bpa)
        } else {
            breakCondition := true
        }
        cols := this.Headers.Length
        innerLoopLen := cols - 2
        f.Pos := this.__ConvertCharsToFilePos(parsedChars - recordDelimiterLen)
        str := f.Read(chars)
        loop {
            loop bp {
                fields := []
                this.Push(fields)
                fields.Capacity := cols
                loop {
                    if RegExMatch(str, pattern1, &match) {
                        if match.Pos = matchPos1 {
                            fields.Push(match)
                            str := SubStr(str, matchPos1 + match.Len - fieldDelimiterLen)
                            parsedChars += match.Len
                            break
                        } else {
                            _Throw(2)
                        }
                    } else if f.AtEoF {
                        _Throw(1)
                    } else {
                        str .= f.Read(chars)
                        chars := Round(chars * 1.5, 0)
                    }
                }
                i := 0
                loop {
                    if RegExMatch(str, pattern2, &match) {
                        if match.Pos = matchPos2 {
                            fields.Push(match)
                            str := SubStr(str, matchPos2 + match.Len - fieldDelimiterLen)
                            parsedChars += match.Len
                            if ++i >= innerLoopLen {
                                break
                            }
                        } else {
                            _Throw(2)
                        }
                    } else if f.AtEoF {
                        _Throw(1)
                    } else {
                        str .= f.Read(chars)
                        chars := Round(chars * 1.5, 0)
                    }
                }
                loop {
                    if RegExMatch(str, pattern3, &match) {
                        if match.Mark = 'end' {
                            if f.AtEof {
                                fields.Push(match)
                                constructor(fields, this)
                                this.ReadChars := chars
                                this.ParsedChars := parsedChars + match.Len
                                this.__End()
                                return
                            } else {
                                str .= f.Read(chars)
                                chars := Round(chars * 1.5, 0)
                            }
                        } else if match.Pos = matchPos3 {
                            fields.Push(match)
                            parsedChars += match.Len
                            str := SubStr(str, matchPos3 + match.Len - recordDelimiterLen)
                            break
                        } else {
                            _Throw(2)
                        }
                    } else if f.AtEoF {
                        _Throw(1)
                    } else {
                        str .= f.Read(chars)
                        chars := Round(chars * 1.5, 0)
                    }
                }
                constructor(fields, this)
                str .= f.Read(chars - StrLen(str))
            }
            if breakCondition || bpa(this) {
                this.ReadChars := chars
                this.ParsedChars := parsedChars
                this.Paused := true
                return
            }
        }

        return

        _Throw(context) {
            err := Error(this.ErrorMessage[1], -1, this.ErrorContext[context] ' Near char: ' parsedChars '; Record: ' this.Length '; After field ' fields.Length)
            err.Near := parsedChars
            err.Record := this.Length
            err.Field := fields.Length
            err.Context := context
            throw err
        }
    }
    __ParseQuoteSplit() {
        options := this.Options
        constructor := this.Constructor
        f := this.File
        cols := this.Headers.Length
        content := this.Content
        index := this.Index
        pattern1 := this.Pattern1
        pattern2 := this.Pattern2
        pattern3 := this.Pattern3
        matchPos1 := 1
        matchPos2 := matchPos3 := StrLen(options.FieldDelimiter) + 1
        parsedChars := this.ParsedChars
        recordDelimiterLen := StrLen(options.RecordDelimiter)
        fieldDelimiterLen := StrLen(options.fieldDelimiter)
        loopLen := Min(options.Breakpoint || this.__LargeBreakpoint, content.Length - index)
        if bpa := options.BreakpointAction {
            breakCondition := !IsObject(bpa)
        } else {
            breakCondition := true
        }
        innerLoopLen := cols - 2
        loop {
            loop loopLen {
                if index >= content.Length {
                    break
                }
                fields := []
                this.Push(fields)
                fields.Capacity := cols
                if str := content.Delete(++index) {
                    if RegExMatch(str, pattern1, &match) {
                        if match.Pos = matchPos1 {
                            fields.Push(match)
                            str := SubStr(str, matchPos1 + match.Len - fieldDelimiterLen)
                            parsedChars += match.Len
                        } else {
                            _Throw(2)
                        }
                    } else {
                        _Throw(1)
                    }
                    loop innerLoopLen {
                        if RegExMatch(str, pattern2, &match) {
                            if match.Pos = matchPos2 {
                                fields.Push(match)
                                str := SubStr(str, matchPos2 + match.Len - fieldDelimiterLen)
                                parsedChars += match.Len
                            } else {
                                _Throw(2)
                            }
                        } else {
                            _Throw(1)
                        }
                    }
                    if RegExMatch(str, pattern3, &match) {
                        if match.Pos = matchPos3 {
                            fields.Push(match)
                            parsedChars += match.Len + recordDelimiterLen
                            str := SubStr(str, matchPos3 + match.Len - recordDelimiterLen)
                        } else {
                            _Throw(2)
                        }
                    } else {
                        _Throw(1)
                    }
                    constructor(fields, this)
                } else {
                    parsedChars += recordDelimiterLen
                }
            }
            if index >= content.Length {
                this.__End()
                this.Index := index
                this.ParsedChars := parsedChars - recordDelimiterLen
                return
            } else if breakCondition || bpa(this) {
                this.Index := index
                this.ParsedChars := parsedChars
                this.Paused := true
                return
            }
        }

        return

        _Throw(context) {
            err := Error(this.ErrorMessage[1], -1, this.ErrorContext[context] ' Near char: ' parsedChars '; Record: ' this.Length '; After field ' fields.Length)
            err.Near := parsedChars
            err.Record := this.Length
            err.Field := fields.Length
            err.Context := context
            throw err
        }
    }
    /**
     * @description This loops the input value using `StrSplit`. This is used when `LoopReadLine` is
     * not possible, and the fields are not quoted. This is faster than `LoopReadQuote`.
     */
    __ParseSplit() {
        options := this.Options
        constructor := this.Constructor
        content := this.Content
        index := this.Index
        fd := options.FieldDelimiter
        bp := options.Breakpoint || this.__LargeBreakpoint
        if bpa := options.BreakpointAction {
            breakCondition := !IsObject(bpa)
        }
        loop {
            loop bp {
                if index >= content.Length {
                    this.__End()
                    this.Index := index
                    return
                }
                fields := StrSplit(content.Delete(++index), fd)
                if fields.Length {
                    if record := constructor(fields, this) {
                        this.Push(record)
                    }
                }
            }
            if breakCondition || bpa(this) {
                this.Paused := true
                this.Index := index
                return
            }
        }
    }
    /**
     * @description An instantiation function that handles preparing the content for parsing.
     * Content is accessible from `instance.Content`. This must be called for {@link ParseCsv} to function,
     * but is generally handled automatically by `instance.Call`.
     */
    __PrepareContent() {
        options := this.Options
        if options.Headers {
            if options.Headers is Array {
                this.Headers := options.Headers
            } else {
                throw TypeError('``Options.Headers`` must be an array of strings.')
            }
            switch this.ReadStyle, 0 {
                case 'Line':
                    this.File := FileOpen(options.PathIn, 'r', options.encoding)
                    pos := this.File.Pos
                    this.File.Read(1)
                    ratio := this.File.Pos - pos
                    this.File.Pos := pos
                    this.ContentLen := (this.File.Length - pos) / ratio
                case 'Split':
                    if !this.InputString {
                        _OpenFileSplit()
                    }
                    this.Index := 0
                case 'Quote':
                    if this.parse_quote_incremental {
                        _OpenFile()
                    } else if this.parse_quote_split {
                        if !this.InputString {
                            _OpenFileSplit()
                        }
                        this.Index := 0
                    }
            }
        } else {
            switch this.ReadStyle, 0 {
                case 'Line':
                    this.File := FileOpen(options.PathIn, 'r', options.encoding)
                    pos := this.File.Pos
                    this.File.Read(1)
                    ratio := this.File.Pos - pos
                    this.File.Pos := pos
                    this.ContentLen := (this.File.Length - pos) / ratio
                    this.Headers := StrSplit(this.File.ReadLine(), options.FieldDelimiter)
                case 'Split':
                    if !this.InputString {
                        _OpenFileSplit()
                    }
                    this.Headers := StrSplit(this.Content.Delete(1), options.FieldDelimiter)
                    this.Index := 1
                case 'Quote':
                    headers := this.Headers := []
                    headers.Capacity := 20
                    pattern1 := this.Pattern1
                    pattern2 := this.Pattern2
                    pattern3 := this.Pattern3
                    fieldDelimiterLen := StrLen(options.FieldDelimiter)
                    matchPos1 := 1
                    matchPos2 := matchPos3 := fieldDelimiterLen + 1
                    parsedChars := 0
                    if this.parse_quote_incremental {
                        _OpenFile()
                        f := this.File
                        ratio := this.BytesToCharRatio
                        chars := this.ReadChars
                        str := f.Read(chars)
                        loop {
                            if RegExMatch(str, pattern1, &match) {
                                if match.Pos = matchPos1 {
                                    headers.Push(match['value'])
                                    str := SubStr(str, matchPos1 + match.Len - fieldDelimiterLen)
                                    parsedChars += match.Len
                                    break
                                } else {
                                    _Throw(2)
                                }
                            } else if f.AtEoF {
                                _Throw(1)
                            } else {
                                str .= f.Read(chars)
                                chars := Round(chars * 1.5, 0)
                            }
                        }
                        loop {
                            if RegExMatch(str, pattern2, &match) {
                                if match.Pos = matchPos2 {
                                    headers.Push(match['value'])
                                    str := SubStr(str, matchPos2 + match.Len - fieldDelimiterLen)
                                    parsedChars += match.Len
                                } else if RegExMatch(str, pattern3, &match) {
                                    if match.Mark = 'end' {
                                        if f.AtEof {
                                            throw Error(this.ErrorMessage[2])
                                        } else {
                                            str .= f.Read(chars)
                                            chars := Round(chars * 1.5, 0)
                                        }
                                    } else if match.Pos = matchPos3 {
                                        headers.Push(match['value'])
                                        parsedChars += match.Len
                                        break
                                    } else {
                                        _Throw(2)
                                    }
                                } else {
                                    _Throw(1)
                                }
                            } else if RegExMatch(str, pattern3, &match) {
                                if match.Mark = 'end' {
                                    if f.AtEof {
                                        throw Error(this.ErrorMessage[2])
                                    } else {
                                        str .= f.Read(chars)
                                        chars := Round(chars * 1.5, 0)
                                    }
                                } else if match.Pos = matchPos3 {
                                    headers.Push(match['value'])
                                    parsedChars += match.Len
                                    break
                                } else {
                                    _Throw(2)
                                }
                            } else if f.AtEoF {
                                _Throw(1)
                            } else {
                                str .= f.Read(chars)
                                chars := Round(chars * 1.5, 0)
                            }
                        }
                        this.ReadChars := chars
                        this.ParsedChars := parsedChars
                    } else if this.parse_quote_split {
                        if !this.InputString {
                            _OpenFileSplit()
                        }
                        str := this.Content.Delete(1)
                        if RegExMatch(str, pattern1, &match) {
                            if match.Pos = matchPos1 {
                                headers.Push(match['value'])
                                str := SubStr(str, matchPos1 + match.Len - fieldDelimiterLen)
                                parsedChars += match.Len
                            } else {
                                _Throw(2)
                            }
                        } else {
                            _Throw(1)
                        }
                        loop {
                            if RegExMatch(str, pattern2, &match) {
                                if match.Pos = matchPos2 {
                                    headers.Push(match['value'])
                                    str := SubStr(str, matchPos2 + match.Len - fieldDelimiterLen)
                                    parsedChars += match.Len
                                } else {
                                    _Throw(2)
                                }
                            } else if RegExMatch(str, pattern3, &match) {
                                if match.Pos = matchPos3 {
                                    headers.Push(match['value'])
                                    parsedChars += match.Len
                                    break
                                } else {
                                    _Throw(2)
                                }
                            } else {
                                _Throw(1)
                            }
                        }
                        this.Index := 1
                        this.ParsedChars := parsedChars + StrLen(options.RecordDelimiter)
                    }
            }
        }

        return

        _OpenFile() {
            this.File := FileOpen(options.PathIn, 'r', options.encoding)
            pos := this.FileStartByte := this.File.Pos
            this.File.Read(1)
            this.BytesToCharRatio := this.File.Pos - pos
            this.File.Pos := pos
            this.ContentLen := (this.File.Length - pos) / this.BytesToCharRatio
        }
        _OpenFileSplit() {
            content := FileRead(options.PathIn, options.encoding)
            this.ContentLen := StrLen(content)
            this.Content := StrSplit(content, options.RecordDelimiter)
        }
        _Throw(context) {
            err := Error(this.ErrorMessage[1], -1, this.ErrorContext[context] ' Near char: ' parsedChars '; Record: headers; After field ' headers.Length)
            err.Near := parsedChars
            err.Record := 'Headers'
            err.Field := headers.Length
            err.Context := context
            throw err
        }
    }

    parse_quote_incremental => this.ReadStyle = 'Quote' && this.Options.FieldsContainRecordDelimiter
    parse_quote_split => this.ReadStyle = 'Quote' && !this.Options.FieldsContainRecordDelimiter

    class RecordConstructor extends Class {
        static __New() {
            this.DeleteProp('__New')
            proto := this.Prototype
            proto.QuoteChar := ''
        }
        __New(Headers?, QuoteChar?) {
            this.Prototype := []
            ObjSetBase(this.Prototype, ParseCsv.Record.Prototype)
            if IsSet(QuoteChar) {
                this.QuoteChar := QuoteChar
                __get := _GetMatch.Bind() ; To avoid changing the name of the function object itself
                __set := _SetMatch.Bind()
                __get.DefineProp('Name', { Value: ParseCsv.Record.Prototype.__Class '.Prototype.__Item.Get' })
                __set.DefineProp('Name', { Value: ParseCsv.Record.Prototype.__Class '.Prototype.__Item.Set' })
                this.Prototype.DefineProp('__Item', { Get: __get, Set: __set })
            }
            if IsSet(Headers) {
                this.SetHeaders(Headers)
            }
            this.Prototype.__Class := ParseCsv.Record.Prototype.__Class

            return

            _GetMatch(Self, Header) {
                return Self.Get(Header is Number ? Header : Self.Headers.Get(Header))['value']
            }
            _SetMatch(Self, Header, Value) {
                local __get, desc, h, i, index
                index := Header is Number ? Header : Self.Headers.Get(Header)
                for h, i in Self.Headers {
                    if i = index {
                        name := RegExReplace(h, '\W', '')
                        break
                    }
                }
                desc := Self.Base.GetOwnPropDesc(name)
                __get := _Get.Bind(Index)
                __set := _Set.Bind(Index)
                __get.DefineProp('Name', { Value: desc.Get.Name })
                __set.DefineProp('Name', { Value: desc.Set.Name })
                Self.DefineProp(name, { Get: __get, Set: __set })

                return Self.Set(Value, Self.Headers.Get(Header))

                _Get(Index, Self) {
                    return Self.Get(Index)
                }
                _Set(Index, Self, Value) {
                    return Self.Set(Value, Index)
                }
            }
        }
        Call(arr, *) {
            ObjSetBase(arr, this.Prototype)
            return arr
        }
        SetHeaders(Headers) {
            proto := this.Prototype
            proto.Headers := Map()
            proto.Headers.CaseSense := false
            clsName := ParseCsv.Record.Prototype.__Class
            if this.QuoteChar {
                get := _GetMatch
                set := _SetMatch
            } else {
                get := _Get
                set := _Set
            }
            for header in Headers {
                name := RegExReplace(header, '\W', '')
                __get := get.Bind(A_Index)
                __set := set.Bind(A_Index)
                __get.DefineProp('Name', { Value: clsName '.Prototype.' name '.Get' })
                __set.DefineProp('Name', { Value: clsName '.Prototype.' name '.Set' })
                proto.DefineProp(name, { Get: __get, Set: __set })
                proto.Headers.Set(header, A_Index)
            }

            return

            _Get(Index, Self) {
                return Self.Get(Index)
            }
            _Set(Index, Self, Value) {
                return Self.Set(Value, Index)
            }
            _GetMatch(Index, Self) {
                return Self.Get(Index)['value']
            }
            _SetMatch(Index, Self, Value) {
                local __get, __set, desc, h, i
                for h, i in Self.Headers {
                    if i = Index {
                        name := RegExReplace(h, '\W', '')
                        break
                    }
                }
                desc := Self.Base.GetOwnPropDesc(name)
                __get := _Get.Bind(Index)
                __set := _Set.Bind(Index)
                __get.DefineProp('Name', { Value: desc.Get.Name })
                __set.DefineProp('Name', { Value: desc.Set.Name })
                Self.DefineProp(name, { Get: __get, Set: __set })

                return Self.Set(Value, Index)

                _Get(Index, Self) {
                    return Self.Get(Index)
                }
                _Set(Index, Self, Value) {
                    return Self.Set(Value, Index)
                }
            }
        }
        Headers => this.Prototype.Headers
    }

    class Record extends Array {
        static __New() {
            this.DeleteProp('__New')
            proto := this.Prototype
            desc := Array.Prototype.GetOwnPropDesc('__Item')
            proto.DefineProp('Get', { Call: desc.Get })
            proto.DefineProp('Set', { Call: desc.Set })
        }
        __Item[Header] {
            Get => this.Get(Header is Number ? Header : this.Headers.Get(Header))
            Set => this.Set(Value, Header is Number ? Header : this.Headers.Get(Header))
        }
    }

    class Options {
        static __New() {
            this.DeleteProp('__New')
            proto := this.Prototype
            proto.Breakpoint := 0
            proto.BreakpointAction := ''
            proto.Constructor := ''
            proto.Encoding := ''
            proto.FieldDelimiter := ','
            proto.FieldsContainRecordDelimiter := false
            proto.Headers := ''
            proto.InitialCapacity := 1000
            proto.PathIn := ''
            proto.QuoteChar := ''
            proto.RecordDelimiter := ''
            proto.Start := true
        }
        __New(Options?) {
            if IsSet(Options) {
                if IsSet(ParseCsvConfig) {
                    for prop in this.Base.OwnProps() {
                        if HasProp(Options, prop) {
                            this.%prop% := Options.%prop%
                        } else if HasProp(ParseCsvConfig, prop) {
                            this.%prop% := ParseCsvConfig.%prop%
                        }
                    }
                } else {
                    for prop in this.Base.OwnProps() {
                        if HasProp(Options, prop) {
                            this.%prop% := Options.%prop%
                        }
                    }
                }
            } else if IsSet(ParseCsvConfig) {
                for prop in this.Base.OwnProps() {
                    if HasProp(ParseCsvConfig, prop) {
                        this.%prop% := ParseCsvConfig.%prop%
                    }
                }
            }
            if this.HasOwnProp('__Class') {
                this.DeleteProp('__Class')
            }
        }
    }
}
