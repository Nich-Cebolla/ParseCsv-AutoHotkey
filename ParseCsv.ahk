#Requires AutoHotkey >=2.0.17


class ParseCsv {
    class Params {
        static Breakpoint := 0
        static BreakpointAction := ''
        static CollectionArrayBuffer := 1000
        static Constructor := ''
        static DisableLineEndingCheck := false
        static Encoding := ''
        static FieldDelimiter := ','
        static Headers := ''
        static MaxReadSizeBytes := 0
        static PathIn := ''
        static QuoteChar := ''
        static RecordDelimiter := ''
        static Start := true

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

    Fields := [], Index := 0, CharPos := 1, ContentLength := 0, Paused := 0, InputString := 0, Complete := 0
    __Item[Index] => this.Collection[Index]
    /**
     * @param {Object} [params] - An object with key:val pairs containin input parameters.
     * @property {Integer} [Breakpoint] - The number of records to parse before the `BreakpointAction`
     * is invoked. When using `Breakpoint` but not using `BreakpointAction`, `ParseCsv` returns the
     * instance object when the `Breakpoint` number of records have been parsed. To determine if the
     * function returned because it is finished or on a breakpoint, check `instance.Paused` which
     * will have a true value when the function returned due to a breakpoint, and a false value when
     * returned due to completion. You can loop this with a `while` loop.
     * @example
        while (instance := ParseCsv()).Paused {
            ; do work
        }
     * @
     * Note the above example is only relevant when `Breakpoint` is true and `BreakpointAction` is
     * not a function/unused. When `BreakpointAction` is a function, the process loops anyways, so
     * looping with that example is irrelevant.
     * @see {@link ParseCsv.Params.BreakpointAction}.
     * @property {Func|BoundFunc|Closure} [BreakpointAction] - If set, a function to call when the
     * `Breakpoint` is reached. If unset or any other value that is not a `Func`, `ParseCsv` will
     * return when `Breakpoint` is reached. The value returned is `this` (the instance object).
     * If using a `BreakpointAction` callback, you can direct `ParseCsv` to return by returning a
     * nonzero value.
     * @example
        MyCallback(instance) {
        ; -1 identifies the last record parsed.
        ; For headers that do not contain special characters, dot notation is acceptable, replacing
        ; spaces with underscores as needed.
            if instance.Collection[-1].Header_1 == 'Value'
                return 1
        }
     * @
     * @property {Integer} [CollectionArrayBuffer=1000] - When set, and when the current number of
     * records exceeds the length of the array, `ParseCsv` will add this number to the array's
     * length, so the array does not need to be constantly resized when adding items. This does not
     * impact its usage. You can still enumerate the collection using a standard `for record in
     * collection` loop, and access the items and properties of the array using standard notation
     * and methods. The methods account for the unused space. To clear the array use the
     * `instance.Clear` method or the `instance.Collection.Clear` method. When unset or blank,
     * `CollectionArrayBuffer` defaults to 1000. When `ParseCsv` completes, the array is shrunk to
     * its actual size, and the property `instance.Collection` is reassigned the value of
     * `instance.Collection.__Collection`, which is a standard array. Said in another way, when the
     * procedure finishes, the collection object is a standard array and the information in this
     * paragraph is no longer relevant. This information is relevant only when using `Breakpoint`s.
     * @property {Func|BoundFunc|Closure|Class} [Constructor] - A function or class constructor that
     * will be called every time a complete record is parsed. The function will receive an array
     * of strings, each string representing one of the fields. If you need the headers along with
     * the strings in order for the function to handle the input, you can get the headers from
     * `ParseCsv.Headers` (off the class object, not the instance object). The function should return
     * the object to be added to the `instance.Collection` array. If not set or blank, the
     * default constructor is used. The default constructor creates a `ParseCsv.Record` object.
     * `ParseCsv.Record.Prototype` behaves like a map object. You can enumerate its contents with a
     * `for header, field in record` loop. You can access it's items using `record[header]` notation.
     * I also enabled the ability to use `record.header` notation if you prefer, and if the headers
     * do not contain special characters that are invalid for properties. If you prefer property notation
     * and the header contains spaces, you can replace the spaces with underscores in the property
     * name. The names are still case-sensitive even when using property notation. If you would
     * prefer that these are not case sensitive, see the documentation for a code block that sets
     * all map objects to case-insensitive by default. All other methods available to map objects
     * are also available to `ParseCsv.Record` objects.
     * @property {Boolean} [DisableLineEndingCheck] - If the `RecordDelimiter` does not match the
     * line endings within the input content, an error may occur during procesing. If this occurs,
     * and if `DisableLineEndingCheck` is false, the `ParseCsv` will check the line endings and display
     * the information in a message box prior to throwing the error. If `DisableLineEndingCheck` is
     * true, the error is thrown without the additional message box.
     * @property {string} [Encoding] - The encoding of the file. If not set, the default encoding is used.
     * @property {string} [FieldDelimiter] - The string that separates fields. For example, a comma.
     * @property {string|Array} [Headers] - If set, the headers are used to create the `ParseCsv.Record`
     * objects. If unset, the first record in the CSV input is used.
     * @property {number} [MaxReadSizeBytes] - The maximum number of bytes to read from the file at
     * a time. This is intended to be used for very large CSVs. This setting is ignored when
     * `QuoteChar` is false `RecordDelimiter` is blank or newlines, and a `PathIn` is set, as the
     * document is parsed by line anyway. This is also ignored when `PathIn` is unset and `InputString`
     * is set, as there is no document. When `MaxReadSizeBytes` is used, the input document is parsed
     * in segments until the entire document has been parsed. If the reason for using `MaxReadSizeBytes`
     * is due to memory constraints or very large input, you will also want to see
     * {@link ParseCsv.Params.Breakpoint}. Without a breakpoint, `ParseCsv` will parse the entire
     * document (though in segments), which will still cause memory issues.
     * @property {string} [PathIn] - The path to the file to parse.
     * @property {string} [QuoteChar] - The character used to quote strings. This must be set when
     * some or all fields are quoted. If not set, the fields are assumed to be unquoted, which will
     * most likely result in an error or an incorrectly parsed CSV if the CSV actually does contain
     * quoted fields. If the fields are not quoted and this is set, `ParseCsv` will still parse
     * correctly but will be slower.
     * @property {string} [RecordDelimiter] - The string that separates records. This option impacts
     * some of the behavior of `ParseCsv`. If not set, depending on other options `ParseCsv` will
     * attempt to identify the newline characters used in the document and will use those. If there
     * are mixed endings and `ParseCsv` cannot handle mixed endings in that scenario, you will see
     * an error. But if the endings do not matter for `ParseCsv`'s operation, it will still
     * parse the input.
     * @property {Boolean} [Start] - If true, the procedure is called immediately upon instantiation.
     * If false, you will receive an instance of `ParseCsv` and you can controll it at-will by calling
     * any of its methods.
     * @param {string} [InputString] - A string to parse. If not set, the file at `PathIn` is used.
     */
    __New(params?, InputString?) {
        params := ParseCsv.Params(params??unset)
        this.params := params, this.Collection := ParseCsv.Collection(params.Constructor, params.CollectionArrayBuffer)
        if IsSet(InputString)
            this.Content := InputString, this.InputString := true
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
    /**
     * @description - Returns the count for the respective input.
     * - "Records" or "R" - the count for how many records currently exist in instance.Collection
     * - "Headers" or "H" - the count for how many headers exist in instance.Headers
     */
    Count[Which := 'R'] {
        Get {
            if (W := SubStr(Which, 1, 1)) = 'R'
                return this.Collection.Length
            else if W = 'H'
                return this.Headers.Length
            else
                throw ValueError('Unexpected input for ``Which``.', -1, 'Specifically: ' Which)
        }
    }
    /**
     * @returns {String} - Returns the pattern that will parse the CSV according to the input values.
     */
    static GetPattern(Quote, FieldDelimiter, RecordDelimiter, Columns) {
        RecordDelimiter := StrReplace(StrReplace(RecordDelimiter, '`n', '\n'), '`r', '\r')
        if RecordDelimiter
            RD1 := RegExReplace(RecordDelimiter, '\\[rnR]|`r|`n', ''), RD2 := RD1||'[\r\n]+'
        else
            RD1 := '', RD2 := '[\r\n]+'
        pattern := Format('JS)(?<=^|{1})', RecordDelimiter||'[\r\n]')
        part := Format('(?:({1}(?:[^{1}]*+(?:{1}{1})*+)*+{1}|[^\r\n{1}{2}{3}]*+){2})', Quote, FieldDelimiter, RD1)
        ; I decided to use a loop to dynamically construct the pattern, instead of a recursive pattern,
        ; because it allows us to capture every field in each record all at once. This works for CSV
        ; since we know how many fields there will be per record after getting the headers.
        Loop Columns - 1
            pattern .= part
        pattern .= Format('({1}(?:[^{1}]*+(?:{1}{1})*+)*+{1}|[^\r\n{1}{2}{3}]*+)(?:{4}|$(*MARK:end))'
        , Quote, FieldDelimiter, RD1, RD2)
        A_Clipboard := pattern
        try
            RegExMatch(' ', Pattern)
        catch Error as err {
            if InStr(err.Message, 'Compile error 25')
                throw Error('The procedure received "' err.Message '". To fix this, change the'
                ' ``RecordDelimiter`` and/or ``FieldDelimiter`` to a value that is a fixed length.', -1)
            else
                throw err
        }
        return pattern
    }
    /**
     * @description The general-purpose activation method. If `Start` is false, you should use this
     * method to start the parsing procedure, as it contains two instantiation methods that
     * are necessary. Or call `PrepareContent` and `SetHeaders` individually. `Call` redirects to
     * `Parse`, which itself redirects to the correct loop method. After instantiation, you can
     * call any method to invoke an action, but `PrepareContent` and `SetHeaders` are required first.
     * @returns {ParseCsv} - Returns `this` (the instance object, `ParseCsv.Prototype`).
     */
    Call() {
        if !this.Paused
            this.PrepareContent(), this.SetHeaders()
        this.Paused := false, this.Parse()
        if !this.Paused {
            this.End()
        }
        return this
    }
    /**
     * @description Clears `instance.Collection`. To be used when parsing very large files, or to
     * free memory for any general reason.
     */
    Clear() => this.Collection.Clear()
    /**
     * @description - Redirects `for` loop enumeration to `instance.Collection`.
     */
    __Enum(VarCount) => this.Collection.__Enum(VarCount)
    /**
     * @description Handles two end-of-procedure tasks: Closes the file if one was opened, and
     * assigns `instance.Collection` the value of `instance.Collection.__Collection`, so the
     * collection is a standard Array instead of a `ParseCsv.Collection` object.
     */
    End() {
        if this.HasOwnProp('File')
            this.File.Close()
        this.Collection.__Collection.Length := this.Collection.Count
        this.Collection := this.Collection.__Collection
    }
    /**
     * @description - Searches the CSV for the first field which contains the input string
     * and returns the index number of the record which contains the field.
     * @param {String} StrToFind - The string to find.
     * @param {String|Integer|Array} [Headers] -
     * - If a string, the header name to search within.
     * - If an integer, the index value of the header to search within. If using an integer to refer
     * to a header by index, its type must be `Number` or `Integer`, and not `String`. If it is a
     * `String`, it will be considered a header name and not an index number.
     * - If an array, an array of integers or strings as described above. These will be searched
     * in the order they are in the array.
     * - If unset, all headers will be searched.
     * @param {Integer} [IndexStart=1] - The record index number to start searching from.
     * @param {Integer} [IndexEnd] - The record index number to end searching at. If unset, the
     * all records after and including IndexStart are searched.
     * @param {Boolean} [RequireFullMatch=true] - If true, the field must match the input string
     * exactly. If false, the field must contain the input string.
     * @param {Boolean} [CaseSensitive=false] - If true, the search is case-sensitive. If false,
     * the search is case-insensitive.
     * @param {Integer} [StartingPos] - This is only used when `RequireFullMatch` is false. The
     * position of the field's string (in number of characters) to search within. This is passed
     * to `InStr`.
     * @param {VarRef} [OutField] - This variable will receive the string value of the field that
     * contains the input string.
     * @param {VarRef} [OutHeader] - This variable will receive the string value of the header name
     * that contains the field that contains the input string.
     * @returns {Integer} - The index number of the record which satisfies the conditions set by
     * the input parameters.
     */
    Find(StrToFind, Headers?, IndexStart := 1, IndexEnd?, RequireFullMatch := true
    , CaseSensitive := false, StartingPos := 1, &OutField?, &OutHeader?) {
        if !IsSet(IndexEnd)
            IndexEnd := this.Count['R']
        Headers := this.__GetHeaders(Headers ?? unset)
        if RequireFullMatch
            Process := CaseSensitive ? _Process_RFM_CS : _Process_RFM
        else
            Process := _Process
        i := IndexStart - 1
        while ++i <= IndexEnd {
            for Header in Headers {
                if Process(&Header) {
                    OutHeader := Header
                    OutField := this[i][Header]
                    return i
                }
            }
        }
        _Process(&Header) => InStr(this[i][Header], StrToFind, CaseSensitive, StartingPos)
        _Process_RFM(&Header) => this[i][Header] = StrToFind
        _Process_RFM_CS(&Header) => this[i][Header] == StrToFind
    }
    /**
     * @description - Iterates the fields in the CSV, passing the values to a callback function.
     * When the function returns true, this function returns the index number of the record.
     * @param {Func|BoundFunc|Closure} Callback - The callback function. When the function returns
     * any true value, this function also returns. The function can accept up to four parameters:
     * - {String} The current field's value
     * - {Integer} The current record index number
     * - {String} The current header name
     * @param {String|Integer|Array} [Headers] -
     * - If a string, the header name to search within.
     * - If an integer, the index value of the header to search within. If using an integer to refer
     * to a header by index, its type must be `Number` or `Integer`, and not `String`. If it is a
     * `String`, it will be considered a header name and not an index number.
     * - If an array, an array of integers or strings as described above. These will be searched
     * in the order they are in the array.
     * - If unset, all headers will be searched.
     * @param {Integer} [IndexStart=1] - The record index number to start searching from.
     * @param {Integer} [IndexEnd] - The record index number to end searching at. If unset, the
     * all records after and including IndexStart are searched.
     * @param {VarRef} [OutField] - This variable will receive the string value of the field that
     * contains the input string.
     * @param {VarRef} [OutHeader] - This variable will receive the string value of the header name
     * that contains the field that contains the input string.
     * @returns {Integer} - The index number of the record that contains the field that was passed
     * to the function when the function returned true.
     */
    FindF(Callback, Headers?, IndexStart := 1, IndexEnd?, &OutField?, &OutHeader?) {
        if !IsSet(IndexEnd)
            IndexEnd := this.Count['R']
        Headers := this.__GetHeaders(Headers ?? unset)
        i := IndexStart - 1
        while ++i <= IndexEnd {
            for Header in Headers {
                if Callback(this[i][Header], i, Header) {
                    OutHeader := Header
                    OutField := this[i][Header]
                    return i
                }
            }
        }
    }
    /**
     * @description - Searches the CSV for a field that matches with the input pattern using
     * `RegExMatch`.
     * @param {String} Pattern - The Regular Expression pattern to match with.
     * @param {String|Integer|Array} [Headers] -
     * - If a string, the header name to search within.
     * - If an integer, the index value of the header to search within. If using an integer to refer
     * to a header by index, its type must be `Number` or `Integer`, and not `String`. If it is a
     * `String`, it will be considered a header name and not an index number.
     * - If an array, an array of integers or strings as described above. These will be searched
     * in the order they are in the array.
     * - If unset, all headers will be searched.
     * @param {Integer} [IndexStart=1] - The record index number to start searching from.
     * @param {Integer} [IndexEnd] - The record index number to end searching at. If unset, the
     * all records after and including IndexStart are searched.
     * @param {Integer} [StartingPos=1] - The position within the field (in number of characters)
     * to begin searching for a match. This is passed directly to `RegExMatch`.
     * @param {VarRef} [OutMatch] - This variable will receive the `RegExMatchInfo` object.
     * @param {VarRef} [OutField] - This variable will receive the string value of the field that
     * contains the input string.
     * @param {VarRef} [OutHeader] - This variable will receive the string value of the header name
     * that contains the field that contains the input string.
     * @returns {Integer} - The index number of the record that contains the field that matched
     * the pattern.
     */
    FindR(Pattern, Headers?, IndexStart := 1, IndexEnd?, StartingPos := 1, &OutMatch?, &OutField?, &OutHeader?) {
        if !IsSet(IndexEnd)
            IndexEnd := this.Count['R']
        Headers := this.__GetHeaders(Headers ?? unset)
        i := IndexStart - 1
        while ++i <= IndexEnd {
            for Header in Headers {
                if RegExMatch(this[i][Header], Pattern, &OutMatch, StartingPos) {
                    OutHeader := Header
                    OutField := this[i][Header]
                    return i
                }
            }
        }
    }
    /**
     * @description - Loops the CSV between `IndexStart` and `IndexEnd`, adding the results from
     * `Find` to an array. When multiple headers are included in the search, the csv is
     * iterated by searching all the records between IndexStart and IndexEnd for one header before
     * moving on to the next header. The arrays themselves are added to a `Map` object, where the
     * key is the header as it is passed to the `Headers` parameter (meaning if indices are used, the
     * keys are integers, else the keys are strings), and the value is the array. If the `Find`
     * method does not return a value for a given header, the key is excluded from the resulting
     * `Map` object.
     * @param {String} StrToFind - The string to find.
     * @param {String|Integer|Array} [Headers] -
     * - If a string, the header name to search within.
     * - If an integer, the index value of the header to search within. If using an integer to refer
     * to a header by index, its type must be `Number` or `Integer`, and not `String`. If it is a
     * `String`, it will be considered a header name and not an index number.
     * - If an array, an array of integers or strings as described above. These will be searched
     * in the order they are in the array.
     * - If unset, all headers will be searched.
     * @param {Integer} [IndexStart=1] - The record index number to start searching from.
     * @param {Integer} [IndexEnd] - The record index number to end searching at. If unset, the
     * all records after and including IndexStart are searched.
     * @param {Boolean} [RequireFullMatch=true] - If true, the field must match the input string
     * exactly. If false, the field must contain the input string.
     * @param {Boolean} [CaseSensitive=false] - If true, the search is case-sensitive. If false,
     * the search is case-insensitive.
     * @param {Integer} [StartingPos] - This is only used when `RequireFullMatch` is false. The
     * position of the field's string (in number of characters) to search within. This is passed
     * to `InStr`.
     * @param {Boolean} [IncludeField=true] - If true, the field value is included in the output.
     * If false, the output only includes the index numbers.
     * @returns {Map} - If the method returns a value at least one time, this returns a map object
     * with the following characteristics:
     * - The keys contained in the object are header names, unless indices are passed to `Headers`,
     * then the keys are those values.
     * - The values depend on the value of `IncludeField`. When `IncludeField` is true, the values
     * are arrays of objects with properties { Field, Index }. When false, The values are arrays of
     * integers which are returned by `Find`.
     * If no values are returned by `Find`, this returns an empty string.
     */
    FindAll(StrToFind, Headers?, IndexStart := 1, IndexEnd?, RequireFullMatch := true
    , CaseSensitive := false, StartingPos := 1, IncludeField := true) {
        return this.__FindAll(
            _Process.Bind(StrToFind, IndexStart, IndexEnd ?? this.Count['R']
                , RequireFullMatch, CaseSensitive, StartingPos, IncludeField
            )
            , Headers ?? unset
        )
        _Process(StrToFind, IndexStart, IndexEnd, RequireFullMatch, CaseSensitive, StartingPos, IncludeField, &Header) {
            local Result := [], Field
            Add := IncludeField ? () => Result.Push({ Index: i, Field: Field }) : () => Result.Push(i)
            i := IndexStart
            while i <= IndexEnd {
                if i := this.Find(StrToFind, Header, i, IndexEnd, RequireFullMatch, CaseSensitive, StartingPos, &Field)
                    Add()
                else
                    break
                i++
            }
            return Result.Length ? Result : ''
        }
    }
    
    /**
     * @description - Loops the CSV between `IndexStart` and `IndexEnd`, adding the results from
     * `FindF` to an array. When multiple headers are included in the search, the csv is
     * iterated by searching all the records between IndexStart and IndexEnd for one header before
     * moving on to the next header. The arrays themselves are added to a `Map` object, where the
     * key is the header as it is passed to the `Headers` parameter (meaning if indices are used, the
     * keys are integers, else the keys are strings), and the value is the array. If the `FindF`
     * method does not return a value for a given header, the key is excluded from the resulting
     * `Map` object.
     * @param {Func|BoundFunc|Closure} Callback - The callback function passed to `FindF`. When the
     * function returns any true value, the result from `FindF` is added to an array.
     * The function can accept up to four parameters:
     * - {String} The current field's value
     * - {Integer} The current record index number
     * - {String} The current header name
     * @param {String|Integer|Array} [Headers] -
     * - If a string, the header name to search within.
     * - If an integer, the index value of the header to search within. If using an integer to refer
     * to a header by index, its type must be `Number` or `Integer`, and not `String`. If it is a
     * `String`, it will be considered a header name and not an index number.
     * - If an array, an array of integers or strings as described above. These will be searched
     * in the order they are in the array.
     * - If unset, all headers will be searched.
     * @param {Integer} [IndexStart=1] - The record index number to start searching from.
     * @param {Integer} [IndexEnd] - The record index number to end searching at. If unset, the
     * all records after and including IndexStart are searched.
     * @param {Boolean} [IncludeField=true] - If true, the field value is included in the output.
     * If false, the output only includes the index numbers.
     * @returns {Map} - If the method returns a value at least one time, this returns a map object
     * with the following characteristics:
     * - The keys contained in the object are header names, unless indices are passed to `Headers`,
     * then the keys are those values.
     * - The values depend on the value of `IncludeField`. When `IncludeField` is true, the values
     * are arrays of objects with properties { Field, Index }. When false, The values are arrays of
     * integers which are returned by `FindF`.
     * If no values are returned by `FindF`, this returns an empty string.
     */
    FindAllF(Callback, Headers?, IndexStart := 1, IndexEnd?, IncludeField := true) {
        return this.__FindAll(
            _Process.Bind(Callback, IndexStart, IndexEnd ?? this.Count['R'], IncludeField)
            , Headers ?? unset
        )
        _Process(Callback, IndexStart, IndexEnd, IncludeField, &Header) {
            local Result := [], Field
            i := IndexStart
            Add := IncludeField ? () => Result.Push({ Index: i, Field: Field }) : () => Result.Push(i)
            while i <= IndexEnd {
                i := this.FindF(Callback, Header, i, IndexEnd, &Field)
                if i
                    Add()
                else
                    break
                i++
            }
            return Result.Length ? Result : ''
        }
    }
    
    /**
     * @description - Loops the CSV between `IndexStart` and `IndexEnd`, adding the results from
     * `FindR` to an array. When multiple headers are included in the search, the csv is
     * iterated by searching all the records between IndexStart and IndexEnd for one header before
     * moving on to the next header. The arrays themselves are added to a `Map` object, where the
     * key is the header as it is passed to the `Headers` parameter (meaning if indices are used, the
     * keys are integers, else the keys are strings), and the value is the array. If the `FindR`
     * method does not return a value for a given header, the key is excluded from the resulting
     * `Map` object.
     * @param {String} Pattern - The Regular Expression pattern to match with.
     * @param {String|Integer|Array} [Headers] -
     * - If a string, the header name to search within.
     * - If an integer, the index value of the header to search within. If using an integer to refer
     * to a header by index, its type must be `Number` or `Integer`, and not `String`. If it is a
     * `String`, it will be considered a header name and not an index number.
     * - If an array, an array of integers or strings as described above. These will be searched
     * in the order they are in the array.
     * - If unset, all headers will be searched.
     * @param {Integer} [IndexStart=1] - The record index number to start searching from.
     * @param {Integer} [IndexEnd] - The record index number to end searching at. If unset, the
     * all records after and including IndexStart are searched.
     * @param {Integer} [StartingPos=1] - The position within the field (in number of characters)
     * to begin searching for a match. This is passed directly to `RegExMatch`.
     * @param {Boolean} [IncludeMatch=true] - If true, the match object is included in the output.
     * @param {Boolean} [IncludeField=true] - If true, the field value is included in the output.
     * @returns {Map} - If the method returns a value at least one time, this returns a map object
     * with the following characteristics:
     * - The keys contained in the object are header names, unless indices are passed to `Headers`,
     * then the keys are those values.
     * - The values depend on the value of `IncludeMatch` and `IncludeField`. When one or both of
     * `IncludeMatch` or `IncludeField` is true, the values are arrays of objects with the respective
     * properties { Field, Index, Match }. When both are false, The values are arrays of integers
     * which are returned by `FindR`.
     * If no values are returned by `FindR`, this returns an empty string.
     */
    FindAllR(Pattern, Headers?, IndexStart := 1, IndexEnd?, StartingPos := 1, IncludeMatch := true, IncludeField := true) {
        return this.__FindAll(_Process.Bind(
                Pattern, IndexStart, IndexEnd ?? this.Count['R'], StartingPos, IncludeMatch, IncludeField
            )
            , Headers ?? unset
        )
        _Process(Pattern, IndexStart, IndexEnd, StartingPos, IncludeMatch, IncludeField, &Header) {
            local Result := [], Match, Field
            i := IndexStart
            if IncludeMatch
                Add := IncludeField ? () => Result.Push({ Index: i, Match: Match , Field: Field }) : () => Result.Push({ Index: i, Match: Match })
            else
                Add := IncludeField ? () => Result.Push({ Index: i, Field: Field }) : () => Result.Push(i)
            while i <= IndexEnd {
                i := this.FindR(Pattern, Header, i, IndexEnd, StartingPos, &Match, &Field)
                if i
                    Add()
                else
                    break
                i++
            }
            return Result.Length ? Result : ''
        }
    }

    /** @description - Used internally when any of `FindAll`, `FindAllF`, or `FindAllR` are called. */
    __FindAll(Callback, Headers?) {
        if IsSet(Headers) {
            if not Headers is Array
                Headers := [Headers]
        } else
            Headers := this.Headers
        Result := Map()
        for Header in Headers
            Result.Set(Header, Callback(&Header) || unset)
        return Result.Count ? Result : ''
    }
    /** @description - Used internally when any of `Find`, `FindF`, `FindR`, or `Join` are called */
    __GetHeaders(Headers?) {
        if IsSet(Headers) {
            if not Headers is Array
                Headers := [Headers]
            Result := []
            for Header in Headers
                Result.Push(Header is Number ? this.Headers[Header] : Header)
            return Result
        } else
            return this.Headers
    }

    /**
     * @description - Loops the CSV between IndexStart and IndexEnd. For each record, combines the
     * fields into a string separated by a delimiter. If joining multiple records, the records
     * are separated by a newline. Note that, if a quote character was included in the input parameters
     * when `ParseCsv` was called, the fields will be enclosed by the quote character, and internal
     * quote characters will be escaped by doubling them. This will occur for all fields, whether
     * or not the field was originally quoted in the input content.
     * @param {Integer} [IndexStart=1] - The starting record index.
     * @param {Integer} [IndexEnd] - The ending record index. If unset, the last record is used.
     * @param {String|Integer|Array} [Headers] -
     * - If a string, the header name to search within.
     * - If an integer, the index value of the header to search within. If using an integer to refer
     * to a header by index, its type must be `Number` or `Integer`, and not `String`. If it is a
     * `String`, it will be considered a header name and not an index number.
     * - If an array, an array of integers or strings as described above. These will be searched
     * in the order they are in the array.
     * - If unset, all headers will be searched.
     * @param {Boolean} [IncludeHeaders=false] - If true, the headers are included in the output as
     * the first line.
     * @param {String} [Delimiter] - The string to use to separate the fields. If unset, this uses
     * the `FieldDelimiter` that was assigned when `ParseCsv` was called
     * @param {String} [Newline='`r`n'] - The string to use to separate the records.
     * @returns {String} - The combined string.
     */
    Join(IndexStart := 1, IndexEnd?, Headers?, IncludeHeaders := false, Delimiter := this.Params.FieldDelimiter, Newline := '`r`n') {
        if !IsSet(IndexEnd)
            IndexEnd := this.Count['R']
        Headers := this.__GetHeaders(Headers ?? unset)
        i := IndexStart - 1
        quote := this.Params.QuoteChar
        Add := this.Params.QuoteChar ? _AddWithQuote : _Add
        if IncludeHeaders {
            if this.Params.QuoteChar {
                for Header in Headers
                    Str .= Delimiter quote StrReplace(Header, quote, quote quote) quote
            } else {
                for Header in Headers {
                    Str .= Delimiter Header
                }
            }
            Str := SubStr(Str, StrLen(Delimiter) + 1) Newline
        }
        while ++i <= IndexEnd {
            for header in Headers
                Str .= Add(&Header)
            Str .= Newline
        }
        return SubStr(Str, 1, StrLen(Str) - StrLen(Newline))

        _Add(&Header) => (A_Index == 1 ? '' : Delimiter) this[i][header]
        _AddWithQuote(&Header) => (A_Index == 1 ? '' : Delimiter) quote StrReplace(this[i][header], quote, quote quote) quote
    }

    /** @returns {Float} - Returns the progress of the parsing procedure as a float between 0 and 1. */
    GetProgress() {
        if this.ReadStyle == 'Line'
            return this.File.Pos / this.File.Length
        else if this.ReadStyle == 'Quote' {
            if this.HasOwnProp('File')
                return (this.File.Pos - StrPut(SubStr(this.Content, this.CharPos + 1), this.params.Encoding||'UTF-8')) / this.File.Length
            else
                return this.CharPos / this.ContentLength
        } else if this.ReadStyle == 'Split' {
            if this.HasOwnProp('File') {
                len := 0
                Loop this.Content.Length
                    len += StrLen(this.Content[A_Index])
                return (this.File.Pos - StrPut(len, this.params.Encoding||'UTF-8')) / this.File.Length
            } else
                return this.Content.Length / this.ContentLength
        } else
            throw ValueError('Unexpected read style: ' this.ReadStyle, -1)
    }
    /**
     * @description This loops the input file one line at a time. It is used when the following are true:
     * - The `RecordDelimiter` is a newline.
     * - The fields are not quoted.
     * - `PathIn` is set.
     */
    LoopReadLine() {
        local params := this.params, BPA := params.BreakpointAction, f := this.File, Collection := this.Collection
        if params.Breakpoint {
            loop params.Breakpoint {
                Collection.__Add(StrSplit(f.ReadLine(), params.FieldDelimiter))
                if f.AtEOF
                    break
            }
            this.Paused := true
            if not BPA is Func || BPA(this)
                return
            this.Paused := false
        } else {
            while !f.AtEOF {
                Collection.__Add(StrSplit(f.ReadLine(), params.FieldDelimiter))
            }
        }
    }
    /**
     * @description This loops the input value using `RegExMatch`. This is the slowest of the methods,
     * but necessary for correctly handling quoted fields. This is used when `QuoteChar` is set.
     */
    LoopReadQuote() {
        local params := this.params, Pattern := this.Pattern, BPA := params.BreakpointAction, Collection := this.Collection
        , LastMatch
        if params.Breakpoint {
            Loop params.Breakpoint {
                if RegExMatch(this.content, Pattern, &match, this.CharPos) {
                    if match.mark == 'end' {
                        if this.HasOwnProp('File') && this.File.AtEOF {
                            _Process()
                            return
                        } else
                            _Read()
                    } else
                        _Process()
                    LastMatch := match
                } else
                    _Read()
            }
            this.Paused := true
            if not BPA is Func || BPA(this)
                return
            this.Paused := false
        } else {
            while RegExMatch(this.content, Pattern, &match, this.CharPos) {
                if match.mark == 'end' {
                    if !this.HasOwnProp('File') || this.File.AtEOF {
                        _Process()
                        return
                    } else
                        _Read()
                } else
                    _Process()
                LastMatch := match
            }
            _Read()
        }
        _Process() {
            if match.Pos != this.CharPos
                throw ValueError(Format('Invalid CSV format. `nThe current match should have began'
                ' at position {1}, but the match occurred at position {2}.`nThe invalid matched'
                ' content:`n{3}', this.CharPos, match.Pos, match[0]), -1)
            Fields := [], Fields.Length := this.RecordLength
            loop this.RecordLength {
                if SubStr(match[A_Index], 1, 1) == this.params.QuoteChar && SubStr(match[A_Index], -1, 1) == this.params.QuoteChar
                    Fields[A_Index] :=  SubStr(match[A_Index], 2, match.Len[A_Index] - 2)
                else
                    Fields[A_Index] := match[A_Index]
            }
            Collection.__Add(Fields)
            this.CharPos := match.Pos + match.len
        }
        _Read() {
            if !this.HasOwnProp('File')
                return
            this.File.Pos -= StrPut(SubStr(this.Content, LastMatch.pos + LastMatch.len + 1), params.Encoding||'UTF-8')
            this.ReadNextQuote()
        }
    }
    /**
     * @description This loops the input value using `StrSplit`. This is used when `LoopReadLine` is
     * not possible, and the fields are not quoted. This is faster than `LoopReadQuote`.
     */
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
                    this.Collection.__Add(StrSplit(this.Content.RemoveAt(1), params.FieldDelimiter))
            } else {
                Loop params.Breakpoint
                    this.Collection.__Add(StrSplit(this.Content.RemoveAt(1), params.FieldDelimiter))
            }
            this.Paused := true
            if not BPA is Func || BPA(this)
                return
            this.Paused := false
        } else
            while !_LoopContentLength()
                continue

        _LoopContentLength(&len?) {
            Loop this.Content.Length
                this.Collection.__Add(StrSplit(this.Content[A_Index], params.FieldDelimiter))
            if !this.HasOwnProp('File') || this.File.AtEOF
                return 1
            this.ReadNextSplit()
            len := this.Content.Length
        }
    }
    /** @description Redirects the function to the correct LoopRead method. */
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
    /**
     * @description An instantiation function that handles preparing the content for parsing.
     * Content is accessible from `instance.Content`. This must be called for `ParseCsv` to function,
     * but is generally handled automatically by `instance.Call`.
     */
    PrepareContent() {
        local params := this.params
        if this.InputString {
            _GetLineEndings()
            if this.ReadStyle == 'Split'
                this.Content := StrSplit(this.Content, params.RecordDelimiter)
            else
                this.ContentLength := StrLen(this.Content)
            return
        }
        if this.ReadStyle == 'Line' || params.MaxReadSizeBytes {
            this.File := FileOpen(params.PathIn, 'r', params.Encoding||unset)
            this.File.Read(1), this.File.Pos := 0
            if this.ReadStyle == 'Line'
                return
            this.Content := this.File.Read(params.MaxReadSizeBytes)
        } else
            this.Content := FileRead(params.PathIn, params.Encoding||unset)
        _GetLineEndings()
        if this.ReadStyle == 'Split' {
            this.Content := StrSplit(this.Content, params.RecordDelimiter)
            if params.MaxReadSizeBytes {
                if !this.File.AtEOF
                    this.File.Pos -= StrPut(this.Content.Pop(), params.Encoding||'UTF-8')
            } else
                this.ContentLength := this.Content.Length
        } else if !params.MaxReadSizeBytes
            this.ContentLength := StrLen(this.Content)
        _GetLineEndings() {
            if !params.RecordDelimiter {
                StrReplace(this.Content, '`r`n', , , &CRLFCount)
                StrReplace(this.Content, '`n', , , &LFCount), StrReplace(this.Content, '`r', , , &CRCount)
                if CRLFCount && CRCount == LFCount
                    params.RecordDelimiter := '`r`n'
                else if CRCount > LFCount
                    params.RecordDelimiter := '`r'
                else
                    params.RecordDelimiter := '`n'
            }
        }
    }
    /** @description Reads the next segment of the content, depending on the read style. */
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
     * @description Constructs a record from a line of text in a file. The intended use for this
     * method is in a situation where one wishes to check each individual item at a time for a
     * value or condition. Looping with `ReadNextLine` requires slightly fewer calculations to
     * accomplish the task, compared to setting `Breakpoint` to 1 and `BreakpointAction` to the
     * function that handles this. To get the last parsed record, use `instance.Collection[-1]`.
     */
    ReadNextLine() {
        this.Collection.__Add(StrSplit(this.File.ReadLine(), this.params.FieldDelimiter))
    }
    /**
     * @description Reads the next segment of the content when the content contains quoted fields.
     * You shouldn't need to call this method directly, as the loop read methods will call it
     * automatically.
     */
    ReadNextQuote() {
        this.Content := this.File.Read(this.File.Pos + this.params.MaxReadSizeBytes > this.File.Length
        ? this.File.Length - this.File.Pos : this.params.MaxReadSizeBytes)
        this.CharPos := 1
    }
    /**
     * @description Reads the next segment of the content when the content is split by a delimiter.
     * You shouldn't need to call this method directly, as the loop read methods will call it
     * automatically.
     */
    ReadNextSplit() {
        this.Content := this.File.Read(this.File.Pos + this.params.MaxReadSizeBytes > this.File.Length
        ? this.File.Length - this.File.Pos : this.params.MaxReadSizeBytes)
        this.ContentLength := StrLen(this.Content)
        this.Content := StrSplit(this.Content, this.params.RecordDelimiter||this.__GetLineEndings(true))
        ; The last item in the array is probably not a complete record, so it is removed.
        if !this.File.AtEOF {
            this.File.Pos -= StrPut((len:=StrLen(this.Content.Pop())), this.params.Encoding||'UTF-8')
            this.ContentLength -= len
        }
    }
    /**
     * @description Handles setting the header values. This must be called for `ParseCsv` to function,
     * but is generally handled automatically by `instance.Call`.
    */
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

        this.RecordLength := this.Collection.Headers.Length, this.Headers := this.Collection.Headers
        if this.ReadStyle == 'Quote'
            this.Pattern := ParseCsv.GetPattern(params.QuoteChar, params.FieldDelimiter, params.RecordDelimiter, this.RecordLength)

        _SetHeadersLoopReadQuote() {
            PatternHeader := Format('JS)(?:{1}(?<value>(?:[^{1}]*+(?:{1}{1})*+)*+){1}|'
            '(?<value>[^\r\n{1}{2}{3}]*+))(?:{2}|{4}(*MARK:record)|$(*MARK:end))|(?:{2}$(*MARK:end))'
            , params.QuoteChar, params.FieldDelimiter
            , params.RecordDelimiter ? RegExReplace(params.RecordDelimiter, '\\[rnR]|`r|`n', '') : ''
            , params.RecordDelimiter ? StrReplace(StrReplace(params.RecordDelimiter, '`n', '\n'), '`r', '\r') : '[\r\n]+')
            A_Clipboard := PatternHeader
            sleep 1
            headers := [], pos := 1
            while RegExMatch(this.Content, PatternHeader, &match, pos) {
                if match.Pos != pos {
                    if !this.Params.DisableLineEndingCheck {
                        MsgBox((
                            'There is a formatting error near position ' pos '`r`n'
                            'This may be caused by incorrect line endings.`r`n'
                            this.__CheckLineEndings()
                            'To disable this message (so only the error is thrown), set'
                            ' ``DisableLineEndingCheck`` to true on the input parameters.`r`n'
                            'The script will now exit.'
                        ))
                    }
                    throw ValueError('There is a formatting error near position ' pos '. This may be'
                    ' caused by line endings that do not match with ``RecordDelimiter``.', -1)
                }
                pos := match.Pos + match.len
                headers.Push(match['value'])
                h .= match[0]
                if match.mark {
                    if match.mark == 'end' {
                        throw Error('The content ended before all of the headers were collected.'
                        ' If you are using ``MaxReadSizeBytes``, make sure the allotted space is'
                        ' enough to capture all of the headers.', -1)
                    } else if match.mark == 'record' {
                        this.Content := StrReplace(this.Content, h, '')
                        this.Collection.SetHeaders(headers)
                        break
                    }
                }
            }
        }
    }
    __CheckLineEndings() {
        StrReplace(this.Content, '`r', , , &CRCount)
        StrReplace(this.Content, '`n', , , &LFCount)
        return (
            '``RecordDelimiter`` = "' StrReplace(StrReplace(this.Params.RecordDelimiter, '`r', '``r')
            , '`n', '``n') '".`r`n'
            'CR Count = ' CRCount '.`r`n'
            'LF Count = ' LFCount '.`r`n'
        )
    }

    class Collection {
        Count := 0, BaseObj := {}, __Collection := Array()
        __New(Constructor, BufferLength) {
            if Constructor
                this.DefineProp('__MakeRecord', {Call: _MakeRecord.Bind(Constructor)})
            this.BufferLength := this.__Collection.Length := BufferLength||1000
            _MakeRecord(Constructor, self, arr) => Constructor(arr)
        }

        __Item[index] {
            Get {
                i := index < 0 ? this.Count + index + 1 : index
                if this.__Collection.Has(i)
                    return this.__Collection[i]
                if this.__Collection.HasOwnProp('Default')
                    return this.__Collection.Default
                if i > this.__Collection.Length
                    throw IndexError('Index out of range: ' (index == i ? index : index ' | ' i), -1)
                throw UnsetItemError('The array does not have a value at index ' (index == i ? index : index ' | ' i), -1)
            }
            Set {
                i := index < 0 ? this.Count + index + 1 : index
                if i > this.__Collection.Length
                    this.__Collection.Length += this.BufferLength
                this.__Collection[i] := value
            }
        }
        /** @description Handles adding records to the collection */
        __Add(Record) {
            this[++this.Count] := this.__MakeRecord(Record)
        }
        /** @description Use this method to clear the collection array, for example, when parsing very large files. */
        Clear() {
            this.__Collection := [], this.Count := 0
        }
        /** @description Handles the production of records. When `Constructor` is set, this method is overridden */
        __MakeRecord(RecordArray) {
            ObjSetBase(Rec := {}, this.BaseObj)
            Rec.__Item := Map()
            for Field in RecordArray
                Rec.Set(Rec.Headers[A_Index], Field)
            return Rec
        }
        /** @description Sets the headers and creates the base object for the record objects. */
        SetHeaders(Headers) {
            ObjSetBase(this.BaseObj, ParseCsv.Record.Prototype)
            this.BaseObj.Headers := this.Headers := ParseCsv.Headers := Headers
        }

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
                val := this.__Collection[i]
                return 1
            }
            enum2(&index, &val) {
                if ++i > this.Count
                    return false
                index := i, val := this.__Collection[i]
                return 1
            }
        }
        Clone() => this.__Collection.Clone()
        Delete(index) => this.__Collection[index < 0 ? this.Count + index + 1 : index] := unset
        Get(index) {
            i := index < 0 ? this.Count + index + 1 : index
            if this.__Collection.Has(i)
                return this.__Collection[i]
            else if this.__Collection.HasOwnProp('Default')
                return this.__Collection.Default
            else if index > this.__Collection.length
                throw IndexError('The index ' index ' is out of range.', -1)
            else
                throw UnsetItemError('The array does not have a value at index ' index, -1)
        }
        Has(index) => index && this.__Collection.Has(index < 0 ? this.Count + index + 1 : index)
        InsertAt(index, val*) {
            this.Count += val.Length
            this.__Collection.InsertAt(index < 0 ? this.Count + index + 1 : index, val*)
        }
        Pop() => this.__Collection.RemoveAt(this.Count--)
        Push(val*) {
            this.__Collection.InsertAt(this.Count, val*)
            this.Count += val.Length
        }
        RemoveAt(index, length?) {
            this.Count -= length??1
            return this.__Collection.RemoveAt(index < 0 ? this.Count + index + 1 : index, length??unset)
        }
        Length => this.Count
        Capacity {
            Get => this.__Collection.Capacity
            Set => this.__Collection.Capacity := value
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
}
