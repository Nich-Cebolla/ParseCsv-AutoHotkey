

generateCsvOpt := {
    Columns: 10
  , Encoding: ''
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
    ; When Options.QuoteChar is set, two output strings are generated, one has the cr and lf
    ; characters replaced with `r and `n for better readability. If OutputDisplayStr = true,
    ; that string is output instead of the standard string.
  , OutputDisplayStr: false
  , Overwrite: false
  , PathOut: ''
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
    ; When Options.QuoteChar is set, two output strings are generated, one has the cr and lf
    ; characters replaced with `r and `n for better readability. If OutputDisplayStr = true,
    ; that string is output instead of the standard string.
    static OutputDisplayStr := false
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
    static ShowTooltip := false
}
