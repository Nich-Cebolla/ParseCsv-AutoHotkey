


; The function creates a randomly generated CSV that can contain quoted fields. When it contains quoted fields, it includes
; escaped quotes, field delimiters, CR, LF, and item delimiters mixed into the quoted values to ensure the parser correctly handles
; quoted fields. You can disable these by setting the related "Random" option to false.

; To create a CSV without quoted fields, set `QuoteChar` to blank. This also inherently disables the "Random" options.

; If your parser needs to handle specific types of other characters that aren't words, set `OtherChars`
; to a string containing those characters, no spaces or separation characters, set
; `OtherCharsProbability` to a value between 0 and 1, and set `OtherCharsOnlyInQuotedStrings`
; if you need them to only be within quoted strings.
; To disable `OtherChars`, set it to blank.

class GenerateCSV {

    __New(options?) {
        options := this.Options := GenerateCsv.Options(options ?? unset)
        if !options.Overwrite && options.PathOut && FileExist(options.PathOut) {
            throw Error('File already exists at the output path. You can disable this warning by setting ``Options.Overwrite := true``.')
        }
        OutStr := ''
        VarSetStrCapacity(&OutStr, 131702)
        fd := options.FieldDelimiter
        rd := options.RecordDelimiter
        qc := options.QuoteChar
        le := options.LineEnding
        if oc := options.OtherChars {
            if RegExMatch(oc, '[' StrReplace(rd fd qc, ']', '\]') '\r\n]') {
                ; Options.OtherChars cannot contain the record delimiter, field delimiter, quote char, or lfcr
                throw ValueError('``Options.OtherChars`` contains one or more invalid characters.', , oc)
            }
            OtherChars := OtherCharsGenerator(oc, options.MinOtherCharsPerGroup, options.MaxOtherCharsPerGroup)
        }
        minWordsPerField := options.MinWordsPerField
        maxWordsPerField := options.MaxWordsPerField
        probabilityQuotedString := options.ProbabilityQuotedString
        randomEscapedQuotes := options.RandomEscapedQuotes
        randomLineEnding := options.RandomLineEnding
        randomRecordDelimiters := options.RandomRecordDelimiters
        randomFieldDelimiters := options.RandomFieldDelimiters
        otherCharsProbability := options.OtherCharsProbability
        otherCharsOnlyInQuotedStrings := options.OtherCharsOnlyInQuotedStrings
        showTooltip := options.ShowTooltip
        words := GetWords()
        wordCount := words.Length
        if !options.NoHeaders {
            if options.Headers {
                if options.Headers is Array {
                    this.Headers := options.Headers
                } else {
                    this.Headers := StrSplit(options.Headers, fd)
                }
                headers := this.Headers
            } else {
                headers := this.Headers := []
                Loop options.Columns {
                    headers.Push('Column ' A_Index)
                }
            }
            cols := this.Columns := headers.Length
            OutStr .= headers[1]
            i := 1
            loop headers.Length - 1 {
                OutStr .= fd headers[++i]
            }
            OutStr .= rd
        }
        rows := this.Rows := options.Rows
        records := this.records := []
        maxFieldLen := maxRecordLen := 0
        if qc {
            ; Records with cr and lf replaced with `r and `n for better readability
            displayRecords := this.displayRecords := []
            displayStr := ''
            VarSetStrCapacity(&displayStr, 131702)
            displayStr .= OutStr
            displayMaxFieldLen := displayMaxRecordLen := 0
            loop rows {
                fields := []
                displayFields := []
                records.Push(fields)
                displayRecords.Push(displayFields)
                record := displayRecord := ''
                loop cols - 1 {
                    _GetFieldQuoted()
                    record .= fd
                    displayRecord .= fd
                }
                _GetFieldQuoted()
                maxRecordLen := Max(maxRecordLen, StrLen(record rd))
                displayMaxRecordLen := Max(displayMaxRecordLen, StrLen(displayRecord rd))
                OutStr .= record rd
                displayStr .= displayRecord rd
                if !Mod(A_Index, 5000) && showTooltip {
                    _ShowTooltip(A_Index ' / ' rows)
                }
            }
            OutStr := SubStr(OutStr, 1, -StrLen(rd))
            displayStr := SubStr(displayStr, 1, -StrLen(rd))
        } else {
            if oc && !otherCharsOnlyInQuotedStrings {
                get := _GetFieldOtherChars
            } else {
                get := _GetField
            }
            loop rows {
                fields := []
                records.Push(fields)
                record := ''
                loop cols - 1 {
                    get()
                    record .= fd
                }
                get()
                maxRecordLen := Max(maxRecordLen, StrLen(record rd))
                OutStr .= record rd
                if !Mod(A_Index, 5000) && showTooltip {
                    _ShowTooltip(A_Index ' / ' rows)
                }
            }
            OutStr := SubStr(OutStr, 1, -StrLen(rd))
        }

        if options.PathOut {
            f := FileOpen(options.PathOut, 'w', options.encoding)
            if IsSet(displayStr) && options.OutputDisplayStr {
                f.Write(displayStr)
            } else {
                f.Write(OutStr)
            }
            f.Close()
        }
        this.MaxFieldLen := maxFieldLen
        this.MaxRecordLen := maxRecordLen
        this.Content := OutStr
        this.ContentLen := StrLen(OutStr)
        if IsSet(displayStr) {
            this.DisplayContent := displayStr
            this.DisplayMaxFieldLen := displayMaxFieldLen
            this.DisplayMaxRecordLen := displayMaxRecordLen
            this.DisplayContentLen := StrLen(displayStr)
        }
        if options.ShowTooltip {
            _ShowTooltip('Done')
        }

        return

        _GetField() {
            local s := ''
            VarSetStrCapacity(&s, 8192)
            loop Random(minWordsPerField, maxWordsPerField) {
                s .= words[Random(1, wordCount)] ' '
            }
            s := SubStr(s, 1, -1)
            fields.Push(s)
            maxFieldLen := Max(maxFieldLen, StrLen(s fd))
            record .= s
        }
        _GetFieldQuoted() {
            local s := ''
            VarSetStrCapacity(&s, 8192)
            VarSetStrCapacity(&ds, 8192)
            if Random() <= probabilityQuotedString {
                s .= qc
                loop Random(minWordsPerField, maxWordsPerField) {
                    if randomEscapedQuotes {
                        s .= _GetRandom(qc qc, 0.3, 0.5)
                    }
                    if randomFieldDelimiters {
                        s .= _GetRandom(fd, 0.1, 0.2)
                    }
                    if randomRecordDelimiters {
                        s .= _GetRandom(rd, 0.1, 0.2)
                    }
                    if randomLineEnding {
                        s .= _GetRandom(le, 0.1, 0.15)
                    }
                    if oc && Random() <= otherCharsProbability {
                        s .= OtherChars()
                    } else {
                        s .= words[Random(1, wordCount)]
                    }
                    s .= ' '
                }
                if StrLen(s) > 1 {
                    s := SubStr(s, 1, -1)
                }
                s .= qc
            } else if otherCharsOnlyInQuotedStrings {
                loop Random(minWordsPerField, maxWordsPerField) {
                    s .= words[Random(1, wordCount)] ' '
                }
                s := SubStr(s, 1, -1)
            } else {
                loop Random(minWordsPerField, maxWordsPerField) {
                    if oc && Random() <= otherCharsProbability {
                        s .= OtherChars()
                    } else {
                        s .= words[Random(1, wordCount)]
                    }
                    s .= ' '
                }
                s := SubStr(s, 1, -1)
            }
            ds := StrReplace(StrReplace(s, '`r', '``r'), '`n', '``n')
            fields.Push(s)
            displayFields.Push(ds)
            record .= s
            displayRecord .= ds
            maxFieldLen := Max(maxFieldLen, StrLen(s fd))
            displayMaxFieldLen := Max(displayMaxFieldLen, StrLen(ds fd))
        }

        _GetFieldOtherChars() {
            local s := ''
            VarSetStrCapacity(&s, 8192)
            loop Random(minWordsPerField, maxWordsPerField) {
                if Random() <= otherCharsProbability {
                    s .= OtherChars()
                } else {
                    s .= words[Random(1, wordCount)]
                }
                s .= ' '
            }
            s := SubStr(s, 1, -1)
            fields.Push(s)
            maxFieldLen := Max(maxFieldLen, StrLen(s fd))
            record .= s
        }

        _GetRandom(str, _min, _max) {
            local s := ''
            VarSetStrCapacity(&s, 64)
            loop {
                n := Random()
                if n > _min && n < _max {
                    s .= str
                } else {
                    return s
                }
            }
        }
        _ShowTooltip(str) {
            MouseO := CoordMode('Mouse', 'Screen')
            TTO := CoordMode('Tooltip', 'Screen')
            MouseGetPos(&mx, &my)
            Tooltip(str, mx + 10, my + 10)
            SetTimer(Tooltip, -2000)
            CoordMode('Mouse', MouseO)
            CoordMode('Tooltip', TTO)
        }
    }

    class Options {
        static __New() {
            this.DeleteProp('__New')
            proto := this.Prototype
            proto.Columns := 10
            proto.Encoding := 'cp1200'
            proto.FieldDelimiter := ','
            proto.Headers := ''
            proto.LineEnding := '`n'
            proto.MaxOtherCharsPerGroup := 3 ; The maximum number of other chars in a group
            proto.MaxWordsPerField := 10
            proto.MinOtherCharsPerGroup := 1 ; The minimum number of other chars in a group
            proto.MinWordsPerField := 0
            proto.NoHeaders := false
            proto.OtherChars := '1234567890+_)(*&^%$#@!~``;/\|><.?'
            proto.OtherCharsOnlyInQuotedStrings := false
            proto.OtherCharsProbability := 0.3 ; The probability a group of other chars is used instead of a word
            ; When Options.QuoteChar is set, two output strings are generated, one has the cr and lf
            ; characters replaced with `r and `n for better readability. If OutputDisplayStr = true,
            ; that string is output instead of the standard string.
            proto.OutputDisplayStr := false
            proto.Overwrite := false
            proto.PathOut := ''
            proto.ProbabilityQuotedString := 0.6
            proto.QuoteChar := '"'
            proto.RandomEscapedQuotes := true
            proto.RandomFieldDelimiters := true
            proto.RandomLineEnding := true
            proto.RandomRecordDelimiters := true
            proto.RecordDelimiter := '`n'
            proto.Rows := 100
            proto.ShowTooltip := false
        }

        __New(options?) {
            if IsSet(options) {
                if IsSet(GenerateCsvConfig) {
                    for prop in GenerateCsv.Options.Prototype.OwnProps() {
                        if HasProp(options, prop) {
                            this.%prop% := options.%prop%
                        } else if HasProp(GenerateCsvConfig, prop) {
                            this.%prop% := GenerateCsvConfig.%prop%
                        }
                    }
                } else {
                    for prop in GenerateCsv.Options.Prototype.OwnProps() {
                        if HasProp(options, prop) {
                            this.%prop% := options.%prop%
                        }
                    }
                }
            } else if IsSet(GenerateCsvConfig) {
                for prop in GenerateCsv.Options.Prototype.OwnProps() {
                    if HasProp(GenerateCsvConfig, prop) {
                        this.%prop% := GenerateCsvConfig.%prop%
                    }
                }
            }
            if this.HasOwnProp('__Class') {
                this.DeleteProp('__Class')
            }
        }
    }
}

class OtherCharsGenerator {
    __New(chars, MinOtherCharsPerGroup, MaxOtherCharsPerGroup) {
        this.list := StrSplit(chars)
        this.MinOtherCharsPerGroup := MinOtherCharsPerGroup
        this.MaxOtherCharsPerGroup := MaxOtherCharsPerGroup
    }
    Call() {
        list := this.list
        s := ''
        Loop Random(this.MinOtherCharsPerGroup, this.MaxOtherCharsPerGroup) {
            s .= list[Random(1, list.Length)]
        }
        return s
    }
}

GetWords() {
    return [
        'reuse','cancellation','pesto','redefine','methodist','bohemian','suspense','publish','supportive',
        'anonymous','earring','zuri','keira','coach','wheeze','midday','stimulus','pup','keep','cower','junior',
        'gathered','yell','ascension','cellar','constantly','helena','good','vile','gleefully','rustic',
        'trophy','invalid','biblical','pixel','liberating','roil','cartel','blended','cheating','convoluted','storm','surgeon',
        'sense','peeling','dotted','normalcy','adage','elianna','objective','outland','scare',
        'counsel','holster','molded','winter','loom','whom','farm','bowl','sinister','consultative','constitution','onlooker',
        'scotch','accomplish','poll','sustenance','dominance','aside','ceramic','capitalist','slug','trojan',
        'basically','stream','loose','mightily','brute','splay','fixate','tights','croatian','blackmail','credential','salted',
        'betrayal','foundation','ilk','supplant','conqueror','snowboard','chairperson','preferred','gravely',
        'afoot','brothel','fully','lupus','elongated','paper','fell','unmanned','body','octavia','class','amount','liquor',
        'swatch','norfolk','commence','glimmer','wallet','barber','bait','gala','worrying','behemoth','cadence','spoken',
        'speaker','repeatedly','borough','juxtapose','suspend','judge','pain','blame','sobering','erode','too','distract',
        'borrow','camouflage','russian','genesis','least','stirrup','kansas','crash','rippling','present','like','suitability',
        'coin','deputy','practicing','least','trinket','undulating','significantly','periodical','playful','dictum','flier',
        'itself','database','parapet','concurrent','interrupt','fiddle','technically','cardiac','athena','eight',
        'borrower','prosecution','faulty','polite','buddy','store','later','waver','presentation','empire','institutional',
        'post','pine','biodiversity','terrorist','blown','inflated','purify','square','harmonic','workhorse','animation',
        'exhaustive','fabric','aspire','ensemble','instantaneously','glitter','stalinist','superbly','pursuit',
        'titanic','partner','polyp','intercede','provocation','accuracy','hang','fencing','careful','photon','waning',
        'proportion','amirah','stingy','sandwich','tattoo','dart','stifle','reflector','intercontinental','resignation',
        'talking','populate','make','chatter','shimmer','sadistic','corkscrew','medallion','pizza','informational',
        'spoon','seattle','freak','motion','contrive','dole','dune','alloy','madeleine','heartache','unlucky','romanian',
        'cheesy','climax','thin','respect','music','cod','anthropological','rest','unanticipated','pipes','poisoned','teacup',
        'snack','monitoring','flooded','crass','black','inland','firepower','stabilization','tertiary','curtail',
        'indispensable','reconstruction','cradle','georgia','lasagna','spread','retool','criteria','responder','grad','variant',
        'survivor','cave','insurgent','libido','cow','illustrative','ricotta','spinal','competitor','arousal',
        'growl','great','conduit','kayla','mandatory','harlee','graft','repeating','murder','scheme','age',
        'lillie','predator','stage','benefit','contamination','driveway','madly','pedagogy','fundamentalism','bushy',
        'theologian','schizophrenia','altitude','adventurer','army','pit','mortgage','crest','mate','complementary','hardy',
        'pancreas','tuft','deploy','hollow','sodium','sneaky','hem','freedom','accent','haven','mutilated',
        'unraveled','makeup','thou','midtown','transducer','apricot','scramble','madness','irving','niece','badland','part',
        'panther','overwhelm','cherokee','christian','obsessively','uncontrolled','functionally','fun','vowel','freya','prone',
        'purse','informer','steering','ace','luminosity','shyly','solvent','arousal','beyond','consolidation','pluralism',
        'marsh','bar','asian','some','orlando','instructive','thanksgiving','opaque','intercollegiate','honor',
        'hope','oregon','integrated','summer','baths','elastic','elaborately','pathological','delight','scant','motivational',
        'riser','exec','commemorative','concern','classic','bacterial','turpentine','region','chuckle','lease',
        'maliyah','elected','hastily','kenyan','untitled','impossible','sting','junkie','mouse','distinction','discharge',
        'specialize','testament','thermos','earthly','rewarding','conditioning','hook','lupus','synopsis','acquisition',
        'sheathe','bead','overrun','before','unique','khakis','frond','laurel','hand','domestic','nonstick','swagger',
        'forecast','hole','creature','relentlessly','recreational','contest','coincide','extrinsic','bargaining',
        'poplar','trek','roomy','geographer','seriousness','navigational','itinerary','multiplication','underline','return',
        'directional','snag','idealized','relieved','hookup','vivid','overseas','detention','affect','penetrating',
        'computational','minefield','constant','mole','inverse','slug','checklist','corpus','karsyn','crusade',
        'trail','pelt','noteworthy','homecoming','increment','incur','snowshoe','separatist','bar','van',
        'rattle','illiteracy','defining','indulge','sensibly','opportunity','pistachio','alleviate','languish'
    ]
}
