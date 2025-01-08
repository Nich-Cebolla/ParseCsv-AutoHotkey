


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
    class Params {
        static PathOut := ''
        static QuoteChar := '"'
        static FieldDelimiter := ','
        static RecordDelimiter := '`n'
        static Columns := 10
        static Rows := 100
        static MinWordsPerField := 0
        static MaxWordsPerField := 10
        static Headers := ''
        static Encoding := ''
        static OtherChars := '1234567890+_)(*&^%$#@!~``;/\|><.?'
        static OtherCharsProbability := 0.3 ; The probability a group of other chars is used instead of a word
        static MinOtherCharsPerGroup := 1 ; The minimum number of other chars in a group
        static MaxOtherCharsPerGroup := 3 ; The maximum number of other chars in a group
        static OtherCharsOnlyInQuotedStrings := false
        static LineEndings := '`n'
        static RandomEscapedQuotes := true
        static RandomFieldDelimiters := true
        static RandomRecordDelimiters := true
        static RandomLineEndings := true

        __New(params) {
            for Name, Val in GenerateCSV.Params.OwnProps() {
                if IsSet(params) && params.HasOwnProp(Name)
                   this.DefineProp(Name, {Value: params.%Name%})
                else if IsSet(GenerateCSVConfig) && GenerateCSVConfig.HasOwnProp(Name)
                   this.DefineProp(Name, {Value: GenerateCSVConfig.%Name%})
                else
                   this.DefineProp(Name, {Value: Val})
            }
        }
    }

    static Call(params?, &OutStr?) {
        local r, c, OtherChars
        params := GenerateCsv.Params(params??{})
        OutStr := ''
        if params.OtherChars {
            if RegExMatch(params.OtherChars, '[' params.RecordDelimiter params.FieldDelimiter params.QuoteChar '\r\n]')
                throw ValueError('params.OtherChars cannot contain any of the following characters: ' params.RecordDelimiter params.FieldDelimiter params.QuoteChar '\r\n')
            OtherChars := OtherCharsGenerator(params.OtherChars, params.OtherCharsProbability, params.MinOtherCharsPerGroup, params.MaxOtherCharsPerGroup)
        }
        if !params.RecordDelimiter
            params.RecordDelimiter := '`n'
        if params.Headers {
            OutStr .= params.Headers
        } else {
            Loop params.Columns
                OutStr .= 'Column ' A_Index (A_Index == params.Columns ? '' : params.FieldDelimiter)
        }
        OutStr .= params.RecordDelimiter
        if params.QuoteChar {
            Loop params.Rows {
                r := A_Index
                Loop params.Columns {
                    c := A_Index
                    _WriteQuoted(A_Index == params.Columns)
                }
            }
        } else {
            if OtherChars && !params.OtherCharsOnlyInQuotedStrings {
                Loop params.Rows {
                    r := A_Index
                    Loop params.Columns {
                        c := A_Index
                        Loop Random(params.MinWordsPerField, params.MaxWordsPerField) {
                            if Random() <= params.OtherCharsProbability
                                OutStr .= (A_Index == 1 ? '' : ' ') OtherChars.Generate()
                            else
                                OutStr .= (A_Index == 1 ? '' : ' ') Words.__Item[Random(1, Words.__Item.Length)]
                        }
                        OutStr .= GetDelimiter(r, c)
                    }
                }
            } else {
                Loop params.Rows {
                    r := A_Index
                    Loop params.Columns {
                        c := A_Index
                        Loop Random(params.MinWordsPerField, params.MaxWordsPerField)
                            OutStr .= (c == 1 ? '' : ' ') Words.__Item[Random(1, Words.__Item.Length)]
                        OutStr .= GetDelimiter(r, c)
                    }
                }
            }
        }
        if params.PathOut {
            if FileExist(params.PathOut) {
                if MsgBox(params.PathOut ' already exists. Overwrite?', , 'YN') == 'No' {
                    A_Clipboard := OutStr
                    MsgBox('Content copied to clipboard.')
                    return
                }
            }
            f := FileOpen(params.PathOut, 'w', params.Encoding||unset)
            f.Write(OutStr)
            f.Close()
        }
        
        MouseO := CoordMode('Mouse', 'Screen')
        TTO := CoordMode('Tooltip', 'Screen')
        MouseGetPos(&mx, &my)
        Tooltip('Done', mx, my)
        SetTimer(Tooltip, -2000)
        CoordMode('Mouse', MouseO)
        CoordMode('Tooltip', TTO)
        return OutStr

        _WriteQuoted(IsEnd) {
            QtyWords := Random(params.MinWordsPerField, params.MaxWordsPerField)
            if Random() > 0.6 {
                OutStr .= params.QuoteChar
                Loop QtyWords {
                    if A_Index != 1
                        OutStr .= ' '
                    if params.RandomEscapedQuotes && _RandomEscapedQuotes := GetRandom(params.QuoteChar params.QuoteChar, 0.3, 0.5)
                        OutStr .= _RandomEscapedQuotes
                    if params.RandomFieldDelimiters && _RandomFieldDelimiters := GetRandom(params.FieldDelimiter, 0.1, 0.2)
                        OutStr .= _RandomFieldDelimiters
                    if params.RandomRecordDelimiters && _RandomRecordDelimiters := GetRandom(params.RecordDelimiter, 0.1, 0.2)
                        OutStr .= _RandomRecordDelimiters
                    if params.RandomLineEndings && _RandomLineEndings := GetRandom(params.LineEndings, 0.1, 0.15)
                        OutStr .= _RandomLineEndings
                    if OtherChars && Random() <= params.OtherCharsProbability
                        OutStr .= OtherChars.Generate()
                    else
                        OutStr .= Words.__Item[Random(1, Words.__Item.Length)]
                }
                OutStr .= params.QuoteChar
                OutStr .= GetDelimiter(r, c)
            } else {
                Loop Random(params.MinWordsPerField, params.MaxWordsPerField)
                    OutStr .= (A_Index == 1 ? '' : ' ') Words.__Item[Random(1, Words.__Item.Length)]
                OutStr .= GetDelimiter(r, c)
            }
        }
        GetRandom(thing, _min, _max) {
            result := ''
            Loop {
                n := Random()
                if !n
                    MsgBox('n was zero!')
                if n > _min && n < _max
                    result .= thing
                else
                    return result
            }
        }
        GetDelimiter(r, c) {
            if r == params.Rows {
                if c == params.Columns
                    return ''
                else
                    return params.FieldDelimiter
            } else {
                if c == params.Columns
                    return params.RecordDelimiter
                else
                    return params.FieldDelimiter
            }
        }
    }
}
; This pattern correctly handles CSV.
GetCsvPattern(Quote, FieldDelimiter, RecordDelimiter?) {
    if RecordDelimiter
        pattern := Format('JS)(?<=^|{2}|{3})(?:{1}(?<value>(?:[^{1}]*(?:{1}{1}){0,1})*){1}'
        '|(?<value>[^\r\n{1}{2}{4}]*?))(?={2}|{3}(*MARK:item)|$(*MARK:end))'
        , Quote, FieldDelimiter, RecordDelimiter, RegExReplace(RecordDelimiter, '(?:\\[rnR])+|[\[\]+*]', ''))
    else
        pattern := Format('JS)(?<=^|{2})(?:{1}(?<value>(?:[^{1}]*(?:{1}{1}){0,1})*){1}'
        '|(?<value>[^{1}{2}{3}{4}]*?))(?={2}|$(*MARK:item))', Quote, FieldDelimiter)
    try
        RegExMatch(' ', pattern)
    catch Error as err {
        if err.message == 'Compile error 25 at offset 6: lookbehind assertion is not fixed length'
            throw Error('The procedure received "Compile error 25 at offset 6: lookbehind assertion is not fixed length".'
            ' To fix this, change the RecordDelimiter and/or FieldDelimiter to a value that is a fixed length.', -1)
        else
            throw err
    }
    return pattern
}

class OtherCharsGenerator {
    __New(chars, probability, MinOtherCharsPerGroup, MaxOtherCharsPerGroup) {
        this.list := StrSplit(chars)
        this.probability := probability
        this.MinOtherCharsPerGroup := MinOtherCharsPerGroup
        this.MaxOtherCharsPerGroup := MaxOtherCharsPerGroup
    }
    Generate() {
        n := Random(this.MinOtherCharsPerGroup, this.MaxOtherCharsPerGroup)
        str := ''
        Loop n
            str .= this.list[Random(1, this.list.Length)]
        return str
    }
}

class Words {
    static __New() {
        Words.__Item :=  [
            'reuse','re-examine','cancellation','pesto','redefine','methodist','bohemian','suspense','publish','supportive',
            'anonymous','earring','zuri','keira','coach','wheeze','midday','stimulus','pup','keep','cower','junior','re-create',
            'gathered','yell','ascension','cellar','constantly','three-dimensional','helena','good','vile','gleefully','rustic',
            'trophy','invalid','biblical','pixel','liberating','roil','cartel','blended','cheating','convoluted','storm','surgeon',
            'sense','peeling','dotted','normalcy','adage','centuries-old','elianna','objective','outland','deep-sea','scare',
            'counsel','holster','molded','winter','loom','whom','farm','bowl','sinister','consultative','constitution','onlooker',
            'scotch','accomplish','poll','sustenance','dominance','aside','ceramic','capitalist','slug','medium-high','trojan',
            'basically','stream','loose','mightily','brute','splay','fixate','tights','croatian','blackmail','credential','salted',
            'deep-seated','betrayal','foundation','ilk','supplant','conqueror','snowboard','chairperson','preferred','gravely',
            'afoot','brothel','fully','lupus','elongated','paper','fell','unmanned','body','octavia','class','amount','liquor',
            'swatch','norfolk','commence','glimmer','wallet','barber','bait','gala','worrying','behemoth','cadence','spoken',
            'speaker','repeatedly','borough','juxtapose','suspend','judge','pain','blame','sobering','erode','too','distract',
            'borrow','camouflage','russian','genesis','least','stirrup','kansas','crash','rippling','present','like','suitability',
            'coin','deputy','practicing','least','trinket','undulating','significantly','periodical','playful','dictum','flier',
            'itself','database','parapet','concurrent','interrupt','fiddle','technically','cardiac','athena','eight','forty-three',
            'borrower','prosecution','faulty','polite','buddy','store','later','waver','presentation','empire','institutional',
            'post','pine','biodiversity','terrorist','blown','inflated','purify','square','harmonic','workhorse','animation',
            'exhaustive','fabric','ground-based','aspire','ensemble','instantaneously','glitter','stalinist','superbly','pursuit',
            'titanic','partner','polyp','intercede','provocation','accuracy','hang','fencing','careful','photon','waning',
            'proportion','amirah','stingy','sandwich','tattoo','dart','stifle','reflector','intercontinental','resignation',
            'talking','populate','make','chatter','shimmer','sadistic','corkscrew','medallion','pizza','informational',
            'spoon','seattle','freak','motion','contrive','dole','dune','alloy','madeleine','heartache','unlucky','romanian',
            'cheesy','climax','thin','respect','music','cod','anthropological','rest','unanticipated','pipes','poisoned','teacup',
            'snack','monitoring','flooded','crass','black','inland','firepower','stabilization','tertiary','curtail',
            'indispensable','reconstruction','cradle','georgia','lasagna','spread','retool','criteria','responder','grad','variant',
            'twenty-nine','survivor','cave','insurgent','libido','cow','illustrative','ricotta','spinal','competitor','arousal',
            'growl','great','conduit','kayla','mandatory','harlee','flood plain','graft','repeating','murder','scheme','age',
            'lillie','predator','stage','benefit','contamination','driveway','madly','pedagogy','fundamentalism','bushy',
            'theologian','schizophrenia','altitude','adventurer','army','pit','mortgage','crest','mate','complementary','hardy',
            'pancreas','tuft','deploy','hollow','sodium','sneaky','hem','freedom','accent','haven','round-trip','mutilated',
            'unraveled','makeup','thou','midtown','transducer','apricot','scramble','madness','irving','niece','badland','part',
            'panther','overwhelm','cherokee','christian','obsessively','uncontrolled','functionally','fun','vowel','freya','prone',
            'purse','informer','steering','ace','luminosity','shyly','solvent','arousal','beyond','consolidation','pluralism',
            'marsh','bar','laid-back','asian','some','orlando','instructive','thanksgiving','opaque','intercollegiate','honor',
            'hope','oregon','integrated','summer','baths','elastic','elaborately','pathological','delight','scant','motivational',
            'riser','fire station','exec','commemorative','concern','classic','bacterial','turpentine','region','chuckle','lease',
            'maliyah','elected','hastily','kenyan','untitled','impossible','sting','junkie','mouse','distinction','discharge',
            'specialize','testament','thermos','earthly','rewarding','conditioning','hook','lupus','synopsis','acquisition',
            'sheathe','bead','overrun','before','unique','khakis','frond','laurel','hand','domestic','nonstick','swagger',
            'forecast','hole','creature','relentlessly','recreational','contest','coincide','extrinsic','medium-sized','bargaining',
            'poplar','trek','roomy','geographer','seriousness','navigational','itinerary','multiplication','underline','return',
            'directional','snag','idealized','relieved','hookup','vivid','overseas','detention','affect','penetrating',
            'computational','minefield','constant','mole','inverse','slug','rhode island','checklist','corpus','karsyn','crusade',
            'trail','pelt','noteworthy','homecoming','increment','incur','snowshoe','separatist','bar','computer-based','van',
            'rattle','illiteracy','defining','indulge','sensibly','opportunity','pistachio','alleviate','languish'
        ]
    }
}