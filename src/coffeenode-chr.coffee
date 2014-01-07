



############################################################################################################
TYPES                     = require 'coffeenode-types'
# TRM                       = require 'coffeenode-trm'
# rpr                       = TRM.rpr.bind TRM
# log                       = TRM.log.bind TRM
#...........................................................................................................
character_sets_and_ranges = require 'character-sets-and-ranges'
@_names_and_ranges_by_csg = character_sets_and_ranges[ 'names-and-ranges-by-csg' ]
binary_interval_search    = require 'binary-interval-search'



#===========================================================================================================
# SPLIT TEXT INTO CHARACTERS
#-----------------------------------------------------------------------------------------------------------
@chrs_from_text = ( text, options ) ->
  return [] if text.length is 0
  #.........................................................................................................
  switch input_mode = options?[ 'input' ] ? 'plain'
    when 'plain'  then splitter = @_plain_splitter
    when 'ncr'    then splitter = @_ncr_splitter
    when 'xncr'   then splitter = @_xncr_splitter
    else throw new Error "unknown input mode: #{rpr input_mode}"
  #.........................................................................................................
  return ( text.split splitter ).filter ( element, idx ) -> return element.length isnt 0

#-----------------------------------------------------------------------------------------------------------
@_new_chunk = ( csg, rsg, chrs ) ->
  R =
    '~isa':     'CHR/chunk'
    'csg':      csg
    'rsg':      rsg
    # 'chrs':     chrs
    'text':     chrs.join ''
  #.........................................................................................................
  return R

#-----------------------------------------------------------------------------------------------------------
@chunks_from_text = ( text, options ) ->
  ### Given a `text` and `options` (of which `csg` is irrelevant here), return a list of `CHR/chunk`
  objects (as returned by `CHR._new_chunk`) that describes stretches of characters with codepoints in the
  same 'range' (Unicode block).
  ###
  R           = []
  return R if text.length is 0
  last_csg    = 'u'
  last_rsg    = null
  chrs        = []
  #.........................................................................................................
  switch output_mode = options?[ 'output' ] ? 'plain'
    when 'plain'
      transform_output = ( chr ) ->
        return chr
    when 'html'
      transform_output = ( chr ) ->
        return switch chr
          when '&' then '&amp;'
          when '<' then '&lt;'
          when '>' then '&gt;'
          else chr
    else
      throw new Error "unknown output mode: #{rpr output_mode}"
  #.........................................................................................................
  for chr in @chrs_from_text text, options
    description = @analyze chr, options
    { csg
      rsg }     = description
    chr         = description[ if csg is 'u' then 'chr' else 'ncr' ]
    if rsg isnt last_rsg
      R.push @_new_chunk last_csg, last_rsg, chrs if chrs.length > 0
      last_csg    = csg
      last_rsg    = rsg
      chrs        = []
    #.......................................................................................................
    chrs.push transform_output chr
  #.........................................................................................................
  R.push @_new_chunk last_csg, last_rsg, chrs if chrs.length > 0
  return R

#-----------------------------------------------------------------------------------------------------------
@html_from_text = ( text, options ) ->
  R = []
  #.........................................................................................................
  input_mode  = options?[ 'input' ] ? 'plain'
  chunks      = @chunks_from_text text, input: input_mode, output: 'html'
  for chunk in chunks
    R.push """<span class="#{chunk[ 'rsg' ]}">#{chunk[ 'text' ]}</span>"""
  #.........................................................................................................
  return R.join ''

#===========================================================================================================
# CONVERTING TO CID
#-----------------------------------------------------------------------------------------------------------
@cid_from_chr = ( chr, options ) ->
  input_mode = options?[ 'input' ] ? 'plain'
  return ( @_chr_csg_cid_from_chr chr, input_mode )[ 2 ]

#-----------------------------------------------------------------------------------------------------------
@csg_cid_from_chr = ( chr, options ) ->
  input_mode = options?[ 'input' ] ? 'plain'
  return ( @_chr_csg_cid_from_chr chr, input_mode )[ 1 .. ]

#-----------------------------------------------------------------------------------------------------------
@_chr_csg_cid_from_chr = ( chr, input_mode ) ->
  ### Given a text with one or more characters, return the first character, its CSG, and its CID (as a
  non-negative integer). Additionally, an input mode may be given as either `plain`, `ncr`, or `xncr`.
  ###
  #.........................................................................................................
  throw new Error "unable to obtain CID from empty string" if chr.length is 0
  #.........................................................................................................
  input_mode ?= 'plain'
  switch input_mode
    when 'plain'  then matcher = @_first_chr_matcher_plain
    when 'ncr'    then matcher = @_first_chr_matcher_ncr
    when 'xncr'   then matcher = @_first_chr_matcher_xncr
    else throw new Error "unknown input mode: #{rpr input_mode}"
  #.........................................................................................................
  match     = chr.match matcher
  throw new Error "illegal character sequence in #{rpr chr}" unless match?
  first_chr = match[ 0 ]
  #.........................................................................................................
  switch first_chr.length
    #.......................................................................................................
    when 1
      return [ first_chr, 'u', first_chr.charCodeAt 0 ]
    #.......................................................................................................
    when 2
      ### thx to http://perldoc.perl.org/Encode/Unicode.html ###
      hi  = first_chr.charCodeAt 0
      lo  = first_chr.charCodeAt 1
      cid = ( hi - 0xD800 ) * 0x400 + ( lo - 0xDC00 ) + 0x10000
      return [ first_chr, 'u', cid ]
    #.......................................................................................................
    else
      [ chr
        csg
        cid_hex
        cid_dec ] = match
      cid = if cid_hex? then parseInt cid_hex, 16 else parseInt cid_dec, 10
      csg = 'u' if csg.length is 0
      return [ first_chr, csg, cid ]


# #-----------------------------------------------------------------------------------------------------------
# @cid_from_ncr = ( ) ->

# #-----------------------------------------------------------------------------------------------------------
# @cid_from_xncr = ( ) ->

# #-----------------------------------------------------------------------------------------------------------
# @cid_from_fncr = ( ) ->


#===========================================================================================================
# CONVERTING FROM CID &c
#-----------------------------------------------------------------------------------------------------------
@as_csg         = ( cid_hint, O ) -> return ( @_csg_cid_from_hint cid_hint, O )[ 0 ]
@as_cid         = ( cid_hint, O ) -> return ( @_csg_cid_from_hint cid_hint, O )[ 1 ]
#...........................................................................................................
@as_chr         = ( cid_hint, O ) -> return @_as_chr.apply        @, @_csg_cid_from_hint cid_hint, O
@as_fncr        = ( cid_hint, O ) -> return @_as_fncr.apply       @, @_csg_cid_from_hint cid_hint, O
@as_sfncr       = ( cid_hint, O ) -> return @_as_sfncr.apply      @, @_csg_cid_from_hint cid_hint, O
@as_xncr        = ( cid_hint, O ) -> return @_as_xncr.apply       @, @_csg_cid_from_hint cid_hint, O
@as_ncr         = ( cid_hint, O ) -> return @_as_xncr.apply       @, @_csg_cid_from_hint cid_hint, O
@as_rsg         = ( cid_hint, O ) -> return @_as_rsg.apply        @, @_csg_cid_from_hint cid_hint, O
@as_range_name  = ( cid_hint, O ) -> return @_as_range_name.apply @, @_csg_cid_from_hint cid_hint, O
#...........................................................................................................
@analyze        = ( cid_hint, O ) -> return @_analyze.apply       @, @_csg_cid_from_hint cid_hint, O

#-----------------------------------------------------------------------------------------------------------
@_analyze = ( csg, cid ) ->
  if csg is 'u'
    chr         = @_unicode_chr_from_cid cid
    ncr = xncr  = @_as_xncr csg, cid
  else
    chr         = @_as_xncr csg, cid
    xncr        = @_as_xncr csg, cid
    ncr         = @_as_xncr 'u', cid
  #.........................................................................................................
  R =
    '~isa':     'CHR/info'
    'chr':      chr
    'csg':      csg
    'cid':      cid
    'fncr':     @_as_fncr  csg, cid
    'sfncr':    @_as_sfncr csg, cid
    'ncr':      ncr
    'xncr':     xncr
    'rsg':      @_as_rsg   csg, cid
  #.........................................................................................................
  return R

#-----------------------------------------------------------------------------------------------------------
@_as_chr = ( csg, cid ) ->
  return @_unicode_chr_from_cid cid if csg is 'u'
  retrun ( @_analyze csg, cid )[ 'chr' ]

#-----------------------------------------------------------------------------------------------------------
@_unicode_chr_from_cid = ( cid ) ->
  return String.fromCharCode cid if cid <= 0xffff
  ### thx to http://perldoc.perl.org/Encode/Unicode.html ###
  hi = ( Math.floor ( cid - 0x10000 ) / 0x400 ) + 0xD800
  lo =              ( cid - 0x10000 ) % 0x400   + 0xDC00
  return ( String.fromCharCode hi ) + ( String.fromCharCode lo )

#-----------------------------------------------------------------------------------------------------------
@_as_fncr = ( csg, cid ) ->
  rsg = ( @_as_rsg csg, cid ) ? csg
  return "#{rsg}-#{cid.toString 16}"

#-----------------------------------------------------------------------------------------------------------
@_as_sfncr = ( csg, cid ) ->
  return "#{csg}-#{cid.toString 16}"

#-----------------------------------------------------------------------------------------------------------
@_as_xncr = ( csg, cid ) ->
  csg = '' if csg is 'u' or not csg?
  return "&#{csg}#x#{cid.toString 16};"

#-----------------------------------------------------------------------------------------------------------
@_as_rsg = ( csg, cid ) ->
  return binary_interval_search @_names_and_ranges_by_csg[ csg ], 'first-cid', 'last-cid', 'rsg', cid

#-----------------------------------------------------------------------------------------------------------
@_as_range_name = ( csg, cid ) ->
  return binary_interval_search @_names_and_ranges_by_csg[ csg ], 'first-cid', 'last-cid', 'range-name', cid


#===========================================================================================================
# ANALYZE ARGUMENTS
#-----------------------------------------------------------------------------------------------------------
@_csg_cid_from_hint = ( cid_hint, options ) ->
  ### This helper is used to derive the correct CSG and CID from arguments as accepted by the `as_*` family
  of methods, such as `CHR.as_fncr`, `CHR.as_rsg` and so on; its output may be directly applied to the
  respective namesake private method (`CHR._as_fncr`, `CHR._as_rsg` and so on). The method arguments should
  obey the following rules:

  * Methods may be called with one or two arguments; the first is known as the 'CID hint', the second as
    'options'.

  * The CID hint may be a number or a text; if it is a number, it is understood as a CID; if it
    is a text, its interpretation is subject to the `options[ 'input' ]` setting.

  * Options must be a POD with the optional members `input` and `csg`.

  * `options[ 'input' ]` is *only* observed if the CID hint is a text; it governs which kinds of character
    references are recognized in the text. `input` may be one of `plain`, `ncr`, or `xncr`; it defaults to
    `plain` (no character references will be recognized).

  * `options[ 'csg' ]` sets the character set sigil. If `csg` is set in the options, then it will override
    whatever the outcome of `CHR.csg_cid_from_chr` w.r.t. CSG is—in other words, if you call
    `CHR.as_sfncr '&jzr#xe100', input: 'xncr', csg: 'u'`, you will get `u-e100`, with the numerically
    equivalent codepoint from the `u` (Unicode) character set.

  * Before CSG and CID are returned, they will be validated for plausibility.

  ###
  #.........................................................................................................
  switch type = TYPES.type_of options
    when 'null', 'jsundefined'
      csg_of_options  = null
      input_mode      = null
    when 'pod'
      csg_of_options  = options[ 'csg' ]
      input_mode      = options[ 'input' ]
    else
      throw new Error "expected a POD as second argument, got a #{type}"
  #.........................................................................................................
  switch type = TYPES.type_of cid_hint
    when 'number'
      csg_of_cid_hint = null
      cid             = cid_hint
    when 'text'
      [ csg_of_cid_hint
        cid             ] = @csg_cid_from_chr cid_hint, input: input_mode
    else
      throw new Error "expected a text or a number as first argument, got a #{type}"
  #.........................................................................................................
  if csg_of_options?
    csg = csg_of_options
  else if csg_of_cid_hint?
    csg = csg_of_cid_hint
  else
    csg = 'u'
  #.........................................................................................................
  @validate_is_csg csg
  @validate_is_cid cid
  return [ csg, cid, ]


#===========================================================================================================
# PATTERNS
#-----------------------------------------------------------------------------------------------------------
# G: grouped
# O: optional
name                      = ( /// (?:     [a-z][a-z0-9]*     ) /// ).source
# nameG                     = ( /// (   (?: [a-z][a-z0-9]* ) | ) /// ).source
nameO                     = ( /// (?: (?: [a-z][a-z0-9]* ) | ) /// ).source
nameOG                    = ( /// (   (?: [a-z][a-z0-9]* ) | ) /// ).source
hex                       = ( /// (?: x   [a-fA-F0-9]+       ) /// ).source
hexG                      = ( /// (?: x  ([a-fA-F0-9]+)      ) /// ).source
dec                       = ( /// (?:     [      0-9]+       ) /// ).source
decG                      = ( /// (?:    ([      0-9]+)      ) /// ).source
#...........................................................................................................
@_csg_matcher             = /// ^ #{name} $ ///
@_ncr_matcher             = /// (?: &           \# (?: #{hex}  | #{dec}  ) ; ) ///
@_xncr_matcher            = /// (?: & #{nameO}  \# (?: #{hex}  | #{dec}  ) ; ) ///
@_ncr_csg_cid_matcher     = /// (?: & ()        \# (?: #{hexG} | #{decG} ) ; ) ///
@_xncr_csg_cid_matcher    = /// (?: & #{nameOG} \# (?: #{hexG} | #{decG} ) ; ) ///
#...........................................................................................................
### Matchers for surrogate sequences and non-surrogate, 'ordinary' characters: ###
@_surrogate_matcher       = /// (?: [  \ud800-\udbff ] [ \udc00-\udfff ] ) ///
@_nonsurrogate_matcher    = ///     [^ \ud800-\udbff     \udc00-\udfff ]   ///
#...........................................................................................................
### Matchers for the first character of a string, in three modes (`plain`, `ncr`, `xncr`): ###
@_first_chr_matcher_plain = /// ^ (?: #{@_surrogate_matcher.source}     |
                                      #{@_nonsurrogate_matcher.source}    ) ///
@_first_chr_matcher_ncr   = /// ^ (?: #{@_surrogate_matcher.source}     |
                                      #{@_ncr_csg_cid_matcher.source}   |
                                      #{@_nonsurrogate_matcher.source}    ) ///
@_first_chr_matcher_xncr  = /// ^ (?: #{@_surrogate_matcher.source}     |
                                      #{@_xncr_csg_cid_matcher.source}  |
                                      #{@_nonsurrogate_matcher.source}    ) ///
#...........................................................................................................
@_plain_splitter          = /// ( #{@_surrogate_matcher.source}     |
                                  #{@_nonsurrogate_matcher.source}    ) ///
@_ncr_splitter            = /// ( #{@_ncr_matcher.source}           |
                                  #{@_surrogate_matcher.source}     |
                                  #{@_nonsurrogate_matcher.source}    ) ///
@_xncr_splitter           = /// ( #{@_xncr_matcher.source}          |
                                  #{@_surrogate_matcher.source}     |
                                  #{@_nonsurrogate_matcher.source}    ) ///


#===========================================================================================================
# VALIDATION
#-----------------------------------------------------------------------------------------------------------
@validate_is_csg = ( x ) ->
  TYPES.validate_isa_text x
  throw new Error "not a valid CSG: #{rpr x}" unless ( x.match @_csg_matcher )?
  throw new Error "unknown CSG: #{rpr x}"     unless @_names_and_ranges_by_csg[ x ]?
  return null

#-----------------------------------------------------------------------------------------------------------
@validate_is_cid = ( x ) ->
  TYPES.validate_isa_number x
  # if x < 0 or x > 0x10ffff or ( parseInt x ) != x
  if x < 0 or x > 0xffffffff or ( parseInt x ) != x
    throw new Error "expected an integer between 0x0 and 0x10ffff, got 0x#{x.toString 16}"
  return null






# console.log name for name of @
# console.log String.fromCharCode 0x61
# console.log String.fromCharCode 0x24563



