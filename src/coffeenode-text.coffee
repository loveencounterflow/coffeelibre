



############################################################################################################
TYPES                     = require 'coffeenode-types'
TRM                       = require 'coffeelibre-trm'
rpr                       = TRM.rpr.bind TRM
log                       = TRM.log.bind TRM

#-----------------------------------------------------------------------------------------------------------
# @_first_chr_matcher      = /// ^ ([\ud800-\udbff].)|. ///


#-----------------------------------------------------------------------------------------------------------
@repeat = ( me, count ) ->
  return ( new Array count + 1 ).join me

#-----------------------------------------------------------------------------------------------------------
@push = ( me, you ) ->
  return me + you

#-----------------------------------------------------------------------------------------------------------
@copy = ( me ) ->
  return me

#-----------------------------------------------------------------------------------------------------------
@starts_with = ( me, probe, idx = 0 ) ->
  # thx to Mark Byers via http://stackoverflow.com/questions/646628/javascript-startswith/4579228#4579228
  return ( me.lastIndexOf probe, idx ) is idx

#-----------------------------------------------------------------------------------------------------------
@ends_with = ( me, search_text ) ->
  # from prototype 1.6.0_rc0
  delta = me.length - search_text.length
  return delta >= 0 and ( me.lastIndexOf search_text ) == delta

#-----------------------------------------------------------------------------------------------------------
@drop_prefix = ( me, probe ) ->
  return me.substr probe.length if @starts_with me, probe
  return me

#-----------------------------------------------------------------------------------------------------------
@drop_suffix = ( me, probe ) ->
  return me.substr 0, me.length - probe.length if @ends_with me, probe
  return me

# #-----------------------------------------------------------------------------------------------------------
# @count = ( me, probe ) ->
#   #.........................................................................................................
#   switch type_of_probe = TYPES.type_of probe
#     #.......................................................................................................
#     when 'text'
#       escape  = require 'COFFEENODE/REGEX/escape'
#       probe   = new RegExp ( escape probe ), 'g'
#     #.......................................................................................................
#     when 'jsregex'
#       probe = new RegExp probe.source, 'g' unless probe.global
#     #.......................................................................................................
#     when 'REGEX/regex'
#       probe = probe[ '%self' ]
#     #.......................................................................................................
#     else
#       throw new Error "encountered illegal probe type: #{rpr type_of_probe}"
#   #.........................................................................................................
#   R = 0
#   me.replace probe, -> R += 1
#   #.........................................................................................................
#   return R

#-----------------------------------------------------------------------------------------------------------
@contains = ( me, probe ) ->
  switch type_of_probe = TYPES.type_of probe
    #.......................................................................................................
    when 'text'
      return true if probe.length == 0
      return ( me.indexOf probe ) >= 0
    #.......................................................................................................
    when 'jsregex'
      return ( me.match probe )?
    #.......................................................................................................
    else
      throw new Error "unknown type for TEXT/contains: #{type_of_probe}"

#-----------------------------------------------------------------------------------------------------------
@contains_only_digits = ( me ) ->
  return false if @is_empty me
  return ( me.match /^[0-9]+$/ )?

#-----------------------------------------------------------------------------------------------------------
@is_empty = ( me ) ->
  return me.length == 0

#-----------------------------------------------------------------------------------------------------------
### TAINT.TODO will only split once when given a text probe; unexpected results when starts
  or ends with probe; will fail with 32bit codepoints
###
@split = ( me, probe ) ->
  return @words_of me unless probe?
  # probe = if isa_regex probe then probe[ '%self' ] else probe
  R     = me.split probe
  return R.filter ( element ) -> return element isnt undefined

#-----------------------------------------------------------------------------------------------------------
### TAINT.TODO "might not recognize all Unicode whitespace codepoints"
###
@words_of = ( me ) ->
  return ( me.replace /^\s*(.*?)\s*$/g, '$1' ).split /\s+/g

#-----------------------------------------------------------------------------------------------------------
@as_text = ( me ) ->
  return me

#-----------------------------------------------------------------------------------------------------------
@trim = ( me ) ->
  return @_trim_whitespace me, true, true

#-----------------------------------------------------------------------------------------------------------
@trim_left = ( me ) ->
  return @_trim_whitespace me, true, false

#-----------------------------------------------------------------------------------------------------------
@trim_right = ( me ) ->
  return @_trim_whitespace me, false, true

#-----------------------------------------------------------------------------------------------------------
@_trim_whitespace = ( me, trim_left, trim_right ) ->
  ###Faster whitespace trimming; adapted from
  http://blog.stevenlevithan.com/archives/faster-trim-javascript.###
  if trim_left
    me          = me.replace /^\s\s*/, ''
  if trim_right
    whitespace  = /\s/
    i           = me.length
    while whitespace.test me.charAt --i
      null
    me = me.slice 0, i + 1
  return me

# #-----------------------------------------------------------------------------------------------------------
# @cut = TAINT.UNICODE ###will split codepoints beyond u/ffff###,
#   ( me, Q ) ->
#     validate_isa_text me
#     #.........................................................................................................
#     [ Q, info, ] = analyze_named_arguments Q,
#       'start':        0
#       'rests':        null
#     #.........................................................................................................
#     start       = Q[ 'start' ]
#     rests       = Q[ 'rests' ]
#     #.........................................................................................................
#     validate_isa_nonnegative_integer start
#     #.........................................................................................................
#     unless start < me.length
#       throw new Error "index error: text has last character at index #{rpr me.length - 1}, got index #{rpr start}"
#     #.........................................................................................................
#     R = me.substring start
#     #.........................................................................................................
#     if rests?
#       validate_isa_list rests
#       unless rests.length == 0 then throw new Error ###expected an empty list,
#         got one with #{rpr rests.length} elements.###
#       rests.push me.substring 0, start
#       rests.push ''
#     #.........................................................................................................
#     return R

#-----------------------------------------------------------------------------------------------------------
@flush_left = ( x, width = 25, filler = ' ' ) ->
  return @flush x, width, 'left', filler

#-----------------------------------------------------------------------------------------------------------
@flush_right = ( x, width = 25, filler = ' ' ) ->
  return @flush x, width, 'right', filler

#-----------------------------------------------------------------------------------------------------------
@flush = ( x, width, align, filler = ' ' ) ->
  ###Given a value, a non-negative integer ``width``, and an optional, non-empty text ``filler`` (which
  defaults to a single space), return a string that starts with the text (or the text of the representation
  of the value), and is padded with as many fillers as needed to make the string ``width`` characters long.
  If ``width`` is zero or smaller than the length of the text, the text is simply returned as-is. No
  clipping of text is ever done.###
  #.........................................................................................................
  unless align == 'left' or align == 'right'
    throw new Error "expected ``left`` or ``right`` for ``align``, got #{rpr align}"
  #.........................................................................................................
  x                 = rpr x unless TYPES.isa_text x
  filler_length     = filler.length
  text_length       = x.length
  #.........................................................................................................
  return x if text_length >= width
  padding = @repeat filler, width - text_length
  return if align == 'left' then x + padding else padding + x

#-----------------------------------------------------------------------------------------------------------
@lower_case = ( me ) ->
  return me.toLowerCase()

#-----------------------------------------------------------------------------------------------------------
@upper_case = ( me ) ->
  return me.toUpperCase()

#-----------------------------------------------------------------------------------------------------------
### TAINT.UNICODE will incorrectly count codepoints above u/ffff as two chrs
###
@length_of = ( me ) ->
  return me.length

#-----------------------------------------------------------------------------------------------------------
# @replace = ( me, probe, replacement ) ->
#   if TYPES.isa_text probe
#     probe = new RegExp ( ( require 'COFFEENODE/REGEX/escape' ) probe ), 'g'
#   return me.replace probe, replacement

#-----------------------------------------------------------------------------------------------------------
@lines_of = ( me, handler ) ->
  ###Given a text and an optional handler, either return a list of lines (without line endings) if handler
  is not given, or call handler as `handler error, line` for each line in the text, and one more call where
  `line` is `null` after the end of the text has been encountered. In either case, all lines will be
  stripped of trailine whitespace, including newline characters; recognized newlines are the usual suspects
  for Unix, Windows, and MacOS systems (namely, `\\n`, `\\r\\n`, and `\\r`).

  **Implementation Note**: Implementation of asynchronous version postponed.###
  #.........................................................................................................
  return me.split @_line_splitter unless handler?
  throw new Error "asynchronous TEXT.lines_of not yet supported"

#-----------------------------------------------------------------------------------------------------------
@_line_splitter = /// \r\n | [\n\v\f\r\x85\u2028\u2029] ///g


#-----------------------------------------------------------------------------------------------------------
# @chrs_of = ( me, handler = null ) ->
#   ###
#   **Implementation Note**: This function has been somewhat hastily put together to make calls with a handler
#   argument process in an asynchronous fashion and to avoid the building of intermediate lists. It should be
#   rewritten to become more conherent.###
#   #.........................................................................................................
#   unless handler?
#     chunks_of     = require 'COFFEENODE/UNICODE/chunks-of'
#     is_smp_chunk  = yes
#     R             = []
#     #.......................................................................................................
#     for chunk in chunks_of me
#       is_smp_chunk = not is_smp_chunk
#       continue if chunk.length == 0
#       ( R.push chunk; continue ) if is_smp_chunk
#       R.push.apply R, chunk.split '' # LIST/extend
#     #.......................................................................................................
#     return R
#   #.........................................................................................................
#   # RegEx to match CESU-8 surrogate and non-surrogate characters. This knowledge really belongs into the
#   # UNICODE library, not here; it is also a partial code duplication of both the synchronous code, above,
#   # and the code used in `first_chr_of`, below.
#   REGEX       = require 'COFFEENODE/REGEX/implementation'
#   chr_matcher = REGEX.new_regex /[\ud800-\udbff].|./
#   REGEX.find chr_matcher, me, ( error, chr ) ->
#     return handler error if error?
#     handler null, chr

#-----------------------------------------------------------------------------------------------------------
# @first_chr_of = ( me ) ->
#   ###Like `( TEXT/chrs_of x )[ 0 ]`, but more efficient, since no intermediate list with all the characters
#   in the text is built.###
#   R = me.match @_first_chr_matcher
#   return R[ 1 ]

#-----------------------------------------------------------------------------------------------------------
@reverse = ( me ) ->
  return ( @chrs_of me ).reverse().join ''

#-----------------------------------------------------------------------------------------------------------
@add = ( me_and_you... ) ->
  return me_and_you.join ''

#-----------------------------------------------------------------------------------------------------------
### TAINT.UNICODE "will split codepoints beyond u/ffff"
###
@partition = ( me, partitioner ) ->
  #.........................................................................................................
  n           = partitioner
  m           = me.length
  idx         = - n
  max_idx     = m - n
  remainder   = m % n
  R           = []
  #.........................................................................................................
  unless remainder == 0
    throw new Error "expected length of text to be a multiple of #{n}, got text with #{m} characters"
  #.........................................................................................................
  for idx in [ 0 .. max_idx ] by n
    R.push @slice me, idx, idx + n
  #.........................................................................................................
  return R

#-----------------------------------------------------------------------------------------------------------
@slice = ( me, start, stop ) ->
  # short form ``?=`` buggy; see https://github.com/jashkenas/coffee-script/issues/1627
  # but ok here since these are variables from the method signature
  start    ?= 0
  stop     ?= me.length
  return '' if start >= stop
  return me.slice start, stop


#===========================================================================================================
# RANDOMIZATION
#-----------------------------------------------------------------------------------------------------------
@shuffle = ( me ) ->
  return ( LIST.shuffle @chrs_of me ).join ''


#===========================================================================================================
# TRANSFORMATIONS
#-----------------------------------------------------------------------------------------------------------
# @as_set = ( me ) ->
#   return ( require 'COFFEENODE/SET/implementation' ).new me


#===========================================================================================================
# VALIDATION
#-----------------------------------------------------------------------------------------------------------
@validate_is_nonempty_text = ( me ) ->
  TYPES.validate_isa_text me
  unless me.length > 0 then throw new Error "expected a non-empty text, got an empty one"

#-----------------------------------------------------------------------------------------------------------
@validate_is_word = ( me ) ->
  @validate_is_nonempty_text me
  unless ( me.match /^\S+$/ )? then throw new Error "expected a non-empty text without whitespace"




