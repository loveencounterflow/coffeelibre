
############################################################################################################
# log                       = ( P ... ) -> java.lang.System.err.println p for p in P
# rpr                       = ( x ) -> return x # .toString() ( require 'util' ).inspect x, false, 22
#-----------------------------------------------------------------------------------------------------------
# njs_util                  = require 'util'
js_type_of                = ( x ) -> return Object::toString.call x
# type_features             = require 'COFFEENODE/Λ/registry/type-features'
#...........................................................................................................
### There appear to have been some changes in NodeJS concerning where to find `isXY` methods: ###
# isBuffer                  = Buffer.isBuffer ? njs_util.isBuffer


#===========================================================================================================
# TYPE ELUCIDATION
#-----------------------------------------------------------------------------------------------------------
@_coffeenode_type_by_js_type =
  '[object Array]':                     'list'
  '[object Boolean]':                   'boolean'
  '[object Function]':                  'function'
  '[object Null]':                      'null'
  '[object String]':                    'text'
  #.........................................................................................................
  '[object Generator]':                 'generator'
  #.........................................................................................................
  '[object Undefined]':                 'jsundefined'
  '[object Arguments]':                 'jsarguments'
  '[object Date]':                      'jsdate'
  '[object Error]':                     'jserror'
  '[object global]':                    'jsglobal'
  '[object RegExp]':                    'jsregex'
  '[object DOMWindow]':                 'jswindow'
  '[object CanvasRenderingContext2D]':  'jsctx'
  '[object ArrayBuffer]':               'jsarraybuffer'
  #.........................................................................................................
  '[object Object]': ( x ) ->
    # return 'jsbuffer' if isBuffer x
    return 'pod'
  #.........................................................................................................
  '[object Number]': ( x ) ->
    return 'jsnotanumber' if isNaN x
    return 'jsinfinity'   if x == Infinity or x == -Infinity
    return 'number'

#-----------------------------------------------------------------------------------------------------------
@type_of = ( x ) ->
  """Given any kind of value ``x``, return its type."""
  #.........................................................................................................
  # validate_argument_count_equals 1
  #.........................................................................................................
  return 'null'         if x is null
  return 'jsundefined'  if x is undefined
  #.........................................................................................................
  R = x[ '~isa' ]
  return R if R?
  #.........................................................................................................
  js_type = js_type_of x
  R       = @_coffeenode_type_by_js_type[ js_type ]
  return js_type.replace /^\[object (.+)\]$/, '$1' unless R?
  return if @isa_function R then R x else R

#-----------------------------------------------------------------------------------------------------------
@isa = ( x, probe ) ->
  """Given any value ``x`` and a non-empty text ``probe``, return whether ``TYPES/type_of x`` equals
  ``probe``."""
  # validate_name probe
  return ( @type_of x ) == probe


#===========================================================================================================
# TYPE TESTING
#-----------------------------------------------------------------------------------------------------------
# It is outright incredible, some would think frightening, how much manpower has gone into reliable
# JavaScript type checking. Here is the latest and greatest for a language that can claim to be second
# to none when it comes to things that should be easy but aren’t: the ‘Miller Device’ by Mark Miller of
# Google (http://www.caplet.com), popularized by James Crockford of Yahoo!.*
#
# As per https://groups.google.com/d/msg/nodejs/P_RzSyPkjkI/NvP28SXvf24J, now also called the 'Flanagan
# Device'
#
# http://ajaxian.com/archives/isarray-why-is-it-so-bloody-hard-to-get-right
# http://blog.360.yahoo.com/blog-TBPekxc1dLNy5DOloPfzVvFIVOWMB0li?p=916 # page gone
# http://zaa.ch/past/2009/1/31/the_miller_device_on_null_and_other_lowly_unvalues/ # moved to:
# http://zaa.ch/post/918977126/the-miller-device-on-null-and-other-lowly-unvalues
#...........................................................................................................
@isa_list          = ( x ) -> return ( js_type_of x ) == '[object Array]'
@isa_boolean       = ( x ) -> return ( js_type_of x ) == '[object Boolean]'
@isa_function      = ( x ) -> return ( js_type_of x ) == '[object Function]'
@isa_pod           = ( x ) -> return ( js_type_of x ) == '[object Object]' # and not isBuffer x
@isa_text          = ( x ) -> return ( js_type_of x ) == '[object String]'
@isa_number        = ( x ) -> return ( js_type_of x ) == '[object Number]' and isFinite x
@isa_null          = ( x ) -> return x is null
@isa_jsundefined   = ( x ) -> return x is undefined
@isa_infinity      = ( x ) -> return x == Infinity or x == -Infinity
#...........................................................................................................
@isa_jsarguments   = ( x ) -> return ( js_type_of x ) == '[object Arguments]'
@isa_jsnotanumber  = ( x ) -> return isNaN x
@isa_jsdate        = ( x ) -> return ( js_type_of x ) == '[object Date]'
@isa_jsglobal      = ( x ) -> return ( js_type_of x ) == '[object global]'
@isa_jsregex       = ( x ) -> return ( js_type_of x ) == '[object RegExp]'
@isa_jserror       = ( x ) -> return ( js_type_of x ) == '[object Error]'
@isa_jswindow      = ( x ) -> return ( js_type_of x ) == '[object DOMWindow]'
@isa_jsctx         = ( x ) -> return ( js_type_of x ) == '[object CanvasRenderingContext2D]'
@isa_jsarraybuffer = ( x ) -> return ( js_type_of x ) == '[object ArrayBuffer]'
#...........................................................................................................
# @isa_jsbuffer      = isBuffer

#-----------------------------------------------------------------------------------------------------------
# Replace some of our ``isa_*`` methods by the ≈6× faster methods provided by NodeJS ≥ 0.6.0, where
# available:
# @isa_list             = njs_util.isArray            if njs_util.isArray?
# @isa_jsregex          = njs_util.isRegExp           if njs_util.isRegExp?
# @isa_jsdate           = njs_util.isDate             if njs_util.isDate?
# @isa_boolean          = njs_util.isBoolean          if njs_util.isBoolean?
# @isa_jserror          = njs_util.isError            if njs_util.isError?
# @isa_function         = njs_util.isFunction         if njs_util.isFunction?
# @isa_primitive        = njs_util.isPrimitive        if njs_util.isPrimitive?
# @isa_text             = njs_util.isString           if njs_util.isString?
# @isa_jsundefined      = njs_util.isUndefined        if njs_util.isUndefined?
# @isa_null             = njs_util.isNull             if njs_util.isNull?
# @isa_nullorundefined  = njs_util.isNullOrUndefined  if njs_util.isNullOrUndefined?
# @isa_number           = njs_util.isNumber           if njs_util.isNumber?
# @isa_object           = njs_util.isObject           if njs_util.isObject?
# @isa_symbol           = njs_util.isSymbol           if njs_util.isSymbol?


#===========================================================================================================
# TYPE FEATURES
#-----------------------------------------------------------------------------------------------------------
# these await further elaboration

# @is_mutable          = ( x ) -> return type_features[ @type_of x ][ 'mutable'        ] is true
# @is_indexed          = ( x ) -> return type_features[ @type_of x ][ 'indexed'        ] is true
# @is_facetted         = ( x ) -> return type_features[ @type_of x ][ 'facetted'       ] is true
# @is_ordered          = ( x ) -> return type_features[ @type_of x ][ 'ordered'        ] is true
# @is_repetitive       = ( x ) -> return type_features[ @type_of x ][ 'repetitive'     ] is true
# @is_single_valued    = ( x ) -> return type_features[ @type_of x ][ 'single_valued'  ] is true
# @is_dense            = ( x ) -> return type_features[ @type_of x ][ 'dense'          ] is true
# @is_callable         = ( x ) -> return type_features[ @type_of x ][ 'callable'       ] is true
# @is_numeric          = ( x ) -> return type_features[ @type_of x ][ 'numeric'        ] is true
# @is_basic            = ( x ) -> return type_features[ @type_of x ][ 'basic'          ] is true
# @is_ecma             = ( x ) -> return type_features[ @type_of x ][ 'ecma'           ] is true
# @is_covered_by_json  = ( x ) -> return type_features[ @type_of x ][ 'json'           ] is true


#===========================================================================================================
# HELPERS
#-----------------------------------------------------------------------------------------------------------
# vaq = validate_argument_count_equals = ( count ) ->
#   a = arguments.callee.caller.arguments
#   unless a.length == count then throw new Error "expected #{count} arguments, got #{a.length}"

# #-----------------------------------------------------------------------------------------------------------
# validate_name = ( x ) ->
#   unless @isa_text x   then throw new Error "expected a text, got a #{@type_of x}"
#   unless x.length > 0   then throw new Error "expected a non-empty text, got an empty one"

#-----------------------------------------------------------------------------------------------------------
# This registry lists all types that can be meaningfully compared using JS's ``===`` / CS's ``==`` strict
# equality operator; conspicuously absent here are lists and PODs, for which the Δ method ``equals`` should
# be used instead:
@simple_equality_types =
  'number':       true
  'infinity':     true
  'text':         true
  'boolean':      true
  'null':         true
  'jsundefined':  true

#-----------------------------------------------------------------------------------------------------------
# This registry lists all types that can be meaningfully compared using `<` and `>`:
@simple_comparison_types =
  'number':       true
  'infinity':     true
  'text':         true
  'boolean':      true
  'null':         true
  # 'jsundefined':  true


#-----------------------------------------------------------------------------------------------------------
# 'ISA' VALIDATION
#...........................................................................................................
@validate_isa = ( x, types... ) ->
  throw new Error "expected one or more types, got none" if types.length == 0
  #.........................................................................................................
  probe_type  = @type_of x
  for type in types
    return null if type is probe_type
  #.........................................................................................................
  if types.length == 1
    message = "expected a #{types[ 0 ]}, got a #{probe_type}"
  else
    types   = types.join ', '
    message = "expected value to have one of these types: #{types}, got a #{probe_type}"
  #.........................................................................................................
  return null

#-----------------------------------------------------------------------------------------------------------
do =>
  for name of @
    match = name.match /^isa_(.+)/
    continue unless match?
    type  = match[ 1 ]
    #.......................................................................................................
    do ( name, type ) =>
      @[ "validate_#{name}" ] = ( x ) ->
        return null if @[ name ] x
        throw new Error "expected a #{type}, got a #{@type_of x}"

#-----------------------------------------------------------------------------------------------------------
# TAG VALIDATION
#...........................................................................................................
# do ->
#   for name of TYPES
#     continue unless name.match /^is_/
#     do ( name ) ->
#       $[ "validate_#{name}" ] = ( x ) ->
#         return null if TYPES[ name ] x
#         throw new Error "expected a x, got a #{TYPES.type_of x}"
