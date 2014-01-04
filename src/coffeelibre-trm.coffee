


############################################################################################################
njs_util                  = require prefix + 'nodejs-util'


#-----------------------------------------------------------------------------------------------------------
@dir = ( message, x ) ->
  log ''
  log '----------------------------------'
  log message
  log ''
  keys = ( key for key of x ).sort()
  for key in keys
    log '*', key #, TYPES.type_of x[ key ]
  return null

#-----------------------------------------------------------------------------------------------------------
@log = ( P... ) ->
  GLOBAL.print ( ( if TYPES.isa_text p then p else rpr p ) for p in P ).join ' '

#-----------------------------------------------------------------------------------------------------------
@rpr = ( x ) ->
  # return njs_util.inspect x, rpr.options
  try
    return njs_util.inspect x, rpr.options
  catch error
    # log x.toString()
    # log TYPES.type_of x[ 0 ]
    # log '' + x[ 0 ].type
    type = TYPES.type_of x
    return "unable to serialize object of type #{type}; message: #{rpr error[ 'message' ]}"
#...........................................................................................................
@rpr.options =
  showHidden:     no
  depth:          4
  colors:         yes
  customInspect:  no

#-----------------------------------------------------------------------------------------------------------
@xray = ( x ) ->
  return njs_util.inspect x, @xray.options
#...........................................................................................................
@xray.options =
  showHidden:     yes
  depth:          42
  colors:         yes
  customInspect:  no
