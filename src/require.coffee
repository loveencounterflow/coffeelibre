
#-----------------------------------------------------------------------------------------------------------
join = ( prefix, suffix ) ->
  ### Poor Man's `path.join`—not recommended for serious use. ###
  return suffix if /^\/$/.test suffix
  return ( prefix.replace /\/$/, '' ).concat '/', suffix

#-----------------------------------------------------------------------------------------------------------
@require = ( route ) ->
  throw new Error "must set `require.prefix` to route before using `require()`" unless @require.prefix?
  #.........................................................................................................
  route         = join @require.prefix, route
  route         = route + '.js' unless /\js$/.test route
  return R if ( R = @require.cache[ route ] )?
  # source        = GLOBAL.readFile route
  source        = GLOBAL.readFile route
  source        = @require.head.concat source, @require.tail
  #.........................................................................................................
  try
    R = eval source
  catch error
    ### TAINT need better error handling ###
    for line, idx in source.split '\n'
      # log ( TRM.grey idx + 1 ), '  ', ( TRM.gold line )
      GLOBAL.print idx + 1, '  ', line
    GLOBAL.print()
    GLOBAL.print "an error occurred when trying to evaluate #{route}:"
    GLOBAL.print error[ 'message' ]
    GLOBAL.print error[ 'stack' ]
    GLOBAL.print()
  #.........................................................................................................
  @require.cache[ route ] = R
  return R
#...........................................................................................................
@require.head          = """
  (function() {
  var exports = {};
  var module  = { 'exports': exports };
  (function() {
  \n"""
#...........................................................................................................
@require.tail = """
  \n}).call( module.exports );
  return module.exports;
  })();
  """
#...........................................................................................................
@require.cache  = {}
@require.prefix = undefined
