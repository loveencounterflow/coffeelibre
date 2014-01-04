


#-----------------------------------------------------------------------------------------------------------
@require = ( route ) ->
  #.........................................................................................................
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
    # throw error
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
@require.cache = {}

