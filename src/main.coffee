



#-----------------------------------------------------------------------------------------------------------
### NB in OpenOffice:   importClass          org.mozilla.javascript.Context ??? ###
### NB in LibreOffice:  importClass Packages.org.mozilla.javascript.Context ??? ###
importClass Packages.org.mozilla.javascript.Context
importClass Packages.org.mozilla.javascript.tools.shell.Global
#...........................................................................................................
GLOBAL        = new Global Context.enter()
#-----------------------------------------------------------------------------------------------------------
prefix = '/Applications/OpenOffice.app/Contents/share/Scripts/javascript/CoffeeLibreDemo/'
#...........................................................................................................
### Globals ###
eval GLOBAL.readFile prefix + 'require.js'
eval GLOBAL.readFile prefix + 'import-classes.js'
require.prefix = prefix
#...........................................................................................................
### Locals ###
#...........................................................................................................
TRM                       = require 'coffeelibre-trm'
CHR                       = require 'coffeenode-chr'
TEXT                      = require 'coffeenode-text'
TYPES                     = require 'coffeenode-types'
font_name_by_rsg          = require 'font-name-by-rsg'
#...........................................................................................................
CL                        = require 'coffeelibre'
#...........................................................................................................
# TRM.log TRM.log.bind
log                       = TRM.log.bind TRM
rpr                       = TRM.rpr.bind TRM
xray                      = TRM.xray.bind TRM


# log '' + java.lang.System.getenv( 'USER' )

#-----------------------------------------------------------------------------------------------------------
@get_styles = ->
  log ''
  log '--------------------------------------------------------------------------------------------'
  log 'styles'
  log ( new Date() ).toString()
  # ThisComponent.StyleFamilies
  doc         = CL.get_current_doc()
  sheet       = CL.get_current_sheet doc
  cell        = CL.get_cell sheet, 0, 0
  pv          = UnoRuntime.queryInterface XPropertySet, cell
  psi         = pv.getPropertySetInfo()
  properties  = psi.getProperties()
  TRM.dir 'pv', pv
  TRM.dir 'psi', psi
  TRM.dir 'properties[ 0 ]', properties[ 0 ]
  # for idx in [ 0 ... properties.length ]
  #   log '' + properties[ idx ].Name
  log '' + psi.getPropertyByName 'CellStyle'
  pv.setPropertyValue 'CellStyle',       'glyph'


#===========================================================================================================
#
#-----------------------------------------------------------------------------------------------------------
@main = ->
  # log rpr ( name for name of GLOBAL )
  # log GLOBAL.readFile
  # log xray GLOBAL.readFile
  # log TYPES.type_of GLOBAL.readFile
  log ''
  log '--------------------------------------------------------------------------------------------'
  log 'format tree'
  log ( new Date() ).toString()
  #.........................................................................................................
  doc   = CL.get_current_doc()
  CL.step doc, 'format tree', -> @format_tree doc
  #.........................................................................................................
  return null


#-----------------------------------------------------------------------------------------------------------
@font_name_from_chr_info = ( chr_info, fallback ) ->
  rsg = chr_info[ 'rsg' ]
  R   = font_name_by_rsg[ rsg ]
  #.........................................................................................................
  unless R?
    return fallback unless fallback is undefined
    throw new Error "unable to find a suitable font for #{chr_info[ 'fncr' ]} #{chr_info[ 'chr' ]}"
  #.........................................................................................................
  return R


#-----------------------------------------------------------------------------------------------------------
@format_tree = ( doc ) ->
  #.........................................................................................................
  doc                 = CL.get_current_doc()
  sheet               = CL.get_current_sheet doc
  selection           = CL.get_current_selection doc
  fallback_font_name  = 'Sun-ExtA'
  #.........................................................................................................
  format_options_by_cell_type =
    empty:
      'cell-style-name':  'empty'
    missing:
      'cell-style-name':  'missing'
    glyph:
      'cell-style-name':  'glyph'
    tree:
      'cell-style-name':  'tree'
    fncr:
      'cell-style-name':  'fncr'
    strokecode:
      'cell-style-name':  'strokecode'
  #.........................................................................................................
  [ xy0, xy1 ]      = CL.get_current_selection doc
  #.........................................................................................................
  [ [ x0, y0, ],
    [ x1, y1, ], ]  = [ xy0, xy1, ]
  x_fncr            = x1 + 2
  #.........................................................................................................
  # log xy0, xy1
  log "current selection: #{CL.range_ref_from_xy xy0, xy1}"
  #.........................................................................................................
  for y in [ y0 .. y1 ]
    fncr_cell = CL.get_cell sheet, x_fncr, y
    #.......................................................................................................
    for x in [ x0 .. x1 ]
      source_cell = CL.get_cell sheet, x, y
      source_text = CL.get_cell_text source_cell
      cell_ref    = CL.cell_ref_from_xy x, y
      #.....................................................................................................
      ### empty cells: ###
      if source_text.length is 0
        CL.format_cell source_cell, format_options_by_cell_type[ 'empty' ]
        continue
      #.....................................................................................................
      ### empty cells: ###
      if source_text is '?'
        CL.format_cell source_cell, format_options_by_cell_type[ 'missing' ]
        continue
      #.....................................................................................................
      ### non-empty cells: ###
      cid         = CHR.as_cid source_text, input: 'xncr'
      chr_info    = CHR.analyze cid
      fncr        = chr_info[ 'fncr' ]
      rsg         = chr_info[ 'rsg' ]
      fncr        = fncr.replace /^u-pua-/, 'jzr-fig-'
      rsg         = rsg.replace  /^u-pua$/, 'jzr-fig'
      is_cjk      = /^(u-cjk|jzr-fig|u-pua)/.test rsg
      #.....................................................................................................
      if rsg is 'jzr-fig'
       source_cell.setFormula CHR._unicode_chr_from_cid cid
      #.....................................................................................................
      ### glyph cells: ###
      if is_cjk
        fncr_cell.setFormula fncr
        CL.format_cell fncr_cell, format_options_by_cell_type[ 'fncr' ]
        format_options                = format_options_by_cell_type[ 'glyph' ]
        font_name                     = @font_name_from_chr_info chr_info, fallback_font_name
        format_options[ 'font-name' ] = font_name
        CL.format_cell source_cell, format_options
        continue
      #.....................................................................................................
      ### tree cells: ###
      CL.format_cell source_cell, format_options_by_cell_type[ 'tree' ]
  #.........................................................................................................
  return null


############################################################################################################
@main()
# @get_styles()



