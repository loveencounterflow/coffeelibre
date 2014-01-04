



#-----------------------------------------------------------------------------------------------------------
### NB in OpenOffice:   importClass          org.mozilla.javascript.Context ??? ###
### NB in LibreOffice:  importClass Packages.org.mozilla.javascript.Context ??? ###
importClass Packages.org.mozilla.javascript.Context
importClass Packages.org.mozilla.javascript.tools.shell.Global
#...........................................................................................................
GLOBAL        = new Global Context.enter()
#-----------------------------------------------------------------------------------------------------------
prefix = '/Volumes/Storage/cnd/node_modules/coffeelibre/lib/'
#...........................................................................................................
### Globals ###
eval GLOBAL.readFile prefix + 'require.js'
eval GLOBAL.readFile prefix + 'import-classes.js'
#...........................................................................................................
### Locals ###
TRM                       = require prefix + 'coffeelibre-trm'
log                       = TRM.log.bind TRM
rpr                       = TRM.rpr.bind TRM
xray                      = TRM.xray.bind TRM
#...........................................................................................................
CHR                       = require prefix + 'coffeenode-chr'
TYPES                     = require prefix + 'coffeenode-types'
font_name_by_rsg          = require prefix + 'font-name-by-rsg'
#...........................................................................................................
CL                        = require prefix + 'coffeelibre'




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
  fallback_font_name      = 'Sun-ExtA'
  fallback_font_size      = 14
  tree_background_color   = 'transparent'
  cjk_background_color    = 0x6ec1cc
  empty_background_color  = 'transparent'
  #.........................................................................................................
  fncr_font_name          = 'Adobe Garamond Pro'
  fncr_font_size          = 8
  #.........................................................................................................
  fncr_format_options     =
    'font-name':            fncr_font_name
    'font-size':            fncr_font_size
    'text-wrap':            no
    'vertical-align':       CellVertJustify.CENTER
    'horizontal-align':     CellHoriJustify.LEFT
  #.........................................................................................................
  doc   = CL.get_current_doc()
  sheet = CL.get_current_sheet doc
  cell  = CL.get_cell sheet, 0, 0
  # cell.setFormula 'helo'
  selection = CL.get_current_selection doc
  #.........................................................................................................
  [ xy0, xy1 ]      = CL.get_current_selection doc
  #.........................................................................................................
  [ [ x0, y0, ],
    [ x1, y1, ], ]  = [ xy0, xy1, ]
  #.........................................................................................................
  log xy0, xy1
  log "current selection: #{CL.range_ref_from_xy xy0, xy1}"
  #.........................................................................................................
  # cells = CL.get_current_cells doc
  x2 = x1 + 2
  #.........................................................................................................
  for y in [ y0 .. y1 ]
    fncr_cell = CL.get_cell sheet, x2, y
    #.......................................................................................................
    for x in [ x0 .. x1 ]
      source_cell = CL.get_cell sheet, x, y
      source_text = CL.get_cell_text source_cell
      cell_ref    = CL.cell_ref_from_xy x, y
      #.....................................................................................................
      if source_text.length is 0
        format_options =
          'font-name':            fallback_font_name
          'font-size':            fallback_font_size
          'background-color':     empty_background_color
          'text-wrap':            no
          'vertical-align':       CellVertJustify.CENTER
          'horizontal-align':     CellHoriJustify.CENTER
      #.....................................................................................................
      else
        cid         = CHR.as_cid source_text, input: 'xncr'
        chr_info    = CHR.analyze cid
        fncr        = chr_info[ 'fncr' ]
        rsg         = chr_info[ 'rsg' ]
        fncr        = fncr.replace /^u-pua-/, 'jzr-fig-'
        rsg         = rsg.replace  /^u-pua$/, 'jzr-fig'
        is_cjk      = /^(u-cjk|jzr-fig|u-pua)/.test rsg
        if rsg is 'jzr-fig'
         source_cell.setFormula CHR._unicode_chr_from_cid cid
        #...................................................................................................
        if is_cjk
          fncr_cell.setFormula fncr
          CL.format_cell fncr_cell, fncr_format_options
        #...................................................................................................
        font_name   = @font_name_from_chr_info chr_info, fallback_font_name
        log '©32s', rpr source_text
        log '©32s', cell_ref, fncr, rsg, is_cjk, font_name
        #...................................................................................................
        format_options =
          'font-name':            font_name
          'font-size':            fallback_font_size
          'text-wrap':            no
          'vertical-align':       CellVertJustify.CENTER
          'horizontal-align':     CellHoriJustify.CENTER
        #...................................................................................................
        ### provide background color if cell contains a CJK character: ###
        if is_cjk
          format_options[ 'background-color'  ] = cjk_background_color
        else
          format_options[ 'font-size'         ] = 16
          format_options[ 'background-color'  ] = tree_background_color
      #.....................................................................................................
      CL.format_cell source_cell, format_options
  #.........................................................................................................
  return null


############################################################################################################
@main()



