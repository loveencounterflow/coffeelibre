
# http://openoffice3.web.fc2.com/JavaScript_general.html#OOoGFh1
# http://classdoc.sourceforge.net/examples/so52apidoc/index-all.html
# http://classdoc.sourceforge.net/examples/so52apidoc/


  # TRM.dir 'sheet', UnoRuntime.queryInterface XNamed, sheet
  # xChartData = UnoRuntime.queryInterface(XChartData, xSpreadsheet)
  # fNotANumber = xChartData.getNotANumber()
  # xNamed = UnoRuntime.queryInterface(XNamed, xSpreadsheet)
  # oName = xNamed.getName()


#-----------------------------------------------------------------------------------------------------------
# defineClass
# deserialize
# doctest
# gc
# getConsole
# getErr
# getIn
# getOut
# getPrompts
# help
# init
# initQuitAction
# installRequire
# isInitialized
# load
# loadClass
# pipe
# print
# quit
# readFile
# readUrl
# runCommand
# runDoctest
# seal
# serialize
# setErr
# setIn
# setOut
# setSealedStdLib
# spawn
# sync
# toint32
# version
### NB in OpenOffice:   importClass          org.mozilla.javascript.Context ??? ###
### NB in LibreOffice:  importClass Packages.org.mozilla.javascript.Context ??? ###
importClass Packages.org.mozilla.javascript.Context
importClass Packages.org.mozilla.javascript.tools.shell.Global
#...........................................................................................................
GLOBAL        = new Global Context.enter()


#-----------------------------------------------------------------------------------------------------------
require = ( route ) ->
  #.........................................................................................................
  route         = route + '.js' unless /\js$/.test route
  # source        = GLOBAL.readFile route
  source        = GLOBAL.readFile route
  source        = require.head.concat source, require.tail
  #.........................................................................................................
  try
    R = eval source
  catch error
    for line, idx in source.split '\n'
      # log ( TRM.grey idx + 1 ), '  ', ( TRM.gold line )
      GLOBAL.print idx + 1, '  ', line
    # throw error
  return eval source
#...........................................................................................................
require.head          = """
  (function() {
  var exports = {};
  var module  = { 'exports': exports };
  (function() {
  \n"""
#...........................................................................................................
require.tail = """
  \n}).call( module.exports );
  return module.exports;
  })();
  """

#-----------------------------------------------------------------------------------------------------------
prefix = '/Volumes/Storage/cnd/node_modules/coffeelibre/lib/'
require prefix + 'import-classes'
CHR                       = require prefix + 'coffeenode-chr'
TYPES                     = require prefix + 'coffeenode-types'
font_name_by_rsg          = require prefix + 'font-name-by-rsg'
njs_util                  = require prefix + 'nodejs-util'





# #===========================================================================================================
# # FS
# #-----------------------------------------------------------------------------------------------------------
# FS    = {}
# #-----------------------------------------------------------------------------------------------------------
# FS.read_file = ( route, encoding = 'utf-8' ) ->
#   R = []
#   try
#     input_file    = new File              route
#     stream        = new FileInputStream   input_file
#     stream_reader = new InputStreamReader stream, 'UTF-8'
#     reader        = new BufferedReader    stream_reader
#     R.push line while ( line = reader.readLine() )?
#   finally
#     reader.close() if reader?
#   return R.join '\n'
# #-----------------------------------------------------------------------------------------------------------
# FS.file_exists = ( route ) ->
#   return ( File route ).exists()



#===========================================================================================================
# OBJECT INSPECTION
#-----------------------------------------------------------------------------------------------------------
TRM   = {}
TRM.dir = ( message, x ) ->
  log ''
  log '----------------------------------'
  log message
  log ''
  keys = ( key for key of x ).sort()
  for key in keys
    log '*', key #, TYPES.type_of x[ key ]
  return null

#-----------------------------------------------------------------------------------------------------------
log = ( P... ) ->
  GLOBAL.print ( ( if TYPES.isa_text p then p else rpr p ) for p in P ).join ' '

#-----------------------------------------------------------------------------------------------------------
rpr = ( x ) ->
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
rpr.options =
  showHidden:     no
  depth:          4
  colors:         yes
  customInspect:  no

#-----------------------------------------------------------------------------------------------------------
xray = ( x ) ->
  return njs_util.inspect x, xray.options
#...........................................................................................................
xray.options =
  showHidden:     yes
  depth:          42
  colors:         yes
  customInspect:  no


#===========================================================================================================
# DOCUMENTS
#-----------------------------------------------------------------------------------------------------------
@get_current_doc = ->
  return oDoc = XSCRIPTCONTEXT.getDocument()

#-----------------------------------------------------------------------------------------------------------
@get_undo_manager = ( doc ) ->
  return ( UnoRuntime.queryInterface XUndoManagerSupplier, doc ).getUndoManager()


#===========================================================================================================
# SPREADSHEETS
#-----------------------------------------------------------------------------------------------------------
@_get_spreadsheet_doc = ( doc ) ->
  return UnoRuntime.queryInterface XSpreadsheetDocument, doc

#-----------------------------------------------------------------------------------------------------------
@get_sheets = ( doc ) ->
  R               = []
  sheets_by_name  = ( @_get_spreadsheet_doc doc ).getSheets()
  # sheets_by_index = UnoRuntime.queryInterface XIndexAccess, sheets_by_name
  #.........................................................................................................
  for idx in [ 0 ... sheets_by_name.elementNames.length ]
    sheet_name      = sheets_by_name.elementNames[ idx ]
    sheet           = sheets_by_name.getByName sheet_name
    R[ sheet_name ] = sheet
    R[ idx        ] = sheet
  #.........................................................................................................
  return R

# #-----------------------------------------------------------------------------------------------------------
# @get_sheets2 = ( doc ) ->
#   log 'get_sheets2'
#   log TYPES.type_of sheets
#   TRM.dir 'sheets2', sheets
#   return sheets

#-----------------------------------------------------------------------------------------------------------
@get_sheet = ( doc_or_sheets, name_or_idx ) ->
  throw 'XXXXXXXX'

#-----------------------------------------------------------------------------------------------------------
@_get_sheet_from_name = ( doc, name ) ->
  throw 'XXXXXXXX'

#-----------------------------------------------------------------------------------------------------------
@_get_sheet_from_idx = ( doc, idx ) ->
  return ( @get_sheets doc )[ idx ]

#-----------------------------------------------------------------------------------------------------------
@get_current_sheet_name = ( doc ) ->
  ### Not sure whether it has to be **this** convoluted, but then, here we are, doing OOo... ###
  model             = UnoRuntime.queryInterface XModel, doc
  controller        = model.getCurrentController()
  view              = UnoRuntime.queryInterface XSpreadsheetView, controller
  sheet             = view.getActiveSheet()
  sheet             = UnoRuntime.queryInterface XNamed, sheet
  return sheet.name

#-----------------------------------------------------------------------------------------------------------
@get_current_sheet = ( doc ) ->
  return ( @get_sheets doc )[ @get_current_sheet_name doc ]

#-----------------------------------------------------------------------------------------------------------
@get_current_selection = ( doc ) ->
  R = @_get_current_selection doc
  return [ [ R.StartColumn, R.StartRow, ], [ R.EndColumn, R.EndRow, ], ]

#-----------------------------------------------------------------------------------------------------------
@_get_current_selection = ( doc ) ->
  ### OOo's API is sometimes quite terse and almost intuitive. To obtain the value displayed in the tool bar
  coordinates box—the address of the current selection—a simple one-liner suffices:

      UnoRuntime.queryInterface XCellRangeAddressable, \
        ( UnoRuntime.queryInterface XModel, doc ).getCurrentSelection()

  ###
  model = UnoRuntime.queryInterface XModel, doc
  R     = model.getCurrentSelection()
  R     = UnoRuntime.queryInterface XCellRangeAddressable, R
  return R.rangeAddress

#-----------------------------------------------------------------------------------------------------------
@get_current_cells = ( doc ) ->
  R                 = []
  sheet             = @get_current_sheet doc
  #.........................................................................................................
  [ [ x0, y0, ],
    [ x1, y1, ], ]  = @get_current_selection doc
  #.........................................................................................................
  for y in [ y0 .. y1 ]
    for x in [ x0 .. x1 ]
      # log @cell_ref_from_xy x, y
      R.push @get_cell sheet, x, y
  #.........................................................................................................
  return R

#-----------------------------------------------------------------------------------------------------------
@get_cell = ( sheet, x, y ) ->
  return sheet.getObject().getCellByPosition x, y

#-----------------------------------------------------------------------------------------------------------
@get_cell_text = ( cell ) ->
  ### Return the cell value as text. If the formula text starts with a single-quote character, we elide
  it; it is just OOo's fancy way of saying 'this is a digit but not a number'. ###
  return ( '' + cell.getFormula() ).replace /^'/, ''

#-----------------------------------------------------------------------------------------------------------
@format_cell = ( cell, options ) ->
  pv = UnoRuntime.queryInterface XPropertySet, cell
  # log options[ 'font-name' ]
  #.........................................................................................................
  if ( value = options[ 'font-name' ] )?
    pv.setPropertyValue 'CharFontName',       value
    pv.setPropertyValue 'CharFontNameAsian',  value
  #.........................................................................................................
  if ( value = options[ 'font-size' ] )?
    pv.setPropertyValue 'CharHeight',         value
    pv.setPropertyValue 'CharHeightAsian',    value
  #.........................................................................................................
  if ( value = options[ 'background-color' ] )?
    if value is 'transparent'
      ### NB no need to wrap Booleans with new java.lang.Boolean ###
      pv.setPropertyValue 'IsCellBackgroundTransparent', true
    else
      pv.setPropertyValue 'IsCellBackgroundTransparent', false
      pv.setPropertyValue 'CellBackColor', new java.lang.Integer value
  #.........................................................................................................
  if ( value = options[ 'text-wrap' ] )?
    pv.setPropertyValue 'IsTextWrapped', value
  #.........................................................................................................
  # BLOCK CENTER LEFT REPEAT RIGHT STANDARD
  if ( value = options[ 'horizontal-align' ] )?
    pv.setPropertyValue 'HoriJustify', value
  #.........................................................................................................
  # STANDARD TOP CENTER BOTTOM
  if ( value = options[ 'vertical-align' ] )?
    pv.setPropertyValue 'VertJustify', value
  #.........................................................................................................
  return null
#-----------------------------------------------------------------------------------------------------------
@range_ref_from_xy = ( xy0, xy1 ) ->
  return ( @cell_ref_from_xy xy0... ).concat ':', ( @cell_ref_from_xy xy1... )

#-----------------------------------------------------------------------------------------------------------
@cell_ref_from_xy = ( x, y, x_is_relative = yes, y_is_relative = yes ) ->
  return ( @xref_from_x x, x_is_relative ) + ( @yref_from_y y, y_is_relative )

#-----------------------------------------------------------------------------------------------------------
@xref_from_x = ( x, is_relative = yes ) ->
  R = x.toString 26
  last_idx = R.length - 1
  #.........................................................................................................
  R = R.replace /[0-9]/g, ( digit, idx ) ->
    return String.fromCharCode 0x11 + ( digit.charCodeAt 0 ) - if idx is last_idx then 0 else 1
  #.........................................................................................................
  R = R.replace /[a-z]/g, ( letter, idx ) ->
    return String.fromCharCode ( letter.charCodeAt 0 ) - 0x16 - if idx is last_idx then 0 else 1
  #.........................................................................................................
  return if is_relative then R else '$' + R

#-----------------------------------------------------------------------------------------------------------
@yref_from_y = ( y, is_relative = yes ) ->
  R = ( y + 1 ).toString 10
  return if is_relative then R else '$' + R


#===========================================================================================================
# TEXT DOCUMENTS
#-----------------------------------------------------------------------------------------------------------
@get_text_range = ->
  throw 'XXXXXXXX'
  oDoc = @get_current_doc()
  # get the XTextDocument interface
  xTextDoc = UnoRuntime.queryInterface XTextDocument, oDoc
  # get the XText interface
  xText = xTextDoc.getText()
  # get an (empty) XTextRange interface at the end of the text
  return xTextRange = xText.getEnd()

#-----------------------------------------------------------------------------------------------------------
@set_text = ( text ) ->
  throw 'XXXXXXXX'
  text_range  = @get_text_range()
  text_range.setString text
  return null

#===========================================================================================================
#
#-----------------------------------------------------------------------------------------------------------
test_format_cells = ->
  column = 2
  for row in [ 20 .. 32 ]
    cell = get_cell column, row
    # log [ column, row, get_cell_text cell ].join ' '
    # cid       = CHR.as_cid cell_text
    format_cell_by_content cell
    cell_text = get_cell_text cell
    if cell_text.length is 0
      log "skipped empty cell #{@cell_ref_from_xy column, row}"
      continue
    chr_info  = CHR.analyze cell_text, input: 'xncr'
    cell.setFormula CHR._unicode_chr_from_cid chr_info[ 'cid' ]
    log chr_info[ 'chr' ]
    #.......................................................................................................
    cell      = get_cell column + 1, row
    fncr      = chr_info[ 'fncr' ]
    fncr      = fncr.replace /^u-pua-/, 'jzr-fig-'
    cell.setFormula fncr
  #.........................................................................................................
  # cell      = get_cell 2, 16
  # cell_text = get_cell_text cell
  # cell.setFormula cell_text
  # cell.setFormula '\ue100'
  # @format_cell cell, 'font-name': 'jizura2'
  # show_text_cids cell_text
  # show_text_cids 'helo 中'

#-----------------------------------------------------------------------------------------------------------
show_text_cids = ( text ) ->
  for idx in [ 0 ... text.length ]
    log text[ idx ] + ': 0x' + ( text.charCodeAt idx ).toString 16
  # log '中國皇帝'


#-----------------------------------------------------------------------------------------------------------
show_cid = ->
  for row in [ 3 .. 4 ]
    source_cell = get_cell 1, row
    target_cell = get_cell 2, row
    text_cesu8  = get_cell_text source_cell
    # log TYPES.type_of text_cesu8
    cid         = CHR.as_cid text_cesu8
    cid_hex     = cid.toString 16
    fncr        = CHR.as_fncr cid
    # target_cell.setFormula cid_hex
    chr_info    = CHR.analyze cid
    target_cell.setFormula fncr

#-----------------------------------------------------------------------------------------------------------
font_name_from_rsg = ( rsg, fallback ) ->
  ### TAINT should use chr-info or CID, since we might have to format with codepoint granularity. ###
  R = font_name_by_rsg[ rsg ]
  #.........................................................................................................
  unless R?
    return fallback unless fallback is undefined
    throw new Error "unable to find a suitable font for RSG #{rpr rsg}"
  #.........................................................................................................
  return R

#-----------------------------------------------------------------------------------------------------------
font_name_from_chr_info = ( chr_info, fallback ) ->
  rsg = chr_info[ 'rsg' ]
  R   = font_name_by_rsg[ rsg ]
  #.........................................................................................................
  unless R?
    return fallback unless fallback is undefined
    throw new Error "unable to find a suitable font for #{chr_info[ 'fncr' ]} #{chr_info[ 'chr' ]}"
  #.........................................................................................................
  return R

#-----------------------------------------------------------------------------------------------------------
format_cell_by_content = ( cell ) ->
  text        = get_cell_text cell
  return null if text.length is 0
  chr_info    = CHR.analyze text, input: 'xncr'
  rsg         = chr_info[ 'rsg' ]
  font_name   = font_name_from_rsg rsg
  #.........................................................................................................
  return @format_cell cell, 'font-name': font_name, 'font-size': 14


#===========================================================================================================
#
#-----------------------------------------------------------------------------------------------------------
@show_formatting = ->

  # Cell properties:
  #  oid
  #  formula
  #  setValue
  #  notifyAll
  #  setFormula
  #  getValue
  #  getError
  #  type
  #  queryInterface
  #  equals
  #  notify
  #  getOid
  #  class
  #  wait
  #  value
  #  toString
  #  hashCode
  #  getType
  #  error
  #  getFormula
  #  isSame
  #  getClass

  #.........................................................................................................
  cell    = get_cell 1, 1
  cell.setFormula ( name for name of cell ).join ', '
  # cell.value = 42
  #.........................................................................................................
  cell    = get_cell 1, 2
  pv = UnoRuntime.queryInterface XPropertySet, cell
  cell.setFormula ( name for name of pv ).join ', '
  pv.setPropertyValue 'CharHeight', 14
  pv.setPropertyValue 'CharFontName', 'Sun-ExtA'
  #.........................................................................................................
  cell    = get_cell 1, 3
  pv = UnoRuntime.queryInterface XPropertySet, cell
  pv.setPropertyValue 'CharHeight', 14
  log "CharFontName:              #{pv.getPropertyValue 'CharFontName'}"
  log "CharFontStyleName:         #{pv.getPropertyValue 'CharFontStyleName'}"
  log "CharFontFamily:            #{pv.getPropertyValue 'CharFontFamily'}"
  log "CharFontCharSet:           #{pv.getPropertyValue 'CharFontCharSet'}"
  log "CharFontNameAsian:         #{pv.getPropertyValue 'CharFontNameAsian'}"
  log "CharFontStyleNameAsian:    #{pv.getPropertyValue 'CharFontStyleNameAsian'}"
  log "CharFontFamilyAsian:       #{pv.getPropertyValue 'CharFontFamilyAsian'}"
  log "CharFontCharSetAsian:      #{pv.getPropertyValue 'CharFontCharSetAsian'}"
  log "CharFontNameComplex:       #{pv.getPropertyValue 'CharFontNameComplex'}"
  log "CharFontStyleNameComplex:  #{pv.getPropertyValue 'CharFontStyleNameComplex'}"
  log "CharFontFamilyComplex:     #{pv.getPropertyValue 'CharFontFamilyComplex'}"
  log "CharFontCharSetComplex:    #{pv.getPropertyValue 'CharFontCharSetComplex'}"
  cell_text = get_cell_text cell
  # cell.setFormula 'x'
  # pv.setPropertyValue 'CharFontName', 'Sun-ExtA'
  # pv.setPropertyValue 'CharFontFamilyAsian',    new Float 0
  # pv.setPropertyValue 'CharFontCharSetAsian',   -1
  # pv.setPropertyValue 'CharFontName',           'Courier'
  # pv.setPropertyValue 'CharFontNameAsian',      'Courier'
  # pv.setPropertyValue 'CharFontNameComplex',    'Courier'
  # pv.setPropertyValue 'CharFontName',           'Adobe FangSong Std R'
  # pv.setPropertyValue 'CharFontNameAsian',      'Adobe FangSong Std R'
  pv.setPropertyValue 'CharFontName',           'Sun-ExtA'
  pv.setPropertyValue 'CharFontNameAsian',      'Sun-ExtA'
  pv.setPropertyValue 'CharHeight',             14
  pv.setPropertyValue 'CharHeightAsian',        14
  # pv.setPropertyValue 'CharFontNameComplex',    'Adobe FangSong Std R'
  # cell.setString '丁'
  show_cid()
  dir 'cell', cell

  pv = UnoRuntime.queryInterface XPropertySet, cell
  # oid
  # notifyAll
  # addPropertyChangeListener
  # removePropertyChangeListener
  # addVetoableChangeListener
  # removeVetoableChangeListener
  # queryInterface
  # equals
  # notify
  # getOid
  # class
  # propertySetInfo
  # wait
  # toString
  # hashCode
  # getPropertyValue
  # isSame
  # getClass
  # setPropertyValue
  # getPropertySetInfo



# When a spreadsheet is open in OO.o, this macro will
# loop over a given number of rows and columns and
# summarize those cells in a JEditorPane
#
# Copyleft 2010 by Kas Thomas
# http://asserttrue.blogspot.com/

# go thru the sheet one row at a time
# and collect cell data into an array of
# records, where each record is an array
# of cell data for a given row
read_cells = (sheet, rows, columns) ->
  masterArray = []
  i = 0

  while i < rows
    ar = []
    k = 0

    while k < columns
      cell = sheet.getObject().getCellByPosition(k, i)
      content = cell.getFormula()
      unless content.indexOf(",") is -1
        ar.push "\"" + content + "\""
      else
        ar.push content
      k++
    masterArray.push ar
    i++
  masterArray


#===========================================================================================================
#
#-----------------------------------------------------------------------------------------------------------
show_objects = ->
  log '###################################################################################################'
  # xTextRange.setString '\n\n' + ( name for name of Packages.com.sun.star.text ).join '\n'
  # xTextRange.setString '\n\n' + ( name for name of xTextRange ).join '\n'
  # dir 'xTextRange', get_text_range()
  dir 'Packages', Packages
  # dir '_get_spreadsheet_doc()', _get_spreadsheet_doc()
  # dir 'Packages.com.sun.star', Packages.com.sun.star
  # dir 'Packages.com.sun.star.uno', Packages.com.sun.star.uno
  # dir 'Packages.com.sun.star.uno.UnoRuntime.getCurrentContext()', Packages.com.sun.star.uno.UnoRuntime.getCurrentContext()
  # dir 'Packages.java', Packages.java
  # dir 'Packages.javax', Packages.javax
  # dir 'ScriptContext', ScriptContext
  # dir 'XSCRIPTCONTEXT.getDocument()', XSCRIPTCONTEXT.getDocument()
  dir 'Packages.com.sun.star.text', Packages.com.sun.star.text
  dir 'Packages.com.sun.star.sheet', Packages.com.sun.star.sheet
  dir 'Packages.com.sun.star.script', Packages.com.sun.star.uno.script
  dir 'Packages.com.sun.star.uno.UnoRuntime', Packages.com.sun.star.uno.UnoRuntime
  dir 'XSCRIPTCONTEXT', XSCRIPTCONTEXT
  dir 'CHR', CHR
  log ( new Date() ).toString()


#===========================================================================================================
#
#-----------------------------------------------------------------------------------------------------------
@main2 = ->

  #get the doc object from the scripting context
  oDoc = XSCRIPTCONTEXT.getDocument()

  #get the XSpreadsheetDocument interface from the doc
  xSDoc = UnoRuntime.queryInterface(XSpreadsheetDocument, oDoc)

  # get a reference to the sheets for this doc
  sheets = xSDoc.getSheets()

  # get Sheet1
  sheet1 = sheets.getByName("Sheet1")

  # construct a new EditorPane
  editor = new EditorPane()
  pane = editor.getPane()

  # harvest cell data (from sheet, rows, cols)
  masterArray = read_cells(sheet1, 100, 8)

  # display the data
  text = masterArray.join("\n")
  pane.setText text


#===========================================================================================================
#
#-----------------------------------------------------------------------------------------------------------
@main = ->
  # log rpr ( name for name of GLOBAL )
  # log GLOBAL.readFile
  # log xray GLOBAL.readFile
  # log TYPES.type_of GLOBAL.readFile
  log ''
  log '©42 --------------------------------------------------------------------------------------------'
  log ( new Date() ).toString()
  #.........................................................................................................
  doc   = @get_current_doc()
  UNDO  = @get_undo_manager doc
  UNDO.enterUndoContext 'format tree'
  #.........................................................................................................
  try
    @format_tree()
  finally
    UNDO.leaveUndoContext()
  #.........................................................................................................
  return null

#-----------------------------------------------------------------------------------------------------------
@format_tree = ->
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
  doc   = @get_current_doc()
  sheet = @get_current_sheet doc
  cell  = @get_cell sheet, 0, 0
  # cell.setFormula 'helo'
  selection = @get_current_selection doc
  #.........................................................................................................
  [ xy0, xy1 ]      = @get_current_selection doc
  #.........................................................................................................
  [ [ x0, y0, ],
    [ x1, y1, ], ]  = [ xy0, xy1, ]
  #.........................................................................................................
  log xy0, xy1
  log "current selection: #{@range_ref_from_xy xy0, xy1}"
  #.........................................................................................................
  # cells = @get_current_cells doc
  x2 = x1 + 2
  #.........................................................................................................
  for y in [ y0 .. y1 ]
    fncr_cell = @get_cell sheet, x2, y
    #.......................................................................................................
    for x in [ x0 .. x1 ]
      source_cell = @get_cell sheet, x, y
      source_text = @get_cell_text source_cell
      cell_ref    = @cell_ref_from_xy x, y
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
          @format_cell fncr_cell, fncr_format_options
        #...................................................................................................
        font_name   = font_name_from_chr_info chr_info, fallback_font_name
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
      @format_cell source_cell, format_options
  #.........................................................................................................
  # cell = @get_cell sheet, 3, 4
  # ctx   = XSCRIPTCONTEXT.getComponentContext()
  # smgr  = ctx.getServiceManager()
  # TRM.dir 'smgr', smgr
  # service_names = smgr.getAvailableServiceNames()
  # for idx in [ 0 ... service_names.length ]
  #   service_name = service_names[ idx ]
  #   GLOBAL.print service_name if /^com\.sun\.star\.script/.test service_name
  # GLOBAL.print service_names.length
  # # script_provider = smgr.createInstanceWithContext 'com.sun.star.script.provider.ScriptProviderForJavaScript', ctx
  # script_provider = smgr.createInstanceWithContext 'com.sun.star.script.provider.ScriptProviderForBasic', ctx
  # TRM.dir 'script_provider', script_provider
  # script_provider = UnoRuntime.queryInterface XScriptProvider, script_provider
  # TRM.dir 'script_provider', script_provider
  # # smgr.createInstanceWithContext name, ctx
  # # cell = UnoRuntime.queryInterface XShape, cell
  # # x = UnoRuntime.queryInterface XDrawPage, sheet
  # # log TYPES.type_of cell.add
  # # TRM.dir 'XSCRIPTCONTEXT', XSCRIPTCONTEXT
  # # ThisComponent = XSCRIPTCONTEXT.getComponentContext()
  # # ControlShape
  # # oLShape  = ThisComponent.CreateInstance("com.sun.star.drawing.ControlShape")
  # # ThisComponent.Drawpage.Add(oLShape)
  # # log CellHoriJustify.CENTER
  # # log CellHoriJustify.CENTER_value
  log ( new Date() ).toString()
  log 'ok'


#-----------------------------------------------------------------------------------------------------------
@f = ->
  # service_names = service_manager.getAvailableServiceNames()
  # for idx in [ 0 ... service_names.length ]
  #   service_name = service_names[ idx ]
  #   GLOBAL.print service_name if /^com\.sun\.star\.script/.test service_name
  # GLOBAL.print service_names.length
  # # script_provider = service_manager.createInstanceWithContext 'com.sun.star.script.provider.ScriptProviderForJavaScript', ctx
  # script_provider = service_manager.createInstanceWithContext 'com.sun.star.script.provider.ScriptProviderForBasic', ctx
  # path_settings = UnoRuntime.queryInterface XMultiPropertySet, path_settings
  # TRM.dir 'path_settings', path_settings
  # TRM.dir 'path_settings.getPropertySetInfo()', path_settings.getPropertySetInfo()
  # TRM.dir 'path_settings.getPropertySetInfo().getProperties()', path_settings.getPropertySetInfo().getProperties()
  # log '©45f', '' + path_settings.getPropertyValue 'Work'
  # log '©45f', '' + path_settings.getPropertySetInfo().getProperties().length
  # TRM.dir 'service_manager', service_manager

  context           = XSCRIPTCONTEXT.getComponentContext()
  service_manager   = context.getServiceManager()
  path_settings     = service_manager.createInstanceWithContext 'com.sun.star.util.PathSettings', context
  path_settings     = UnoRuntime.queryInterface XPropertySet, path_settings
  #.........................................................................................................
  # for property in path_settings.getPropertySetInfo().getProperties()
  #   # TRM.dir 'property', property
  #   name  = '' + property.Name
  #   continue if /_internal$/.test name
  #   continue if     /_user$/.test name
  #   continue if /_writable$/.test name
  #   routes = '' + path_settings.getPropertyValue name
  #   routes = routes.split ';'
  #   for route in routes
  #     route = route.replace /^file:\/\//, ''
  #     route = route.replace /%([0-9a-f]{2})/, ( $0, $1 ) -> return String.fromCharCode parseInt $1, 16
  #     log "#{TEXT.flush_left name, 20} #{route}"
  #.........................................................................................................
  script_provider     = service_manager.createInstanceWithContext 'com.sun.star.script.provider.ScriptProviderForJavaScript', context
  meta_data           = service_manager.createInstanceWithContext 'com.sun.star.script.framework.container.ScriptMetaData', context
  script_uri_helper   = service_manager.createInstanceWithContext 'com.sun.star.script.provider.ScriptURIHelper', context
  script_uri_helper   = UnoRuntime.queryInterface XScriptURIHelper, script_uri_helper
  script_context      = script_provider.getScriptingContext()
  invocation_context  = XSCRIPTCONTEXT.getInvocationContext()
  TRM.dir 'XSCRIPTCONTEXT', XSCRIPTCONTEXT
  script_container    = invocation_context.getScriptContainer()
  script_container    = UnoRuntime.queryInterface XInvocation2, script_container
  TRM.dir 'script_container', script_container
  TRM.dir 'script_uri_helper', script_uri_helper
  TRM.dir 'meta_data', meta_data
  # invocation_context  = UnoRuntime.queryInterface XInvocation2, invocation_context
  # TRM.dir 'invocation_context', invocation_context
  # log 'script_uri_helper.rootStorageURI: ' + script_uri_helper.rootStorageURI
  # log 'script_uri_helper.storageURI: ' + script_uri_helper.storageURI
  # log 'getScriptURI:      ' + script_uri_helper.getScriptURI()
  # log 'getStorageURI:     ' + script_uri_helper.getStorageURI()

  # TRM.dir 'script_container', script_container
  # log '' + script_provider.getScriptingContext()

#-----------------------------------------------------------------------------------------------------------
string_from_stream = ( stream, encoding = 'utf-8' ) ->
  # R = []
  # try
  #   stream_reader = new StreamReader stream, 'UTF-8'
  #   reader        = new BufferedReader    stream_reader
  #   R.push line while ( line = reader.readLine() )?
  # finally
  #   reader.close() if reader?
  # return R.join '\n'
  return '' + stream.toString()
my_out = new StringWriter()
# my_out = new ByteArrayOutputStream( 1024 )
log GLOBAL.runCommand 'ls', '/tmp', output: my_out
log '©29f', string_from_stream my_out


############################################################################################################
@main()
# main2()
# show_objects()
# test_format_cells()


# sub drawcircle
# rem ----------------------------------------------------------------------
# rem define variables
# dim document   as object
# dim dispatcher as object
# rem ----------------------------------------------------------------------
# rem get access to the document
# document   = ThisComponent.CurrentController.Frame
# dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

# rem ----------------------------------------------------------------------
# dim args1(0) as new com.sun.star.beans.PropertyValue
# args1(0).Name = "InsertDraw"
# args1(0).Value = 3

# dispatcher.executeDispatch(document, ".uno:InsertDraw", "", 0, args1())
# end sub

# Sub InsertProcessShape
#    Dim oDoc As Object
#    Dim oDrawPage As Object
#    Dim oShape As Object
#    Dim shapeGeometry(0) as new com.sun.star.beans.PropertyValue
#    Dim oSize As new com.sun.star.awt.Size
#    oSize.width = 3000
#    oSize.height = 1000
#    oDoc = ThisComponent
#    odrawPage = oDoc.DrawPages(0)
#    oShape = oDoc.createInstance("com.sun.star.drawing.CustomShape")
#    shapeGeometry(0).Name = "Type"
#    shapeGeometry(0).Value = "flowchart-process"
#    oDrawPage.add(oShape)
#    oShape.CustomShapeGeometry = shapeGeometry
#    oShape.Size = oSize
# End Sub

# sub insertpicture
# rem ----------------------------------------------------------------------
# rem define variables
# dim document   as object
# dim dispatcher as object
# rem ----------------------------------------------------------------------
# rem get access to the document
# document   = ThisComponent.CurrentController.Frame
# dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
# rem ----------------------------------------------------------------------
# dim args1(2) as new com.sun.star.beans.PropertyValue
# args1(0).Name = "FileName"
# args1(0).Value = "file://localhost/Users/flow/Pictures/Scan%201.jpg"
# args1(1).Name = "FilterName"
# args1(1).Value = "JPEG - Joint Photographic Experts Group"
# args1(2).Name = "AsLink"
# args1(2).Value = false
# dispatcher.executeDispatch(document, ".uno:InsertGraphic", "", 0, args1())
# end sub


# function RegRep (sSource as String, sRegExp as String, sGlobUpcase as String, sReplace as String) as String
# oMasterScriptProviderFactory = createUnoService("com.sun.star.script.provider.MasterScriptProviderFactory")
# oMasterScriptProvider = oMasterScriptProviderFactory.createScriptProvider("")
# oScriptReplace = oMasterScriptProvider.getScript("vnd.sun.star.script:Tools.Replace.js?language=JavaScript&location=user")
# RegRep =  oScriptReplace.invoke(Array(sSource, sRegExp, sGlobUpcase, sReplace ), Array(), Array())
# end function
