



#===========================================================================================================
# DOCUMENTS
#-----------------------------------------------------------------------------------------------------------
@get_current_doc = ->
  return oDoc = XSCRIPTCONTEXT.getDocument()

#-----------------------------------------------------------------------------------------------------------
@get_undo_manager = ( doc ) ->
  return ( UnoRuntime.queryInterface XUndoManagerSupplier, doc ).getUndoManager()

#-----------------------------------------------------------------------------------------------------------
@step = ( doc, title, action ) ->
  ### Perform an atomic, undoable action. ###
  UNDO  = @get_undo_manager doc
  UNDO.enterUndoContext title
  #.........................................................................................................
  try
    action()
  finally
    UNDO.leaveUndoContext()
  #.........................................................................................................
  return null


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
  if ( value = options[ 'cell-style-name' ] )?
    pv.setPropertyValue 'CellStyle',          value
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
