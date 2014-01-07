coffeelibre
===========

Scripting Libre(Open/Neo)Office with CoffeeScript

### Cell Properties

    doc         = CL.get_current_doc()
    sheet       = CL.get_current_sheet doc
    cell        = CL.get_cell sheet, 0, 0
    pv          = UnoRuntime.queryInterface XPropertySet, cell
    psi         = pv.getPropertySetInfo()
    properties  = psi.getProperties()
    for idx in [ 0 ... properties.length ]
      log '' + properties[ idx ].Name


    AbsoluteName
    AsianVerticalMode
    BottomBorder
    CellBackColor
    CellProtection
    CellStyle
    CharColor
    CharContoured
    CharCrossedOut
    CharEmphasis
    CharFont
    CharFontCharSet
    CharFontCharSetAsian
    CharFontCharSetComplex
    CharFontFamily
    CharFontFamilyAsian
    CharFontFamilyComplex
    CharFontName
    CharFontNameAsian
    CharFontNameComplex
    CharFontPitch
    CharFontPitchAsian
    CharFontPitchComplex
    CharFontStyleName
    CharFontStyleNameAsian
    CharFontStyleNameComplex
    CharHeight
    CharHeightAsian
    CharHeightComplex
    CharLocale
    CharLocaleAsian
    CharLocaleComplex
    CharOverline
    CharOverlineColor
    CharOverlineHasColor
    CharPosture
    CharPostureAsian
    CharPostureComplex
    CharRelief
    CharShadowed
    CharStrikeout
    ChartColumnAsLabel
    ChartRowAsLabel
    CharUnderline
    CharUnderlineColor
    CharUnderlineHasColor
    CharWeight
    CharWeightAsian
    CharWeightComplex
    CharWordMode
    ConditionalFormat
    ConditionalFormatLocal
    ConditionalFormatXML
    DiagonalBLTR
    DiagonalTLBR
    FormulaLocal
    FormulaResultType
    HoriJustify
    IsCellBackgroundTransparent
    IsTextWrapped
    LeftBorder
    NumberFormat
    NumberingRules
    Orientation
    ParaAdjust
    ParaBottomMargin
    ParaIndent
    ParaIsCharacterDistance
    ParaIsForbiddenRules
    ParaIsHangingPunctuation
    ParaIsHyphenation
    ParaLastLineAdjust
    ParaLeftMargin
    ParaRightMargin
    ParaTopMargin
    Position
    RightBorder
    RotateAngle
    RotateReference
    ShadowFormat
    ShrinkToFit
    Size
    TableBorder
    TopBorder
    UserDefinedAttributes
    Validation
    ValidationLocal
    ValidationXML
    VertJustify
    WritingMode


### The Global Object

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


### Helpful Links

* http://openoffice3.web.fc2.com/JavaScript_general.html#OOoGFh1
* http://classdoc.sourceforge.net/examples/so52apidoc/index-all.html
* http://classdoc.sourceforge.net/examples/so52apidoc/
* https://wiki.openoffice.org/wiki/Framework/Article/Accelerators_Configuration
* https://github.com/mozilla/rhino/blob/master/toolsrc/org/mozilla/javascript/tools/shell/Global.java
* http://www.openoffice.org/api/docs/common/ref/com/sun/star/beans/XPropertySet.html
* http://www.openoffice.org/api/docs/common/ref/index-files/index-19.html

* http://www.pitonyak.org/oo.php
    * http://www.pitonyak.org/OOME_3_0.pdf
    * http://www.pitonyak.org/AndrewMacro.pdf

* http://bernard.marcelly.perso.sfr.fr/index2.html
    * http://bernard.marcelly.perso.sfr.fr/Files_en/XrayTool60_en.odt
    * http://bernard.marcelly.perso.sfr.fr/Files_en/JS_OOo_v11en.zip

