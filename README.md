## coffeelibre


### Why?

There are not too many options when it comes to chosing a free and open spreadsheet solution. Online
solutions i've tried were often too clunky and restricted in their possibilities, and MS Excel—while
arguably the champion in this field—is neither free nor open. This leaves you with having to chose
between one of the `/[a-z]+office\.org/` alternatives; more specifically,

*   http://openoffice.org
*   http://libreoffice.org
*   http://neooffice.org

which are essentially clones of each other.

> Of these, NeoOffice leaves the best impression on OSX; i will still use OpenOffice for the purposes of
> the present discussion.

It so happened that a while ago i had a fancy idea how to catalogue components of Chinese characters in a
tree-like fashion; all i wanted was a software that would be more amenable to try out this idea than a pure
text format (which i had tried). I figured that using Unicode box drawing characters as connectors should be
good enough to visualize that tree, and that i could repurpose a spreadsheet—basically, a cellular
automaton of sorts—to help me allocate spaces and characters in a suitable fashion, so i turned to
OpenOffice.

I quickly found what i always find after opening an OpenOffice application: that its interface is incredibly
clunky, ugly, and backwards; that its developers have a propensity to hide functionalities in deeply nested
menus, submenus, and subsubsubsmenus, that new versions of it are unlikely to come up with anything new,
that it is quite happy to crash any given moment, and that especially OpenOffice Calc is incredibly stupid /
stubborn / incapable of doing any consistent formatting (with or without using named styles).

Lacking alternatives and wanting to get on with my idea, i looked into writing macros for my purpose. The
thinking was that i needed several different fonts for the Chinese characters alone, as they are spread
about Unicode's BMP and SIP, and i have a hundred or so character components in the Private Use Area, which
would have to be displayed in my own custom font. Also, the tree lines should get their own suitable font,
and of course i wanted colors. I knew that i'd have to move around stuff quite a bit to make ends meet, and
that OpenOffice is a bitch when it comes to formatting, so i needed custom macros to help me do that.

I managed to find ready-made macros written in Basic that made moving cells using the keyboard much easier,
but i had no hope i could find prefab macros to do the formatting—i mean, who on earth writes macros to
format characters in spreadsheet cells according to their respective Unicode code points? I already had
implementations in Coffee/JavaScript to do exactly that, except for the part where i had to tell OpenOffice
to apply the formatting. Since OpenOffice does boast support for JavaScript as macro languages, so i decided
to give it a try.



### How?

While there is (on OSX) a folder `~/Library/Application Support/OpenOffice/4/user/Scripts` (which may
or may not work for scripting as described here), i chose to use
`/Applications/OpenOffice.app/Contents/share/Scripts/javascript` as location to link a folder that
contains my scripts—more specifically, you could simply

    cd ~
    git clone https://github.com/loveencounterflow/coffeelibre.git
    cd /Applications/OpenOffice.app/Contents/share/Scripts/javascript
    ln -s ~/coffeelibre CoffeeLibreDemo # or whatever name
    cd -

> Notice that on OSX, applications like OpenOffice are distributed as `*.app` bundles, which are really
> directories that behave as files under many circumstances; you can still `cd` into an `*.app`, but Finder
> will show them as files (which are run on double-click rather than entered into); PathFinder does have a
> `View / Show Package Contents` option in its menu two switch between normal and nerd modes.







# Materials


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

