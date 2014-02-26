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

> Of these, NeoOffice gives the best impression on OSX; i will still use OpenOffice for the purposes of
> the present discussion.

It so happened that a while ago i had a fancy idea how to catalogue components of Chinese characters in a
tree-like fashion; all i wanted was a software that would be more amenable to try out this idea than a pure
text format (which i had tried). I figured that using Unicode box drawing characters as connectors should be
good enough to visualize that tree, and that i could repurpose a spreadsheet—basically, a cellular
automaton of sorts—to help me allocate spaces and characters in a suitable fashion, so i turned to
OpenOffice.

I quickly found what i always find after opening an OpenOffice application:

* that its **interface** is **incredibly clunky, ugly, and backwards**;
* that its developers have a propensity to **hide functionalities in deeply nested
  menus, submenus, and subsubsubmenus**;
* that new versions of it are **unlikely to come up with anything new**;
* that it's quite **happy to crash** any given moment, and
* that (especially Calc) is **incredibly stupid / stubborn / incapable of doing any meaningful / consistent
  formatting** (with or without using named styles).

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

#### The Disappointing Hello World Example

Which is where my troubles started. The OpenOffice folks are quite dexterous at producing a huge pile of
documentation on their website (or at least they were when they stopped updating a lot of the stuff years
ago). They are also very good at providing a Hello World example in no less than five language variants.
They're also masters in letting the matter rest at that; if you're in for some JS-in-OOo macro fun, this
much is about how much you're going to get from a readily available source. They're also lightning fast in
updating their sample snippets. Let's have a look at the code (on
https://wiki.openoffice.org/wiki/Documentation/DevGuide/Scripting/Writing_Macros):

in 2007 (412 characters):

````javascript
importClass(Packages.com.sun.star.uno.UnoRuntime);
importClass(Packages.com.sun.star.text.XTextDocument);
importClass(Packages.com.sun.star.text.XText);
importClass(Packages.com.sun.star.text.XTextRange);

oDoc = XSCRIPTCONTEXT.getDocument();
xTextDoc = UnoRuntime.queryInterface(XTextDocument,oDoc);
xText = xTextDoc.getText();
xTextRange = xText.getEnd();
xTextRange.setString( "Hello World (in JavaScript)" );
````

in 2014 (909 characters, not counting comments):

````javascript
importClass(Packages.com.sun.star.uno.UnoRuntime)
importClass(Packages.com.sun.star.text.XTextDocument)
importClass(Packages.com.sun.star.text.XText)
importClass(Packages.com.sun.star.text.XTextRange)
importClass(Packages.com.sun.star.beans.XPropertySet)
importClass(Packages.com.sun.star.awt.FontSlant)
importClass(Packages.com.sun.star.awt.FontUnderline)

oDoc = XSCRIPTCONTEXT.getDocument()
xTextDoc = UnoRuntime.queryInterface(XTextDocument,oDoc)
xText = xTextDoc.getText()
xTextRange = xText.getEnd()
pv = UnoRuntime.queryInterface(XPropertySet, xTextRange)

pv.setPropertyValue("CharHeight", 16.0) // Double
// CharBackColor receives an Integer
pv.setPropertyValue("CharBackColor", new java.lang.Integer(1234567))
// CharUnderline receives a group constant
pv.setPropertyValue("CharUnderline", new java.lang.Short(Packages.com.sun.star.awt.FontUnderline.WAVE))
// CharPosture receives an enum
pv.setPropertyValue("CharPosture", Packages.com.sun.star.awt.FontSlant.ITALIC)
xTextRange.setString( "Hello World (in JavaScript)" )
````

That's a whopping dev rate of about 71 characters per year (of course i'm cheating here—the actual rate
is much higher, since the last significant changes to the code shown must have occurred more than four years
ago, according to the Wiki's history page).

Alas, elaborate as the code is, it won't work with OpenOffice Calc, only in Writer; also, there is a
likelyhood that it will crash on OSX, since there is a log-standing bug that keeps macros from using
Java AWT stuff (google `openoffice osx awt` for this one).

There are a few points that merit your attention in these snippets:

* there's this strange `queryInterface` call which we'll discuss in a moment;
* while the examples are basically JavaScript, some of the vocabulary is non-standard; in particular,
  * `importClass`
  * `XSCRIPTCONTEXT`
  * `UnoRuntime`
  * `GLOBAL` (not shown here)
  are not part of standard JavaScript. Why the OOo folks chose to distribute their custom facilities in no
  less than four different objects i have no clue.


#### XXXXXXXXX

Still, i managed to get the sample code running, and intense use of search engines turned up various code
snippets. I then set out to translate those snippets into CoffeeScript, isolate pertinent pieces of
functionality, and organize them into functions with meaningful names. It's really very much a matter of
undoing the unholy mess that the OOo API is. I mean, consider this code that essentially just gives you
an object that represents the current spreadsheet:


````coffeescript
@get_current_doc = ->
  return XSCRIPTCONTEXT.getDocument()

@_get_spreadsheet_doc = ( doc ) ->
  return UnoRuntime.queryInterface XSpreadsheetDocument, doc

@get_sheets = ( doc ) ->
  R               = []
  sheets_by_name  = ( @_get_spreadsheet_doc doc ).getSheets()
  for idx in [ 0 ... sheets_by_name.elementNames.length ]
    sheet_name      = sheets_by_name.elementNames[ idx ]
    sheet           = sheets_by_name.getByName sheet_name
    R[ sheet_name ] = sheet
    R[ idx        ] = sheet
  return R

@get_current_sheet_name = ( doc ) ->
  model       = UnoRuntime.queryInterface XModel, doc
  controller  = model.getCurrentController()
  view        = UnoRuntime.queryInterface XSpreadsheetView, controller
  sheet       = view.getActiveSheet()
  sheet       = UnoRuntime.queryInterface XNamed, sheet
  return sheet.name

@get_current_sheet = ( doc ) ->
  return ( @get_sheets doc )[ @get_current_sheet_name doc ]
````

There's a good chance that my code could be made more efficient, but then it does work as it stands. The
important thing to understand here (and the one most annoying thing you will want to abstract away like
crazy) is that insane

````coffeescript
doc = XSCRIPTCONTEXT.getDocument()
UnoRuntime.queryInterface XSpreadsheetDocument, doc
````

monkey business that, if not kept at bay, would permeat each and every step you want to take when scripting.

It would appear that, being the smart guys they are, the OOo folks chose to embrace Java (and XML) to the
fullest. Of course, Java being Java, it does embrace static typing—a feature well known to get into the way
of free-and-easy programmers quite often. So why not abstract the bejeezes out of it?

![very abstract](https://raw.github.com/loveencounterflow/coffeelibre/master/art/extreme-abstracting.jpg)

> art by http://geek-and-poke.com/

Basically, `XSCRIPTCONTEXT.getDocument()` gives you an object that represents *something* in a hand-waving
non-committal fashion, and *nothing in particular* at the same time. To get more specific, we still have to
'wrap' that `NotAnythingInParticular` blob with an interface. To quote from the docs (which are easily
retrievable from https://www.openoffice.org/api/docs/java/ref/com/sun/star/uno/UnoRuntime.html#queryInterface%28com.sun.star.uno.Type,%20java.lang.Object%29)

> [`UnoRuntime.queryInterface`] returns null in case the given UNO object does not support the given UNO
> interface type (or is itself null). Otherwise, a reference to a Java object implementing the Java
> interface type corresponding to the given UNO interface is returned. In the latter case, it is
> unspecified whether the returned Java object is the same as the given object, or is another facet of
> that UNO object.

Anyone shouting 'JUST GIMME THAT DARN OBJECT ALREADY' at this point?

#### XXXXXXXXX

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

