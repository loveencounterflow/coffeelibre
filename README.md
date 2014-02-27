

- [CoffeeLibre](#coffeelibre)
	- [What?](#what)
	- [Breaking News (2014-02-27@11:58 CET)](#breaking-news-2014-02-27@1158-cet)
	- [Why?](#why)
	- [How?](#how)
		- [The Disappointing Hello World Example](#the-disappointing-hello-world-example)
		- [Same Basic Samples](#same-basic-samples)
		- [The Passive-Aggressive Treatment of Java in AOO](#the-passive-aggressive-treatment-of-java-in-aoo)
		- [Bootstrapping](#bootstrapping)
		- [Undo Transactions](#undo-transactions)
		- [XXXXXXXXX](#xxxxxxxxx)
		- [XXXXXXXXX](#xxxxxxxxx-1)
	- [Materials](#materials)
		- [Cell Properties](#cell-properties)
		- [The Global Object](#the-global-object)
		- [Helpful Links](#helpful-links)

> **Table of Contents**  *generated with [DocToc](http://doctoc.herokuapp.com/)*


## CoffeeLibre


### What?

**CoffeeLibre is an attempt to bring easier scripting with CoffeeScript to Apache OpenOffice (AOO).**

OpenOffice macros suffer from a number of problems:

* they are hard to access and maintain from within the OpenOffice GUI;
* writing macros with one of the several(!)* script editors provided by AOO is an unpleasant experience at best;
* writing AOO macros in JavaScript is possible, but results in incredibly convoluted code.

> *) I have no idea what it is that makes application authors assume they will be capable of writing a
> decent embedded code editor; maybe the thinking is that programs like Calc and Writer are 'inherently more
> complex' than 'a meak text editor' so once you're setting out to write something Calc, a code editor will
> come at very low additional costs. Then again, i've met people who insist on using Windows Notepad to edit
> their code...

You can achieve to remedy all of the above—*to a certain degree*—by

* keeping source files out of the AOO directory tree (creating symbolic links instead);
* authoring macros with your favorite text editor; and
* wrapping up pieces of functionality—especially annoyingly long-winded AOO-API incantations—into
  libraries (for which [src/coffeelibre.coffee](src/coffeelibre.coffee) may serve as an example).

### Breaking News (2014-02-27@11:58 CET)

JavaScript macros for AOO are run inside the [Rhino VM](https://www.mozilla.org/rhino/). Now there's a
slight chance this choice may be made an option: Not only is there [DynJS](http://dynjs.org/), "an
ECMAScript runtime for the JVM", there's [Nodyn](http://nodyn.io/), too, which describes itself as "a
node.js compatible framework, running on the JVM powered by the DynJS Javascript runtime running under
vert.x—the polyglot application platform and event bus".

People who are crazy / stupid enough to think seriously about forking AOO / writing an office suite / a
spreadsheet application / a **glorified multidimensional cellular automaton for numbers and texts that
allows for alternative grid and graph layouts such as triangular and polar grids** (which is what i'd
suggest to implement using [node-webkit](https://github.com/rogerwang/node-webkit)) should definitely take a
look.

### Why?

A while ago i had a fancy idea how to catalogue components of Chinese characters in a tree-like fashion; all
i wanted was a software that would be more amenable to try out this idea than a pure text format (which i
had tried). **I figured that using Unicode box drawing characters as connectors should be good enough to
visualize that tree, and that i could repurpose a spreadsheet—basically, a cellular automaton of sorts—to
help me allocate spaces and characters in a suitable fashion.**

**Unfortunately, there are not too many options when it comes to chosing a free and open spreadsheet
solution.** Online spreadsheet i've tried were often too clunky and restricted in their possibilities, and
MS Excel—while arguably the champion in this field—is neither free nor open. This leaves you with a choice
between one of the `/[a-z]+office\.org/` alternatives; more specifically,

*   http://openoffice.org
*   http://libreoffice.org
*   http://neooffice.org

—which are, essentially, clones of each other.

> Of these, NeoOffice gives the best impression on OSX; i will still use OpenOffice for the purposes of
> the present discussion.


I quickly found what i always find after opening an OpenOffice application:

* that its **interface** is **incredibly clunky, ugly, and backwards**;
* that its developers have a propensity to **hide functionalities in deeply nested
  menus, submenus, and subsubsubmenus**;
* that new versions are **unlikely to come up with anything new**;
* that it's quite **happy to crash** any given moment, and
* that especially Calc is **incredibly stupid / stubborn / incapable when it comes to doing any meaningful
  / consistent formatting** (with or without using named styles).

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
to apply the formatting. Since OpenOffice does have support for JavaScript, i decided to give it a try.

I succeeded in so far as i managed to draw a tree and leave most of the formatting chores to a macro
(rather, a family of methods deployed over several files). Here's a screenshot of the current result:

![result](https://raw.github.com/loveencounterflow/coffeelibre/master/art/Screen Shot 2014-02-26 at 20.31.51.png)

Observe how tree lines are drawn in grey, and Chinese characters (and components) have turquoise
backgrounds. What you cannot immediately see is that the Chinese stuff in this picture comes from three
different fonts—characters identified with prefixes `u-cjk-` and `u-cjk-xa` use Sun-ExtA.ttf, those with `u
-cjk-xb` use (a fork of) Sun-ExtB.ttf, and those marked with a `jzr-` use a font i produced with fontforge,
as those glyphs are not encoded in Unicode (as of v6.3).

**It would be quite an exacting task to get these formatting details right in a manual fashion, especially
since AOO Calc has a habit of not doing what you would expect it to do when moving cell contents or when
adding or deleting columns. Having bound the macro to a convenient keyboard shortcut, it was a snap to
simply select a given region of the spreadshit and hit `⌘Y` to get all the details right.**

### How?

#### The Disappointing Hello World Example

The OpenOffice folks are quite dexterous at producing a huge pile of documentation on their website (or at
least they were when they stopped updating a lot of the stuff years ago). They are also very good at
providing a Hello World example in no less than five language variants. They're also masters in letting the
matter rest at that; if you're in for some JS-in-AOO macro fun, this much is about how much you're going to
get from a readily available source. They're also lightning fast in updating their sample snippets. Let's
have a look at the code (on
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
likelyhood that it will crash on OSX, since there is a long-standing bug that keeps AOO macros from using
Java AWT stuff (google `openoffice osx awt` for this one).

There are a few points that merit your attention in these snippets:

* there's this strange `queryInterface` call which we'll discuss in a moment;
* while the examples are basically JavaScript, some of the vocabulary is non-standard; in particular,
  * `importClass`
  * `XSCRIPTCONTEXT`
  * `UnoRuntime`

  are not part of standard JavaScript but so-called 'host objects'.


#### Same Basic Samples

Intense use of search engines turned up various code snippets.*

> By far the most comprehensive archive of AOO JavaScript macros i've found is at http://openoffice3.web.fc2.com;
> although those pages are written in Japanese, they should still prove valuable for the neophyte.
> There's a mirror of the website at https://github.com/loveencounterflow/OpenOffice-macros.

I then set out to translate those snippets into CoffeeScript, isolate pertinent pieces of
functionality, and organize them into functions with meaningful names. **It's really very much a matter of
undoing the unholy mess that the AOO API is.** I mean, consider this code that essentially just gives you
an object that represents the current spreadsheet open in Calc:


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

#### The Passive-Aggressive Treatment of Java in AOO

It would appear that, being the smart guys they are, the AOO folks chose to embrace Java (and XML) to the
fullest. Of course, Java being Java, it does embrace static typing—a feature well known to get into the way
of free-and-easy programmers quite often. So why not abstract the bejeezes out of it?

![very abstract](https://raw.github.com/loveencounterflow/coffeelibre/master/art/extreme-abstracting.jpg)

> art by [Geek & Poke](http://geek-and-poke.com).—At this point
> in the program i always recommended to read up on Steve Yegge's classic,
> [Execution in the Kingdom of Nouns](http://steve-yegge.blogspot.de/2006/03/execution-in-kingdom-of-nouns.html);
> i can also heartily recommend to click through the file tree of
> https://github.com/EnterpriseQualityCoding/FizzBuzzEnterpriseEdition
> (courtesy http://abstractsingletonproxyfactorybean.com).

Basically, `XSCRIPTCONTEXT.getDocument()` gives you an object that represents *something* in a hand-waving
non-committal fashion, and *nothing in particular* at the same time. To get more specific, we still have to
'wrap' that `NotAnythingInParticular` blob with an interface.

**As far as i understand it, this is a weak /
obfuscated / helpless way of bringing a modicum of dynamism into that rigid beast that is Java—not unlike
what people do when implementing, say, JavaScript in Java.**

To quote from the docs (which are easily
retrievable from https://www.openoffice.org/api/docs/java/ref/com/sun/star/uno/UnoRuntime.html#queryInterface%28com.sun.star.uno.Type,%20java.lang.Object%29)

> [`UnoRuntime.queryInterface`] returns `null` in case the given UNO object does not support the given UNO
> interface type (or is itself `null`). Otherwise, a reference to a Java object implementing the Java
> interface type corresponding to the given UNO interface is returned. In the latter case, it is
> unspecified whether the returned Java object is the same as the given object, or is another facet of
> that UNO object.

Anyone feeling the urge to shout 'JUST GIMME THAT DARN OBJECT ALREADY' at this point?


#### Bootstrapping

In order to bring AOO JavaScript macro programming (which happens inside the [Rhino VM](https://www.mozilla.org/rhino/))
more into line with programming in NodeJS, we have to sort of 'bootstrap'
in order to get access to such important facilities as, for example, `require`ing modules and printing object
APIs to the command line. Here are the first few lines of `coffeelibre/src/main.coffee`, showing the
shortest boilerplate i felt able to come up with:

````coffeescript
#-----------------------------------------------------------------------------------------------------------
### NB in OpenOffice:   importClass          org.mozilla.javascript.Context ??? ###
### NB in LibreOffice:  importClass Packages.org.mozilla.javascript.Context ??? ###
importClass Packages.org.mozilla.javascript.Context
importClass Packages.org.mozilla.javascript.tools.shell.Global
#...........................................................................................................
# Why not simply provide a `Global` object? Because that would be too obvious???
GLOBAL = new Global Context.enter()
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
log                       = TRM.log.bind TRM
rpr                       = TRM.rpr.bind TRM
xray                      = TRM.xray.bind TRM
````

> I would've liked very much to get that problematic file locator out of the code, but didn't manage to;
> pull requests welcome!

What happens here is, in a nutshell, that we first avail ourselves of the `Context` and `Global` objects,
which we *both* need to construct a `GLOBAL` object; that `GLOBAL` object in turn has a method `readFile`.

We then `eval GLOBAL.readFile prefix + 'require.js'`, which gives us an approximate re-implementation of the
NodeJS `require` keyword (which lands in the global namespace). We then proceed to load a fair number of
Java classes that come under such delicious names as `com.sun.star.style.XStyleFamiliesSupplier` and
`com.sun.star.script.provider.MasterScriptProviderFactory` (but, alas, no `AbstractSingletonProxyFactoryBean`).

Next, we load some utility libraries: `TRM` is a dumbed-down version of
[`coffeenode-trm`](https://github.com/loveencounterflow/coffeenode-trm) and provides TeRMinal printout
methods, while `CHR`, `TEXT` and `TYPES` are copied without changes from
[`coffeenode-chr`](https://github.com/loveencounterflow/coffeenode-chr),
[`coffeenode-text`](https://github.com/loveencounterflow/coffeenode-text), and
[`coffeenode-types`](https://github.com/loveencounterflow/coffeenode-types), respectively (`TEXT` provides
string manipulation routines; `CHR` is all about Unicode character codepoints, and `TYPES` gives us
sane JS typechecking methods).


#### Undo Transactions

OpenOffice not only provides multi-level undo, the undo functionality is also made available to scripts.
Even better, multiple atomic actions can be grouped together so they appear as single steps in the undo
actions lists. This means you can write complex macros that e.g. format lots and lots of cells in a
spreadsheet and you can still undo the entire transaction with a single `⌘Z` (or `^Z`).

**Initiating an AOO undo transaction is as easy as saying "OK Uno runtime, please query interface X undo
manager supplier doc, get undo manager"**... seriously, though:

````coffeescript
#-----------------------------------------------------------------------------------------------------------
@get_undo_manager = ( doc ) ->
  ### Note: this could very well be made a private method. ###
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
````

> **Note**: i've decided to take `doc = COFFEELIBRE.get_current_doc()` as the first parameter in all of my
> CoffeeLibre methods, whether it is strictly by a given method or not; this is similar to the explicit
> `self` in Python and a consequence of the Data-Centric, Library-Oriented programming methodology that i
> tend to write all my stuff in.

Undo transactions are simple to use with this setup; just write

````coffeescript
  COFFEELIBRE.step doc, 'some descriptive title', -> do something here
````


#### XXXXXXXXX

Each macro needs to have a `parcel-descriptor.xml` in its folder; you'll find one under
`coffeelibre/lib/parcel-descriptor.xml` (since that is the folder we linked into the AOO scripts folder).
Enjoy:

````xml
<parcel language="JavaScript" xmlns:parcel="scripting.dtd">
    <script language="JavaScript">
        <locale lang="en">
            <displayname value="CoffeeLibre Demo"/>
            <description>
                A Demo of CoffeeLibre.
            </description>
        </locale>
        <functionname value="main.js"/>
        <logicalname value="Main.JavaScript"/>
    </script>
</parcel>
````

We've basically wasted a lot of keystrokes to reassure OpenOffice—**five times**, no less—that what we have
here (yes, in AOO's own `Scripts/javascript` folder) is a JavaScript macro indeed. Of course, a simple
convention-over-configuration agreement would have obliterated the need for this configuration file, but
what's not to like about writing redundant XML?


#### XXXXXXXXX

While there is (on OSX) a folder `~/Library/Application Support/OpenOffice/4/user/Scripts` (which may
or may not work for scripting as described here), i chose to use
`/Applications/OpenOffice.app/Contents/share/Scripts/javascript` as location to link a folder that
contains my scripts—more specifically, you could simply

````bash
cd ~
git clone https://github.com/loveencounterflow/coffeelibre.git
cd /Applications/OpenOffice.app/Contents/share/Scripts/javascript
ln -s ~/coffeelibre CoffeeLibreDemo # or whatever name
cd -
````

> Notice that on OSX, applications like OpenOffice are distributed as `*.app` bundles, which are really
> directories that behave as files under many circumstances; you can still `cd` into an `*.app`, but Finder
> will show them as files (which are run on double-click rather than entered into); PathFinder does have a
> `View / Show Package Contents` option in its menu two switch between normal and nerd modes.

Run

````bash
coffee --watch --output ~/coffeelibre/lib --compile ~/coffeelibre/src
````

in a terminal window to ensure all your changes in your CoffeeScript sources will be immediately reflected
in the JavaScript transpilation targets (fortunately, AOO does not appear to cache macros, so its easy to
always work with up-to-date sources).

![Accessing the AOO Macro Administration Facility (a)](https://raw.github.com/loveencounterflow/coffeelibre/master/art/Screen Shot 2014-02-26 at 15.48.58.png "Accessing the AOO Macro Administration Facility (a)")
*Accessing the AOO Macro Administration Facility (a)*

![Accessing the AOO Macro Administration Facility (b)](https://raw.github.com/loveencounterflow/coffeelibre/master/art/Screen Shot 2014-02-26 at 16.40.52.png "Accessing the AOO Macro Administration Facility (b)")
*Accessing the AOO Macro Administration Facility (b)*

![Assigning a keyboard shortcut to your macro](https://raw.github.com/loveencounterflow/coffeelibre/master/art/Screen Shot 2014-02-26 at 16.24.42.png "Assigning a keyboard shortcut to your macro")
*Assigning a keyboard shortcut to your macro*



### Materials


#### Cell Properties

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


#### The Global Object

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


#### Helpful Links

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

