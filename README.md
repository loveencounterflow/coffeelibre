coffeelibre
===========

Scripting Libre(Open/Neo)Office with CoffeeScript


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
