// Generated by CoffeeScript 1.6.3

/* NB in OpenOffice:   importClass          org.mozilla.javascript.Context ??? */


/* NB in LibreOffice:  importClass Packages.org.mozilla.javascript.Context ??? */

(function() {
  var CHR, CL, GLOBAL, TEXT, TRM, TYPES, font_name_by_rsg, log, prefix, rpr, xray;

  importClass(Packages.org.mozilla.javascript.Context);

  importClass(Packages.org.mozilla.javascript.tools.shell.Global);

  GLOBAL = new Global(Context.enter());

  prefix = '/Applications/OpenOffice.app/Contents/share/Scripts/javascript/CoffeeLibreDemo/';


  /* Globals */

  eval(GLOBAL.readFile(prefix + 'require.js'));

  eval(GLOBAL.readFile(prefix + 'import-classes.js'));

  require.prefix = prefix;


  /* Locals */

  TRM = require('coffeelibre-trm');

  CHR = require('coffeenode-chr');

  TEXT = require('coffeenode-text');

  TYPES = require('coffeenode-types');

  font_name_by_rsg = require('font-name-by-rsg');

  CL = require('coffeelibre');

  log = TRM.log.bind(TRM);

  rpr = TRM.rpr.bind(TRM);

  xray = TRM.xray.bind(TRM);

  this.get_styles = function() {
    var cell, doc, properties, psi, pv, sheet;
    log('');
    log('--------------------------------------------------------------------------------------------');
    log('styles');
    log((new Date()).toString());
    doc = CL.get_current_doc();
    sheet = CL.get_current_sheet(doc);
    cell = CL.get_cell(sheet, 0, 0);
    pv = UnoRuntime.queryInterface(XPropertySet, cell);
    psi = pv.getPropertySetInfo();
    properties = psi.getProperties();
    TRM.dir('pv', pv);
    TRM.dir('psi', psi);
    TRM.dir('properties[ 0 ]', properties[0]);
    log('' + psi.getPropertyByName('CellStyle'));
    return pv.setPropertyValue('CellStyle', 'glyph');
  };

  this.main = function() {
    var doc;
    log('');
    log('--------------------------------------------------------------------------------------------');
    log('format tree');
    log((new Date()).toString());
    doc = CL.get_current_doc();
    CL.step(doc, 'format tree', function() {
      return this.format_tree(doc);
    });
    return null;
  };

  this.font_name_from_chr_info = function(chr_info, fallback) {
    var R, rsg;
    rsg = chr_info['rsg'];
    R = font_name_by_rsg[rsg];
    if (R == null) {
      if (fallback !== void 0) {
        return fallback;
      }
      throw new Error("unable to find a suitable font for " + chr_info['fncr'] + " " + chr_info['chr']);
    }
    return R;
  };

  this.format_tree = function(doc) {
    var cell_ref, chr_info, cid, fallback_font_name, fncr, fncr_cell, font_name, format_options, format_options_by_cell_type, is_cjk, rsg, selection, sheet, source_cell, source_text, x, x0, x1, x_fncr, xy0, xy1, y, y0, y1, _i, _j, _ref, _ref1, _ref2, _ref3;
    doc = CL.get_current_doc();
    sheet = CL.get_current_sheet(doc);
    selection = CL.get_current_selection(doc);
    fallback_font_name = 'Sun-ExtA';
    format_options_by_cell_type = {
      empty: {
        'cell-style-name': 'empty'
      },
      missing: {
        'cell-style-name': 'missing'
      },
      glyph: {
        'cell-style-name': 'glyph'
      },
      tree: {
        'cell-style-name': 'tree'
      },
      fncr: {
        'cell-style-name': 'fncr'
      },
      strokecode: {
        'cell-style-name': 'strokecode'
      }
    };
    _ref = CL.get_current_selection(doc), xy0 = _ref[0], xy1 = _ref[1];
    _ref1 = [xy0, xy1], (_ref2 = _ref1[0], x0 = _ref2[0], y0 = _ref2[1]), (_ref3 = _ref1[1], x1 = _ref3[0], y1 = _ref3[1]);
    x_fncr = x1 + 2;
    log("current selection: " + (CL.range_ref_from_xy(xy0, xy1)));
    for (y = _i = y0; y0 <= y1 ? _i <= y1 : _i >= y1; y = y0 <= y1 ? ++_i : --_i) {
      fncr_cell = CL.get_cell(sheet, x_fncr, y);
      for (x = _j = x0; x0 <= x1 ? _j <= x1 : _j >= x1; x = x0 <= x1 ? ++_j : --_j) {
        source_cell = CL.get_cell(sheet, x, y);
        source_text = CL.get_cell_text(source_cell);
        cell_ref = CL.cell_ref_from_xy(x, y);

        /* empty cells: */
        if (source_text.length === 0) {
          CL.format_cell(source_cell, format_options_by_cell_type['empty']);
          continue;
        }

        /* empty cells: */
        if (source_text === '?') {
          CL.format_cell(source_cell, format_options_by_cell_type['missing']);
          continue;
        }

        /* non-empty cells: */
        cid = CHR.as_cid(source_text, {
          input: 'xncr'
        });
        chr_info = CHR.analyze(cid);
        fncr = chr_info['fncr'];
        rsg = chr_info['rsg'];
        fncr = fncr.replace(/^u-pua-/, 'jzr-fig-');
        rsg = rsg.replace(/^u-pua$/, 'jzr-fig');
        is_cjk = /^(u-cjk|jzr-fig|u-pua)/.test(rsg);
        if (rsg === 'jzr-fig') {
          source_cell.setFormula(CHR._unicode_chr_from_cid(cid));
        }

        /* glyph cells: */
        if (is_cjk) {
          fncr_cell.setFormula(fncr);
          CL.format_cell(fncr_cell, format_options_by_cell_type['fncr']);
          format_options = format_options_by_cell_type['glyph'];
          font_name = this.font_name_from_chr_info(chr_info, fallback_font_name);
          format_options['font-name'] = font_name;
          CL.format_cell(source_cell, format_options);
          continue;
        }

        /* tree cells: */
        CL.format_cell(source_cell, format_options_by_cell_type['tree']);
      }
    }
    return null;
  };

  this.main();

}).call(this);