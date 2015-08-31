//Content specific to Word generation

relTypes.numbering = relTypePrefix + "numbering";
relTypes.aFChunk = relTypePrefix + "aFChunk";
contTypes.numbering = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

function _valBuild(e, v) {
    return function(xdoc) {
        return xdoc.el(e, {"w:val": v});
    };
}

function DOCXBuilder (b64Template, title, created, creator) {
    XPkg.call(this, b64Template, title, created, creator);
    this.doc = this.getPart("/word/document.xml");
    this.body = this.doc.one("/w:document/w:body");
    this.sectPr = this.body.one("./w:sectPr");
    var es = this.body.all("./*");
    for (var i = 0; i < es.length; i++) {
        es[i].remove();
    }
    if (this.sectPr) {
        this.body.add(this.sectPr);
    }

    this.rStyles = {};
    this.pStyles = {};
    this.rStyles.Italic = this.rStyleItalic;
    this.rStyles.Bold = this.rStyleBold;
    this.rStyles.Underline = this.rStyleUnderline;

    var styles = this.getPart("/word/styles.xml").all("w:styles/w:style");
    for (i = 0; i < styles.length; i++) {
      var styleType = styles[i].getAttr("w:type");
      var styleId = styles[i].getAttr("w:styleId");
      if (styleType === "paragraph") {
        this.pStyles[styleId] = _valBuild("w:pStyle", styleId);
      } else if (styleType === "character") {
        this.rStyles[styleId] = _valBuild("w:rStyle", styleId);
      }
    }
    this.undoPart("/word/styles.xml");
}
DOCXBuilder.prototype = {
    mimetype : "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    rStyleItalic : function(xdoc) {return xdoc.el("w:i"); },
    rStyleBold   : function (xdoc) {return xdoc.el("w:b"); },
    rStyleUnderline : _valBuild("w:u", "single"),
    rStyleHyperlink : _valBuild("w:rStyle", "Hyperlink"),
    pStyleNoSpacing : _valBuild("w:pStyle", "NoSpacing"),
    pStyleListParagraph : _valBuild("w:pStyle", "ListParagraph"),
    pStyle : function (name) {
        return this.pStyles[name] || _valBuild("w:pStyle", name);
    },
    rStyle : function (name) {
      return this.rStyles[name] || _valBuild("w:rStyle", name);
    },
    _levelIndent : function (level) {
        return 720 * (1 + (level ||0));
    },
    makeUl : function () {
        return this.makeList([{fmt:"bullet", fnt:"Symbol", txt:"\xB7"}, {fmt:"bullet", fnt:"Courier New", txt:"o"}, {fmt:"bullet", fnt:"Wingdings", txt:"\xA7"}], 9);
    },
    makeOl : function () {
        return this.makeList([{fmt:"decimal"}, {fmt:"lowerLetter"}, {fmt:"lowerRoman"}], 9);
    },
    makeOlNum : function () {
        return this.makeList([{fmt:"decimal"}], 9);
    },
    makeOlHier : function () {
        return this.makeList([{fmt:"decimal", hier:true}], 9);
    },
    /*jshint -W071 */
    makeList : function (levels, depth) {
        var numPart = this.getPart("/word/numbering.xml");
        if (!numPart) {
            numPart = this.addPart(contTypes.numbering, "/word/numbering.xml", docTemplate("w:numbering", null, "w"));
            this.doc.addRelationship(this.doc.nextRelationshipId(), relTypes.numbering, "numbering.xml", "Internal");
        }
        var nsid = Math.random().toString(16).substring(2, 10).toUpperCase();
        var anumid = numPart.nextId("//w:abstractNum", "w:abstractNumId", 0);
        var listid = numPart.nextId("//w:num", "w:numId", 1);
        var anum = numPart.el("w:abstractNum").setAttr("w:abstractNumId", anumid);
        anum.add(numPart.el("w:nsid").setAttr("w:val", nsid));
        anum.add(numPart.el("w:multiLevelType").setAttr("w:val", "hybridMultilevel"));
        var lvl, level, lvlText, hLvlText, idt;
        depth = depth || levels.length;
        for (var i = 0; i < depth; i++) {
            level = levels[i % levels.length];
            idt = this._levelIndent(i);
            lvl = numPart.el("w:lvl").setAttr("w:ilvl", i);
            lvl.add(numPart.el("w:start").setAttr("w:val", "1"));
            lvl.add(numPart.el("w:numFmt").setAttr("w:val", level.fmt));
            if (level.fmt === "bullet") {
                lvlText = level.txt || " ";
                hLvlText = "";
            } else {
                lvlText = "%" + (i+1) + ".";
                hLvlText += lvlText;
            }
            lvl.add(numPart.el("w:lvlText").setAttr("w:val", level.heir ? hLvlText : lvlText));
            lvl.add(numPart.el("w:lvlJc").setAttr("w:val", "left"));
            lvl.add(numPart.el("w:pPr").add(
                numPart.el("w:ind", {"w:hanging": "360", "w:left": idt})));
            if (level.fnt) {
                lvl.add(numPart.el("w:rPr").add(
                    numPart.el("w:rFonts", {"w:hint": "default", "w:hAnsi": level.fnt, "w:ascii": level.fnt})));
            }
            anum.add(lvl);
        }
        var num = numPart.el("w:num").setAttr("w:numId", listid);
        num.add(numPart.el("w:abstractNumId").setAttr("w:val", anumid));
        numPart.xdoc.root().add(anum);
        numPart.xdoc.root().add(num);
        return listid;
    },
    /*jshint +W071 */
    docLine : function (html, pStyle, rStyle) {
        var seg = convertHtml(html);
        if (seg.length === 1 && seg[0].type === "paragraph") {
            seg = seg[0];
            this._paragraph(this.body, this.sectPr, seg, pStyle, rStyle);
        } else {
            throw "Content is too complex";
        }
    },
    docContent : function (html) {
        var blocks = convertHtml(html);
        this._blocks(this.body, this.sectPr, blocks);
    },
    docChunk : function (html) {
        html = "<html><head></head><body>" + cleanupHtml(html) + "</body></html>";
        var rId = this.doc.nextRelationshipId();
        var uri = "/word/chunk"+rId+".html";
        this.addFile("text/html", uri, html);
        this.doc.addRelationship(rId, relTypes.aFChunk, uri, "Internal");
        var chunk = this.doc.el("w:altChunk").setAttr("r:id", rId);
        this.body.add(chunk, this.sectPr);
    },
    /*jshint -W071 */
    /*jshint -W072 */
    _block : function (parent, before, block, listid, level) {
        var pPrs;
        if (block.type === "table") {
            this._table(parent, before, block);
        } else if (block.type === "list") {
            if (!listid) {
                listid = block.listType === "ul" ? this.makeUl() : this.makeOlHier();
                level = 0;
            } else {
                level += 1;
            }
            this._blocks(parent, before, block.blocks, listid, level);
        } else if (block.type === "list-item" && listid) {
            pPrs = this.doc.el("w:numPr");
            pPrs.add(_valBuild("w:ilvl", level)(this.doc.xdoc));
            pPrs.add(_valBuild("w:numId", listid)(this.doc.xdoc));
            this._paragraph(parent, before, block, null, null, pPrs);
        } else if (listid) {
            pPrs = this.doc.el("w:ind").setAttr("w:left", this._levelIndent(level));
            this._paragraph(parent, before, block, this.pStyleListParagraph, null, pPrs);
        } else {
            this._paragraph(parent, before, block);
        }
    },
    _blocks : function (parent, before, blocks, listid, level) {
        for (var i = 0; i < blocks.length; i++) {
            this._block(parent, before, blocks[i], listid, level);
        }
    },
    _runs : function (parent, runs, rStyle) {
        var run, r, rPr, t;
        for (var i = 0; i < runs.length; i++) {
            run = runs[i];
            r = this.doc.el("w:r");
            if (run.b || run.i || run.u || run.color || run.link || rStyle) {
                rPr = this.doc.el("w:rPr");
                if (run.link) {
                    if (rStyle) {
                        rPr.add(this.rStyleUnderline(this.doc.xdoc));
                    } else {
                        rPr.add(this.rStyleHyperlink(this.doc.xdoc));
                    }
                }
                if (rStyle) {
                    rPr.add(rStyle(this.doc.xdoc));
                }
                if (run.b) {
                    rPr.add(this.rStyleBold(this.doc.xdoc));
                }
                if (run.i) {
                    rPr.add(this.rStyleItalic(this.doc.xdoc));
                }
                if (run.u) {
                    rPr.add(this.rStyleUnderline(this.doc.xdoc));
                }
                if (run.color) {
                    rPr.add(_valBuild("w:color", run.color)(this.doc.xdoc));
                }
                r.add(rPr);
            }
            t = this.doc.el("w:t", run.text);
            if (/^\s+|\s+$/.test(run.text)) {
                t.setAttr("xml:space", "preserve");
            }
            r.add(t);
            if (run.link) {
                var rId = this.doc.nextRelationshipId();
                this.doc.addRelationship(rId, relTypes.hyperlink, run.link, "External");
                r = this.doc.el("w:hyperlink").setAttr("r:id", rId).add(r);
            }
            parent.add(r);
        }
    },
    _paragraph : function (parent, before, paragraph, pStyle, rStyle, pPrs) {
        var p, pPr;
        p = this.doc.el("w:p");
        if (pStyle || rStyle || pPrs) {
            pPr = this.doc.el("w:pPr");
            if (pStyle) {
                pPr.add(pStyle(this.doc.xdoc));
            }
            if (rStyle) {
                pPr.add(this.doc.el("w:rPr").add(rStyle(this.doc.xdoc)));
            }
            if (paragraph.hr) {
                var hr = this.doc.el("w:pBdr");
                hr.add(this.doc.el("w:bottom").setAttr("w:color", "auto").setAttr("w:space", "1").setAttr("w:sz", "6").setAttr("w:val", "single"));
                pPr.add(hr);
            }
            if (pPrs) {
                pPr.add(pPrs);
            }
            p.add(pPr);
        }
        this._runs(p, paragraph.runs, rStyle);
        parent.add(p, before);
    },
    _table : function (parent, before, table) {
        var i, j, row, cell, r, c, e, cspan, rspan, cind;
        var tbl = this.doc.el("w:tbl");
        /*
        e = this.doc.el("w:tblPr");
        e.add(this.doc.el("w:tblStyle").setAttr("w:val", "TableGrid"));
        e.add(this.doc.el("w:tblW").setAttr("w:w", "0").setAttr("w:type", "auto"));
        e.add(this.doc.el("w:tblBorders").add(
            this.doc.el("w:insideV", {"w:val": "single", "w:themeTint": "BF", "w:themeColor": "accent1", "w:color": "7BA0CD", "w:space": "0", "w:sz": "8"});
        ));
        e.add(this.doc.el("w:tblLook").setAttr("06A0"));
        tbl.add(e);

        e = this.doc.el("w:tblGrid");
        var cwidth = toInt(10296/cols);
        for (i = 0; i < cols; i++) {
            e.add(this.doc.el("w:gridCol").setAttr("w:w", cwidth);
        }
        tbl.add(e);
        */

        var colspans = {};
        var spanned = function () {
            while (colspans[cind]) {
                cspan = colspans[cind][1] || 1;
                rspan = colspans[cind][0];

                c = this.doc.el("w:tc");
                e = this.doc.el("w:tcPr");
                //e.add(this.doc.el("w:tcW", {"w:w": cwidth * cspan, "w:type": "dxa"}));
                if (cspan > 1) {
                    e.add(this.doc.el("w:gridSpan").setAttr("w:val", cspan));
                }
                e.add(this.doc.el("w:vMerge"));
                c.add(e);
                c.add(this.doc.el("w:p"));
                if (colspans[cind][0]) {
                    colspans[cind][0] -= 1;
                } else {
                    delete colspans[cind];
                }
                r.add(c);
                cind += cspan;
            }
        };

        for (i = 0; i < table.blocks.length; i++) {
            row = table.blocks[i];
            r = this.doc.el("w:tr");
            cind = 0;
            for (j = 0; j < row.blocks.length; j++) {
                spanned();

                c = this.doc.el("w:tc");
                cspan = cell.colspan || 1;
                rspan = cell.rowspan;
                cell = row.blocks[j];
                e = this.doc.el("w:tcPr");
                //e.add(this.doc.el("w:tcW", {"w:w": cwidth * cspan, "w:type": "dxa"});
                if (cspan > 1) {
                    e.add(this.doc.el("w:gridSpan").setAttr("w:val", cspan));
                }
                if (rspan > 1) {
                    colspans[cind] = [rspan-1, cspan || 1];
                    e.add(this.doc.el("w:vMerge").setAttr("w:val", "restart"));
                }
                c.add(e);
                this._blocks(c, cell.blocks);
                r.add(c);
                cind += cspan;
            }
            spanned();
            tbl.add(r);
        }

        parent.add(tbl, before);
    }
};
/*jshint +W071 */
/*jshint +W072 */

(function () {
    for (var f in XPkg.prototype) {
        if (XPkg.prototype.hasOwnProperty(f)) {
            DOCXBuilder.prototype[f] = XPkg.prototype[f];
        }
    }
})();

OpenXmlBuilder.DOCXBuilder = DOCXBuilder;
