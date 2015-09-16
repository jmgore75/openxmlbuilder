//Content specific to PowerPoint generation

relTypes.slideMaster = relTypePrefix + "slideMaster";
relTypes.slideLayout = relTypePrefix + "slideLayout";
relTypes.slide = relTypePrefix + "slide";
contTypes.slide = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

function PPTXContentWriter(slide, txBody) {
  this.slide = slide;
  this.txBody = txBody;
  var ps = this.txBody.all("./a:p");
  for (var j = 0; j < ps.length; j++) {
    ps[j].remove();
  }
  this.listStack = [];
  this.listItem = false;
}
PPTXContentWriter.prototype = {
  block: function(block) {
    if (block.type === "list") {
      this.listStack.unshift(block.listType || "ul");
      this.blocks(block.blocks);
      this.listStack.shift();
    } else if (block.type === "table") {
      this.inlineTable(block);
    } else if (block.type === "list-item") {
      this.listItem = true;
      this.blocks(block.blocks);
      this.listItem = false;
    } else {
      this.paragraph(block);
    }
  },
  blocks: function(blocks) {
    for (var i = 0; i < blocks.length; i++) {
      this.block(blocks[i]);
    }
  },
  runs: function(parent, runs) {
    var run, r, rPr, t, rId;
    var lastLink;
    for (var i = 0; i < runs.length; i++) {
      run = runs[i];
      r = this.slide.el("a:r");
      lastLink = run.link;
      if (run.b || run.i || run.u || run.color || run.link) {
        rPr = this.slide.el("a:rPr");
        if (run.b) {
          rPr.setAttr("b", "1");
        }
        if (run.i) {
          rPr.setAttr("i", "1");
        }
        if (run.u) {
          rPr.setAttr("u", "sng");
        }
        if (run.color) {
          rPr.add(this.slide.el("a:solidFill").add(this.slide.el("a:srgbClr", {
            "val": run.color
          })));
        }
        if (run.link) {
          rId = this.slide.nextRelationshipId();
          this.slide.addRelationship(rId, relTypes.hyperlink, run.link, "External");
          rPr.add(this.slide.el("a:hlinkClick", {
            "r:id": rId
          }));
        }
        r.add(rPr);
      }
      t = this.slide.el("a:t", run.text);
      //if (/^\s+|\s+$/.test(run.text)) {
      //    t.setAttr("xml:space", "preserve");
      //}
      r.add(t);
      parent.add(r);
    }
    if (lastLink) {
      parent.add(this.slide.el("a:endParaRPr").add(this.slide.el("a:hlinkClick", {
        "r:id": ""
      })));
    }
  },
  paragraph: function(paragraph, indent) {
    var runs = paragraph.runs || [];
    var level = this.listStack.length;
    var listType;
    if (indent || level) {
      if (level > 0) {
        if (this.listItem) {
          listType = this.listStack[0];
          if (!runs.length) {
            runs = [{
              text: " "
            }];
          }
          this.listItem = false;
        }
      }
      if (indent) {
        level += indent;
      }
    }
    var pPr = this.slide.el("a:pPr");
    if (level > 1) {
      pPr.setAttr("lvl", level - 1);
    }
    if (listType === "ol") {
      pPr.add(this.slide.el("a:buAutoNum", {
        "type": "arabicPeriod"
      }));
    } else if (!listType) {
      pPr.add(this.slide.el("a:buNone"));
    }
    var p = this.slide.el("a:p");
    p.add(pPr);
    if (!listType && level > 0) {
      p.add(this.slide.el("a:r").add(this.slide.el("a:t", "\t")));
    }
    this.runs(p, runs);
    this.txBody.add(p);
  },
  inlineTable: function(table) {
    var i, j, cind, cspan, rspan, row, cell;
    var colspans = {};
    var spanned = function() {
      while (colspans[cind]) {
        cspan = colspans[cind][1] || 1;
        rspan = colspans[cind][0];

        if (colspans[cind][0]) {
          colspans[cind][0] -= 1;
        } else {
          delete colspans[cind];
        }
        cind += cspan;
      }
    };

    this.paragraph({
      runs: [{
        u: true,
        i: true,
        text: "Table"
      }]
    });
    this.listStack.unshift("ul");
    for (i = 0; i < table.blocks.length; i++) {
      row = table.blocks[i];
      this.listItem = true;
      this.paragraph({
        runs: [{
          u: true,
          i: true,
          text: "Row " + (row + 1)
        }]
      });
      this.listStack.unshift("ul");
      cind = 0;
      for (j = 0; j < row.blocks.length; j++) {
        spanned();
        cell = row.blocks[j];
        cspan = cell.colspan || 1;
        rspan = cell.rowspan;
        this.listItem = true;
        this.paragraph({
          runs: [{
            i: true,
            text: "Column" + (cind + 1)
          }]
        });
        this.blocks(cell);
        if (rspan > 1) {
          colspans[cind] = [rspan - 1, cspan || 1];
        }
        cind += cspan;
      }
      spanned();
      this.listStack.shift();
    }
    this.listStack.shift();
  }
};

function PPTXBuilder(b64Template, title, created, creator) {
  XPkg.call(this, b64Template, title, created, creator);

  this.presentation = this.getPart("/ppt/presentation.xml");
  this.sldIdLst = this.presentation.one("/p:presentation/p:sldIdLst");
  this.slideCount = this.sldIdLst.all("./p:sldId").length;

  this.app = this.getPart("/docProps/app.xml").xdoc;
  this.slidesProp = this.app.one("/ep:Properties/ep:Slides");
  this.slidesTitlesProp = this.app.all("/ep:Properties/ep:HeadingPairs/vt:vector/vt:variant/vt:i4");
  this.slidesTitlesProp = this.slidesTitlesProp[this.slidesTitlesProp.length - 1];
  this.titlesVector = this.app.one("/ep:Properties/ep:TitlesOfParts/vt:vector");

  this.removeAllSlides();
  this.slideLayouts = this._findLayouts();
}
PPTXBuilder.prototype = {
  mimetype: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  addSlide: function(title, slideLayout) {
    slideLayout = slideLayout || this.slideLayouts["Title and Content"];
    this.slideCount++;
    var slideNum = this.slideCount;
    var sldId = slideNum + 255;
    title = title || "Slide " + slideNum;

    var slideUri = "slides/slide" + slideNum + ".xml";
    var slide = this.addPart(contTypes.slide, "/ppt/" + slideUri, slideLayout.template);

    slide.addRelationship(slide.nextRelationshipId(), relTypes.slideLayout, slideLayout.relUri, "Internal");
    var rId = this.presentation.nextRelationshipId();
    this.presentation.addRelationship(rId, relTypes.slide, slideUri, "Internal");

    this.sldIdLst.add(this.presentation.el("p:sldId", {
      "r:id": rId,
      "id": sldId
    }));
    this.slidesProp.setValue(slideNum);
    this.slidesTitlesProp.setValue(slideNum);
    this.titlesVector.add(this.app.el("vt:lpstr", title));
    this.titlesVector.setAttr("size", slideNum + 1);

    return slide;
  },
  contentSlide: function(content, slideLayout) {
    var title = content.title || content["Title 1"];
    var slide = this.addSlide(title, slideLayout);

    var shapes = slide.all("//p:sp");
    for (var i = 0; i < shapes.length; i++) {
      var spId = shapes[i].one("./p:nvSpPr/p:cNvPr").getAttr("name");
      var html = content[spId];
      if (html) {
        var blocks = convertHtml(html);
        if (blocks.length === 1 && blocks[0].type === "table") {
          this.pptTable(slide, shapes[i], blocks[0]);
        } else {
          this.pptContent(slide, shapes[i], blocks);
        }
      }
    }
  },
  removeAllSlides: function() {
    var rels = this.presentation.getRelationshipsByRelationshipType(relTypes.slide);
    var rel, part;
    for (var i = 0; i < rels.length; i++) {
      rel = rels[i];
      part = this.getPart(this.fullPath(rel.getAttr("Target"), this.presentation.path));
      this.presentation.deleteRelationship(rel.relationshipId);
      part.deleted = true;
    }

    this.sldIdLst = this.presentation.one("/p:presentation/p:sldIdLst");
    var es = this.sldIdLst.all("./p:sldId");
    for (i = 0; i < es.length; i++) {
      es[i].remove();
    }

    this.slidesProp.setValue("0");
    this.slidesTitlesProp.setValue("0");
    this.titlesVector.setAttr("size", "1");
    es = this.titlesVector.all("./vt:lpstr");
    for (i = 1; i < es.length; i++) {
      es[i].remove();
    }

    this.slideCount = 0;
  },
  /*jshint -W071 */
  /*jshint -W072 */
  _findLayouts: function() {
    var slideLayouts = {};
    var slideLayoutObjects = this.zip.folder("ppt/slideLayouts").file(/^slideLayout\d+\.xml$/);

    for (var j = 0; j < slideLayoutObjects.length; j++) {
      var slo = slideLayoutObjects[j];
      var layoutXmlName = slo.name.slice(17);
      var slideLayoutId = parseInt(layoutXmlName.slice(11, -4), 10);
      var fullUri = "/" + slo.name;
      var relUri = "../slideLayouts/" + layoutXmlName;

      var slideLayout = new XDoc(slo.asText());
      var masterPath = new XDoc(this.zip.file("ppt/slideLayouts/_rels/" + layoutXmlName + ".rels").asText())
        .one("/rel:Relationships/rel:Relationship[@Type='" + relTypes.slideMaster + "']").getAttr("Target");
      var slideMaster = this.getPart(this.fullPath(masterPath, fullUri));
      var layoutSections = [];

      var sld = slideLayout.root();
      sld.removeAttr("preserve");
      sld.removeAttr("type");
      var layoutName = sld.one("p:cSld").getAttr("name");
      sld.one("p:cSld").removeAttr("name");

      var hf = slideMaster.one("//p:hf");
      var shapes = sld.all("//p:sp");
      var ph, phType;
      for (var i = 0; i < shapes.length; i++) {
        if (hf) {
          ph = shapes[i].one("./p:nvSpPr/p:nvPr/p:ph");
          if (ph) {
            phType = ph.getAttr("type");
            if (phType && hf.getAttr(phType) === "0") {
              shapes[i].remove();
              continue;
            }
          }
        }
        var spId = shapes[i].one("./p:nvSpPr/p:cNvPr").getAttr("name");
        if (spId) {
          layoutSections.push(spId);
        }
      }
      var sldTemplate = slideLayout.toXmlString();
      sldTemplate = sldTemplate.replace(/(<\/?p:sld)Layout/g, "$1");

      slideLayouts[layoutName] = {
        name: layoutName,
        index: slideLayoutId,
        fileName: layoutXmlName,
        fullUri: fullUri,
        relUri: relUri,
        sections: layoutSections,
        template: sldTemplate
      };
    }
    return slideLayouts;
  },
  _table: function(slide, parent, table) {
    var i, j, k, row, cell, r, c, e, cspan, rspan, cind;
    var tbl = slide.el("a:tbl");
    //a:tblPr

    var phTxbody = function() {
      return slide.el("a:txBody").add(slide.el("a:bodyPr")).add(slide.el("a:lstStyle"));
    };

    var phTc = function() {
      return slide.el("a:tc").add(
        phTxbody().add(slide.el("a:p"))
      ).add(slide.el("a:tcPr"));
    };

    var colspans = {};
    var spanned = function() {
      while (colspans[cind]) {
        cspan = colspans[cind][1] || 1;
        rspan = colspans[cind][0];

        var c = phTc().setAttr("vMerge", 1);
        r.add(c);
        if (cspan > 1) {
          c.setAttr("gridSpan", cspan);
          for (k = 1; k < cspan; k++) {
            r.add(phTc().setAttr("hMerge", 1).setAttr("vMerge", 1));
          }
        }
        if (colspans[cind][0]) {
          colspans[cind][0] -= 1;
        } else {
          delete colspans[cind];
        }
        cind += cspan;
      }
    };

    for (i = 0; i < table.blocks.length; i++) {
      row = table.blocks[i];
      r = slide.el("a:tr");
      //.setAttr("h", height);
      cind = 0;
      for (j = 0; j < row.blocks.length; j++) {
        spanned();

        cell = row.blocks[j];
        cspan = cell.colspan || 1;
        rspan = cell.rowspan;
        c = slide.el("a:tc");
        r.add(c);
        if (cspan > 1) {
          c.setAttr("gridSpan", cspan);
          for (k = 1; k < cspan; k++) {
            r.add(phTc().setAttr("hMerge", 1));
          }
        }
        if (rspan > 1) {
          colspans[cind] = [rspan - 1, cspan || 1];
        }

        e = phTxbody();
        c.add(e);
        var writer = new PPTXContentWriter(slide, e);
        writer.blocks(cell.blocks);
        e = slide.el("a:tcPr");
        c.add(e);
        cind += cspan;
      }
      spanned();
      tbl.add(r);
    }

    parent.add(tbl);
  },
  pptContent: function(slide, shape, blocks) {
    var writer = new PPTXContentWriter(slide, shape.one("./p:txBody"));
    writer.blocks(blocks);
  },
  pptLine: function(slide, shape, block) {
    var writer = new PPTXContentWriter(slide, shape.one("./p:txBody"));
    if (block.type === "paragraph") {
      writer.paragraph(block);
    } else {
      throw "Block is not a paragraph";
    }
  },
  pptTable: function(slide, shape, block) {
    var graphicFrame = slide.el("p:graphicFrame");
    var a, b, c, i;
    a = slide.el("p:nvGraphicFramePr");
    graphicFrame.add(a);
    b = shape.one("./p:cNvPr");
    if (b) {
      a.add(b);
    }
    b = shape.one("./p:cNvSpPr");
    if (b) {
      c = slide.el("p:cNvGraphicFramePr");
      b = b.all("./*");
      for (i = 0; i < b.length; i++) {
        c.add(b[i]);
      }
      a.add(c);
    }
    b = shape.one("./p:nvPr");
    if (b) {
      b.add(b);
    }

    //p:xfrm //TODO

    a = slide.el("a:graphic");
    graphicFrame.add(a);
    b = slide.el("a:graphicData").setAttr("uri", "http://schemas.openxmlformats.org/drawingml/2006/table");
    a.add(b);

    shape.replace(graphicFrame);

    if (block.type === "table") {
      this._table(slide, b, block);
    } else {
      throw "Block is not a table";
    }
  }
};
/*jshint +W071 */
/*jshint +W072 */

(function() {
  for (var f in XPkg.prototype) {
    if (XPkg.prototype.hasOwnProperty(f)) {
      PPTXBuilder.prototype[f] = XPkg.prototype[f];
    }
  }
})();

OpenXmlBuilder.PPTXBuilder = PPTXBuilder;
