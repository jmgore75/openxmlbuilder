//Content specific to PowerPoint generation

relTypes.slideMaster = relTypePrefix + "slideMaster"; 
relTypes.slideLayout = relTypePrefix + "slideLayout"; 
relTypes.slide = relTypePrefix + "slide"; 
contTypes.slide = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"; 

function PPTXBuilder (b64Template, title, created, creator) {
    XPkg.call(this, b64Template, title, created, creator); 
    this.slideCount = 0; 
    
    this.presentation = this.getPart("/ppt/presentation.xml"); 
    this.sldIdLst = this.presentation.one("/p:presentation/p:sldIdLst"); 
    var es = this.sldIdLst.all("./p:sldId"); 
    for (var i = 0; i < es.length; i++) {
        es[i].remove(); 
    }
    
    this.app = this.getPart("/docProps/app.xml").xdoc; 
    this.slidesProp = this.app.one("/ep:Properties/ep:Slides"); 
    this.slidesProp.setValue("0"); 
    this.slidesTitlesProp = this.app.all("/ep:Properties/ep:HeadingPairs/vt:vector/vt:variant/vt:i4");
    this.slidesTitlesProp = this.slidesTitlesProp[this.slidesTitlesProp.length-1]; 
    this.slidesTitlesProp.setValue("0"); 
    this.titlesVector = this.app.one("/ep:Properties/ep:TitlesOfParts/vt:vector"); 
    this.titlesVector.setAttr("size", "1"); 
    es = this.titlesVector.all("./vt:lpstr"); 
    for (i = 1; i < es.length; i++) {
        es[i].remove(); 
    }
    
    var rels = this.presentation.getRelationshipsByRelationshipType(relTypes.slide); 
    var rel, part; 
    for (i = 0; i < rels.length; i++) {
        rel = rels[i]; 
        part = this.getPart(this.fullPath(rel.getAttr("Target"), this.presentation.path)); 
        this.presentation.deleteRelationship(rel.relationshipId); 
        part.deleted = true; 
    }
}
PPTXBuilder.prototype = {
    mimetype : "application/vnd.openxmlformats-officedocument.presentationml.presentation", 
    addSlide : function (title, slideLayoutId) {
        slideLayoutId = slideLayoutId || 2; 
        this.slideCount++; 
        var slideNum = this.slideCount;
        var sldId = slideNum + 255; 
        title = title || "Slide " + slideNum; 

        var slideUri = "slides/slide" + slideNum + ".xml"; 
        var slideLayoutUri = "slideLayouts/slideLayout" + slideLayoutId + ".xml"; 

        var slideLayout = this.getPart("/ppt/" + slideLayoutUri); 
        var sldTemplate = slideLayout.xdoc.toXmlString(); 
        sldTemplate = sldTemplate.replace(/(<\/?p:sld)Layout/g, "$1"); 
        var slide = this.addPart(contTypes.slide, "/ppt/" + slideUri, sldTemplate); 
 
        slide.addRelationship(slide.nextRelationshipId(), relTypes.slideLayout, "../" + slideLayoutUri, "Internal"); 

        var sld = slide.xdoc.root(); 
        sld.removeAttr("preserve"); 
        sld.removeAttr("type"); 
        sld.one("p:cSld").removeAttr("name"); 
        
        //Remove unused placeholders
        var rel = slideLayout.getRelationshipsByRelationshipType(relTypes.slideMaster)[0]; 
        var slideMaster = this.getPart(this.fullPath(rel.getAttr("Target"), slideLayout.path)); 
        var hf = slideMaster.one("//p:hf"); 
        if (hf) {
            var shapes = slide.all("//p:sp"); 
            var ph, phType; 
            for (var i = 0; i < shapes.length; i++) {
                ph = shapes[i].one("./p:nvSpPr/p:nvPr/p:ph"); 
                if (ph) {
                    phType = ph.getAttr("type"); 
                    if (phType && hf.getAttr(phType) === "0") {
                        shapes[i].remove(); 
                    }
                }
            }
        }

        var rId = this.presentation.nextRelationshipId(); 
        this.presentation.addRelationship(rId, relTypes.slide , slideUri, "Internal"); 
        
        this.sldIdLst.add(this.presentation.el("p:sldId", {"r:id":rId, "id":sldId})); 
        this.slidesProp.setValue(slideNum); 
        this.slidesTitlesProp.setValue(slideNum); 
        this.titlesVector.add(this.app.el("vt:lpstr", title)); 
        this.titlesVector.setAttr("size", slideNum + 1); 
        
        return slide; 
    }, 
    contentSlide : function (content, slideLayoutId) {
        slideLayoutId = slideLayoutId || 2; 
        var title = content.title || content["Title 1"]; 
        var slide = this.addSlide(title, slideLayoutId); 
        
        var shapes = slide.all("//p:sp"); 
        for (var i = 0; i < shapes.length; i++) {
            var spId = shapes[i].one("./p:nvSpPr/p:cNvPr").getAttr("name"); 
            var html = content[spId]; 
            if (html) {
                this.pptContent(slide, shapes[i], html);
            } 
        }
    }, 
    /*jshint -W071 */
    /*jshint -W072 */
    _block : function (slide, txBody, block, listType, level) {
        if (block.type === "list") {
            listType = block.listType || "ul"; 
            level = level >= 0 ? level +=1 : 0; 
            this._blocks(slide, txBody, block.blocks, listType, level); 
        } else if (block.type === "table") {
            this._inlineTable(slide, txBody, block); 
        } else if (block.type === "list-item") {
            this._paragraph(slide, txBody, block, listType, level); 
        } else {
            this._paragraph(slide, txBody, block, null, level); 
        }
    }, 
    _blocks : function (slide, txBody, blocks, listType, level) {
        for (var i = 0; i < blocks.length; i++) {
            this._block(slide, txBody, blocks[i], listType, level); 
        }
    },
    _runs : function (slide, parent, runs) {
        var run, r, rPr, t, rId; 
        var lastLink; 
        for (var i = 0; i < runs.length; i++) {
            run = runs[i]; 
            r = slide.el("a:r"); 
            lastLink = run.link; 
            if (run.b || run.i || run.u || run.color || run.link) {
                rPr = slide.el("a:rPr"); 
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
                    rPr.add(slide.el("a:solidFill").add(slide.el("a:srgbClr", {"val": run.color})));
                }
                if (run.link) {
                    rId = slide.nextRelationshipId(); 
                    slide.addRelationship(rId, relTypes.hyperlink, run.link, "External"); 
                    rPr.add(slide.el("a:hlinkClick", {"r:id": rId})); 
                }
                r.add(rPr); 
            }
            t = slide.el("a:t", run.text); 
            //if (/^\s+|\s+$/.test(run.text)) {
            //    t.setAttr("xml:space", "preserve"); 
            //}
            r.add(t); 
            parent.add(r); 
        }
        if (lastLink) {
            parent.add(slide.el("a:endParaRPr").add(slide.el("a:hlinkClick", {"r:id": ""}))); 
        } 
    },
    _paragraph : function (slide, parent, paragraph, listType, level) {
        var p, pPr; 
        p = slide.el("a:p"); 
        if (level >= 0 || listType !== "ul") {
            pPr = slide.el("a:pPr"); 
            p.add(pPr); 
            if (level > 0) {
                pPr.setAttr("lvl", level); 
            }
            if (listType === "ol") {
                pPr.add(slide.el("a:buAutoNum", {"type":"arabicPeriod"})); 
            } else if (!listType) {
                pPr.add(slide.el("a:buNone")); 
            }
            if (!listType && level >= 0) {
                p.add(slide.el("a:r").add(slide.el("a:t", "\t"))); 
            }
        }
        this._runs(slide, p, paragraph.runs); 
        parent.add(p); 
    }, 
    _inlineTable : function (slide, parent, table, listType, level) {
        var i, j, cind, cspan, rspan, row, cell; 
        var colspans = {}; 
        var spanned = function () {
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
        
        this._paragraph(slide, parent, {runs:[{u:true, i:true, text:"Table"}]}, listType, level); 
        level = (level || 0) + 1; 
        for (i = 0; i < table.blocks.length; i++) {
            row = table.blocks[i]; 
            this._paragraph(slide, parent, {runs:[{u:true, i:true, text:"Row " + (row + 1)}]}, null, level); 
            cind = 0; 
            for (j = 0; j < row.blocks.length; j++) {
                spanned(); 
                cell = row.blocks[j];  
                cspan = cell.colspan || 1; 
                rspan = cell.rowspan;
                this._paragraph(slide, parent, {runs:[{i:true, text:"Column" + (cind + 1)}]}, null, level + 1); 
                this._blocks(slide, parent, cell, null, level+1); 
                if (rspan > 1) {
                    colspans[cind] = [rspan-1, cspan || 1]; 
                }
                cind += cspan; 
            }
            spanned(); 
        }
    }, 
    _table : function (slide, parent, table) {
        var i, j, k, row, cell, r, c, e, cspan, rspan, cind; 
        var tbl = slide.el("a:tbl"); 
        //a:tblPr

        var phTxbody = function () {
            return slide.el("a:txBody").add(slide.el("a:bodyPr")).add(slide.el("a:lstStyle")); 
        };
        
        var phTc = function () {
            return slide.el("a:tc").add(
                phTxbody().add(slide.el("a:p")) 
            ).add(slide.el("a:tcPr")); 
        };

        var colspans = {}; 
        var spanned = function () {
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
                    colspans[cind] = [rspan-1, cspan || 1]; 
                }

                e = phTxbody(); 
                c.add(e); 
                this._blocks(e, cell.blocks); 
                e = slide.el("a:tcPr"); 
                c.add(e); 
                cind += cspan; 
            }
            spanned(); 
            tbl.add(r); 
        }
        
        parent.add(tbl); 
    }, 
    pptContent : function (slide, shape, html) {
        var txBody = shape.one("./p:txBody"); 
        var ps = txBody.all("./a:p"); 
        for (var j = 0; j < ps.length; j++) {
            ps[j].remove(); 
        }
        var blocks = convertHtml(html); 
        this._blocks(slide, txBody, blocks); 
    }, 
    pptLine : function (slide, shape, html) {
        var txBody = shape.one("./p:txBody"); 
        var ps = txBody.all("./a:p"); 
        for (var j = 0; j < ps.length; j++) {
            ps[j].remove(); 
        }
        var blocks = convertHtml(html); 
        if (blocks.length === 1 && blocks[0].type === "paragraph") {
            var block = blocks[0]; 
            this._paragraph(slide, txBody, block); 
        } else {
            throw "Content is not a single paragraph"; 
        } 
    }, 
    pptTable : function (slide, shape, html) {
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

        var blocks = convertHtml(html); 
        if (blocks.length === 1 && blocks[0].type === "table") {
            var block = blocks[0]; 
            this._table(slide, b, block); 
        } else {
            throw "Content is not a table"; 
        } 
    }
};
/*jshint +W071 */
/*jshint +W072 */

(function () {
    for (var f in XPkg.prototype) {
        if (XPkg.prototype.hasOwnProperty(f)) {
            PPTXBuilder.prototype[f] = XPkg.prototype[f]; 
        }
    }
})(); 

OpenXmlBuilder.PPTXBuilder = PPTXBuilder; 
