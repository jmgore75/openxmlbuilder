/*!
 * OpenXmlBuilder
 * The OpenXmlBuilder allows you to build simple OpenXml documents right in the browser.
 * Copyright (c) 2015 Jeremy Gore
 * Licensed MIT
 * https://github.com/jmgore75/openxmlbuilder
 * v0.1.1
 */
(function(window, undefined) {
  "use strict";
  var OpenXmlBuilder = {};
  OpenXmlBuilder.version = "0.1.1";
  var ox = {};
  ox.xmlns = "http://www.w3.org/2000/xmlns/";
  ox.rel = "http://schemas.openxmlformats.org/package/2006/relationships";
  ox.a = "http://schemas.openxmlformats.org/drawingml/2006/main";
  ox.p = "http://schemas.openxmlformats.org/presentationml/2006/main";
  ox.r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  ox.vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
  ox.dcterms = "http://purl.org/dc/terms/";
  ox.dc = "http://purl.org/dc/elements/1.1/";
  ox.cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
  ox.ep = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
  ox.w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  ox.ct = "http://schemas.openxmlformats.org/package/2006/content-types";
  var oxNs = {};
  var selNs = "";
  (function() {
    for (var abbrev in ox) {
      if (ox.hasOwnProperty(abbrev)) {
        oxNs[ox[abbrev]] = abbrev;
        if (abbrev === "xmlns") {
          selNs += " xmlns='" + ox[abbrev] + "'";
        } else {
          selNs += " xmlns:" + abbrev + "='" + ox[abbrev] + "'";
        }
      }
    }
  })();
  selNs = selNs.slice(1);
  function nsResolver(pref) {
    return ox[pref] || null;
  }
  function nameNS(name, rootNs) {
    var r = {
      full: name,
      name: name,
      ns: rootNs
    };
    r.pref = oxNs[rootNs] || "";
    var toks = name.split(":");
    if (toks.length === 2) {
      r.pref = toks[0];
      r.name = toks[1];
      r.ns = ox[r.pref] || null;
    }
    r.use = r.ns === rootNs ? r.name : r.full;
    return r;
  }
  var relTypes = {};
  var relTypePrefix = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
  relTypes.hyperlink = relTypePrefix + "hyperlink";
  var contTypes = {};
  var xmlDoctype = '<?xml version="1.0" encoding="utf-8" standalone="yes"?>';
  function docTemplate() {
    var t = xmlDoctype;
    t += "<" + arguments[0];
    if (arguments.length > 1 && arguments[1]) {
      t += " xmlns='" + ox[arguments[1]] + "'";
    }
    for (var i = 2; i < arguments.length; i++) {
      t += " xmlns:" + arguments[i] + "='" + ox[arguments[i]] + "'";
    }
    t += "></" + arguments[0] + ">";
    return t;
  }
  var relsTemplate = docTemplate("Relationships", "rel");
  var EGHelper = {
    makeDoc: function(xmlStr) {
      return new DOMParser().parseFromString(xmlStr, "text/xml");
    },
    makeEl: function(doc, tag, rootNs) {
      var r = nameNS(tag, rootNs);
      return doc.createElementNS(r.ns || null, r.use);
    },
    toXml: function(doc) {
      return new XMLSerializer().serializeToString(doc);
    },
    xpath: function(node, xpath, first) {
      var doc = node.ownerDocument || node;
      if (first) {
        return doc.evaluate(xpath, node, nsResolver, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
      } else {
        var rs = [];
        var r = doc.evaluate(xpath, node, nsResolver, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
        for (var i = 0; i < r.snapshotLength; i++) {
          rs.push(r.snapshotItem(i));
        }
        return rs;
      }
    },
    setNs: function(node, ns, abbrev) {
      node.setAttributeNS(ox.xmlns, abbrev ? "xmlns:" + abbrev : "xmlns", ns);
    }
  };
  var IEHelper = {
    makeDoc: function(xmlStr) {
      var doc = new window.ActiveXObject("Microsoft.XMLDOM");
      doc.async = "false";
      doc.loadXML(xmlStr);
      doc.setProperty("SelectionNamespaces", selNs);
      return doc;
    },
    makeEl: function(doc, tag, rootNs) {
      var r = nameNS(tag, rootNs);
      return doc.createNode(1, r.use, r.ns || null);
    },
    toXml: function(doc) {
      return doc.xml;
    },
    xpath: function(node, xpath, first) {
      if (first) {
        return node.selectSingleNode(xpath);
      } else {
        return node.selectNodes(xpath);
      }
    },
    setNs: function(node, ns, abbrev) {
      var a = node.doc.createNode(2, abbrev ? "xmlns:" + abbrev : "xmlns", ox.xmlns);
      a.value = ns;
      node.setAttributeNode(a);
    }
  };
  var xhelp = window.DOMParser ? EGHelper : IEHelper;
  OpenXmlBuilder.XMLHelper = xhelp;
  function openXmlISOString(date) {
    function pad2(number) {
      if (number < 10) {
        return "0" + number;
      }
      return number;
    }
    return date.getUTCFullYear() + "-" + pad2(date.getUTCMonth() + 1) + "-" + pad2(date.getUTCDate()) + "T" + pad2(date.getUTCHours()) + ":" + pad2(date.getUTCMinutes()) + ":" + pad2(date.getUTCSeconds()) + "Z";
  }
  function nodeWrap(nodes) {
    var r = [];
    for (var i = 0; i < nodes.length; i++) {
      r.push(new XNode(nodes[i]));
    }
    return r;
  }
  function XNode(node) {
    this.node = node;
  }
  XNode.prototype = {
    all: function(path) {
      return nodeWrap(xhelp.xpath(this.node, path, false));
    },
    one: function(path) {
      var node = xhelp.xpath(this.node, path, true);
      return node ? new XNode(node) : null;
    },
    remove: function() {
      this.node.parentNode.removeChild(this.node);
    },
    setValue: function(value) {
      if (this.node.nodeType === 1) {
        while (this.node.firstChild) {
          this.node.removeChild(this.node.firstChild);
        }
        var txt = this.node.ownerDocument.createTextNode(value);
        this.node.appendChild(txt);
      } else if (this.node.nodeType === 2) {
        this.value = value;
      }
      return this;
    },
    getValue: function() {
      if (this.node.nodeType === 2) {
        return this.value;
      } else {
        return this.node.nodeValue;
      }
    },
    setAttr: function(attr, value) {
      this.node.setAttribute(attr, value);
      return this;
    },
    getAttr: function(attr) {
      return this.node.getAttribute(attr);
    },
    removeAttr: function(attr) {
      this.node.removeAttribute(attr);
      return this;
    },
    add: function(el, before) {
      if (before) {
        this.node.insertBefore(el.node, before.node);
      } else {
        this.node.appendChild(el.node);
      }
      return this;
    },
    replace: function(replacement) {
      this.node.insertBefore(replacement.node);
      this.remove();
      return replacement;
    }
  };
  function XDoc(xmlStr) {
    if (xmlStr.charCodeAt(0) === 65279) {
      xmlStr = xmlStr.substring(1);
    }
    this.doc = xhelp.makeDoc(xmlStr);
    this.rootNs = this.root().getAttr("xmlns");
  }
  XDoc.prototype = {
    el: function(tag, attr, value) {
      var el = new XNode(xhelp.makeEl(this.doc, tag, this.rootNs));
      if (!value && attr && typeof attr !== "object") {
        value = attr;
        attr = null;
      }
      if (attr) {
        for (var a in attr) {
          if (attr.hasOwnProperty(a)) {
            el.setAttr(a, attr[a]);
          }
        }
      }
      if (value) {
        el.setValue(value);
      }
      return el;
    },
    root: function() {
      return new XNode(this.doc.documentElement);
    },
    all: function(path) {
      return nodeWrap(xhelp.xpath(this.doc, path, false));
    },
    one: function(path) {
      var node = xhelp.xpath(this.doc, path, true);
      return node ? new XNode(node) : null;
    },
    nextId: function(path, attr, first, pref) {
      first = first || 0;
      pref = pref || "";
      var ids = {};
      var es = this.all(path);
      for (var i = 0; i < es.length; i++) {
        ids[es[i].getAttr(attr)] = 1;
      }
      while (ids[pref + first]) {
        first += 1;
      }
      return pref + first;
    },
    setNamespace: function(ns, abbrev) {
      xhelp.setNs(this.root(), ns, abbrev);
      if (!abbrev) {
        this.rootNs = ns;
      }
    },
    toXmlString: function() {
      return xhelp.toXml(this.doc);
    }
  };
  function XPart(xpkg, contentType, path, xmlStr, relsXmlStr) {
    this.xpkg = xpkg;
    this.path = path;
    this.contentType = contentType;
    var idx = path.lastIndexOf("/") + 1;
    this.relsPath = path.slice(0, idx) + "_rels/" + path.slice(idx) + ".rels";
    if (xmlStr) {
      this.xdoc = new XDoc(xmlStr);
      if (relsXmlStr) {
        this.relsXdoc = new XDoc(relsXmlStr);
      }
    } else {
      var f = this.xpkg.zip.file(this.path.slice(1));
      if (f) {
        this.xdoc = new XDoc(f.asText());
        f = this.xpkg.zip.file(this.relsPath.slice(1));
        if (f) {
          this.relsXdoc = new XDoc(f.asText());
        }
      }
    }
  }
  XPart.prototype = {
    nextRelationshipId: function() {
      if (!this.relsXdoc) {
        return "rId1";
      }
      return this.relsXdoc.nextId("/rel:Relationships/rel:Relationship", "Id", 1, "rId");
    },
    addRelationship: function(rId, type, target, mode) {
      if (!this.relsXdoc) {
        this.relsXdoc = new XDoc(relsTemplate);
      }
      var attr = {
        Target: target,
        Type: type,
        Id: rId
      };
      mode = mode || "External";
      if (mode !== "Internal") {
        attr.TargetMode = mode;
      }
      var e = this.relsXdoc.el("rel:Relationship", attr);
      this.relsXdoc.root().add(e);
    },
    deleteRelationship: function(rId) {
      if (this.relsXdoc) {
        var es = this.relsXdoc.all("/rel:Relationships/rel:Relationship");
        for (var i = 0; i < es.length; i++) {
          if (es[i].getAttr("Id") === rId) {
            es[i].remove();
            if (es.length === 1) {
              delete this.relsXdoc;
            }
            return;
          }
        }
      }
    },
    el: function() {
      return this.xdoc.el.apply(this.xdoc, arguments);
    },
    root: function() {
      return this.xdoc.root();
    },
    all: function() {
      return this.xdoc.all.apply(this.xdoc, arguments);
    },
    one: function() {
      return this.xdoc.one.apply(this.xdoc, arguments);
    },
    nextId: function() {
      return this.xdoc.nextId.apply(this.xdoc, arguments);
    },
    setNamespace: function() {
      return this.xdoc.setNamespace.apply(this.xdoc, arguments);
    },
    getRelationshipsByRelationshipType: function(rtype) {
      return this.relsXdoc.all("/rel:Relationships/rel:Relationship[@Type='" + rtype + "']");
    },
    loadXml: function(strXml) {
      this.xdoc = new XDoc(strXml);
    }
  };
  function XPkg(b64Template, title, created, creator) {
    this.zip = new JSZip(b64Template, {
      base64: true,
      checkCRC32: false
    });
    this.parts = {};
    created = created || new Date();
    created = openXmlISOString(created);
    title = title || "Document " + created;
    creator = creator || "OpenXmlBuilder";
    this.ctDoc = new XDoc(this.zip.file("[Content_Types].xml").asText());
    var core = this.getPart("/docProps/core.xml").xdoc;
    core.one("//dcterms:created").setValue(created);
    core.one("//dcterms:modified").setValue(created);
    core.one("//cp:revision").setValue("1");
    if (title) {
      var titleElement = core.one("//dc:title");
      if (titleElement) {
        titleElement.setValue(title);
      } else {
        core.root().add(core.el("dc:title", title));
      }
    }
    core.one("//dc:creator").setValue(creator);
    core.one("//cp:lastModifiedBy").setValue(creator);
  }
  XPkg.prototype = {
    getPart: function(path) {
      if (this.parts.hasOwnProperty(path)) {
        return this.parts[path];
      } else {
        var override = this.ctDoc.one("/ct:Types/ct:Override[@PartName='" + path + "']");
        var part = override ? new XPart(this, override.getAttr("ContentType"), path) : null;
        this.parts[path] = part;
        return part;
      }
    },
    addPart: function(contentType, path, xmlStr, relsXmlStr) {
      var part = new XPart(this, contentType, path, xmlStr, relsXmlStr);
      this.parts[path] = part;
      return part;
    },
    addFile: function(ct, path, data) {
      this.zip.file(path.slice(1), data);
      var cts = this.ctDoc.all("/ct:Types/ct:Override");
      for (var i = 0; i < cts.length; i++) {
        if (cts[i].getAttr("PartName") === path) {
          cts[i].remove();
        }
      }
      var e = this.ctDoc.el("ct:Override", {
        ContentType: ct,
        PartName: path
      });
      this.ctDoc.root().add(e);
    },
    deleteFile: function(path) {
      this.zip.remove(path);
      var cts = this.ctDoc.all("/ct:Types/ct:Override");
      for (var i = 0; i < cts.length; i++) {
        if (cts[i].getAttr("PartName") === path) {
          cts[i].remove();
        }
      }
    },
    fullPath: function(relPath, ctxPath) {
      if (relPath.indexOf("/") === 0) {
        return relPath;
      }
      ctxPath = ctxPath || "/";
      var slashIndex = ctxPath.lastIndexOf("/");
      ctxPath = ctxPath.slice(0, slashIndex + 1);
      while (relPath.indexOf("../") === 0) {
        ctxPath = ctxPath.slice(0, ctxPath.length - 1);
        slashIndex = ctxPath.lastIndexOf("/");
        if (slashIndex === -1) {
          throw "internal error when processing relationships";
        }
        ctxPath = ctxPath.slice(0, slashIndex + 1);
        relPath = relPath.slice(3);
      }
      return ctxPath + relPath;
    },
    undoPart: function(path) {
      delete this.parts[path];
    },
    saveChanges: function() {
      var part;
      for (var path in this.parts) {
        if (this.parts.hasOwnProperty(path)) {
          part = this.parts[path];
          if (part.deleted || !part.xdoc) {
            this.deleteFile(part.path);
            this.zip.remove(part.relsPath);
          } else {
            this.addFile(part.contentType, part.path, part.xdoc.toXmlString());
            if (part.relsXdoc) {
              this.zip.file(part.relsPath.slice(1), part.relsXdoc.toXmlString());
            } else {
              this.zip.remove(part.relsPath);
            }
          }
        }
      }
      this.parts = {};
      this.zip.file("[Content_Types].xml", this.ctDoc.toXmlString());
      return this.zip;
    },
    saveToBase64: function() {
      return this.saveChanges().generate({
        compression: "deflate",
        type: "base64"
      });
    },
    saveToBlob: function() {
      return this.saveChanges().generate({
        compression: "deflate",
        type: "blob"
      });
    },
    saveToDataURI: function() {
      return "data:" + this.mimetype + ";base64," + this.saveToBase64();
    }
  };
  var getStyle = window.getComputedStyle || function(e) {
    return e.currentStyle;
  };
  function leftPad(s, len, c) {
    c = c || " ";
    len = len || 2;
    while (s.length < len) {
      s = c + s;
    }
    return s;
  }
  var colornames = {
    AQUA: "00FFFF",
    BLACK: "000000",
    BLUE: "0000FF",
    FUCHSIA: "FF00FF",
    GRAY: "808080",
    GREEN: "008000",
    LIME: "00FF00",
    MAROON: "800000",
    NAVY: "000080",
    OLIVE: "808000",
    PURPLE: "800080",
    RED: "FF0000",
    SILVER: "C0C0C0",
    TEAL: "008080",
    WHITE: "FFFFFF",
    YELLOW: "FFFF00"
  };
  function convertColor(c) {
    var tem, i = 0;
    c = c ? c.toString().toUpperCase() : "";
    if (/^#[A-F0-9]{3,6}$/.test(c)) {
      if (c.length < 7) {
        var A = c.split("");
        c = A[1] + A[1] + A[2] + A[2] + A[3] + A[3];
      } else {
        c = c.substr(1, 8);
      }
      return c;
    }
    if (/^[A-Z]+$/.test(c)) {
      return colornames[c] || "";
    }
    c = c.match(/\d+(\.\d+)?%?/g) || [];
    if (c.length < 3 || c.length > 4) {
      return "";
    }
    for (i = 0; i < c.length; i++) {
      tem = c[i];
      if (tem.indexOf("%") !== -1) {
        tem = Math.round(parseFloat(tem) * 2.55);
      } else {
        tem = parseInt(tem, 10);
      }
      if (tem < 0 || tem > 255) {
        return "";
      } else {
        c[i] = leftPad(tem.toString(16).toUpperCase(), 2, "0");
      }
    }
    if (c.length === 4 && c[3] === "00") {
      return "";
    }
    return c.slice(0, 3).join("");
  }
  function convertStyle(s) {
    var a = {};
    if (s.verticalAlign === "super") {
      a.sup = 1;
    }
    if (s.verticalAlign === "sub") {
      a.sub = 1;
    }
    var i = parseInt(s.fontWeight, 10);
    if (isNaN(i)) {
      if (s.fontWeight === "bold") {
        a.b = 1;
      }
    } else {
      if (i >= 700) {
        a.b = 1;
      }
    }
    if (s.fontStyle === "italic") {
      a.i = 1;
    }
    if (s.textDecoration === "underline") {
      a.u = 1;
    }
    var color = convertColor(s.color);
    if (color && color !== "000000") {
      a.color = color;
    }
    color = convertColor(s.backgroundColor);
    if (color && color !== "FFFFFF") {
      a.bgColor = color;
    }
    return a;
  }
  function Blocker() {
    this.blocks = [];
    this.block = null;
  }
  Blocker.prototype = {
    addRun: function(style, text) {
      var run = {};
      for (var attr in style) {
        if (style.hasOwnProperty(attr)) {
          run[attr] = style[attr];
        }
      }
      run.text = text;
      var block = this.currentBlock();
      block.runs.push(run);
      return run;
    },
    currentBlock: function() {
      if (!this.block) {
        this.block = {
          type: "paragraph",
          runs: []
        };
      }
      return this.block;
    },
    breakBlock: function(force) {
      if (this.block || force) {
        var block = this.currentBlock();
        this.blocks.push(block);
        this.block = null;
        return block;
      } else if (this.blocks.length) {
        return this.blocks[this.blocks.length - 1];
      }
    },
    getBlocks: function() {
      this.breakBlock();
      return this.blocks;
    },
    procNode: function(style, child) {
      var sblocker;
      if (child.nodeType === 3) {
        this.addRun(style, child.nodeValue);
      } else if (child.nodeType === 1) {
        var cssStyle = getStyle(child);
        if (cssStyle.display === "none") {
          return;
        }
        var isBlock = cssStyle.display === "block";
        var cstyle = convertStyle(cssStyle);
        if (isBlock) {
          this.breakBlock();
        }
        if (child.tagName === "LI") {
          this.breakBlock();
          this.currentBlock().type = "list-item";
          this.procNodes(cstyle, child.childNodes);
        } else if (child.tagName === "OL" || child.tagName === "UL") {
          this.breakBlock();
          sblocker = new Blocker();
          sblocker.procNodes(cstyle, child.childNodes);
          this.blocks.push({
            type: "list",
            listType: child.tagName.toLowerCase(),
            blocks: sblocker.getBlocks()
          });
        } else if (child.tagName === "TABLE") {
          sblocker = new Blocker();
          sblocker.procNodes(cstyle, child.childNodes);
          this.blocks.push({
            type: "table",
            blocks: sblocker.blocks
          });
        } else if (child.tagName === "TR") {
          sblocker = new Blocker();
          sblocker.procNodes(cstyle, child.childNodes);
          this.blocks.push({
            type: "table-row",
            blocks: sblocker.blocks
          });
        } else if (child.tagName === "TD" || child.tagName === "TH") {
          sblocker = new Blocker();
          sblocker.procNodes(cstyle, child.childNodes);
          this.blocks.push({
            type: "table-cell",
            blocks: sblocker.getBlocks()
          });
        } else if (child.tagName === "A") {
          cstyle.link = child.href;
          this.procNodes(cstyle, child.childNodes);
        } else if (child.tagName === "HR") {
          this.currentBlock().hr = true;
          this.breakBlock();
        } else if (child.tagName === "BR") {
          this.breakBlock();
        } else if (child.tagName === "IMG") {
          var run = this.addRun(cstyle, "[" + (child.alt || child.title || "IMAGE") + "]");
          run.link = child.src;
        } else {
          this.procNodes(cstyle, child.childNodes);
        }
      }
    },
    procNodes: function(style, children) {
      for (var i = 0; i < children.length; i++) {
        this.procNode(style, children[i]);
      }
    }
  };
  function cleanupHtml(html) {
    if (html) {
      return html.replace(/[\u200B-\u200D\uFEFF]/g, "");
    }
  }
  function convertHtml(html) {
    var d = document.createElement("DIV");
    var p = document.head || document.body;
    p.appendChild(d);
    d.innerHTML = cleanupHtml(html);
    var mblocker = new Blocker();
    mblocker.procNodes({}, d.childNodes);
    p.removeChild(d);
    d = null;
    return mblocker.getBlocks();
  }
  OpenXmlBuilder.convertHtml = convertHtml;
  relTypes.slideMaster = relTypePrefix + "slideMaster";
  relTypes.slideLayout = relTypePrefix + "slideLayout";
  relTypes.slide = relTypePrefix + "slide";
  contTypes.slide = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
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
        id: sldId
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
    _findLayouts: function() {
      var slideLayouts = {};
      var slideLayoutObjects = this.zip.folder("/ppt/slideLayouts").file(/^slideLayout\d+\.xml$/);
      for (var j = 0; j < slideLayoutObjects.length; j++) {
        var slo = slideLayoutObjects[j];
        var slideLayoutId = parseInt(slo.name.slice(11, -4), 10);
        var fullUri = "/ppt/slideLayouts/" + slo.name;
        var relUri = "../slideLayouts/" + slo.name;
        var slideLayout = new XDoc(slo.asText());
        var masterPath = new XDoc(this.zip.file("/ppt/slideLayouts/_rels/" + slo.name + ".rels").asText()).one("/rel:Relationships/rel:Relationship[@Type='" + relTypes.slideMaster + "']").getAttr("Target");
        var slideMaster = this.getPart(this.fullPath(masterPath, slideLayout.path));
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
          fileName: slo.name,
          fullUri: fullUri,
          relUri: relUri,
          sections: layoutSections,
          template: sldTemplate
        };
      }
      return slideLayouts;
    },
    _block: function(slide, txBody, block, listType, level) {
      if (block.type === "list") {
        listType = block.listType || "ul";
        level = level >= 0 ? level += 1 : 0;
        this._blocks(slide, txBody, block.blocks, listType, level);
      } else if (block.type === "table") {
        this._inlineTable(slide, txBody, block);
      } else if (block.type === "list-item") {
        this._paragraph(slide, txBody, block, listType, level);
      } else {
        this._paragraph(slide, txBody, block, null, level);
      }
    },
    _blocks: function(slide, txBody, blocks, listType, level) {
      for (var i = 0; i < blocks.length; i++) {
        this._block(slide, txBody, blocks[i], listType, level);
      }
    },
    _runs: function(slide, parent, runs) {
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
            rPr.add(slide.el("a:solidFill").add(slide.el("a:srgbClr", {
              val: run.color
            })));
          }
          if (run.link) {
            rId = slide.nextRelationshipId();
            slide.addRelationship(rId, relTypes.hyperlink, run.link, "External");
            rPr.add(slide.el("a:hlinkClick", {
              "r:id": rId
            }));
          }
          r.add(rPr);
        }
        t = slide.el("a:t", run.text);
        r.add(t);
        parent.add(r);
      }
      if (lastLink) {
        parent.add(slide.el("a:endParaRPr").add(slide.el("a:hlinkClick", {
          "r:id": ""
        })));
      }
    },
    _paragraph: function(slide, parent, paragraph, listType, level) {
      var p, pPr;
      p = slide.el("a:p");
      if (level >= 0 || listType !== "ul") {
        pPr = slide.el("a:pPr");
        p.add(pPr);
        if (level > 0) {
          pPr.setAttr("lvl", level);
        }
        if (listType === "ol") {
          pPr.add(slide.el("a:buAutoNum", {
            type: "arabicPeriod"
          }));
        } else if (!listType) {
          pPr.add(slide.el("a:buNone"));
        }
        if (!listType && level >= 0) {
          p.add(slide.el("a:r").add(slide.el("a:t", "	")));
        }
      }
      this._runs(slide, p, paragraph.runs);
      parent.add(p);
    },
    _inlineTable: function(slide, parent, table, listType, level) {
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
      this._paragraph(slide, parent, {
        runs: [ {
          u: true,
          i: true,
          text: "Table"
        } ]
      }, listType, level);
      level = (level || 0) + 1;
      for (i = 0; i < table.blocks.length; i++) {
        row = table.blocks[i];
        this._paragraph(slide, parent, {
          runs: [ {
            u: true,
            i: true,
            text: "Row " + (row + 1)
          } ]
        }, null, level);
        cind = 0;
        for (j = 0; j < row.blocks.length; j++) {
          spanned();
          cell = row.blocks[j];
          cspan = cell.colspan || 1;
          rspan = cell.rowspan;
          this._paragraph(slide, parent, {
            runs: [ {
              i: true,
              text: "Column" + (cind + 1)
            } ]
          }, null, level + 1);
          this._blocks(slide, parent, cell, null, level + 1);
          if (rspan > 1) {
            colspans[cind] = [ rspan - 1, cspan || 1 ];
          }
          cind += cspan;
        }
        spanned();
      }
    },
    _table: function(slide, parent, table) {
      var i, j, k, row, cell, r, c, e, cspan, rspan, cind;
      var tbl = slide.el("a:tbl");
      var phTxbody = function() {
        return slide.el("a:txBody").add(slide.el("a:bodyPr")).add(slide.el("a:lstStyle"));
      };
      var phTc = function() {
        return slide.el("a:tc").add(phTxbody().add(slide.el("a:p"))).add(slide.el("a:tcPr"));
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
            colspans[cind] = [ rspan - 1, cspan || 1 ];
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
    pptContent: function(slide, shape, blocks) {
      var txBody = shape.one("./p:txBody");
      var ps = txBody.all("./a:p");
      for (var j = 0; j < ps.length; j++) {
        ps[j].remove();
      }
      this._blocks(slide, txBody, blocks);
    },
    pptLine: function(slide, shape, block) {
      var txBody = shape.one("./p:txBody");
      var ps = txBody.all("./a:p");
      for (var j = 0; j < ps.length; j++) {
        ps[j].remove();
      }
      if (block.type === "paragraph") {
        this._paragraph(slide, txBody, block);
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
  (function() {
    for (var f in XPkg.prototype) {
      if (XPkg.prototype.hasOwnProperty(f)) {
        PPTXBuilder.prototype[f] = XPkg.prototype[f];
      }
    }
  })();
  OpenXmlBuilder.PPTXBuilder = PPTXBuilder;
  relTypes.numbering = relTypePrefix + "numbering";
  relTypes.aFChunk = relTypePrefix + "aFChunk";
  contTypes.numbering = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
  function _valBuild(e, v) {
    return function(xdoc) {
      return xdoc.el(e, {
        "w:val": v
      });
    };
  }
  function DOCXBuilder(b64Template, title, created, creator) {
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
    var styles = this.getPart("/word/document/styles.xml").all("w:styles/w:style");
    for (i = 0; i < styles.length; i++) {
      var styleType = styles[i].getAttr("w:type");
      var styleId = styles[i].getAttr("w:styleId");
      if (styleType === "paragraph") {
        this.pStyles[styleId] = _valBuild("w:pStyle", styleId);
      } else if (styleType === "character") {
        this.rStyles[styleId] = _valBuild("w:rStyle", styleId);
      }
    }
    this.undoPart("/word/document/styles.xml");
  }
  DOCXBuilder.prototype = {
    mimetype: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    rStyleItalic: function(xdoc) {
      return xdoc.el("w:i");
    },
    rStyleBold: function(xdoc) {
      return xdoc.el("w:b");
    },
    rStyleUnderline: _valBuild("w:u", "single"),
    rStyleHyperlink: _valBuild("w:rStyle", "Hyperlink"),
    pStyleNoSpacing: _valBuild("w:pStyle", "NoSpacing"),
    pStyleListParagraph: _valBuild("w:pStyle", "ListParagraph"),
    pStyle: function(name) {
      return this.pStyles[name] || _valBuild("w:pStyle", name);
    },
    rStyle: function(name) {
      return this.rStyles[name] || _valBuild("w:rStyle", name);
    },
    _levelIndent: function(level) {
      return 720 * (1 + (level || 0));
    },
    makeUl: function() {
      return this.makeList([ {
        fmt: "bullet",
        fnt: "Symbol",
        txt: "·"
      }, {
        fmt: "bullet",
        fnt: "Courier New",
        txt: "o"
      }, {
        fmt: "bullet",
        fnt: "Wingdings",
        txt: "§"
      } ], 9);
    },
    makeOl: function() {
      return this.makeList([ {
        fmt: "decimal"
      }, {
        fmt: "lowerLetter"
      }, {
        fmt: "lowerRoman"
      } ], 9);
    },
    makeOlNum: function() {
      return this.makeList([ {
        fmt: "decimal"
      } ], 9);
    },
    makeOlHier: function() {
      return this.makeList([ {
        fmt: "decimal",
        hier: true
      } ], 9);
    },
    makeList: function(levels, depth) {
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
          lvlText = "%" + (i + 1) + ".";
          hLvlText += lvlText;
        }
        lvl.add(numPart.el("w:lvlText").setAttr("w:val", level.heir ? hLvlText : lvlText));
        lvl.add(numPart.el("w:lvlJc").setAttr("w:val", "left"));
        lvl.add(numPart.el("w:pPr").add(numPart.el("w:ind", {
          "w:hanging": "360",
          "w:left": idt
        })));
        if (level.fnt) {
          lvl.add(numPart.el("w:rPr").add(numPart.el("w:rFonts", {
            "w:hint": "default",
            "w:hAnsi": level.fnt,
            "w:ascii": level.fnt
          })));
        }
        anum.add(lvl);
      }
      var num = numPart.el("w:num").setAttr("w:numId", listid);
      num.add(numPart.el("w:abstractNumId").setAttr("w:val", anumid));
      numPart.xdoc.root().add(anum);
      numPart.xdoc.root().add(num);
      return listid;
    },
    docLine: function(html, pStyle, rStyle) {
      var seg = convertHtml(html);
      if (seg.length === 1 && seg[0].type === "paragraph") {
        seg = seg[0];
        this._paragraph(this.body, this.sectPr, seg, pStyle, rStyle);
      } else {
        throw "Content is too complex";
      }
    },
    docContent: function(html) {
      var blocks = convertHtml(html);
      this._blocks(this.body, this.sectPr, blocks);
    },
    docChunk: function(html) {
      html = "<html><head></head><body>" + cleanupHtml(html) + "</body></html>";
      var rId = this.doc.nextRelationshipId();
      var uri = "/word/chunk" + rId + ".html";
      this.addFile("text/html", uri, html);
      this.doc.addRelationship(rId, relTypes.aFChunk, uri, "Internal");
      var chunk = this.doc.el("w:altChunk").setAttr("r:id", rId);
      this.body.add(chunk, this.sectPr);
    },
    _block: function(parent, before, block, listid, level) {
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
    _blocks: function(parent, before, blocks, listid, level) {
      for (var i = 0; i < blocks.length; i++) {
        this._block(parent, before, blocks[i], listid, level);
      }
    },
    _runs: function(parent, runs, rStyle) {
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
    _paragraph: function(parent, before, paragraph, pStyle, rStyle, pPrs) {
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
    _table: function(parent, before, table) {
      var i, j, row, cell, r, c, e, cspan, rspan, cind;
      var tbl = this.doc.el("w:tbl");
      var colspans = {};
      var spanned = function() {
        while (colspans[cind]) {
          cspan = colspans[cind][1] || 1;
          rspan = colspans[cind][0];
          c = this.doc.el("w:tc");
          e = this.doc.el("w:tcPr");
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
          if (cspan > 1) {
            e.add(this.doc.el("w:gridSpan").setAttr("w:val", cspan));
          }
          if (rspan > 1) {
            colspans[cind] = [ rspan - 1, cspan || 1 ];
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
  (function() {
    for (var f in XPkg.prototype) {
      if (XPkg.prototype.hasOwnProperty(f)) {
        DOCXBuilder.prototype[f] = XPkg.prototype[f];
      }
    }
  })();
  OpenXmlBuilder.DOCXBuilder = DOCXBuilder;
  if (typeof define === "function" && define.amd) {
    define(function() {
      return OpenXmlBuilder;
    });
  } else if (typeof module === "object" && module && typeof module.exports === "object" && module.exports) {
    module.exports = OpenXmlBuilder;
  } else {
    window.OpenXmlBuilder = OpenXmlBuilder;
  }
})(function() {
  return this;
}());