/* globals JSZip */
var OpenXmlBuilder = {};
OpenXmlBuilder.version = "<%= version %>";

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

/*jshint quotmark:single */
var xmlDoctype = '<?xml version="1.0" encoding="utf-8" standalone="yes"?>';
/*jshint quotmark:double */

//Given main element, default namespace, and any other namespaces, generate an empty xml document
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

//Standard rels file template
var relsTemplate = docTemplate("Relationships", "rel");

var EGHelper = {
  makeDoc: function(xmlStr) {
    return (new DOMParser()).parseFromString(xmlStr, "text/xml");
  },
  makeEl: function(doc, tag, rootNs) {
    var r = nameNS(tag, rootNs);
    return doc.createElementNS(r.ns || null, r.use);
  },
  toXml: function(doc) {
    return (new XMLSerializer()).serializeToString(doc);
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

  return date.getUTCFullYear() +
    "-" + pad2(date.getUTCMonth() + 1) +
    "-" + pad2(date.getUTCDate()) +
    "T" + pad2(date.getUTCHours()) +
    ":" + pad2(date.getUTCMinutes()) +
    ":" + pad2(date.getUTCSeconds()) +
    "Z";
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
  if (xmlStr.charCodeAt(0) === 0xFEFF) {
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

/*jshint -W072 */
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
/*jshint +W072 */
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
      "Target": target,
      "Type": type,
      "Id": rId
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
  this.zip = new JSZip(
    b64Template, {
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
      "ContentType": ct,
      "PartName": path
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
