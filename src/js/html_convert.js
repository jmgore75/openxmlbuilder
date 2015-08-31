var getStyle = window.getComputedStyle || function (e) {
    return e.currentStyle;
};

function leftPad (s, len, c) {
    c= c || " ";
    len= len || 2;
    while(s.length < len) {
        s= c + s;
    }
    return s;
}

var colornames = {
    AQUA: "00FFFF", BLACK: "000000", BLUE: "0000FF", FUCHSIA: "FF00FF",
    GRAY: "808080", GREEN: "008000", LIME: "00FF00", MAROON: "800000",
    NAVY: "000080", OLIVE: "808000", PURPLE: "800080", RED: "FF0000",
    SILVER: "C0C0C0", TEAL: "008080", WHITE: "FFFFFF", YELLOW: "FFFF00"
};
function convertColor(c) {
    var tem, i= 0;
    c= c? c.toString().toUpperCase(): "";
    if(/^#[A-F0-9]{3,6}$/.test(c)){
        if(c.length< 7){
            var A= c.split("");
            c= A[1]+A[1]+A[2]+A[2]+A[3]+A[3];
        } else {
            c= c.substr(1, 8);
        }
        return c;
    }
    if(/^[A-Z]+$/.test(c)){
        return colornames[c] || "";
    }
    c= c.match(/\d+(\.\d+)?%?/g) || [];
    if(c.length<3 || c.length>4) {
        return "";
    }
    for (i = 0; i < c.length; i++){
        tem= c[i];
        if (tem.indexOf("%") !== -1) {
            tem= Math.round(parseFloat(tem)*2.55);
        }
        else {
            tem= parseInt(tem, 10);
        }
        if( tem < 0 || tem > 255) {
            return "";
        }
        else {
            c[i] = leftPad(tem.toString(16).toUpperCase(), 2, "0");
        }
    }
    if( c.length === 4 && c[3] === "00") {
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
        if ( i >= 700) {
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

function Blocker () {
    this.blocks = [];
    this.block = null;
}
Blocker.prototype = {
    addRun : function (style, text) {
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
    currentBlock : function () {
        if (!this.block) {
            this.block = {type:"paragraph", runs:[]};
        }
        return this.block;
    },
    breakBlock : function (force) {
        if (this.block || force) {
            var block = this.currentBlock();
            this.blocks.push(block);
            this.block = null;
            return block;
        } else if (this.blocks.length) {
            return this.blocks[this.blocks.length - 1];
        }
    },
    getBlocks : function () {
        this.breakBlock();
        return this.blocks;
    }, /*jshint -W071 */
    procNode: function (style, child) {
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
                this.blocks.push({type:"list", listType:child.tagName.toLowerCase(), blocks:sblocker.getBlocks()});
            } else if (child.tagName === "TABLE") {
                sblocker = new Blocker();
                sblocker.procNodes(cstyle, child.childNodes);
                this.blocks.push({type:"table", blocks:sblocker.blocks});
            } else if (child.tagName === "TR") {
                sblocker = new Blocker();
                sblocker.procNodes(cstyle, child.childNodes);
                this.blocks.push({type:"table-row", blocks:sblocker.blocks});
            } else if (child.tagName === "TD" || child.tagName === "TH") {
                sblocker = new Blocker();
                sblocker.procNodes(cstyle, child.childNodes);
                this.blocks.push({type:"table-cell", blocks:sblocker.getBlocks()});
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
    }, /*jshint +W071 */
    procNodes: function (style, children) {
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

function convertHtml (html) {
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
