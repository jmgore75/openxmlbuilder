# OpenXmlBuilder API

This documents details the OpenXmlBuilder API, including various types of properties, methods, and events.


XMLHelper
convertHTML
PPTXBuilder
DOCXBuilder

## Static

### Static Properties

#### `OpenXmlBuilder.version`

_[`String`]_ The version of the OpenXmlBuilder library being used, e.g. `"v1.0.0"`.

### Static Methods

#### `OpenXmlBuilder.convertHTML(...)`

_[`Array`]_ Convert the given `html` to a sequence of blocks.  
OpenXml text sections typically are not as nested as html, and instead are made up of a sequence of 
objects, mainly paragraphs.  This function preprocesses the data into an array of blocks, each of 
which specifies its type and contains the necessary sub items and formatting metadata.  
E.g., a series of paragraphs, each of which contains a sequence of rums.  


## XPkg

All OpenXmlBuilder objects inherit from XPkg

### XPkg Properties

#### `pkg.zip`

_[`JSZip`]_ The internal JSZip object.  You should call `saveChanges` prior to using this object.


### XPkg Methods

#### `pkg.saveChanges()`

_[`this`]_ Save all pending changes to the zip.


#### `pkg.saveToBase64()`

_[`String`]_ Save changes and export document in base64 encoding.


#### `pkg.saveToBlob()`

_[`Blob`]_ Save changes and export document as a blob.


#### `pkg.saveToDataURI()`

_[`String`]_ Save changes and export document as a data URI.



## PPTXBuilder

PowerPoint files are built using `PPTXBuilder` objects. 

### PPTXBuilder Constructor

```js
var pb = new OpenXmlBuilder.PPTXBuilder(OpenXmlB64Templates.pptx, title, created, creator); 
```

_[`PPTXBuilder`]_ Create a PPTXBuilder based on the given base64 encoded `template`, 
optionally with a `title`, `created` date, and `creator` name.


### PPTXBuilder Properties

#### `pb.slideCount`

_[`Integer`]_ The current number of slides.


### PPTXBuilder Methods

#### `pb.contentSlide(...)`

```js
pb.contentSlide({"Title 1":"Title text", "Subtitle 2" : "A subtitle" }, 1); 
```

_[`undefined`]_ Create a new slide with the given `content` and `slideLayoutId`.  
The `slideLayoutId` identifies the slide layout to use.  A given slide layout will have
multiple sections which all have an identifier.  The content should be a plain object with 
keys corresponding to the identifiers and values corresponding to the html to insert.  


## DOCXBuilder

Word files are built using `DOCXBuilder` objects. 

### DOCXBuilder Constructor

```js
var db = new OpenXmlBuilder.DOCXBuilder(OpenXmlB64Templates.docx, title, created, creator); 
```

_[`DOCXBuilder`]_ Create a DOCXBuilder based on the given base64 encoded `template`, 
optionally with a `title`, `created` date, and `creator` name.


### DOCXBuilder Methods

#### `db.pstyle(...)`

```js
  db.docLine(title, db.pStyle("Title")); 
```

_[`function`]_ Create a function which when called with the document will generate the specified paragraph style.  
The paragraph styles are specified in the styles panel of Word.  The particular styles may vary from 
template to template.  


#### `db.rstyle(...)`

```js
  db.docLine(title, null, db.rStyle("Italic")); 
```

_[`function`]_ Create a function which when called with the document will generate the specified run `style`.  The run
styles are specified in the styles panel of Word.  The particular styles may vary from template to 
template.  The special styles "Italic", "Bold", and "Underline" are also supported.  


#### `db.docLine(...)`

```js
  db.docLine(title, db.pStyle("Title"), db.rStyle("Italic")); 
```

_[`undefined`]_ Create a single line in the document for the given `html` content, `paragraphStyle` 
and `runStyle`.  The html should not be converted to multiple lines (i.e. no block element or a single 
p or div element).  


#### `db.docContent(...)`

```js
```

_[`undefined`]_ Insert a section of `html` into the document using the built-in converter.  This is 
probably a less reliable method than docContent and will drop some formatting but should produce cleaner
text that fits better with the template but may lead to carry-over of undesired formatting.  


#### `db.docChunk(...)`

```js
```

_[`undefined`]_ Insert a section of `html` into the document using Word's native ability to convert
html to text.  This is probably a more reliable method than docContent but may lead to carry-over 
of undesired formatting.  


