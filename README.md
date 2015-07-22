# OpenXmlBuilder

The OpenXmlBuilder library allows you to build simple OpenXml documents right in the browser.  It is small and compatible with both modern browsers and Internet Explorer in compatibility view.  Currently it supports PowerPoint and Word documents.  

## Dependencies

OpenXmlBuilder requires the [JSZip](https://stuk.github.io/jszip/) library.  

## Simple Example

```js
  var title = "New Report"; 
  var created = new Date(); 
  var creator = "Me"; 
  var sampleHtml = "<p>Body text with <b>bold</b> <i>italic</i> <span style='color:red'>color</span> <a href='http://www.google.com'>hyperlink</a>. </p><ul><li>i1<br>i1 cont</li><li>i2<ul><li>i2.1</li></ul></li></ul>"; 

// pptx.js
  var pb = new OpenXmlBuild.PPTBuilder(OpenXmlB64Templates.pptx, title, created, creator); 
  pb.contentSlide({"Title 1":"Delete this slide", 
    "Subtitle 2" : "Delete this slide to ensure that contents are scaled to fit within the slides." }, 1); 
  pb.contentSlide({"Title 1":title, "Subtitle 2" : "Created " + created.toString() }, 1); 
  pb.contentSlide({"Title 1":"A question", "Content Placeholder 2" : sampleHtml}, 2); 
  return pb.saveToBase64(); 

// doc.js
  var db = new OpenXmlBuild.DOCBuilder(OpenXmlB64Templates.docx, title, created, creator); 
  db.docLine(title, db.pStyle("Title")); 
  db.docLine("Created " + created.toString(), db.pStyle("Subtitle")); 
  db.docLine("HTML translated to native style", db.pStyle("Heading1")); 
  db.docContent(sampleHtml)
  db.docLine("HTML imported as chunk", db.pStyle("Heading1")); 
  db.docChunk(sampleHtml)
  return db.saveToBase64(); 
```

See [docs/instructions.md](docs/instructions.md) for more advanced options in using the library on your site.
See [docs/api/OpenXmlBuilder.md](docs/api/OpenXmlBuilder.md) for the complete API documentation.


## Contributing

see [CONTRIBUTING.md](CONTRIBUTING.md)


## Releases

see [releases](https://github.com/jmgore75/openxmlbuilder/releases)

