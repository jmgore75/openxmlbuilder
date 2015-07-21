# Overview

The OpenXmlBuilder library allows you to build simple OpenXml documents right in the browser.  It is small and compatible with both modern browsers and Internet Explorer in compatibility view.   

## Dependencies

OpenXmlBuilder requires the [JSZip](https://stuk.github.io/jszip/) library.  

## Setup

To use the library, simply include the following JavaScript declarations in your page:

```html
<script type="text/javascript" src="http://cdnjs.cloudflare.com/ajax/libs/jszip/2.5.0/jszip.min.js"></script>
<script type="text/javascript" src="OpenXmlBuilder.js"></script>
<script type="text/javascript" src="OpenXmlB64Templates.js"></script>
```

OpenXmlBuilder expects you to provide template documents in base64 encoded format.  
Template documents are provided for you in OpenXmlB64Templates.js, but you may wish to generate your 
own templates in base64 format, or even load templates dynamically.  

## API

For the full API documentation, see [api/OpenXmlBuilder.md](api/OpenXmlBuilder.md). The full set of
[Configuration Options](api/OpenXmlBuilder.md#configuration-options) are also documented there.

### Creating a PowerPoint document

  ```js
    function pptReport(title, sections) {
        var created = new Date(); 
        var creator = "OXB"; 

        var pb = new OpenXmlBuilder.PPTXBuilder(OpenXmlB64Templates.pptx, title, created, creator); 
        pb.contentSlide({"Title 1":"<span style='color:red'>DELETE THIS SLIDE</span>", "Subtitle 2" : "Delete this slide to ensure that text content is scaled to fit within the presentation." }, 1); 
        pb.contentSlide({"Title 1":title, "Subtitle 2" : "Created " + created.toString() }, 1); 
        var i, section; 
        for (i = 0; i < sections.length; i++) {
            section = sections[i]; 
            if (section.summary) {
                pb.contentSlide({"Title 1":"<a href='" + section.url + "'>"+section.title+"</a>", "Content Placeholder 2" : section.summary}); 
            }
        }
        return pb.saveToBlob(); 
    }
  ```

### Creating a Word document

  ```js
    function docReport(title, sections) {
        var created = new Date(); 
        var creator = "OXB"; 

        var db = new OpenXmlBuilder.DOCXBuilder(OpenXmlB64Templates.docx, title, created, creator); 
        var heading = db.pStyle("Heading1"); 
        db.docLine(title, db.pStyle("Title")); 
        db.docLine("Created " + created.toString(), db.pStyle("Subtitle")); 
        
        var i, section; 
        for (i = 0; i < sections.length; i++) {
            section = sections[i]; 
            if (section.summary) {
                db.docLine("<a href='" + section.url + "'>"+section.title+"</a>", heading); 
                db.docChunk(section.summary); 
            }
        }
        return db.saveToBlob(); 
    }
  ```


### Creating an Excel document

This is not yet supported.  


### Saving the results

Internally you are building a zip file, which may then be used to generate base64 or a blob.  

You must first save changes, then generate the data to save.  However, there is no standard 
method for saving files from the browser, so you must choose from one of several options.  

#### Using saveAs via [FileSaver.js](https://github.com/eligrey/FileSaver.js)

This is the preferred method of saving files in modern browsers, which exploits a variety of methods depending on the browser.  

  ```js
    saveAs(pb.saveToBlob(), "example.ppt"); 
  ```

#### Generating a dataURI (not recommended)

An even simpler approach but one which is poorly supported is the use of data URIs.  
Data uris are not generally supported in IE, and should generally be avoided. 

A data uri may be used to generate a hyperlink that saves the document: 

  ```js
    var a= document.createElement("A"); 
    a.href = pb.saveToDataURI(); 
    a.download = "example.ppt"; 
    a.innerHTML = "Download File"; 
    document.body.append(a); 
  ```

Alternatively you may tell the browser to download the file (but will not be able to specify a filename): 

  ```js
    document.location.href = pb.saveToDataURI(); 
  ```
#### Using a Flash-based solution

For legacy browsers your only option may be to use a flash-based solution such as [ClipAndSave](https://github.com/jmgore75/clipandsave).


