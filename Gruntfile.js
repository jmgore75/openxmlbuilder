/*jshint -W106 */
/*jshint node:true, maxstatements: false, maxlen: false */

var os = require("os");
var path = require("path");

module.exports = function(grunt) {
  "use strict";

  // Metadata
  var pkg = grunt.file.readJSON("package.json");

  // Shared configuration
  var localPort = 7320;  // "ZERO"

  // Project configuration.
  var config = {
    // Task configuration
    jshint: {
      options: {
        jshintrc: true
      },
      Gruntfile: ["Gruntfile.js"],
      js: ["src/js/**/*.js", "!src/js/start.js", "!src/js/end.js"],
      test: ["test/**/*.js"],
      dist: ["dist/*.js", "!dist/*.min.js"]
    },
    clean: {
      dist: ["OpenXmlBuilder.*", "dist/OpenXmlBuilder.*"],
      meta: ["bower.json", "composer.json", "LICENSE"]
    },
    b64js: {
      client: {
        properties: {
          "docx": "src/templates/template.docx", 
          "pptx": "src/templates/template.pptx"
        }, 
        object : "OpenXmlB64Templates", 
        dest: "dist/OpenXmlB64Templates.js"
      }
    }, 
    concat: {
      options: {
        stripBanners: false,
        process: {
          data: pkg
        }
      },
      client: {
        src: [
          "src/meta/source-banner.tmpl",
          "src/js/start.js",
          "src/js/core.js",
          "src/js/html_convert.js",
          "src/js/pptx.js",
          "src/js/docx.js",
          "src/js/end.js"
        ],
        dest: "dist/OpenXmlBuilder.js"
      }
    },
    uglify: {
      options: {
        report: "min"
      },
      js: {
        options: {
          preserveComments: function(node, comment) {
            return comment &&
              comment.type === "comment2" &&
              /^(!|\*|\*!)\r?\n/.test(comment.value);
          },
          beautify: {
            beautify: true,
            // `indent_level` requires jshint -W106
            indent_level: 2
          },
          mangle: false,
          compress: false
        },
        files: [
          {
            src: ["<%= concat.client.dest %>"],
            dest: "<%= concat.client.dest %>"
          }
        ]
      },
      minjs: {
        options: {
          preserveComments: function(node, comment) {
            return comment &&
              comment.type === "comment2" &&
              /^(!|\*!)\r?\n/.test(comment.value);
          },
          sourceMap: true,
          // Bundles the contents of "`src`" into the "`dest`.map" source map file. This way,
          // consumers only need to host the "*.min.js" and "*.min.map" files rather than
          // needing to host all three files: "*.js", "*.min.js", and "*.min.map".
          sourceMapIncludeSources: true
        },
        files: [
          {
            src: ["<%= concat.client.dest %>"],
            dest: "dist/OpenXmlBuilder.min.js"
          }
        ]
      }
    },
    template: {
      options: {
        data: pkg
      },
      bower: {
        files: {
          "bower.json": ["src/meta/bower.json.tmpl"]
        }
      },
      composer: {
        files: {
          "composer.json": ["src/meta/composer.json.tmpl"]
        }
      },
      LICENSE: {
        files: {
          "LICENSE": ["src/meta/LICENSE.tmpl"]
        }
      }
    },
    chmod: {
      options: {
        mode: "444"
      },
      dist: ["dist/OpenXmlBuilder.*"],
      meta: ["bower.json", "composer.json", "LICENSE"]
    },
    connect: {
      server: {
        options: {
          port: localPort
        }
      }
    },
    qunit: {
      file: [
        "test/shared/private.tests.js.html",
        "test/client/private.tests.js.html",
        "test/client/api.tests.js.html",
        "test/built/OpenXmlBuilder.tests.js.html"
        //"test/**/*.tests.js.html"
      ],
      http: {
        options: {
          urls:
            grunt.file.expand([
              "test/shared/private.tests.js.html",
              "test/core/private.tests.js.html",
              "test/core/api.tests.js.html",
              "test/client/private.tests.js.html",
              "test/client/api.tests.js.html",
              "test/built/OpenXmlBuilder.tests.js.html"
              //"test/**/*.tests.js.html"
            ]).map(function(testPage) {
              return "http://localhost:" + localPort + "/" + testPage + "?noglobals=true";
            })
        }
      }
    },
    watch: {
      options: {
        spawn: false
      },
      Gruntfile: {
        files: "<%= jshint.Gruntfile %>",
        tasks: ["jshint:Gruntfile"]
      },
      js: {
        files: "<%= jshint.js %>",
        tasks: ["jshint:js", "unittest"]
      },
      test: {
        files: "<%= jshint.test %>",
        tasks: ["jshint:test", "unittest"]
      }
    }
  };
  grunt.initConfig(config);

  // These plugins provide necessary tasks
  grunt.loadNpmTasks("grunt-contrib-jshint");
  grunt.loadNpmTasks("grunt-contrib-clean");
  grunt.loadNpmTasks("grunt-contrib-concat");
  grunt.loadNpmTasks("grunt-contrib-uglify");
  grunt.loadNpmTasks("grunt-template");
  grunt.loadNpmTasks("grunt-chmod");
  grunt.loadNpmTasks("grunt-contrib-connect");
  grunt.loadNpmTasks("grunt-contrib-qunit");
  grunt.loadNpmTasks("grunt-contrib-watch");

  // Building b64 template
  grunt.registerMultiTask("b64js", "B64 files into javascript strings", function () {
    var options = this.options({});
    var v, b64, filepath; 

    var output = "var " + this.data.object + " = {};\n"; 
    // Iterate over all specified file groups.
    for (var v in this.data.properties) {
      filepath = this.data.properties[v]; 
      b64 = grunt.file.read(filepath, {
          encoding: null
        }).toString('base64');
      output += this.data.object + '.' + v + '="'; 
      output += b64; 
      output += '";\n'; 
      grunt.log.writeln('Base64 encoded "' + filepath + '".');
    }
    grunt.file.write(this.data.dest, output);
    grunt.log.writeln('Built ' + this.data.dest);
  }); 

  // Task aliases and chains
  grunt.registerTask("jshint-prebuild", ["jshint:Gruntfile", "jshint:js", "jshint:test"]);
  grunt.registerTask("validate",        ["jshint-prebuild"]);
  grunt.registerTask("build",           ["clean", "b64js", "concat", "jshint:dist", "uglify", "template", "chmod"]);
  grunt.registerTask("build-travis",    ["clean:dist", "b64js", "concat", "jshint:dist", "chmod:dist"]);
  grunt.registerTask("test",            ["connect", "qunit"]);

  // Default task
  grunt.registerTask("default", ["validate", "build", "test"]);
  // Travis CI task
  grunt.registerTask("travis",  ["validate", "build-travis", "test"]);

};
