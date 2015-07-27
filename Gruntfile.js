/*jshint -W106 */
/*jshint node:true, maxstatements: false, maxlen: false */

var os = require("os");
var path = require("path");

function extend(tgt, src) {
  var v;
  tgt = tgt || {};
  for (v in src) {
    if (src.hasOwnProperty(v) && !tgt.hasOwnProperty(v)) {
      tgt[v] = src[v];
    }
  }
  return tgt;
}

module.exports = function(grunt) {
  "use strict";

  var b64 = function(filepath) {
    var output = grunt.file.read(filepath, {
      encoding: null
    }).toString("base64");
    grunt.log.writeln("Base64 encoded \"" + filepath + "\".");
    return output;
  };

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
      },
      b64Templates : {
        options : {
          process : {
            data : extend({
              object : "OpenXmlB64Templates",
              b64 : b64,
              templates : [
                {name:"pptx", src:"src/templates/template.pptx"},
                {name:"docx", src:"src/templates/template.docx"}
              ]
            }, pkg)
          }
        },
        files: {
          "dist/OpenXmlB64Templates.js" : ["src/meta/source-banner.tmpl", "src/meta/b64Lib.tmpl"]
        }
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

  // Task aliases and chains
  grunt.registerTask("jshint-prebuild", ["jshint:Gruntfile", "jshint:js", "jshint:test"]);
  grunt.registerTask("validate",        ["jshint-prebuild"]);
  grunt.registerTask("build",           ["clean", "concat", "jshint:dist", "uglify", "template", "chmod"]);
  grunt.registerTask("test",            ["connect", "qunit"]);

  // Default task
  grunt.registerTask("default", ["validate", "build", "test"]);

};
