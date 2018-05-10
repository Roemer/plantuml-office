'use strict';

var FS = require('fs'),
  PATH = require('path'),
  SVGO = require('../lib/svgo'),
  dir = PATH.resolve(__dirname, "../../Clouds/"),
  svgo = new SVGO({
    js2svg: { pretty: true, indent: '  ' },
    plugins: [{
      cleanupAttrs: true,
    }, {
      removeDoctype: true,
    }, {
      removeXMLProcInst: true,
    }, {
      removeComments: true,
    }, {
      removeMetadata: true,
    }, {
      removeTitle: true,
    }, {
      removeDesc: true,
    }, {
      removeUselessDefs: true,
    }, {
      removeEditorsNSData: true,
    }, {
      removeEmptyAttrs: true,
    }, {
      removeHiddenElems: true,
    }, {
      removeEmptyText: true,
    }, {
      removeEmptyContainers: true,
    }, {
      removeViewBox: true,
    }, {
      cleanUpEnableBackground: true,
    }, {
      convertStyleToAttrs: true,
    }, {
      convertColors: true,
    }, {
      convertPathData: true,
    }, {
      convertTransform: true,
    }, {
      removeUnknownsAndDefaults: true,
    }, {
      removeNonInheritableGroupAttrs: true,
    }, {
      removeUselessStrokeAndFill: true,
    }, {
      removeUnusedNS: true,
    }, {
      cleanupIDs: true,
    }, {
      cleanupNumericValues: true,
    }, {
      moveElemsAttrsToGroup: true,
    }, {
      moveGroupAttrsToElems: true,
    }, {
      collapseGroups: true,
    }, {
      removeRasterImages: false,
    }, {
      mergePaths: true,
    }, {
      convertShapeToPath: true,
    }, {
      sortAttrs: false,
    }, {
      transformsWithOnePath: true,
    }, {
      removeDimensions: false,
    }]
  });

FS.readdir(dir, (err, files) => {
  files.forEach(file => {
    var filepath = PATH.resolve(__dirname, dir + "\\" + file);
    console.log(filepath);
    FS.readFile(filepath, 'utf8', function (err, data) {

      if (err) {
        throw err;
      }

      svgo.optimize(data, { path: filepath }).then(function (result) {

        console.log(result.info);
        FS.writeFile(filepath, result.data);

      });

    });
  });
})
