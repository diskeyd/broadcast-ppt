// Read the text off an image with the macOS recogniser, one line per output line.
//   osascript -l JavaScript scripts/ocr_vision.js <image>...
// Used to identify a hymn score page; nothing is installed for it and it only runs on the Mac.
ObjC.import('Vision');
ObjC.import('Foundation');

function run(argv) {
  var out = [];
  argv.forEach(function (p) {
    var req = $.VNRecognizeTextRequest.alloc.init;
    req.recognitionLanguages = $(['ko-KR', 'en-US']);
    req.recognitionLevel = 0; // accurate
    var handler = $.VNImageRequestHandler.alloc.initWithURLOptions($.NSURL.fileURLWithPath(p), $());
    handler.performRequestsError($([req]), null);
    var res = req.results;
    for (var i = 0; i < res.count; i++) {
      out.push(ObjC.unwrap(res.objectAtIndex(i).topCandidates(1).objectAtIndex(0).string));
    }
  });
  return out.join('\n');
}
