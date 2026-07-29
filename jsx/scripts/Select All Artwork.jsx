//#target illustrator

if(!app.documents.length) {
  'No open document.';
} else {
  app.activeDocument.selectObjectsOnActiveArtboard();
  'Selected artwork on the active artboard.';
}
