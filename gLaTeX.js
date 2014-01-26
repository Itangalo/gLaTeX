/**
 * Add a menu to the document.
 */
function onOpen() {
  DocumentApp.getUi().createMenu('LaTeX')
      .addItem('Insert LaTeX expression', 'latexDialog')
      .addToUi();
}

/**
 * Shows a sidebar where LaTeX can be entered.
 */
function latexDialog() {
  // Get the last added formula, or use the fallback if none exists.
  var fallback = '\\large\\left\\{\\begin{matrix}1x &+& 2x & =3\\\\ -x &+&y  & =7\\end{matrix}\\right.';
  var LaTeX = ScriptProperties.getProperty('latex') || fallback;

  var app = UiApp.createApplication().setTitle('LaTeX editor');
  var area = app.createTextArea().setName('formula').setText(LaTeX).setWidth('100%').setHeight('300px');
  app.add(area);

  var saveFormula = app.createServerHandler('saveFormula');
  saveFormula.addCallbackElement(area);
  app.add(app.createButton('Save formula', saveFormula));

  DocumentApp.getUi().showSidebar(app);
}

/**
 * Inserts the parsed LaTeX experssion to the document.
 */
function saveFormula(eventInfo) {
  // Save the experssion, so it can be fetched next time the sidebar opens.
  ScriptProperties.setProperty('latex', eventInfo.parameter.formula);

  var cursor = DocumentApp.getActiveDocument().getCursor();
  if (cursor) {
    // Attempt to insert an image at the cursor position. If insertion returns null,
    // then the cursor's containing element doesn't allow text insertions.
    var url = UrlFetchApp.fetch(encodeURI('http://latex.codecogs.com/gif.latex?' + eventInfo.parameter.formula));
    // This 'code' thing is an attempt to store the LaTeX expression with the
    // document element, so it can be edited later on. Work in progress.
    var code = {};
    code[DocumentApp.Attribute.CODE] = eventInfo.parameter.formula;
    // Actual insertion of LaTeX image happens here.
    var element = cursor.insertInlineImage(url.getBlob()).setAttributes(code);
  } else {
    DocumentApp.getUi().alert('Cannot find a cursor in the document.');
  }
}

// (For development only.)
function debug(variable) {
  Browser.msgBox(typeof variable, variable, {});
}
