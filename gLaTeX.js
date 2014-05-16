/**
 * Add a menu to the document.
 */
function onOpen() {
  DocumentApp.getUi().createMenu('LaTeX')
      .addItem('Insert/edit LaTeX expression', 'latexDialog')
      .addItem('What is this?', 'latexHelp')
      .addItem('Settings (alpha)','latexSettings')
      .addToUi();
}


/**
 * Fetches the current LaTeX expression display size, or uses a fallback.
 */
function latexFontSize()
{
  return PropertiesService.getUserProperties().getProperty('gLatexDefaultFontSize') || 'Huge';
}


/**
 * Shows a sidebar where LaTeX settings can be changed.
 */
function latexSettings() {
  // Change the settings for the script 
  // Options:    Default size of outputted latex
  // Possible latex settings: tiny, small, normal, large, Large, LARGE, huge, Huge
  
  var app = UiApp.createApplication().setTitle('LaTeX Settings');
  
  //Default Size Setting
  var lbSize = app.createListBox(false).setId('lbDefaultFontSize').setName('defaultFontSize');
  lbSize.addItem('tiny');
  lbSize.addItem('small');
  lbSize.addItem('normal');
  lbSize.addItem('large');
  lbSize.addItem('Large');
  lbSize.addItem('LARGE');
  lbSize.addItem('huge');
  lbSize.addItem('Huge');
  
  app.add(app.createLabel("Default Formula Size"));
  app.add(lbSize);
  
  var saveSettings = app.createServerHandler('saveSettings');
  saveSettings.addCallbackElement(lbSize);
  
  app.add(app.createButton('Save Settings', saveSettings));
  
  DocumentApp.getUi().showSidebar(app);
}

/**
 * Save setting changes to the script parameters.
 */
function saveSettings(eventInfo) {
  
  var selectedFontSize = eventInfo.parameter.defaultFontSize;
  
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('gLatexDefaultFontSize', selectedFontSize);
  
  return UiApp.getActiveApplication().close();
}

/**
 * Fetches the last entered LaTeX expression, or uses a fallback.
 */
function latestExpression() {
  return PropertiesService.getUserProperties().getProperty('gLatexLatest') || '\\left\\{\\begin{matrix}1x &+& 2x & =3\\\\ -x &+&y  & =7\\end{matrix}\\right.';
}

/**
 * Shows a sidebar where LaTeX can be entered.
 */
function latexDialog() {
  // Get the formula from any gLaTeX image being edited, or fall back to the latest used expression.
  var fontSize = latexFontSize();
  var imageprefix = 'http://www.texify.com/img/%5C'+fontSize+'%5C%21';
  var linkprefix = 'http://www.texify.com/%5C'+fontSize+'%5C%21';
  var suffix = '.gif';

  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection != null) {
    var link = decodeURI(selection.getSelectedElements()[0].getElement().getAttributes()[DocumentApp.Attribute.LINK_URL]);
    var formula = link.substring(linkprefix.length, link.length);
  }
  else {
    formula = latestExpression();
  }

  var app = UiApp.createApplication().setTitle('LaTeX editor');
  var area = app.createTextArea().setName('formula').setText(formula).setWidth('100%').setHeight('300px');
  app.add(area);

  var saveFormula = app.createServerHandler('saveFormula');
  saveFormula.addCallbackElement(area);
  app.add(app.createButton('Save and insert', saveFormula));
  
  DocumentApp.getUi().showSidebar(app);
}

/**
 * Inserts the parsed LaTeX experssion to the document.
 */
function saveFormula(eventInfo) {
  var fontSize = latexFontSize();
  var imageprefix = 'http://www.texify.com/img/%5C'+fontSize+'%5C%21';
  var linkprefix = 'http://www.texify.com/%5C'+fontSize+'%5C%21';
  var suffix = '.gif';

  // Check if there is something selected, that should be replaced.
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection != null) {
    // We will lose the position of the cursor when removing the selected elements, so we record a new position for the cursor here.
    var position = DocumentApp.getActiveDocument().newPosition(selection.getRangeElements()[0].getElement().getParent(), 1);
    // Now, go ahead and remove all the selected elements.
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      elements[i].getElement().removeFromParent();
    }
    // Set the new cursor position, as is nothing happened.
    DocumentApp.getActiveDocument().setCursor(position);
  }

  // Insert the image at the cursor.
  var cursor = DocumentApp.getActiveDocument().getCursor();
  var formula = eventInfo.parameter.formula
  PropertiesService.getUserProperties().setProperty('gLatexLatest', formula); //!!
  var image = UrlFetchApp.fetch(encodeURI(imageprefix + formula.replace('\n', ' ') + suffix)).getBlob();
  // We link the image to the web service generating the image. This is not only to be nice,
  // it is also how the expression is saved in raw format -- allowing us to open and edit it.
  var attributes = {};
  attributes[DocumentApp.Attribute.LINK_URL] = encodeURI(linkprefix + eventInfo.parameter.formula);

  cursor.insertInlineImage(image).setAttributes(attributes);
  
  return UiApp.getActiveApplication().close();
}

// Displays a popup with help.
function latexHelp() {  
  var app = UiApp.createApplication().setTitle('gLaTeX (work in progress)');
  var message = 'gLaTeX is a tool for inserting LaTeX-experssions into Google documents, thereby allowing complex formulas and stuff. \n\n You can find the source code at https://github.com/Itangalo/gLaTeX. A demo video is found at http://www.youtube.com/watch?v=bV75ZlmZit4 \n\n The formula images are generated by texify: http://www.texify.com/. Many thanks to the people running the service.';
  app.add(app.createHTML(message).setWidth('100%'));

  DocumentApp.getUi().showSidebar(app);

}

// (For development only.)
function dev() {
}

function devHandler(eventInfo) {
}

// (For development only.)
function debug(variable) {
  var app = UiApp.createApplication().setTitle('Debug');
  var message = 'type: ' + typeof variable + '<br /><br />';
  message += 'content: ' + variable + '<br /><br />';

  if (typeof variable == 'object' || typeof variable == 'array') {
    for (var i in variable) {
      message += i + ': ' + variable[i] + '<br /><br />';
    }
  }
  app.add(app.createHTML(message).setWidth('100%'));

  DocumentApp.getUi().showSidebar(app);
}
