function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Format Markdown')
      .addItem('Fix All', 'fixMarkdown')
      .addItem('Highlight Gitbook links', 'highlightGitbookLinks')
      .addToUi();
}

var doc;
var body;
var searchType;
var searchResult;
var imgSrc;
var imgAlt;
var docTitle;
var foundTag;
var foundIconTag;
var foundIcon;
var isTable;
var isDiv ;
var isCode ;
var codeType ;
var rangeBuilder;
var par;
var nextPar;
var lastPar;
var parText;

function fixMarkdown() {
  doc = DocumentApp.getActiveDocument();
  body = doc.getBody();
  searchType = DocumentApp.ElementType.PARAGRAPH;
  searchResult = null;
  imgSrc = null;
  imgAlt = null;
  docTitle = null;
  foundTag = null;
  foundIconTag = null;
  foundIcon = null;
  isTable = false;
  isDiv = false;
  isCode = false;
  codeType = null;
  rangeBuilder = doc.newRange();
  while (searchResult = body.findElement(searchType, searchResult)) {
    par = searchResult.getElement().asParagraph();
    nextPar = null;
    lastPar = null;
    parText = par.getText();
    if (!docTitle) {
      getDocTitle();
    }
    checkIfCode();
    // highlightTODO();
    checkIfDiv();
    checkIfTable();
    // putCodeInBackticks();
    addSpacesInCode();
    processImageInfo();
    removeImageWarning();
    placeIcons();
    fixExternalLinksInDivs();
    fixInternalLinksInDivs();
    fixCodeInDivs();
    fixBoldInDivs();
    fixItalicsInDivs();
    fixExternalLinksContainingParens();
    fixCodefontedExternalLinks();
    fixCodefontedInternalLinks();
    removeExtraLinesInListsAndDivs();
    removeTags();
    // fixAnchors();
    // fixIndentedCode();
    // highlightCode();
    // addImageBorderAndCenter();
    // formatCodeAndTables();
  }
}

function highlightGitbookLinks() {
  doc = DocumentApp.getActiveDocument();
  body = doc.getBody();
  searchType = DocumentApp.ElementType.PARAGRAPH;
  searchResult = null;
  imgSrc = null;
  imgAlt = null;
  docTitle = null;
  foundTag = null;
  foundIconTag = null;
  foundIcon = null;
  isTable = false;
  isDiv = false;
  isCode = false;
  codeType = null;
  rangeBuilder = doc.newRange();
  while (searchResult = body.findElement(searchType, searchResult)) {
    par = searchResult.getElement().asParagraph();
    nextPar = null;
    lastPar = null;
    parText = par.getText();

    var foundMd = allIndexOf(parText, '.md');
    var foundGitbookIO = allIndexOf(parText, 'gitbooks.io');

    for (var i = 0; i < foundMd.length; i++) {
      par.editAsText().setBackgroundColor(foundMd[i], foundMd[i] + 2,
        '#ff0000');
    }

    for (var i = 0; i < foundGitbookIO.length; i++) {
      par.editAsText().setBackgroundColor(foundGitbookIO[i],
        foundGitbookIO[i] + 10, '#ff0000');
    }
  }
}

function allIndexOf(str, toSearch) {
  var indices = [];
  for (var pos = str.indexOf(toSearch); pos !== -1;
    pos = str.indexOf(toSearch, pos + 1)) {
    indices.push(pos);
  }
  return indices;
}

function getDocTitle() {
  if (parText.indexOf('# ') === 0) {
    docTitle = parText.substring(2);
    Logger.log('DOCTITLE:   ' + docTitle);
  }
}

function checkIfCode() {
  if (!isCode) {
    if (parText.indexOf('```') === 0) {
      isCode = true;
    }
  } else {
    if (parText.indexOf('```') === 0) {
      isCode = false;
    }
  }
}

function highlightTODO() {
  if (parText.indexOf('TODO') > -1 && !isCode) {
    nextPar = par.getNextSibling();
    par.editAsText().insertText(0, '<div class=\"todo\">');
    parText = par.getText();
    nextPar.editAsText().appendText('</div>\n');
  }
}

function checkIfDiv() {

  // This should be after highlightTODO()
  // TODO </div>s are in the same paragraph as the opening <div>,
  // so isDiv will always be true after TODO paragraphs

  if (!isCode && parText.indexOf('<div') === 0) {
    isDiv = true;
  }
  if (!isCode && parText.indexOf('</div>') === 0) {
    isDiv = false;
  }
}

function checkIfTable() {
  if (!isCode && parText.indexOf('<table>') === 0) {
    isTable = true;
  }
  if (!isCode && parText.indexOf('</table>') === 0) {
    isTable = false;
  }
}

function putCodeInBackticks() {
  while (parText.indexOf('<code>') > -1 && !isTable && !isDiv) {
    par.replaceText('<code>', '`');
    parText = par.getText();
  }

  while (parText.indexOf('</code>') > -1 && !isTable && !isDiv) {
    par.replaceText('</code>', '`');
    parText = par.getText();
  }
}

function addSpacesInCode() {
  if (parText == '' && isCode) {
    par.editAsText().insertText(0,  ' ');
    parText = par.getText();
  }
}

function processImageInfo() {
  if (parText.indexOf('!\[alt_text\]') === 0) {
    var nextPar = par.getNextSibling();
    var nextParText = nextPar.asParagraph().getText();
    while (nextParText.indexOf('IMAGEINFO') == -1) {
      nextPar = nextPar.getNextSibling();
      if (nextPar != null) {
        nextParText = nextPar.asParagraph().getText();
      }
    }
    var imgInfo = nextParText;
    if (imgInfo.indexOf('IMAGEINFO') > -1) {
      imgInfo = imgInfo.replace('\\', '');
      imgInfoArray = imgInfo.substring(imgInfo.indexOf(':') + 2).split(', ');
    } else if (par.getNextSibling().getNextSibling().asParagraph().getText()
    .indexOf('IMAGEINFO') > -1) {
      imgInfo = imgInfo.getNextSibling().asParagraph().getText();
      imgInfoArray = imgInfo.substring(imgInfo.indexOf(':') + 2).split(', ');
    } else {
      imgInfoArray[0] = 'No Source Provided';
      imgInfoArray[1] = 'No Title Provided';
    }

    imgSrc = '../images/'.concat(docTitle).concat('/').concat(imgInfoArray[0]);
    imgAlt = imgInfoArray[1];
    nextPar.clear();
    nextParText = null;
    par.clear();
    par.appendText('<img src=\"' + imgSrc + '\" alt=\"' + imgAlt +
    '\" title=\"' + imgAlt + '\">');
    imgSrc = imgAlt = null;
  }
}

function removeImageWarning() {
  if (parText.indexOf('style=\"color: red') > -1) {
    par.clear();
    par = par.getNextSibling();
  }
}

function placeIcons() {
  parText = par.getText();
  if (parText.indexOf('[ICON HERE]') > -1) {
    foundIconTag = par.findText('\\[ICON HERE\\]');
  }
  if (parText.indexOf('<img') > -1 && parText.indexOf('ic_') > -1) {
    foundIcon = par.findText('<img');
  }

  if (foundIconTag != null && foundIcon != null) {
    var tagStart = foundIconTag.getStartOffset();
    var tagEnd = foundIconTag.getEndOffsetInclusive() ;
    foundIconTag.getElement().editAsText().deleteText(tagStart, tagEnd);
    foundIconTag.getElement().editAsText().insertText(tagStart,
      foundIcon.getElement().asText().getText());
    foundIcon.getElement().removeFromParent();
    foundIconTag = null;
    foundIcon = null;
  }
}

function fixExternalLinksInDivs() {

  // TODO - handle the inline `<div>` case

  if (parText.indexOf('<div') > -1 && !isCode) {
    isDiv = true;
    nextPar = par.getNextSibling();
    if (parText.indexOf('<div class=\"todo\"') > -1) {
      nextPar = par;
    }
    while (isDiv) {
      if (nextPar.getText().indexOf('</div>') === 0) {
        isDiv = false;
      }
      while (nextPar.getText().indexOf('\(http') > -1) {
        var foundLinkText = nextPar.findText('\\[.*?]');
        var foundLink = nextPar.findText('\\(http([^\\s]*?\\))*');
        if (foundLink != null && foundLinkText != null) {
          var linkText = nextPar.getText().substring(foundLinkText
            .getStartOffset() + 1, foundLinkText.getEndOffsetInclusive());
          var link = nextPar.getText().substring(foundLink.getStartOffset() + 1,
          foundLink.getEndOffsetInclusive());
          nextPar.editAsText().deleteText(foundLinkText.getStartOffset(),
          foundLink.getEndOffsetInclusive());
          var builtLink = '<a href=\"' + link + '\">' + linkText + '</a>';
          nextPar.editAsText().insertText(foundLinkText.getStartOffset(),
          builtLink);
        }
      }
      nextPar = nextPar.getNextSibling();
    }
    nextPar = null;
  }
}

function fixInternalLinksInDivs() {

  // TODO - handle the inline `<div>` case

  if (parText.indexOf('<div') > -1 && !isCode) {
    isDiv = true;
    nextPar = par.getNextSibling();
    if (parText.indexOf('<div class=\"todo\"') > -1) {
      nextPar = par;
    }
    while (isDiv) {
      if (nextPar.getText().indexOf('</div>') === 0) {
        isDiv = false;
      }
      while (nextPar.getText().indexOf('\(#') > -1) {
        var foundLinkText = nextPar.findText('\\[.*?]');
        var foundLink = nextPar.findText('\\(#.*?\\)');
        if (foundLink != null && foundLinkText != null) {
          var linkText = nextPar.getText().substring(foundLinkText
            .getStartOffset() + 1, foundLinkText.getEndOffsetInclusive());
          var link = nextPar.getText().substring(foundLink.getStartOffset() + 1,
          foundLink.getEndOffsetInclusive());
          nextPar.editAsText().deleteText(foundLinkText.getStartOffset(),
          foundLink.getEndOffsetInclusive());
          var builtLink = '<a href=\"' + link + '\">' + linkText + '</a>';
          nextPar.editAsText().insertText(foundLinkText.getStartOffset(),
          builtLink);
        } else {
          break;
        }
      }
      nextPar = nextPar.getNextSibling();
    }
    nextPar = null;
  }
}

function fixCodeInDivs() {

  // TODO - handle the inline `<div>` case

  if (parText.indexOf('<div') > -1 && !isCode) {
    isDiv = true;
    nextPar = par.getNextSibling();
    if (parText.indexOf('<div class=\"todo\"') > -1) {
      nextPar = par;
    }
    while (isDiv) {
      if (nextPar.getText().indexOf('</div>') === 0) {
        isDiv = false;
      }
      while (nextPar.getText().indexOf('`') > -1) {
        var foundCodeText = nextPar.findText('\\`(.*?)\\`');
        if (foundCodeText != null) {
          var codeText = nextPar.getText().substring(foundCodeText
            .getStartOffset() + 1, foundCodeText.getEndOffsetInclusive());
          nextPar.editAsText().deleteText(foundCodeText.getStartOffset(),
          foundCodeText.getEndOffsetInclusive());
          var replaceCodeText = '<code>' + codeText + '</code>';
          nextPar.editAsText().insertText(foundCodeText.getStartOffset(),
          replaceCodeText);
        } else {
          break;
        }
      }
      nextPar = nextPar.getNextSibling();
    }
    nextPar = null;
  }
}

function fixBoldInDivs() {

  // TODO - handle the inline `<div>` case

  if (parText.indexOf('<div') > -1 && !isCode) {
    isDiv = true;
    nextPar = par.getNextSibling();
    if (parText.indexOf('<div class=\"todo\"') > -1) {
      nextPar = par;
    }
    while (isDiv) {
      if (nextPar.getText().indexOf('</div>') === 0) {
        isDiv = false;
      }
      while (nextPar.getText().indexOf('**') > -1) {
        var foundBoldText = nextPar.findText('\\*\\*(.*?)\\*\\*');
        if (foundBoldText != null) {
          var boldText = nextPar.getText().substring(foundBoldText
            .getStartOffset() + 2, foundBoldText.getEndOffsetInclusive() - 1);
          nextPar.editAsText().deleteText(foundBoldText.getStartOffset(),
          foundBoldText.getEndOffsetInclusive());
          var replaceBoldText = '<strong>' + boldText + '</strong>';
          nextPar.editAsText().insertText(foundBoldText.getStartOffset(),
          replaceBoldText);
        } else {
          break;
        }
      }
      nextPar = nextPar.getNextSibling();
    }
    nextPar = null;
  }
}

function fixItalicsInDivs() {

  // TODO - handle the inline `<div>` case

  if (parText.indexOf('<div') > -1 && !isCode) {
    isDiv = true;
    nextPar = par.getNextSibling();
    if (parText.indexOf('<div class=\"todo\"') > -1) {
      nextPar = par;
    }
    while (isDiv) {
      if (nextPar.getText().indexOf('</div>') === 0) {
        isDiv = false;
      }
      while (nextPar.getText().indexOf('*') > -1) {
        var foundItalicizedText = nextPar.findText('\\*(.*?)\\*');
        if (foundItalicizedText != null) {
          var italicizedText = nextPar.getText().substring(foundItalicizedText
            .getStartOffset() + 1, foundItalicizedText.getEndOffsetInclusive());
          nextPar.editAsText().deleteText(foundItalicizedText.getStartOffset(),
          foundItalicizedText.getEndOffsetInclusive());
          var replaceItalicizedText = '<em>' + italicizedText + '</em>';
          nextPar.editAsText().insertText(foundItalicizedText.getStartOffset(),
          replaceItalicizedText);
        } else {
          break;
        }
      }
      nextPar = nextPar.getNextSibling();
    }
    nextPar = null;
  }
}

function fixExternalLinksContainingParens() {
  while (parText.indexOf('\(http') > -1) {
    var foundLinkText = par.findText('\\[.*?]');
    var foundLink = par.findText('\\(http.*?\\)([^\\s]+?\\))+');
    var linkContainsParens = par.findText('\\(http([^\\s]*?)\\(');
    if (linkContainsParens != null && foundLink != null &&
      foundLinkText != null) {
      var linkText = parText.substring(foundLinkText.getStartOffset() + 1,
      foundLinkText.getEndOffsetInclusive());
      var link = parText.substring(foundLink.getStartOffset() + 1,
      foundLink.getEndOffsetInclusive());
      par.editAsText().deleteText(foundLinkText.getStartOffset(),
      foundLink.getEndOffsetInclusive());
      var builtLink = '<a href=\"' + link + '\">' + linkText + '</a>';
      par.editAsText().insertText(foundLinkText.getStartOffset(), builtLink);
      parText = par.getText();
    } else {
      break;
    }
  }
}

function fixCodefontedExternalLinks() {
  while ((parText.indexOf('`[') > -1 && parText.indexOf('\(http') > -1)) {
    var foundLinkText = par.findText('`\\[.*?]');
    var foundLink = par.findText('\\(http([^\\s]*?\\))*');
    if (foundLink != null && foundLinkText != null) {
      var linkText = parText.substring(foundLinkText.getStartOffset() + 2,
      foundLinkText.getEndOffsetInclusive());
      var link = parText.substring(foundLink.getStartOffset() + 1,
      foundLink.getEndOffsetInclusive());
      Logger.log('LINK: ' + link + ' LINKTEXT: ' + linkText + ' FULLREMOVE: ' +
      parText.substring(foundLinkText.getStartOffset(),
      foundLink.getEndOffsetInclusive() + 1));
      if (link.indexOf('\(') > -1) {
        link = parText.substring(foundLink.getStartOffset() + 1,
        foundLink.getEndOffsetInclusive() + 1);
        par.editAsText().deleteText(foundLinkText.getStartOffset(),
        foundLink.getEndOffsetInclusive() + 2);
      } else {
        par.editAsText().deleteText(foundLinkText.getStartOffset(),
        foundLink.getEndOffsetInclusive() + 1);
      }
      var builtLink = '<a href=\"' + link + '\">`' + linkText + '`</a>';
      par.editAsText().insertText(foundLinkText.getStartOffset(), builtLink);
      parText = par.getText();
    }
  }
}

function fixCodefontedInternalLinks() {
  while ((parText.indexOf('`[') > -1 && parText.indexOf('\(#') > -1)) {
    var foundLinkText = par.findText('`\\[.*?]');
    var foundLink = par.findText('\\(#.*?\\)');
    if (foundLink != null && foundLinkText != null) {
      var linkText = parText.substring(foundLinkText.getStartOffset() + 2,
      foundLinkText.getEndOffsetInclusive());
      var link = parText.substring(foundLink.getStartOffset() + 1,
      foundLink.getEndOffsetInclusive());
      Logger.log('LINK: ' + link + ' LINKTEXT: ' + linkText);
      if (link.indexOf('\(') > 0) {
        link = parText.substring(foundLink.getStartOffset() + 1,
        foundLink.getEndOffsetInclusive() + 1);
        par.editAsText().deleteText(foundLinkText.getStartOffset(),
        foundLink.getEndOffsetInclusive() + 2);
      } else {
        par.editAsText().deleteText(foundLinkText.getStartOffset(),
        foundLink.getEndOffsetInclusive() + 1);
      }
      var builtLink = '<a href=\"' + link + '\"><code>' +
      linkText + '</code></a>';
      par.editAsText().insertText(foundLinkText.getStartOffset(), builtLink);
      parText = par.getText();
    }
  }
}

function removeExtraLinesInListsAndDivs() {
  if (parText.indexOf('1.') === 0 || parText.indexOf('</div') === 0 ||
  parText.indexOf('<div') === 0 || parText.indexOf('<strong>') === 0 ||
    (parText.indexOf(' ') === 0 && parText.indexOf(' * ') === -1 && !isCode) ||
    parText.indexOf('```') > -1 || parText.indexOf('<img') === 0 ||
    parText.indexOf('<table>') === 0 || parText.indexOf('* ') === 0 ||
      parText.indexOf('    * ') === 0 || parText.indexOf('**') === 0) {

    lastPar = par.getPreviousSibling();

    while (lastPar.getText() == '' || lastPar.getText() == ' ' ||
    lastPar.getText() == '    ' || lastPar.getText() == '        ') {
      lastPar.removeFromParent();
      lastPar = par.getPreviousSibling();
    }

    //Fixes extra paragraphs that start bold
    if (parText.indexOf('** ') === 0) {
      par.insertText(0, ' ');
    }

    //Fixes numbered list after bulleted list
    if (parText.indexOf('1. ') === 0 &&
    lastPar.getText().indexOf('* ') === 0) {
      par.insertText(0, '\n');
      par.insertText(0, '\n');
    }

    //Fixes numbering for notes inside lists
    if ((parText.indexOf('1.') === 0 &&
    lastPar.getText().indexOf('</div>') > -1) ||
      (parText.indexOf('```') === 0 &&
      lastPar.getText().indexOf('</div>') > -1)) {
      par.insertText(0, '\n');
    }

    //Fixes extra paragraphs in lists
    if (parText.indexOf(' ') === 0 &&
    parText.indexOf(' * ') === -1 && !isCode && !isTable) {
      par.insertText(0, '\n');
    }
    nextPar = lastPar = null;
  }
}

function fixAnchors() {
  var foundName = par.findText('name');
  var foundConclusion = par.findText('conclusion');
  if (foundName != null) {
    foundName.getElement().editAsText().replaceText('name', 'id');
  }
  if (foundConclusion != null) {
    foundConclusion.getElement().editAsText()
    .replaceText('conclusion', 'summary');
  }
}

function removeTags() {
  var tagStart;
  var tagEnd;
  if (parText.indexOf('[LINK') > -1) {
    foundTag = par.findText('\\[LINK.*\\]');

    // All regexes in brackets need to replace .* with 0 or more expression

    tagStart = foundTag.getStartOffset();
    tagEnd = foundTag.getEndOffsetInclusive();
    foundTag.getElement().editAsText().deleteText(tagStart, tagEnd);
  }
  if (parText.indexOf('[ROOTVIEW') > -1) {
    foundTag = par.findText('\\[ROOTVIEW.*\\]');
    tagStart = foundTag.getStartOffset();
    tagEnd = foundTag.getEndOffsetInclusive();
    foundTag.getElement().editAsText().deleteText(tagStart, tagEnd);
  }
  if (parText.indexOf('[PROJECT_TEMPLATE') > -1) {
    foundTag = par.findText('\\[PROJECT_TEMPLATE.*\\]');
    tagStart = foundTag.getStartOffset();
    tagEnd = foundTag.getEndOffsetInclusive();
    foundTag.getElement().editAsText().deleteText(tagStart, tagEnd);
  }
  if (parText.indexOf('[COMPONENT_TEMPLATE') > -1) {
    foundTag = par.findText('\\[COMPONENT_TEMPLATE.*\\]');
    tagStart = foundTag.getStartOffset();
    tagEnd = foundTag.getEndOffsetInclusive();
    foundTag.getElement().editAsText().deleteText(tagStart, tagEnd);
  }
  if (parText.indexOf('[EMULATOR') > -1) {
    foundTag = par.findText('\\[EMULATOR.*\\]');
    tagStart = foundTag.getStartOffset();
    tagEnd = foundTag.getEndOffsetInclusive();
    foundTag.getElement().editAsText().deleteText(tagStart, tagEnd);
  }
  if (parText.indexOf('[UPCOMING_CHANGE') > -1) {
    foundTag = par.findText('\\[UPCOMING_CHANGE.*\\]');
    tagStart = foundTag.getStartOffset();
    tagEnd = foundTag.getEndOffsetInclusive();
    foundTag.getElement().editAsText().deleteText(tagStart, tagEnd);
  }
  if (parText.indexOf('[APP LINK') > -1) {
    foundTag = par.findText('\\[APP LINK.*\\]');
    tagStart = foundTag.getStartOffset();
    tagEnd = foundTag.getEndOffsetInclusive();
    foundTag.getElement().editAsText().deleteText(tagStart, tagEnd);
  }

  if (parText.indexOf('REVIEWERS:') > -1 ||
  parText.indexOf('[IMAGEINFO]') > -1 ||
  parText.indexOf('[CODE_HIGHLIGHT') > -1 ||
  parText.indexOf('[NESTED CODE') > -1) {
    rangeBuilder.addElement(par);
  }

  for (i = 0; i < rangeBuilder.getRangeElements().length; i++) {
    rangeBuilder.getRangeElements()[i].getElement().removeFromParent();
  }
}

function fixIndentedCode() {
  if (parText.indexOf('[NESTED CODE') === 0) {
    var numIndent = parText.charAt(par.findText('\"')
    .getEndOffsetInclusive() + 1);
    var nextPar = par.getNextSibling();
    var isCode = false;
    while (!isCode) {
      if (nextPar.getText().indexOf('```') === 0 && !isCode) {
        isCode = true;
      }
      nextPar = nextPar.getNextSibling();
    }

    while (isCode) {
      for (var i = 0; i < numIndent * 4; i++) {
        nextPar.editAsText().insertText(0, ' ');
      }

      nextPar = nextPar.getNextSibling();

      if (nextPar.getText().indexOf('```') === 0) {
        isCode = false;
      }
    }
  }
}

function highlightCode() {
  if (parText.indexOf('[CODE_HIGHLIGHT') > -1) {
    if (parText.indexOf('JAVA') > -1) {
      codeType = 'java';
    } else if (parText.indexOf('XML') > -1) {
      codeType = 'xml';
    }
  }
  if (parText.indexOf('```') === 0 && codeType != null) {
    par.appendText(codeType);
    par.editAsText().insertText(0, '\n');
    codeType = null;
  }
}

function addImageBorderAndCenter() {
  if (parText.indexOf('<img') > -1 && parText.indexOf('ic_') == -1) {
    par.editAsText().insertText(parText.indexOf('<img') + 5,
    'class=\"center\" ');
  } else if (parText.indexOf('<img') > -1 && parText.indexOf('ic_') > -1) {
    par.editAsText().insertText(parText.indexOf('<img') + 5, 'class=\"ic\" ');
  }

  if (parText.indexOf('pv_') > -1) {
    par.editAsText().insertText(parText.indexOf('<img') + 18, ' pv');
  } else if (parText.indexOf('ph_') > -1) {
    par.editAsText().insertText(parText.indexOf('<img') + 18, ' ph');
  } else if (parText.indexOf('tv_') > -1) {
    par.editAsText().insertText(parText.indexOf('<img') + 18, ' tv');
  } else if (parText.indexOf('th_') > -1) {
    par.editAsText().insertText(parText.indexOf('<img') + 18, ' th');
  }

  if (parText.indexOf('pv_') > -1 || parText.indexOf('ph_') > -1 ||
  parText.indexOf('th_') > -1 || parText.indexOf('tv_') > -1) {
    par.editAsText().insertText(parText.indexOf('<img') + 5,
    'style=\"border:1px solid black\" ');
  }
}

function formatCodeAndTables() {
  var lastParText;
  var nextParText;
  if (lastPar != null) {
    lastParText = lastPar.asParagraph().getText();
  } else {
    lastParText = 'START';
  }
  var nextPar = par.getNextSibling();
  if (nextPar != null) {
    var nextParText = nextPar.asParagraph().getText();
  }

  if (parText.indexOf('<table>') === 0 || parText.indexOf('```') > -1 ||
  parText.indexOf('<div class=\"note\">') === 0 ||
  parText.indexOf('</div>') === 0 || parText.indexOf('<img') === 0 ||
  parText.indexOf('* ') === 0 || parText.indexOf('    * ') === 0 ||
    parText.indexOf('1.') === 0 || parText.indexOf('    1.') === 0) {
    while (lastParText == '' || lastParText == ' ' || lastParText == null) {
      lastPar = lastPar.getPreviousSibling();
      lastPar.getNextSibling().removeFromParent();
      if (lastPar != null) {
        lastParText = lastPar.asParagraph().getText();
      }
    }
    if (parText.indexOf('```xml') > -1) {
      par.clear();
      par.appendText('```xml');
    }

    if (parText.indexOf('```java') > -1) {
      par.clear();
      par.appendText('```java');
    }
  }

  if (parText.indexOf('```') === 0 || parText.indexOf('</table>') === 0 ||
  parText.indexOf('<img') === 0 || parText.indexOf('<div class') === 0 ||
  parText.indexOf('</div>') === 0) {
    while ((nextParText == '' || nextParText == ' ' ||
    nextParText == null) && !nextPar.isAtDocumentEnd()) {
      nextPar = nextPar.getNextSibling();
      nextPar.getPreviousSibling().removeFromParent();
      if (nextPar != null) {
        nextParText = nextPar.asParagraph().getText();
      }
    }
  }

  if (parText.indexOf('**') === 0) {
    par.editAsText().insertText(0, '\n');
  }
}

function logAll(lastParText, parText, nextParText) {
  Logger.log('lastParText: ' + lastParText + ' parText: ' +
  parText + ' nextParText: ' + nextParText)
}
