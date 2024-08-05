function onOpen() {
  Logger.log('start')
  let ui = SlidesApp.getUi();
  ui.createMenu('Translate')
    .addItem('Grab Text', 'getAllTextFromSlides')
    .addItem('Translate', 'translateHandler')
    .addItem('view ui', 'logSlideDetails')
    .addItem('foo', '_translate')

    .addToUi();
}

// input is pageElement
// https://developers.google.com/apps-script/reference/slides/page-element
class textBoxContents {
  constructor(pageElementContent, pageId) {
    const shape = pageElementContent.asShape()
    const tRange = shape.getText()

    this.pageId = pageId
    this.id = pageElementContent.getObjectId()
    this.content = tRange.asString()
    this.fontSize = tRange.getTextStyle().getFontSize()
  }

  info() {
    Logger.log('Item start')
    Logger.log('slide id: ' + this.pageId)
    Logger.log('id: ' + this.id)
    Logger.log('content: ' + this.content)
    Logger.log('fontSize: ' + this.fontSize)
  }
}

// known limitation: different sizes in same textRange
function extractContents() {
  const slides = SlidesApp.getActivePresentation().getSlides()
  const contents = []
  slides.forEach(function(slide) {
    const slideId = slide.getObjectId()
    const pageElements = slide.getPageElements()
    pageElements.forEach(function(ele) {
      if (ele.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        contents.push(new textBoxContents(ele, slideId))
      }
    })

  })

return contents
}

function translateContents(arr, startLang, endLang){
  const translatedArr = []
  arr.forEach(function(ele){
    const spanishContents = LanguageApp.translate(ele.content, startLang, endLang)
    ele.content = spanishContents
    translatedArr.push(ele)
  })
  return translatedArr
}

function replaceContent(contents){
  const slides = SlidesApp.getActivePresentation()
  contents.forEach(function(ele){
    const oldVal = slides.getSlideById(ele.pageId).getPageElementById(ele.id).asShape().getText()
    Logger.log(oldVal)
    oldVal.setText(ele.content)
  })

}


function _translate() {
  Logger.log('translate called')
  contents = extractContents()
  Logger.log('translate start')
  const updatedContents = translateContents(contents, 'en', 'es')
  Logger.log('translate end')
  replaceContent(contents, updatedContents)
  Logger.log('replace end')
}

function _getAllTextFromSlides() {
  let presentation = SlidesApp.getActivePresentation();
  let slides = presentation.getSlides();
  let allTextArray = [];

  // Iterate through each slide
  slides.forEach(function(slide) {
    let pageElements = slide.getPageElements();

    // Iterate through each element on the slide
    pageElements.forEach(function(element) {
      try {
        if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          let shape = element.asShape();
          let textRange = shape.getText();
          if (textRange) {
            let shapeText = textRange.asString();
            if (shapeText) {
              allTextArray.push(shapeText);
            }
          }
        }
      } catch (e) {
        // If there's an error processing the element, log it
        Logger.log('Error processing element: ' + e.message);
      }
    });
  });

  // Log the results to the Apps Script log
  Logger.log('Text collected from slides: ' + JSON.stringify(allTextArray));

  // Optionally, you can return the array if needed
  return allTextArray;
}


function translateHandler() {
  let count = 1
  // const textArray = getAllTextFromSlides()

  const slides = SlidesApp.getActivePresentation().getSlides()
  /*
  slides.forEach(function(slide){
    asdf
*/
  slides.forEach(function(slide) {
    const pageElements = slide.getPageElements()
    Logger.log('page element start')
    pageElements.forEach(function(ele) {
      Logger.log(ele)
    })
    Logger.log('page element end')
    pageElements.forEach(function(element) {
      try {
        if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          let shape = element.asShape();
          let textRange = shape.getText();


          // Check if the shape has text
          if (textRange) {
            let currentText = textRange.asString();
            // Append the letter "A" to the current text
            textRange.setText(count + currentText);
            // const textStyle = textRange.getTextStyle();
            // textRange.getTextStyle().setBold(textStyle.isBold());
            // textRange.getTextStyle().setItalic(textStyle.isItalic());
            // textRange.getTextStyle().setUnderline(textStyle.isUnderline());
            // textRange.getTextStyle().setFontSize(3);

            count++
          }
        }
      } catch (e) {
        // Log any errors that occur
        Logger.log('Error processing element: ' + e.message);
      }
    });
  });
}



function logSlideDetails() {
  let presentation = SlidesApp.getActivePresentation();
  let slides = presentation.getSlides();

  slides.forEach(function(slide, index) {
    try {
      Logger.log('Slide ' + (index + 1));
      let pageElements = slide.getPageElements();
      pageElements.forEach(function(element) {
        try {
          if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
            let shape = element.asShape();
            Logger.log('Shape Text: ' + shape.getText().asString());
          }
        } catch (e) {
          Logger.log('Error processing element: ' + e.message);
        }
      });
    } catch (e) {
      Logger.log('error: ' + e)
    }
  });
}


function getAllTextFromSlides() {
  let presentation = SlidesApp.getActivePresentation();
  let slides = presentation.getSlides();
  let allTextArray = [];

  // Iterate through each slide
  slides.forEach(function(slide) {
    let pageElements = slide.getPageElements();

    // Iterate through each element on the slide
    pageElements.forEach(function(element) {
      try {
        if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          let shape = element.asShape();
          let textRange = shape.getText();
          if (textRange) {
            let shapeText = textRange.asString();
            if (shapeText) {
              allTextArray.push(shapeText);
            }
          }
        }
      } catch (e) {
        // If there's an error processing the element, log it
        Logger.log('Error processing element: ' + e.message);
      }
    });
  });

  // Log the results to the Apps Script log
  Logger.log('Text collected from slides: ' + JSON.stringify(allTextArray));

  // Optionally, you can return the array if needed
  return allTextArray;
}
