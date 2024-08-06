function onOpen() {
  Logger.log('start')
  let ui = SlidesApp.getUi();
  ui.createMenu('Translate')
    .addItem('Spanish', 'toSpanish')
    .addItem('Arabic', 'toArabic')
    .addToUi();
}

function toSpanish(){
  translate('es')
}

function toArabic(){
  translate('ar')
}



function translate(desLanguage) {
  const slides = SlidesApp.getActivePresentation().getSlides()
  const texts = []
  
  slides.forEach(function(slide) {
        let newRanges = []
    const pageElements = slide.getPageElements()
    pageElements.forEach(function(ele) {
      if (ele.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const tRange = ele.asShape().getText()

        // gathering text run info
        const tRuns = tRange.getRuns()
        tRuns.forEach(function(run){
          if (run.asString().trim().length > 0) {
    
          newRanges.push({info:run, text:run.asString().trim(), size:run.getTextStyle().getFontSize(), listStyle: run.getListStyle().isInList()})
      }})

        
      }
      }
    )
  newRanges.forEach(function(data){
    if (data.text !== ''){
    const translatedText = LanguageApp.translate(data.text.trim(), 'en', desLanguage)
    slide.replaceAllText(data.text.trim(), translatedText)
  }
  })
  Logger.log(newRanges)

  })
}



