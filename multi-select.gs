function onOpen(e) {
    SpreadsheetApp.getUi()
    .createMenu('Scripts')
    .addItem('Multi-select for this cell...', 'showDialog')
    .addToUi();
  }
  function showDialog() {  
    var html = HtmlService.createHtmlOutputFromFile('dialog').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi()
    .showSidebar(html);
  }
  function getValidationData(){
    try {
      return SpreadsheetApp.getActiveRange().getDataValidation().getCriteriaValues()[0].getValues();
    } catch(e) {
      return null
    }
  }
  
  function setValues_(e, update) {
    var selectedValues = [];
    
    for (var i in e) {
      selectedValues.push(i);
    }
    var separator = ','
    var total = selectedValues.length
    if (total > 0) {
      var range = SpreadsheetApp.getActiveRange()
      var value = selectedValues.join(separator)
      if (update) {
        var values = range.getValues()
        // check every cell in range
        for (var row = 0; row < values.length; ++row) {
          for (var column = 0; column < values[row].length; ++column) {
            var currentValues = values[row][column].split(separator);//typeof values[row][column] === Array ? values[row][column].split(separator) : [values[row][column]+'']
            // find same values and remove them
            var newValues = []
            for (var j = 0; j < currentValues.length; ++j) {
              var uniqueValue = true
              for(var i = 0; i < total; ++i) {
                if (selectedValues[i] == currentValues[j]) {
                  uniqueValue = false
                  break
                }
              }
              
              if (uniqueValue && currentValues[j].trim() != '') {
                newValues.push(currentValues[j])
              }
            }
            
            if (newValues.length > 0) {
              range.getCell(row+1, column+1).setValue(newValues.join(separator)+separator+value)
            } else {
              range.getCell(row+1, column+1).setValue(value);
            }
          }
        }
      } else {
        range.setValue(value);
      }
    }
  }
  
  function updateCell(e) {
    return setValues_(e, true)
  }
  
  function fillCell(e) {
    setValues_(e)
  }
