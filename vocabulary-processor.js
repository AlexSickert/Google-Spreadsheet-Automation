    
// ================================================================
// to Dos: 
// get array of all ignored words or better hash of hash (by language) 
// convert all words to lowercase

// ================================================================
// make menue
// ================================================================

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    ui.createMenu('VOCABULARY PROCESSOR')
            .addItem('FRENCH', 'runF')
            .addItem('SPANISH', 'runS')
            .addItem('RUSSIAN', 'runR')
            .addItem('ITALIAN', 'runI')
            .addItem('PORTUGUESE', 'runP')
            .addToUi();
}

function runF() {
    runByLanguage("French");
}

function runS() {
    runByLanguage("Spanish");
}

function runI() {
    runByLanguage("Italian");
}

function runP() {
    runByLanguage("Portuguese");
}

function runR() {
    runByLanguage("Russian");
}

// ================================================================
//
// ================================================================

var languages = ["French", "Italian", "Russian", "Spanish", "Portuguese"];

var numProcessRows = 100;
var ignoreColumns = {};
ignoreColumns["Russian"] = 1;
ignoreColumns["French"] = 2;
ignoreColumns["Italian"] = 4;
ignoreColumns["Spanish"] = 5;
ignoreColumns["Portuguese"] = 3;

var db = {}; // will be 3-dimensional

for (i = 0; i < languages.length; i++) {
    db[languages[i]] = {};
    db[languages[i]]["ignore"] = {};
    db[languages[i]]["words"] = {};
    db[languages[i]]["new"] = {};
}

var translateCodes = {};
translateCodes["Russian"] = "ru";
translateCodes["French"] = "fr";
translateCodes["Italian"] = "it";
translateCodes["Spanish"] = "es";
translateCodes["Portuguese"] = "pt";

var newWordColumns = {};
newWordColumns["Russian"] = 1;
newWordColumns["French"] = 3;
newWordColumns["Italian"] = 7;
newWordColumns["Spanish"] = 9;
newWordColumns["Portuguese"] = 5;

var ignoredCharacter = {};

// ================================================================
//
// ================================================================

function exportIgnoredChars() {
   // ignoredCharacter
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Unicode");
  
  var startRow = 2;
  
  for (var key in ignoredCharacter){
    
    sheet.getRange(startRow, 1).setValue(key);
    sheet.getRange(startRow, 2).setValue(ignoredCharacter[key]);
    startRow += 1;
    
  }
  
  
}

// ================================================================
//
// ================================================================

function runByLanguage(l) {

    // clean up manual action and add stuff to ignore list
    cleanUpOne(l);
    // append stuff to ignore table
    writeIgnorToTable(ignoreColumns[l], db[l]["ignore"]);
  
    // add the new word to the voc list 
    addWordsToLearn(l);

    // clean up other stuff
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(l);
    sheet.getRange('A2:D200').clearContent();

    // make ignore list 
    makeIgnoreList(l);

    //now we load new stuff
    //var sheet = SpreadsheetApp.getActiveSheet();
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Links");

    // row, colum
    Logger.log(sheet.getRange(1, 2).getValue());


    var currRow;
    var stop;

    stop = false;
    currRow = 2;
    var tst;

    while (stop == false) {

        tst = sheet.getRange(currRow, 1).getValue();
      
        if (tst === null){
            stop = true;
            break;
        }

        if (tst.length < 1) {
            stop = true;
            break;
        }

        // only process right language
        if (tst === l) {
            sheet.getRange(currRow, 4).setValue("loading...");
            processOneLink(sheet.getRange(currRow, 1).getValue(), sheet.getRange(currRow, 2).getValue())
            sheet.getRange(currRow, 4).setValue("done");
        }
        currRow += 1;
    }

    // export sorted hash to table
    Logger.log(" exportHash(l)");
    exportHash(l);

    // translate the new vocs
    Logger.log(" translateOneLanguage(l);");
    translateOneLanguage(l);

    exportIgnoredChars();

    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
            .alert('Done !');

}


// ================================================================
//
// ================================================================
function addWordsToLearn(language) {

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LEARN");
    stop = false;
    var currRow = 2;
    var word;
    var translation;

    //get the existing words
    currRow = 3;

    while (stop === false) {
        word = sheet.getRange(currRow, newWordColumns[language]).getValue();

        if (word.length < 1) {
            stop = true;
            break;
        }

        translation = sheet.getRange(currRow, newWordColumns[language] + 1).getValue();
        db[language]["new"][word] = translation;
        currRow += 1;
    }

    // add new words
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(language);
    stop = false;
    currRow = 2;

    while (stop === false) {

        tst = sheet.getRange(currRow, 3).getValue();
        word = sheet.getRange(currRow, 1).getValue();
        translation = sheet.getRange(currRow, 2).getValue();

        if (currRow > numProcessRows + 2) {
            stop = true;
            break;
        }

        // add only if not already there
        if (tst.trim() === "a") {
            Logger.log("adding new: " + word);
            if (db[language]["new"][word] === undefined) {
                db[language]["new"][word] = translation;
            }
        }
        currRow += 1;
    }


    // add the word to the table
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LEARN");
    currRow = 3;
    for(var newWord in db[language]["new"]){
        sheet.getRange(currRow, newWordColumns[language]).setValue(newWord);
        sheet.getRange(currRow, newWordColumns[language] + 1).setValue(db[language]["new"][newWord]);
        currRow += 1;
    }


}


// ================================================================
//
// ================================================================
function cleanUpOne(language) {

    // load exisitng ignore list
    // then add new
    // write back to the table

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IGNORE");

    var stop;
    stop = false;
    var currRow = 2;
    var tst;
    var col;

    col = ignoreColumns[language];


    // load the exisitng ignore list
    Logger.log("cleanUpOne(language)  load the exisitng ignore list");
    while (stop == false) {

        tst = sheet.getRange(currRow, col).getValue();
        Logger.log("ignore from existing: " + tst);

        if (tst.length < 1) {
            stop = true;
            break;
        }

        db[language]["ignore"][tst] = true;

        currRow += 1;
    }


    // now we add the new words
    Logger.log("cleanUpOne(language) now we add the new words");
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(language);
    stop = false;
    var currRow = 2;
    var word;
    while (stop == false) {

        tst = sheet.getRange(currRow, 3).getValue();
        word = sheet.getRange(currRow, 1).getValue();

        if (currRow > numProcessRows + 2) {
            stop = true;
            break;
        }

        if (tst.trim() == "" || tst.trim() == "i" || tst.trim() == undefined) {
            Logger.log("ignore new: " + word);
            db[language]["ignore"][word] = true;
        }
        currRow += 1;
    }
  
    // we also need to exclude the words that are already on the new voc list
    Logger.log("cleanUpOne(language) we also need to exclude the words that are already on the new voc list");
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LEARN");
    currRow = 3;
    stop = false;
  
    while (stop === false) {
        word = sheet.getRange(currRow, newWordColumns[language]).getValue();

        if (word.length < 1) {
            stop = true;
            break;
        }
        db[language]["ignore"][word] = true;
        currRow += 1;
    }  
}
// ================================================================
//
// ================================================================

function writeIgnorToTable(column, hash) {

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IGNORE");

    var row = 2;

    for (var i in hash) {
        sheet.getRange(row, column).setValue(i);
        row += 1;
    }
    return;
}

// ================================================================
// make the ignore lists
// ================================================================
function makeIgnoreList(l) {

    Logger.log("ignore: ");

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IGNORE");

    var y = 2;
    var continueFlag;
    continueFlag = true;
    var oneWord;
    var x = ignoreColumns[l];

    while (continueFlag) {
        oneWord = sheet.getRange(y, x).getValue();
        if (oneWord == null){
          continueFlag = false;
        }else{
            Logger.log("makeIgnoreList(l) ignore: " + oneWord);
            db[l]["ignore"][oneWord] = true;
            if (oneWord.length < 1) {
                continueFlag = false;
            }
      }        
       
      y += 1;
    }
  
  
    // add the words from the vocabulary list as well as we dont went to see them again 
  
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LEARN");
    var currRow = 3;
    var stop = false;
    var word;
  
    while (stop === false) {
        word = sheet.getRange(currRow, newWordColumns[l]).getValue();

        if (word.length < 1) {
            stop = true;
            break;
        }
        db[l]["ignore"][word] = true;
        currRow += 1;
    }    

    return;
}


// ================================================================
//
// ================================================================

function translateOneLanguage(lang) {
    //translate("Нравится", "Russian");
    Logger.log("translateOneLanguage " + lang);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(lang);
    var searchWord;
    var restult;

    for (var i = 2; i < numProcessRows + 2; i++) {

        searchWord = sheet.getRange(i, 1).getValue();
        Logger.log("translateOneLanguage " + searchWord);
        if (searchWord.length > 0) {
            restult = translate(searchWord, lang);
            sheet.getRange(i, 2).setValue(restult);
        }
    }
    ;
}

// ================================================================
//
// ================================================================

function translate(word, language) {

    var from;
    var url;

    try {       
        
        from = translateCodes[language];

        url = 'https://translate.googleapis.com/translate_a/single?client=gtx&sl=' + from + '&tl=de&dt=t&q=' + word.trim();

        var t = UrlFetchApp.fetch(url).getContentText();

        var e = eval(t);

        Logger.log(t);
        Logger.log(e[0][0][0]);
        //return HtmlService.createHtmlOutput(output[0]);
        return(e[0][0][0]);

    } catch (e) {
        return "";
    }
}



// ================================================================
//
// ================================================================

function processOneLink(language, link) {
    Logger.log(language);
    Logger.log(link);
    Logger.log("-----------------------");
    //var content = getContentStringCleanSimple(link);
    var content = getContentStringClean(link);
    //getContentStringClean
    content.replace(/[\n\r]/g, ' ');
    content.replace(".", ' ');
    addToHash(content, language);

}

// ================================================================
// getContentStringCleanSimple
// ================================================================
function getContentStringCleanSimple(link) {
    try {
        var html = UrlFetchApp.fetch(link).getContentText();
        var doc = XmlService.parse(html);
        var html = doc.getRootElement();
        //var menu = getElementsByClassName(html, 'vertical-navbox nowraplinks')[0];
        var menu = getElementsByTagName(html, 'body')[0];
        //var output = XmlService.getRawFormat().format(menu);
        return(menu.getValue());
    } catch (e) {
        Logger.log("error: " + e);
        return(" ");
    }
}
// ================================================================
// export one hash 
// ================================================================
function exportHash(l) {

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(l);
    var row;
    row = 2;
    var hash = db[l]["words"];
    var list = getSortedKeys(hash);

    var arrayLength = list.length;
    for (var i = arrayLength - 1; i >= 0; i--) {
        if (row < numProcessRows + 2) {
            sheet.getRange(row, 1).setValue(list[i]);
            sheet.getRange(row, 4).setValue(hash[list[i]]);
        }
        row += 1;
    }
}
// ================================================================
// sort keys of has by value and return array of keys
// ================================================================
function getSortedKeys(obj) {
    var keys = [];
    for (var key in obj)
        keys.push(key);
    return keys.sort(function (a, b) {
        return obj[a] - obj[b];
    });
}
// ================================================================
// add a word to hash
// ================================================================
function addToHash(content, language) {

    Logger.log("in addToHash - language is " + language);
  
    content = cleanUnicode(content);
  
    var res = content.split(" ");
    var arrayLength = res.length;
    var x;

    for (var i = 0; i < arrayLength; i++) {

        x = res[i].trim();
        var ignore = false;
        if (x.length > 0) {

            //here we need to extent. we need to check if it is in ignored list or weird content

            if (db[language]["ignore"][x] !== undefined) {
                ignore = true;
            }

            if (ignore === false) {
                if (db[language]["words"][x] === undefined) {
                    db[language]["words"][x] = 1;
                } else {
                    db[language]["words"][x] += 1;
                }
                ;
            }
            ;
        }
        ;
    }
    ;
}


// ================================================================
//
// ================================================================

function getContentStringClean(link) {

    var response = UrlFetchApp.fetch(link);
    var content = response.getContentText("UTF-8");
    content = content.replace(/\n/gi, " ");
    content = content.replace(/\r/gi, " ");
    content = content.replace(/\t/gi, " ");
    var pattern = /<body[^>]*>((.|[\n\r])*)<\/body>/im
    var array_matches = pattern.exec(content);
    var res2 = "";
    Logger.log(" getContentStringClean(link) array_matches");
  
    if(array_matches !== null){
    
      
    for (var i = 0; i < array_matches.length; i++) {
        res2 += (array_matches[i]);
    }

    // get rid of script and style blocks
    res2 = res2.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "");
    res2 = res2.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, "");

    //remove all tags
    res2 = res2.replace(/<.+?>/gi, " ");
    // remove clutter
    res2 = res2.replace(/,/gi, " ");
    //res2 = res2.replace(/;/gi, " ");
    res2 = res2.replace(/"/gi, " ");
    res2 = res2.replace(/»/gi, " ");
    res2 = res2.replace(/«/gi, " ");
    res2 = res2.replace(/\./gi, " ");
    res2 = res2.replace(/:/gi, " ");
    res2 = res2.replace(/\(/gi, " ");
    res2 = res2.replace(/\)/gi, " ");
    res2 = res2.replace(/\[/gi, " ");
    res2 = res2.replace(/\]/gi, " ");
    res2 = res2.replace(/"/gi, " ");
    res2 = res2.replace(/&rsquo;/gi, "'");      
    res2 = res2.replace(/&raquo;/g, " ");
    res2 = res2.replace(/&laquo;/g, " ");
    res2 = res2.replace(/&Agrave;/g, "À");
    res2 = res2.replace(/&Aacute;/g, "Á");
    res2 = res2.replace(/&Acirc;/g, "Â");
    res2 = res2.replace(/&Atilde;/g, "Ã");
    res2 = res2.replace(/&Auml;/g, "Ä");
    res2 = res2.replace(/&Aring;/g, "Å");
    res2 = res2.replace(/&AElig;/g, "Æ");
    res2 = res2.replace(/&Ccedil;/g, "Ç");
    res2 = res2.replace(/&Egrave;/g, "È");
    res2 = res2.replace(/&Eacute;/g, "É");
    res2 = res2.replace(/&Ecirc;/g, "Ê");
    res2 = res2.replace(/&Euml;/g, "Ë");
    res2 = res2.replace(/&Igrave;/g, "Ì");
    res2 = res2.replace(/&Iacute;/g, "Í");
    res2 = res2.replace(/&Icirc;/g, "Î");
    res2 = res2.replace(/&Iuml;/g, "Ï");
    res2 = res2.replace(/&ETH;/g, "Ð");
    res2 = res2.replace(/&Ntilde;/g, "Ñ");
    res2 = res2.replace(/&Ograve;/g, "Ò");
    res2 = res2.replace(/&Oacute;/g, "Ó");
    res2 = res2.replace(/&Ocirc;/g, "Ô");
    res2 = res2.replace(/&Otilde;/g, "Õ");
    res2 = res2.replace(/&Ouml;/g, "Ö");
    res2 = res2.replace(/&times;/g, "×");
    res2 = res2.replace(/&Oslash;/g, "Ø");
    res2 = res2.replace(/&Ugrave;/g, "Ù");
    res2 = res2.replace(/&Uacute;/g, "Ú");
    res2 = res2.replace(/&Ucirc;/g, "Û");
    res2 = res2.replace(/&Uuml;/g, "Ü");
    res2 = res2.replace(/&Yacute;/g, "Ý");
    res2 = res2.replace(/&THORN;/g, "Þ");
    res2 = res2.replace(/&szlig;/g, "ß");
    res2 = res2.replace(/&agrave;/g, "à");
    res2 = res2.replace(/&aacute;/g, "á");
    res2 = res2.replace(/&acirc;/g, "â");
    res2 = res2.replace(/&atilde;/g, "ã");
    res2 = res2.replace(/&auml;/g, "ä");
    res2 = res2.replace(/&aring;/g, "å");
    res2 = res2.replace(/&aelig;/g, "æ");
    res2 = res2.replace(/&ccedil;/g, "ç");
    res2 = res2.replace(/&egrave;/g, "è");
    res2 = res2.replace(/&eacute;/g, "é");
    res2 = res2.replace(/&ecirc;/g, "ê");
    res2 = res2.replace(/&euml;/g, "ë");
    res2 = res2.replace(/&igrave;/g, "ì");
    res2 = res2.replace(/&iacute;/g, "í");
    res2 = res2.replace(/&icirc;/g, "î");
    res2 = res2.replace(/&iuml;/g, "ï");
    res2 = res2.replace(/&eth;/g, "ð");
    res2 = res2.replace(/&ntilde;/g, "ñ");
    res2 = res2.replace(/&ograve;/g, "ò");
    res2 = res2.replace(/&oacute;/g, "ó");
    res2 = res2.replace(/&ocirc;/g, "ô");
    res2 = res2.replace(/&otilde;/g, "õ");
    res2 = res2.replace(/&ouml;/g, "ö");
    res2 = res2.replace(/&divide;/g, "÷");
    res2 = res2.replace(/&oslash;/g, "ø");
    res2 = res2.replace(/&ugrave;/g, "ù");
    res2 = res2.replace(/&uacute;/g, "ú");
    res2 = res2.replace(/&ucirc;/g, "û");
    res2 = res2.replace(/&uuml;/g, "ü");
    res2 = res2.replace(/&yacute;/g, "ý");
    res2 = res2.replace(/&thorn;/g, "þ");
    res2 = res2.replace(/&yuml;/g, "ÿ");  
      
    res2 = res2.replace(/&nbsp;/g, " ");  
      
 

    //Logger.log(res2.trim()); 

    //res2 = res2.replace(/<(?:[^>'"]*|(['"]).*?\1)*>/g, ' ');      
      
      
    }else{
      res2 = "";
    }
  
    
    return res2;

}

// ================================================================
//
// ================================================================
function cleanUnicode(str) {

                var oneChar;
                var numOfchar;
                var tst;
                var ret = "";

                for (var i = 0; i < str.length; i++) {
                    oneChar = str.substring(i, i + 1);
                    numOfchar = oneChar.charCodeAt(0);

                    tst = false;

                    if (numOfchar >= 65 && numOfchar <= 90) {
                        tst = true;
                    }
                
                    if (numOfchar >= 97 && numOfchar <= 122) {
                        tst = true;
                    }
                  
                    if (numOfchar >= 1072 && numOfchar <= 1103) {
                        tst = true;
                    }
                 
                    if (numOfchar >= 1040 && numOfchar <= 1071) {
                        tst = true;
                    }
                    
                    if (numOfchar >= 192 && numOfchar <= 255) {
                        tst = true;
                    }
              
                    if (numOfchar === 246) {
                        tst = true;
                    }
                   
                    if (numOfchar === 228) {
                        tst = true;
                    }
                 
                    if (numOfchar === 252) {
                        tst = true;
                    }
                 
                    if (numOfchar === 223) {
                        tst = true;
                    }
                   
                    if (numOfchar === 214) {
                        tst = true;
                    }
                  
                    if (numOfchar === 196) {
                        tst = true;
                    }
                 
                    if (numOfchar === 220) {
                        tst = true;
                    }  
                  
                    if (numOfchar === 38) {
                        tst = true;
                    }
                  
                    if (numOfchar === 59) {tst = true;}
                    if (numOfchar === 8211) {tst = true;}
                    if (numOfchar === 45) {tst = true;}
                    if (numOfchar === 8217) {tst = true;}
                    if (numOfchar === 1105) {tst = true;}
                    if (numOfchar === 1110) {tst = true;}
                    if (numOfchar === 1111) {tst = true;}
                    if (numOfchar === 96) {tst = true;}
                    if (numOfchar === 8722) {tst = true;}
                  
                    if (tst) {
                        ret += oneChar;
                    } else {
                        ignoredCharacter[oneChar] = oneChar + " = " + numOfchar
                        ret += " ";
                    }
                }


                return ret;

            };





