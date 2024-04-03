/**
 * The Script that generates test variants from the list of test questions stored in the Google Document.
 *
 * Extending Google Docs developer guide:
 *     https://developers.google.com/apps-script/guides/docs
 *
 * Document service reference documentation:
 *     https://developers.google.com/apps-script/reference/document/
 */
function onOpen() {
  DocumentApp.getUi()
      .createMenu('Колоквіум')
      .addItem('Генератор...', 'generateDialog')
      .addItem('Статистика...', 'topicsDialog')      
      .addToUi();
}

function generateDialog() {
  var html = HtmlService
    .createTemplateFromFile('Form')
    .evaluate()

  DocumentApp.getUi()
      .showModalDialog(html, 'Генерація колоквіуму')
}

function topicsDialog() {
  var html = HtmlService
    .createTemplateFromFile('Topics')
    .evaluate()

  DocumentApp.getUi()
      .showModalDialog(html, 'Звіт по темам')
}

function generateDocument(documentName, variantsCount, questionsPerVariantCount, pagesPerTestCount, stats) {  
  var docName = documentName ?? 'col-test'
  var varsCount = variantsCount ?? 50
  var questionsCount = questionsPerVariantCount ?? 15
  var pagesPerTest = pagesPerTestCount ?? 4

  var document = DocumentApp.create(docName)
  
  try {
    generate(document, varsCount, questionsCount, pagesPerTest, stats)
    document.saveAndClose()
    
    var generatedId = document.getId()
    var generatedFile = DriveApp.getFileById(generatedId)
    var directParents = DriveApp.getFileById(DocumentApp.getActiveDocument().getId()).getParents()
    
    while( directParents.hasNext()) {
      generatedFile.moveTo(directParents.next())
    }

    var recipient = Session.getActiveUser().getEmail()
    var subject = 'Колоквіум ' +  docName + ' готовий'
    var body = "Лінк на тести " + documentName + "\n" + generatedFile.getUrl() + "\n"  
    MailApp.sendEmail(recipient, subject, body)
    
    return generatedFile.getUrl();
  } catch(e) {
    Logger.log("Exception occured ", e)
    DocumentApp.getUi().alert("Помилка при генерації документу: " + e)
    DriveApp.getFileById(document.getId()).setTrashed(true)
  }
}

// ------------------------------------- Document generation secion -------------------------------------
function testParse() {
  Logger.log("Document: " + generateDocument('delete-me', 3, 5, 1))
}


function Question(id, text) {
	var self = this;
	self.id = id;
	self.text = text;
	self.options = [];
	self.usedTimes = 0;
	self.incompatible = [];
	self.tags = [];
  self.topics = [];
  self.correctOptions = [1] // reserved
	self.text = replaceFunc(text, self).trim();
  self.index = 0;
    
  self.addOption = function(text) {
    self.options.push(replaceFunc(text, self));
	}
    
	self.printAndGetAnswer = function(no, body) {
      var text = self.text;      
      var paragraph;
      
      if(text.indexOf("```") == -1) {
        paragraph = body.appendParagraph(no + '. ' + self.text).setBold(true).setFontSize(12).setSpacingBefore(12);
      } else {
        // We have code formatted as ```code```
         paragraph = body.appendParagraph(no + '. ');
         paragraph.setBold(true).setFontSize(12).setSpacingBefore(12);
         var pos = 0;

         while (pos < text.length) {
            var start = text.indexOf("```", pos);
            
            if (start < 0) {
              paragraph.appendText(text.substring(pos)).setBold(true).setFontSize(12);
              pos = text.length;
            } else {
              var startText = text.substring(pos, start)

              if(startText) {
                paragraph.appendText(startText).setBold(true).setFontSize(12);
              }
                                    
              var end = text.indexOf("```", start + 3) // 3 is len of ```
              var endText = text.substring(start + 3, end)

              if (endText) {
                paragraph.appendText(endText).setBold(false).setFontSize(11).setFontFamily('Consolas');
              } else {
                pos = text.length;
              }

              pos = end + 3;
            }
         }
      }
      
      // Retrieve the paragraph's attributes.
      //var atts = paragraph.getAttributes();
      
      // Log the paragraph attributes.
      //for (var att in atts) {
      //  Logger.log(att + ":" + atts[att]);
      //}
      
      return self.printOptions(body);		
	}		
    
    self.regenerateIndex = function() {      
      self.index = Math.ceil(Math.random()*100) + 35*self.usedTimes;
    }
    
    self.printOptions = function(body) {
      if(self.options.length == 0) 
        throw 'Питання не має жодного варіанту відповіді: ' + self.text;
      
      var opts = [...self.options]
      var multipleOpts = opts.length > 1
      var multipleCorrectOptions = self.correctOptions.length > 1
      var result = ['-']
      
      if (multipleOpts) {
        const correctOpts = self.correctOptions.map(i => opts[i-1]);
        // randomize
        opts.sort(function() { return Math.random() - 0.5; });

        // find correct option indexes
        result = correctOpts.map(co => opts.indexOf(co)+1)
      } 
      
      opts.forEach(function(o, idx) {
        self.printOpt(o.trim(), idx, body, multipleOpts, multipleCorrectOptions);
      });

      return result
    }
    
    self.printOpt = function(o, idx, body, multipleOpts, multipleCorrectOptions) {
      // 25EF - ◯, 2610 - ☐
      var checkboxSymb = multipleCorrectOptions ? String.fromCharCode(parseInt('2610', 16)) : String.fromCharCode(parseInt('25EF', 16))
      var text = multipleOpts ? ( checkboxSymb + ' ' + (idx + 1) + ') ' + o) : o;
      
      if(multipleOpts) {
        body.appendParagraph(text)
          .setBold(false).setFontSize(11);
      } else {
        body.appendParagraph(text).setBold(false).setFontSize(11);  
      }
    }
}

function parseQuestions() {
  var body = DocumentApp.getActiveDocument().getBody();
  var bodyText = body.getText().replace(/\n\s*\n/g, '\n');

  // cut all text to questions and answers
  var globalRegExp = /(^\d(.|\n[^(^\d)])*)/gmi;
  
  // match question or answer with digit and text
  var questionRegExp = /^([\d.]+)\s((.|\n)*)$/i;
  
  var questionTokens = bodyText.match(globalRegExp);

  if (!questionTokens || questionTokens.length == 0) {
    throw 'Не знайдено жодного запитання!';
  }
  
  // form questions list
  var questions = [];

  for (var i = 0; i < questionTokens.length; i++ ) {
    var qStr = questionTokens[i];
    try {
      var tokens = questionRegExp.exec(qStr);
      
      if(isQuestionIndex(tokens[1])) {       // question
        questions.push(new Question(parseInt(tokens[1]), tokens[2]));
      } else {      // variant	
        if(questions.length == 0) 
          throw 'Помилка(' + i + ") - не можна починати з варіантів: " + questionTokens[i];
        
        questions[questions.length - 1].addOption(tokens[2]);
      }
    } catch (e) {
      Logger.log("Error on parsing question: " + qStr + ", exception: " + e)
      throw 'Помилка при розборі питання ' + qStr
    }
  }
  
  questions.forEach(function(q) {
  	q.incompatible.forEach(function(v, idx) {
   	  questions[idx-1].incompatible[q.id] = v;
   	});
  });

  return questions;
}

function generate(document, variantsCount, questionsCount, pagesPerTest, stats) {
  var questions = parseQuestions();
  
  // Form document
  if(questionsCount > questions.length/2) 
  	throw 'Кількість запитань (' + questions.length + ') недостатня для формування тестів. Необхідно хоча б ' + (questionsCount*2) + ' запитань';
  
  var answers = [];
  var outBody = document.getBody();
  
  for(var varNo = 1; varNo <= variantsCount; varNo++) {
    answers.push(formVariant(varNo, outBody, mixQuestions(varNo, questions, questionsCount)));
  }
  
  // Report with answers
  
  outBody.appendParagraph("Відповіді:").setHeading(DocumentApp.ParagraphHeading.HEADING3);
  
  answers.forEach(function(variant, varIdx) {
    outBody.appendParagraph('Варіант: ' + (varIdx + 1) + ':').setBold(true);
    
    var answText = '';
    
    variant.forEach(function(a, idx) {
      answText += (idx+1)  + ' : ' + a.join(", ") + '\n';
    });
    
    outBody.appendParagraph(answText).setFontSize(10).setBold(false);
  });

  if (stats) {
    var statPar = outBody.appendParagraph("Статистика:");
    statPar.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    
    var statText = ''
    
    questions.forEach(function(e, idx) {
      statText += 'Запитання: ' + (idx + 1) + ' використано ' + e.usedTimes + ' раз\n';
    });
    
    outBody.appendParagraph(statText).setBold(false).setFontSize(10);
  }
}

function questionsStats() {
  try {
    var questions = parseQuestions()
  }  catch(e) {
    Logger.log("Exception occured ", e)
    DocumentApp.getUi().alert("Помилка при генерації документу: " + e)
    return {total: 0, topics: []}
  }

  var topics = {}

  questions.forEach(q => q.topics.forEach(topic => { 
    topics[topic] = (topics[topic] ?? 0) + 1
  }))

  var list = []

  for (const [topic, count] of Object.entries(topics)) {
    list.push({ topic, count })
  }

  list.sort((a,b) => b.count - a.count)

  return {
    total: questions.length,
    topics: list
  }
}

// variant contains dot
function isQuestionIndex(s) {
  return s.indexOf(".") == -1;
}

function replaceFunc(text, question) {
  return text.replace(/@(\w+)\((.*)\)/g, function (match, fname, argsStr) {
    try {
      Logger.log('Q' + question.id + ' : Replacing "' + match + ', fname=' + fname + ', args=' + argsStr);
      
      var args = argsStr.split(',');
      
      switch(fname) {
        case 'TextField': 
          var lines = args[0]? parseInt(args[0]) : 1;
          var lineTpl = '___________________________________________________________________________';
          var text = lineTpl;
          for(var i = 1; i < lines; i++) {
            text += '\n' + lineTpl;
          } 
          return text;
        case 'NotWith': 
          args.forEach(function(e){ question.incompatible[ parseInt(e.trim()) ] = 1; });
          return '';
        case 'Tag':
          args.forEach(function(e){ question.tags.push(parseInt(e.trim())); });
          return '';
        case 'MultipleOptions':
          question.correctOptions = args.map(o => parseInt(o.trim()))
          return ''
        case 'Topic':
          args.forEach(function(a) { 
            question.topics.push(a.trim()) 
          })
          return ''
        default:
          Logger.log('Q' + question.id + ' : Replacing "' + match + ' – Not matched a function')
          return match;
      }
    } catch (err) {
       Logger.log(err)
       return match
    }
  })
}

function containsIncompatible(arr, itm) {
	for(var i = 0, len=arr.length; i < len; i++) {
		if (arr[i].incompatible[itm.id] == 1) return true;
        if (hasCommonElements(arr[i].tags, itm.tags)) return true;
	}
	return false;
} 

function hasCommonElements(array1, array2) {       
    for (var i = 0, len=array1.length; i< len; i++) {
       if (array2.indexOf(array1[i]) != -1) return true;
    }
          
    return false;
}
           
function mixQuestions(varNo, questions, questionsCount) {
  	var clone = questions.slice(0);

    clone.forEach(function(el) { el.regenerateIndex(); });
  
	// mix questions simultaneously trying to put most used to the end.
	clone.sort(function(q1, q2) { return q1.index - q2.index; });

	var copy = [];
	
	clone.forEach(function(itm) {
		if(copy.length == questionsCount) return;
		
		if(!containsIncompatible(copy, itm)) {
			copy.push(itm);
			itm.usedTimes++;
		}
	});

	if(copy.length < questionsCount)
		throw "Кількість запитань менша за необхідну, можливо занадто багато несумісних питань";

	return copy;
}

// Builds a single variant section
function formVariant(number, body, questions) {
    var titlePar = body.appendParagraph('Варіант ' + number);
    titlePar.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    titlePar.setBold(true);
    titlePar.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    titlePar.setFontFamily('Arial');
  
    var namePar = body.appendParagraph('ПІБ ____________________________________________________   Група _______________');
    namePar.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    namePar.setBold(false);
    namePar.setFontSize(12);
  
	var answers = [];

	questions.forEach(function(e, idx) { 
		var rightAnswer = e.printAndGetAnswer(idx + 1, body);
		answers.push(rightAnswer);
	});    
  
  body.appendPageBreak();
  
	return answers;
}
