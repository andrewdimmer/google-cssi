// Google CSSI Involvement Manager
// Created by Andrew Dimmer
// v0.1.2

var EMAIL_LINK = "https://docs.google.com/document/d/1vCxeFkInr7jjpeNTb-0WD9KTIeciA9mY8JJ0-erLj8I/edit?usp=sharing";         // Replace this link with your email
var FORM_PREFILL_LINK = "";         // Replace this link with the prefilled link for your form.

var SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Names and Info");
var STUDENTS = SHEET.getRange(2, 1, SHEET.getLastRow() - 1, 4).getValues();

// Test sending emails
function test() {
  var first = "YOUR_FIRST_NAME";          // Replace this with your first name
  var last = "YOUR_LAST_NAME";          // Replace this with your last name
  var email = "YOUR_EMAIL";          // Replace this with your email
  var studentData = [first, last, email, "Finished"];
  
  var email = parseDocToEmail(EMAIL_LINK);
  var link = getFormLink(studentData);
  email.htmlBody = email.htmlBody.replace("[FIRST_NAME_SHORTCODE]",studentData[0]);
  email.plainText = email.plainText.replace("[FIRST_NAME_SHORTCODE]",studentData[0]);
  email.htmlBody = email.htmlBody.replace("[LAST_NAME_SHORTCODE]",studentData[1]);
  email.plainText = email.plainText.replace("[LAST_NAME_SHORTCODE]",studentData[1]);
  email.htmlBody = email.htmlBody.replace("[LINK_SHORTCODE]",link.htmlBody);
  email.plainText = email.plainText.replace("[LINK_SHORTCODE]",link.plainText);
  GmailApp.sendEmail(studentData[2], email.subject, email.plainText, {"htmlBody": email.htmlBody});
}

// Sends emails
function sendEmails() {
  var emailTemplate = parseDocToEmail(EMAIL_LINK);
  for (var i = 0; i < STUDENTS.length; i++) {
    if (STUDENTS[i][2].length > 0) {
      var link = getFormLink(STUDENTS[i]);
      var currentEmailHTML = emailTemplate.htmlBody.replace("[FIRST_NAME_SHORTCODE]",STUDENTS[i][0]).replace("[LAST_NAME_SHORTCODE]",STUDENTS[i][1]).replace("[LINK_SHORTCODE]",link.htmlBody);
      var currentEmailText = emailTemplate.plainText.replace("[FIRST_NAME_SHORTCODE]",STUDENTS[i][0]).replace("[LAST_NAME_SHORTCODE]",STUDENTS[i][1]).replace("[LINK_SHORTCODE]",link.htmlBody);
      GmailApp.sendEmail(STUDENTS[i][2], emailTemplate.subject, currentEmailText, {"htmlBody": currentEmailHTML});
    }
  }
}

// Parses a Google Doc into a plain text and HTML email
function parseDocToEmail(url) {
  var lists = {
    "endStack": [],
    "ids": [],
    "number": []
  };
  // Note: if performance becomes an issue in the future, the DriveApp can get the date last modified and the parsed version can be stored in a sheet, but for the sake of limiting permissions that feature has not yet been implimented.
  var emailContent = {"subject": "", "plainText": "", "htmlBody": ""};
  var document = DocumentApp.openByUrl(url);
  var paragraphs = document.getBody().getParagraphs();
  var listItems = document.getBody().getListItems();
  var l = 0; currentLevel = 0;
  emailContent.subject = paragraphs[0].getText();
  var pastBold = null, pastItalic = null, pastLink = null, pastUnderline = null;
  for (var p = 1; p < paragraphs.length; p++) {
    var textString = paragraphs[p].getText();
    var textObject = paragraphs[p].editAsText();
    if (paragraphs[p].getType() == DocumentApp.ElementType.LIST_ITEM) {
      if (currentLevel < listItems[l].getNestingLevel()+1) {
        currentLevel = listItems[l].getNestingLevel()+1;

        if (lists.ids.indexOf(listItems[l].getListId()) < 0) {
          lists.ids.push(listItems[l].getListId());
          lists.number.push([1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1]);         // Note: only handles up to 20 nested layers, then throws an error. Make dynamic in future? Possibly get number function...
        }
        
        if (listItems[l].getGlyphType() == DocumentApp.GlyphType.BULLET) {
          emailContent.htmlBody += "<ul style='list-style-type:disc;'>";
          lists.endStack.push("</ul>");
        } else if (listItems[l].getGlyphType() == DocumentApp.GlyphType.HOLLOW_BULLET) {
          emailContent.htmlBody += "<ul style='list-style-type:circle;'>";
          lists.endStack.push("</ul>");
        } else if (listItems[l].getGlyphType() == DocumentApp.GlyphType.SQUARE_BULLET) {
          emailContent.htmlBody += "<ul style='list-style-type:square;'>";
          lists.endStack.push("</ul>");
        } else if (listItems[l].getGlyphType() == DocumentApp.GlyphType.NUMBER) {
          emailContent.htmlBody += "<ol type='1'>"; // "Add start='" + lists.number[lists.ids.indexOf(listItems[l].getListId())][currentLevel-1] + "'" in the future to restart a list... not a priority now!
          lists.endStack.push("</ol>");
        } else if (listItems[l].getGlyphType() == DocumentApp.GlyphType.LATIN_UPPER) {
          emailContent.htmlBody += "<ol type='A'>"; // "Add start='" + lists.number[lists.ids.indexOf(listItems[l].getListId())][currentLevel-1] + "'" in the future to restart a list... not a priority now!
          lists.endStack.push("</ol>");
        } else if (listItems[l].getGlyphType() == DocumentApp.GlyphType.LATIN_LOWER) {
          emailContent.htmlBody += "<ol type='a'>"; // "Add start='" + lists.number[lists.ids.indexOf(listItems[l].getListId())][currentLevel-1] + "'" in the future to restart a list... not a priority now!
          lists.endStack.push("</ol>");
        } else if (listItems[l].getGlyphType() == DocumentApp.GlyphType.ROMAN_UPPER) {
          emailContent.htmlBody += "<ol type='I'>"; // "Add start='" + lists.number[lists.ids.indexOf(listItems[l].getListId())][currentLevel-1] + "'" in the future to restart a list... not a priority now!
          lists.endStack.push("</ol>");
        } else if (listItems[l].getGlyphType() == DocumentApp.GlyphType.ROMAN_LOWER) {
          emailContent.htmlBody += "<ol type='i'>"; // "Add start='" + lists.number[lists.ids.indexOf(listItems[l].getListId())][currentLevel-1] + "'" in the future to restart a list... not a priority now!
          lists.endStack.push("</ol>");
        }
      }
      emailContent.htmlBody += "<li>";
      for (var numtabs = 0; numtabs < currentLevel; numtabs++) {
        emailContent.plainText += "\t"
      }
      if (listItems[l].getGlyphType() == DocumentApp.GlyphType.BULLET) {
        emailContent.plainText += "* ";
      } else if (listItems[l].getGlyphType() == DocumentApp.GlyphType.HOLLOW_BULLET) {
        emailContent.plainText += "+ ";
      } else if (listItems[l].getGlyphType() == DocumentApp.GlyphType.SQUARE_BULLET) {
        emailContent.plainText += "- ";
      } else if (listItems[l].getGlyphType() == DocumentApp.GlyphType.NUMBER) {
        emailContent.plainText += lists.number[lists.ids.indexOf(listItems[l].getListId())][currentLevel-1] + ". ";
      } else if (listItems[l].getGlyphType() == DocumentApp.GlyphType.LATIN_UPPER) {
        emailContent.plainText += plainTextListA(lists.number[lists.ids.indexOf(listItems[l].getListId())][currentLevel-1], false);
      } else if (listItems[l].getGlyphType() == DocumentApp.GlyphType.LATIN_LOWER) {
        emailContent.plainText += plainTextListA(lists.number[lists.ids.indexOf(listItems[l].getListId())][currentLevel-1], true);
      } else if (listItems[l].getGlyphType() == DocumentApp.GlyphType.ROMAN_UPPER) {
        emailContent.plainText += plainTextListI(lists.number[lists.ids.indexOf(listItems[l].getListId())][currentLevel-1], false);
      } else if (listItems[l].getGlyphType() == DocumentApp.GlyphType.ROMAN_LOWER) {
        emailContent.plainText += plainTextListI(lists.number[lists.ids.indexOf(listItems[l].getListId())][currentLevel-1], true);
      }
      lists.number[lists.ids.indexOf(listItems[l].getListId())][currentLevel-1]++;
    }
    for (var i = textString.indexOf("\n")+1; i < textString.length; i++) {
      if (pastBold == true && textObject.isBold(i) == null) {         // No Longer Bold
        pastBold = null;
        emailContent.htmlBody += "</b>";
      }
      if (pastItalic == true && textObject.isItalic(i) == null) {         // No Longer Italic
        pastItalic = null;
        emailContent.htmlBody += "</i>";
      }
      if (pastUnderline == true && textObject.isUnderline(i) == null) {         // No Longer Underlined
        pastUnderline = null;
        emailContent.htmlBody += "</u>";
      }
      if (pastLink != null && textObject.getLinkUrl(i) == null) {         // No Longer a Link
        emailContent.plainText += " (" + pastLink + ")";
        pastLink = null;
        emailContent.htmlBody += "</a>";
      }
      if (pastLink != null && textObject.getLinkUrl(i) != null && (pastLink.length != textObject.getLinkUrl(i).length || pastLink.indexOf(textObject.getLinkUrl(i)) != 0)) {         // Two different links
        emailContent.plainText += " (" + pastLink + ")";
        pastLink = textObject.getLinkUrl(i);
        emailContent.htmlBody += "</a><a href='" + pastLink + "' target='_blank'>";
      }
      if (pastLink == null && textObject.getLinkUrl(i) != null) {         // Becomes a Link
        pastLink = textObject.getLinkUrl(i);
        emailContent.htmlBody += "<a href='" + pastLink + "' target='_blank'>";
      }
      if (pastUnderline == null && textObject.isUnderline(i) == true) {         // Becomes Underlined
        pastUnderline = true;
        emailContent.htmlBody += "<u>";
      }
      if (pastItalic == null && textObject.isItalic(i) == true) {         // Becomes Italic
        pastItalic = true;
        emailContent.htmlBody += "<i>";
      }
      if (pastBold == null && textObject.isBold(i) == true) {         // Becomes Bold
        pastBold = true;
        emailContent.htmlBody += "<b>";
      }
      emailContent.htmlBody += textString.charAt(i);
      emailContent.plainText += textString.charAt(i);
    }
    if (paragraphs[p].getType() == DocumentApp.ElementType.LIST_ITEM) {
      emailContent.htmlBody += "</li>";
      l++;
      if (paragraphs[p].isAtDocumentEnd() || paragraphs[p+1].getType() != DocumentApp.ElementType.LIST_ITEM) {
        while (currentLevel != 0) {
          emailContent.htmlBody += lists.endStack.pop();
          lists.number[lists.ids.indexOf(listItems[l-1].getListId())][currentLevel-1] = 1;
          currentLevel--;
        }
      } else {
        if ((listItems[l].getNestingLevel() + 1) < currentLevel) {
          while (currentLevel != (listItems[l].getNestingLevel()+1)) {
            emailContent.htmlBody += lists.endStack.pop();
            lists.number[lists.ids.indexOf(listItems[l-1].getListId())][currentLevel-1] = 1;
            currentLevel--;
          }
        }
        if (listItems[l].getListId().indexOf(listItems[l-1].getListId()) < 0 || listItems[l-1].getListId().indexOf(listItems[l].getListId()) < 0) {
          emailContent.htmlBody += lists.endStack.pop();
          currentLevel--;
        }
      }
    } else {
      emailContent.htmlBody += "<br />";
    }
    emailContent.plainText += "\n";
  }
  // Logger.log(emailContent);          // Debug
  return emailContent
}

// Helper function to parse ordered lists using letters
function plainTextListA(number, lowerCase) {
  var letters = "ZABCDEFGHIJKLMNOPQRSTUVWXY";
  var returnString = "";
  for (var i = 0; i <= Math.floor((number-1)/26); i++) {
    returnString += letters.charAt(number%26);
  }
  returnString += ". ";
  if (lowerCase) {
    return returnString.toLowerCase();
  }
  return returnString;
}

// Helper function to parse ordered lists using roman numerials
function plainTextListI(number, lowerCase) {
  if (number < 1001 && number > 0) {
    var hundreds = ["","C","CC","CCC","CD","D","DC","DCC","DCCC","CM","M"];
    var tens = ["","X","XX","XXX","XL","L","LX","LXX","LXXX","XC"];
    var ones = ["","I","II","III","IV","V","VI","VII","VIII","IX"];
    var returnString = hundreds[Math.floor(number/100)] + tens[Math.floor(number/10)%10] + ones[number%10] + ". ";
    if (lowerCase) {
      return returnString.toLowerCase();
    }
    return returnString;
  }
  return number + ". ";
}

// Creates the PreFilled Link for each student
function getFormLink(studentData) {
  var link = FORM_PREFILL_LINK;
  link = link.replace("FIRST", clean(studentData[0]));
  link = link.replace("LAST", clean(studentData[1]));
  link = link.replace("EMAIL", clean(studentData[2]));
  link = link.replace("Finished", clean(studentData[3]));
  var returnData = {
    "htmlBody": "<a href='" + link + "' target='_blank'>Click here to check in.</a>",
    "plainText": link
  }
  //Logger.log(returnData);         // Debug
  return returnData;
}

// Cleans out special characters so it is safe to pass to the form as a URL
function clean(input) {
  input = new String(input);
  if(input.length > 0) {
    if (input.indexOf("%") > -1) {
      input = input.replace(new RegExp("%", 'g'), "%25");
    }
    if (input.indexOf("#") > -1) {
      input = input.replace(new RegExp("#", 'g'), "%23");
    }
    if (input.indexOf("^") > -1) {
      input = input.replace(new RegExp("\\^", 'g'), "%5E");
    }
    if (input.indexOf("&") > -1) {
      input = input.replace(new RegExp("&", 'g'), "%26");
    }
    if (input.indexOf("+") > -1) {
      input = input.replace(new RegExp("\\+", 'g'), "%2B");
    }
    if (input.indexOf("`") > -1) {
      input = input.replace(new RegExp("`", 'g'), "%60");
    }
    if (input.indexOf("=") > -1) {
      input = input.replace(new RegExp("=", 'g'), "%3D");
    }
    if (input.indexOf("[") > -1) {
      input = input.replace(new RegExp("\\[", 'g'), "%5B");
    }
    if (input.indexOf("]") > -1) {
      input = input.replace(new RegExp("]", 'g'), "%5D");
    }
    if (input.indexOf("\\") > -1) {
      input = input.replace(new RegExp("\\\\", 'g'), "%5C");
    }
    if (input.indexOf("{") > -1) {
      input = input.replace(new RegExp("\\{", 'g'), "%7B");
    }
    if (input.indexOf("}") > -1) {
      input = input.replace(new RegExp("\\}", 'g'), "%7D");
    }
    if (input.indexOf("|") > -1) {
      input = input.replace(new RegExp("\\|", 'g'), "%7C");
    }
    if (input.indexOf("\"") > -1) {
      input = input.replace(new RegExp("\"", 'g'), "%22");
    }
    if (input.indexOf("<") > -1) {
      input = input.replace(new RegExp("<", 'g'), "%3C");
    }
    if (input.indexOf(">") > -1) {
      input = input.replace(new RegExp(">", 'g'), "%3E");
    }
    if (input.indexOf(" ") > -1) {
      input = input.replace(new RegExp(" ", 'g'), "+");
    }
  }
  return input;
}