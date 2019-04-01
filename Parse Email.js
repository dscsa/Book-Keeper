

//Used by amtRegEx and findAmts
//Test Case "Fwd: 2017-12-27 One-Time GLG $1000, Long Foundation $25,000, Cecilia Henig Individual $100, total $26,100"
//Make sure 33.33% is not registered as an amount
var amtRegEx = /-?\$?-?[\d,]*\d\.\d{2}(?!%)\b|-?\$-?[\d,]*\d\b/g
var isTotal = /(\btotal:? |= ?)\$?([\d,.]*\d\b)|\$?([\d,.]*\d\b) total(?!:? \d|:? \$)/i

function findTotal(parsed, body) {
  //Use the ending "(?! \d| \$)" to make sure "Fwd: 2017-11-13 Background Check One-Time GP 14.93 501c3 14.93 Total 29.86" match 29.86 and not 14.93.
  //"2019-03-11 Invoices 1654, 1655, 1647, 1688, 1537, 1637, 1661 , 1511, 1524, Marilyn Groves $5000 one-time, total $6345"
  //Ex 2"Invoices 1691, 1598, 1642, 1684, 1656, 1673, 1633, 1671, 1653, 1613, 1681, 1690, 1604 TOTAL: $5110"
  var subject  = parsed.subject
  var isMatch  = subject.match(isTotal) //match totals e.g, $0.89+$0.44 = $1.33 or total $133 or $133 total

  if (isMatch)
    var total = cleanAmts([isMatch[2] || isMatch[3]])[0] //total will always be the 2nd or third capture group

  if ( ! total) return

  parsed.subject   = subject.replace(isMatch[0], '') //replace totals so they don't get confused with amts later
  parsed.total     = +total
  parsed.totalType = "specified"
  //debugEmail(subject, parsed.subject, isMatch, total, parsed)
}

function findPercents(parsed, body) {

  parsed.percents = ['100']  //if no $s and no $s specified, then assume 100% of the total

  var pctRegEx = /\b-?\d{1,2}%/g
  var matches  = parsed.subject.match(pctRegEx)

  if ( ! matches) return

  parsed.percents = cleanPercents(matches)

  if (parsed.amts.length < 2) {//don't allow mixing of %s and $s right now.  Need way to preserve order if we mix.  Note default of 100% should fill in an empty amt if there is a total available
    parsed.amts = parsed.percents.map(function(percent) { return +(parsed.total*percent/100).toFixed(2) })
    var sumAmts = sum(parsed.amts)
    if (sumAmts != parsed.total) parsed.amts[0] = +(+parsed.total - sumAmts + parsed.amts[0]).toFixed(2) //Fix Partial Cent Rounding Issues
  }
}

function findAmts(parsed, body) {

  var matches  = parsed.subject.match(amtRegEx) || []

  parsed.amts = cleanAmts(matches)

  //debugEmail('findAmts', parsed.subject, matches, parsed, body)
}

function findInvoiceNos(parsed, subject) {
  //Errant space after 1661: "2019-03-11 Invoices 1654, 1655, 1647, 1688, 1537, 1637, 1661 , 1511, 1524, Marilyn Groves $5000 one-time, total $6345"
  var invoiceNos = subject.match(/\bInvoices? *#?(\d{1,4})([, ]+#?\d{1,4})*/ig)
  return parsed.invoiceNos =  invoiceNos ? invoiceNos[0].split(/[, ]+#?\b/).slice(1) : []
}

function findInvoiceAmts(parsed) {

  if ( ! parsed.invoiceNos.length) return

  var invoices = getInvoices(parsed.invoiceNos) || []

  var DocNumbers = invoices.map(function(invoice) { return invoice.DocNumber })
  for (var i in parsed.invoiceNos) {
    var invoiceNo = parsed.invoiceNos[i]
    if ( ! ~ DocNumbers.indexOf(invoiceNo))
      parsed.errors.push("Are you Invoice #"+invoiceNo+" exists?  I couldn't find it")
  }

  parsed.invoiceAmts = invoices.map(function(invoice) {

    if (invoice.LinkedTxn.length)
      parsed.errors.push('<a href="https://c30.qbo.intuit.com/app/invoice?txnId='+invoice.Id+'">Invoice #'+invoice.DocNumber+'</a> appears to already be paid.  Please select another invoice or delete this payment') // +' <pre>'+prettyJson(invoice)+'</pre>'

    return invoice.TotalAmt
  })
}

function defaultTotal(parsed, body) {

  if (parsed.total) return

  //Only Searching BEFORE "We hope to see you again soon." Avoids getting higher prices for "Buy It Again", "Bargain Recommendations", and "Customers who bought ... also bought"

  var matches = body.split(/Top Picks for You|We hope to see you again soon|Buy It Again|Bargain Recommendations|Customers who bought|Recommended for you/i)[0].match(amtRegEx)
  parsed.inEmail = matches ? cleanAmts(matches) : []

  var allAmts = parsed.amts.concat(parsed.invoiceAmts)

  //Amt labeled total in the body of the email
  var shortBody = body.split('- Forwarded message -')
  shortBody = shortBody[0]+shortBody[1] //Don't include everything in forwarding chain since it might have lots of stuff
  var bodyTotal = shortBody.match(isTotal)

  var sumAmts = sum(allAmts)

  //debugEmail('defaultTotal', 'allAmts.length', allAmts.length, 'parsed.inEmail', parsed.inEmail, 'sumAmt', sum(allAmts), 'parsed', parsed, 'shortBody', shortBody, 'fullBody', body)

  //debugEmail('defaultTotal invoked', allAmts, parsed)

  if (sumAmts && parsed.attachments) {
    parsed.total = sumAmts
    parsed.totalType = "attachment"
  }

  else if (allAmts.length == 1 && parsed.inEmail.length == 0) { //Assume it's in an email attachment
    parsed.total = allAmts[0]
    parsed.totalType = "single amt"
  }

  else if (allAmts.length == 1 && inOrSum(parsed.inEmail, allAmts[0])) {
    parsed.total = allAmts[0] //if the one amt in subject is specified is in the email body then trust it.
    parsed.totalType = "single amt matching email amt or total"
  }

  else if (allAmts.length > 1 && inOrSum(parsed.inEmail, sumAmts)) {//several amts in smight but no total have been explicly sent in the subject
    parsed.total = sumAmts
    parsed.totalType = "sum of amts matching email"
  }

  else if (bodyTotal) {
    parsed.total = bodyTotal[2] || bodyTotal[3]
    if (allAmts.length == 0) parsed.amts.push(parsed.total)
    parsed.totalType = "total in email"
  }

  else if (allAmts.length == 0) { //otherwise if no amts and no total provided, assume the total is the max amt in the body
    parsed.total = Math.max.apply(null, parsed.inEmail)
    parsed.amts.push(parsed.total)
    parsed.totalType = "max amt in email"
  }
}

//Remove $ and , in amts e.g. $26,000 -> 26000
function cleanAmts(amts) {
  return amts.map(function(amt) { return amt.replace(/\$|,|\.00/g, '') })
}

function cleanPercents(percents) {
  return percents.map(function(percent) { return percent.slice(0, -1) })
}

function inOrSum(inEmail, val) {
  return ~ inEmail.indexOf(val) || (val == sum(inEmail)) //we can't scan attachments so if there are some, then just assume the user's amt/total is correct.
}

function sum(arr) {

  if ( ! arr || ! arr.reduce) {
    //debugEmail('no array given sum()', new Error('no array given sum()').stack, arr)
    return 0
  }

  return arr.reduce(function(sum, amt) { return +amt+sum }, 0)
}



/*
function findInvoices(parsed) {
  if (subject.match(/\binvoices?([ ,#]\d{4})+\b/i))
    parsed.errors.push("Is this an invoice?  I have not yet been trained on how to handle invoices.")


}*/

function findPrograms(parsed, programs) {
  var shortPrograms = pullDataFromColumn(0,programs)
  var fullPrograms = pullDataFromColumn(1,programs)
  var matches = findMatches(parsed, shortPrograms, fullPrograms)
  parsed.programs = matches.length ? matches : ['501c3'] //default value
}

function findClasses(parsed, classes) {
  var shortClasses = pullDataFromColumn(0,classes)
  var fullClasses = pullDataFromColumn(1,classes)
  parsed.classes = findMatches(parsed, shortClasses, fullClasses)

  //Allow a default value only for certain non-split deposits
  if (parsed.classes.length || parsed.accounts.length != 1) return

  if ( ~ parsed.accounts[0].indexOf('Contributions & Grants'))
    parsed.classes = ['One-Time'] //We consider grants one-time since we want to exlcude them from our financial sustainability calculations

  if ( ~ parsed.accounts[0].indexOf('Program Service Revenue'))
    parsed.classes = ['On-Going'] //We consider earned income on-going since we do want to inlcude them in our financial sustainability calculations
}

function findVendors(parsed, vendors, body) {
  var shortVendors = pullDataFromColumn(0,vendors)
  var fullVendors = pullDataFromColumn(1,vendors)

  Logger.log(['findVendors before subject match'])

  var matches = findMatches(parsed, shortVendors, fullVendors)

  Logger.log(['findVendors after subject match', matches])

  if ( ! matches.length)
    matches = findMatches(body, shortVendors, fullVendors)

  Logger.log(['findVendors finished', matches])
  parsed.vendors = matches.length ? matches : []
}


function findAccounts(parsed, accounts) {
  var shortAccounts = pullDataFromColumn(0,accounts)
  var fullAccounts = pullDataFromColumn(1,accounts)
  parsed.accounts = findMatches(parsed, shortAccounts, fullAccounts, '(\\b[a-zA-Z][a-zA-Z& ]+:)?') // capture (and remove) meta expense category if present.
}

//Look for exact matches listed in spreadsheet.
function findMatches(parsed, list, full, prefix){

  full = full || list

  var matches = []
  for(var i = 0; i < list.length; i++){
    var lookup = [
      prefix || '',
      '\\b',                                   // force start of word (no prefixes)
      list[i].replace(/([ \-:])/g, '[ \-:]?'), // spaces, colons, hyphens are optional.
      '\\w*'                                   // allow for suffixes e.g. Registration(s)?
    ]

    var regex = new RegExp(lookup.join(''), 'i')
    var match = (parsed.subject || parsed).match(regex) //make it work for email bodies as well
    if(match) {
      //debugEmail(parsed.subject, 'list[i]', list[i], "lookup.join('')", lookup.join(''), 'full[i]', full[i])
      matches.push({index:match.index, match:full[i]})
      i = i - replaceMatch(regex, match, parsed) //remove this match and repeat search again instead of a /g global regex flag which gets rid of the index
    }
  }

  //Return matches in order of appearence not in order of the keywords
  return matches.sort(function(a,b) { return a.index - b.index }).map(function(match) { return match.match })
}

//If match found, remove it from the subject (not submitted) so that we won't accidentally count the same keyword twice in case something else matches it)
function replaceMatch(regex, match, parsed) {
   if ( ! parsed.subject) return 0 //make it work for email bodies as well
   var replacement = fillDefaults(' ', match[0].length) //maintain indexes by filling in with space.
   parsed.subject = parsed.subject.replace(regex, replacement)
   return 1
}

//returns a column of index |ind| from the sheet data |data| as an array
//ignore column headings
function pullDataFromColumn(col,data){
  var res = []
  for(var i = 1; i < data.length;i++){
    var text = data[i][col].toString()
    if(text.length > 0) res.push(text)
  }
  return res
}

//Is it overkill to save the entire thread as PDF rather than just the most recent message
//https://ctrlq.org/code/19117-save-gmail-as-pdf?_ga=2.229660250.733420608.1521668974-273298195.1521668974
function thread2attachments(thread) {
  var attachments = []
  var msgs = thread.getMessages()
  var html = ""
  for (var i in msgs) {
   var msg = msgs[i]
    html += "From: " + msg.getFrom() + "<br />"
    html += "To: " + msg.getTo() + "<br />"
    html += "Date: " + msg.getDate() + "<br />"
    html += "Subject: " + msg.getSubject() + "<br />"
    html += "<hr />"
    html += msg.getBody().replace(/<img[^>]*>/g,"")
    html += "<hr />"

    var atts = msg.getAttachments()
    for (var j in atts) {
      attachments.push(atts[j])
    }
  }

  /* Save the attachment files and create links in the document's footer */
  if (attachments.length > 0) {
    var footer = "<strong>Attachments:</strong><ul>"
    for (var k in attachments) {
      footer += "<li>" + attachments[k].getName() + "</li>"
      /*
      try {
        attachments[k] = attachments[k].getAs("application/pdf")
      } catch (e) {}*/
    }
    html += footer + "</ul>"
  }

  /* Convert the Email Thread into a PDF File */
  var tempFile = DriveApp.createFile(thread.getFirstMessageSubject()+'.html', html, "text/html")
  attachments.unshift(tempFile.getAs("application/pdf"))
  tempFile.setTrashed(true)
  return attachments

}

//Works with Arrays and Strings
function fillDefaults(toFill, length) {
  for (var i = 1; i<length; i++)
    toFill = toFill.concat(toFill[0])

  return toFill
}

/*
function findAmts(parsed, body) {
  //Test Case "Fwd: 2017-12-27 One-Time GLG $1000, Long Foundation $25,000, Cecilia Henig Individual $100, total $26,100"
  var subject  = parsed.subject.replace(/(total|=) \$?[\d,.]*\d|\$?[\d,.]*\d total/i, '') //remove totals e.g, $0.89+$0.44 = $1.33 or total $133 or $133 total
  var matches  = cleanAmts(subject.match(amtRegEx) || [])

  if (matches.length) {
    parsed.amts = parsed.amts.concat(matches)
    parsed.total = parsed.amts.reduce(function(sum, amt) { return +amt+sum }, 0)
  } else { //if amts not in the subject then parse the email body for the max value.
    //Only Searching BEFORE "We hope to see you again soon." Avoids getting higher prices for "Buy It Again", "Bargain Recommendations", and "Customers who bought ... also bought"
    body = body.split(/We hope to see you again soon|Buy It Again|Bargain Recommendations|Customers who bought/i)[0]
    parsed.total = Math.max.apply(null, cleanAmts(body.match(amtRegEx) || []))
    if (parsed.total) parsed.amts = [parsed.total]
  }
}*/
