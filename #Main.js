//Highest-level function: Reads through emails with the label "Receipts",
//which have not been processed yet. 
//Also some emails labeled "Pending Correction", which is applied if there is an error in the
//subject. An email is sent, and looks for a reply with the proper formatting.

/*
Setup timer trigger and email for unmatched emails
Setup timer trigger and email reminder for unmatched expenses
Setup timer trigger for new & corrected emails
Response when email is matched (ideally show categorizations, and QBO data)
Response when email is parsed but unmatched 
Response when email cannot be parsed
*/

// 1) Trigger 1 min, go through all new emails in inbox, try to parse, match, and label them in 
//"Book Keeper: Successful Match", 
//"Book Keeper: Parse Error", 
//"Book Keeper: Other Error", 
//"Book Keeper: Awaiting Match", 
//"Book Keeper: Cancelled", 
//"Book Keeper: Multiple Matches"
// 2) Trigger 1 day, go through all "successful match" emails, add them QBO, and label them as "Added to QBO"
// 2) Trigger 1 day, go through all uncategorized quickbooks expenses and send a company wide email asking for receipts
// 3) Trigger 1 week, go through all "awaiting match" emails and see if there is a match, send email or "successful match" or "awaiting match" (which allows user to cancel)
var LIVE_MODE = true

function missingReceipts() {
  
  var txns = getRecentExpenses(365).concat(getRecentDeposits(365))

  if ( ! txns.length) return sendEmail("No Missing Receipts Today!")
  
  var csv = txns.reduce(function(line, txn) {
    return line+'<br>'+txn.date+('       $'+txn.amt).slice(txn.amt.length-3)+' <a href="https://qbo.intuit.com/app/'+getEntityName(txn.id)+'?txnId='+txn.id.slice(1)+'">'+getEntityName(txn.id)+'</a>, '+txn.bank+', '+(txn.memo && txn.memo.replace(/ {2,}/g, ' '))
  }, '')
  
  // 'office@sirum.org', 
  sendEmail('paloalto@sirum.org', "This Week's Missing Receipts", [
    'Hello Team,',
    '',
    'I hope you all are well! I was going through the books and saw a couple of transactions that are missing receipts.',
    'Could each of you please review the transactions below and submit any receipts for which you are responsible?',
    '<pre>'+csv+'</pre>',
    'Thanks',
    '<a href="https://docs.google.com/spreadsheets/d/1klEQQ7u73D8y1UdPLu2C3xChQ1ZlLfEpGfhACe9WNXQ/edit">Mr. Keeper</a>'
  ])
}

function scanInbox() {
  
  importCategories()
  
  var threads = GmailApp.search('in:inbox after:2018/03/22') //don't try to clear out historic backlog yet
  
  Log('scanInbox', threads)
  for (var i in threads)
    processNewThread(threads[i])
 
}

function succesfulMatches() {
  
  importCategories()
  
  var threads = GmailApp.search('label:The Book Keeper/Successful Match') //
  
  for (var i in threads)
    processPendingThread(threads[i], 'Successful Match')
}

function awaitingMatches() {
  
  importCategories()
  
  var threads = GmailApp.search('label:The Book Keeper/Awaiting Match')
  
  //debugEmail('Awaiting Matches', threads.length)
  
  for (var i in threads)
    processPendingThread(threads[i], 'Awaiting Match')
}

function processPendingThread(thread, label) {
  var message = getLastMessage(thread)
  var subject = parseTo(message) || message.getSubject()
  var parsed  = parseSubject(subject, message)
  
  try {
    var txns = searchTxns(parsed)
  } catch (e) {
    debugEmail('Search Txns Other Error:', e, parsed)
    removeLabel(thread, label)
    reply(message, 'Unfortunately, I encountered an unknown error when searching for a matching expense/deposit for this receipt.  Please report this to management<br><pre>'+prettyJson(parsed)+'</pre>')
    return addLabel(thread, 'Other Error')
  }
  
  var now  = new Date()
      
  if (now.getHours() == '10')
    debugEmail('processPendingThread', 'label', label, 'message.getSubject()', message.getSubject(), 'parsed', parsed, 'expenses', txns)
  
  if (txns.length == 1 && parsed.invoiceNos)
    return matchDeposit2Invoices(txns[0], parsed, thread, label)
    
  if (txns.length == 1 && validVendor(parsed, txns[0].id))
    return addToQBO(txns[0].id, parsed, thread, label)
   
  var error = parsed.errors[0]
  delete parsed.errors
  delete parsed.subject
  
  if (txns.length == 1) {
    removeLabel(thread, label)
    reply(message, error+'<br><br>This is how I understand your receipt:<br><pre>'+prettyJson(parsed)+'</pre>I found the matching expense/deposit below but can only add it if you resend me the receipt with exactly one vendor:<br><pre>'+prettyJson(txns)+'</pre>')
    //debugEmail('parsed.errors.length', 'subject', subject, 'parsed', parsed, 'message.getFrom()', message.getFrom())
    return addLabel(thread, 'Parse Error')
  }
  
  if (txns.length > 1) {
    multipleMatches(message, parsed, txns, thread)
    return removeLabel(thread, label)
  }
  
  //if (txns.length < 1)
  removeLabel(thread, label)
  addLabel(thread, 'Awaiting Match') 
    
  if (label == 'Awaiting Match' && (now.getDate() != 'Monday' || now.getHours() != '10')) return //Check Awaiting Match every hour but only Email on Mondays at 10am
  
  var cannotFind = label == 'Awaiting Match' ? 'still cannot not' : 'can no longer'
  reply(message, 'I just wanted to send you a brief update on your receipt:<br><pre>'+prettyJson(parsed)+'</pre>Unfortunately, I '+cannotFind+' find a transaction matching '+txns.query.split('WHERE')[1]+', but will keep checking each day until I find one or you reply telling me to "cancel"<br>') 
}

function processNewThread(thread) {
  
  if ( ! thread) Log('You cannot call this method directly from the editor')
  
  Log('processNewThread', typeof thread, thread, new Error().stack)
  var message   = getLastMessage(thread) 
  var subject   = parseTo(message) || message.getSubject()
  var plainBody = message.getPlainBody()
  var toAddress = message.getTo()
  
  //debugEmail('processNewThread', 'thread.getFirstMessageSubject()', thread.getFirstMessageSubject(), 'message.getTo()', message.getTo(), 'subject', subject, 'plainBody', plainBody, 'cancel', parseBody(subject, plainBody, /(^|[^"])\bcancel\b($|[^"])/i), 'TxnId', parseBody(subject, plainBody, /.{0,50}\b(\d{5})\b/), "hasLabel(thread, 'Multiple Matches')", hasLabel(thread, 'Multiple Matches'), 'thread.getLabels()', thread.getLabels())
  
  if (parseBody(subject, plainBody, /(^|[^"])\bcancel\b($|[^"])/i)) {
    removeLabel(thread, 'Multiple Matches')    
    removeLabel(thread, 'Awaiting Match')
    removeLabel(thread, 'Successful Match')
    reply(message, 'Ok, I understand. My apologies if this is because of an error I made.<br>I have cancelled this receipt, please resubmit if you still want it added.<br>')
    return addLabel(thread, 'Cancelled')
  } 
  
  var parsed = parseSubject(subject, message)
  
  //Someone responded with the ID of the matching expense.  Need to check after "cancel" since ID might be in quoted text below
  var txnId = parseBody(subject, plainBody, /(^|[^"])\b([DdEeCc]\d{4,5})\b($|[^"])/)
  if (txnId && validVendor(parsed, txnId) && hasLabel(thread, 'Multiple Matches')) {
    addToQBO(txnId[2], parsed, thread, 'Multiple Matches')
    return reply(message, 'Thanks for letting me know! I just added it to QuickBooks<br>')
  }
  
  try {
    var txns = searchTxns(parsed)
  } catch (e) {
    debugEmail('Search Txns Other Error:', e, parsed)
    removeLabel(thread, 'Multiple Matches')    
    removeLabel(thread, 'Awaiting Match')
    removeLabel(thread, 'Successful Match')
    reply(message, 'Unfortunately, I encountered an unknown error when searching for a matching expense/deposit for this receipt.  Please report this to management<br><pre>'+prettyJson(parsed)+'</pre>')
    return addLabel(thread, 'Other Error')
  }
  
  if ( ! txnId && txns.length == 1)
    validVendor(parsed, txns[0].id)
    
  if (parsed.invoiceNos) {
  
    var invoices = getInvoices(parsed.invoiceNos) || []
    
    var DocNumbers = invoices.map(function(invoice) { return invoice.DocNumber })
    for (var i in parsed.invoiceNos) {
      var invoiceNo = parsed.invoiceNos[i]
      if ( ! ~ DocNumbers.indexOf(invoiceNo))
        parsed.errors.push("Are you Invoice #"+invoiceNo+" exists?  I couldn't find it")
    }

    var invoiceTotals = invoices.map(function(invoice) { 
      
      if (invoice.LinkedTxn.length)
        parsed.errors.push('<a href="https://c30.qbo.intuit.com/app/invoice?txnId='+invoice.Id+'">Invoice #'+invoice.DocNumber+'</a> appears to already be paid.  Please select another invoice or delete this payment') // +' <pre>'+prettyJson(invoice)+'</pre>'
              
      return invoice.TotalAmt 
      
    })
    
    var invoiceTotal  = sum(invoiceTotals)
    if (parsed.total != invoiceTotal)
      parsed.errors.push('Did you specify the correct amount and invoices because $'+parsed.total+' does not match '+invoiceTotals.join('+')+' = $'+invoiceTotal+'?') 
  }
  
  txns.endDate = new Date(parsed.date)
  txns.endDate.setDate(txns.endDate.getDate()+5+7) //One week past match date just in case QBO bank feed is slow to pick things up

  if (Math.abs(new Date() - txns.endDate)/1000/60/60/24 > 365) parsed.errors.push('Are you sure you specified the correct year?')
    
  var errors = parsed.errors
  delete parsed.errors
  delete parsed.subject
  
  if (errors.length) {
    removeLabel(thread, 'Multiple Matches')    
    removeLabel(thread, 'Awaiting Match')
    removeLabel(thread, 'Successful Match')
    reply(message, 'Thanks for submitting the receipt below. Unfortunately, I have a couple of questions before I can categorize it correctly:<br> – '+errors.join('<br> – ')+'<br><br>Here is my best attempt to understand your current receipt:<br><pre>'+prettyJson(parsed)+'</pre>Sorry for the inconvenience, but could you please reply to this email changing your subject line to address the issue(s) above? Here are additional '+getSheetLink('instructions'))
    //debugEmail('parsed.errors.length', 'subject', subject, 'parsed', parsed, 'message.getFrom()', message.getFrom())
    return addLabel(thread, 'Parse Error')
  }
  
  //debugEmail('parsed', 'subject', subject, 'parsed', parsed, 'txns', txns)
  
  if ( ! txns.length && txns.endDate > new Date()) 
    noMatchesKeepLooking(message, parsed, txns, thread)
  
  else if ( ! txns.length) 
    noMatchesStopLooking(message, parsed, txns, thread)
  
  else if (txns.length > 1) 
    multipleMatches(message, parsed, txns, thread)
    
  else //txns.length == 1
    successfulMatch(message, parsed, txns, thread)
}

function getInvoiceNos(subject) {
   var invoiceNos = subject.match(/\bInvoices? *#?(\d{1,4})(,? *#?\d{1,4})*/ig)
   return invoiceNos && invoiceNos[0].split(/,? *#?\b/).slice(1)
}

function searchTxns(parsed) {
  if ( ! parsed.total || ! parsed.date) return []
  var expenses = (parsed.invoiceNos ? null : searchExpenses(parsed.total, parsed.date)) || []
  var deposits = searchDeposits(parsed.total, parsed.date) || []
  var txns = exactMatchesOrAll(expenses.concat(deposits))
  txns.query = expenses.query || deposits.query //Concat gets rid of our secret query property
  return txns
}

function validVendor(parsed, txnId) {
  if (txnId[0] != 'C' || parsed.vendors.length == 1 || ~ parsed.subject.indexOf('reimbursement')) return true //e.g. Reimbursement to Kiah for destruction doesn't need a vendor or a 1099
  parsed.errors.push('Did you specify the '+getSheetLink('vendors', false)+' correctly? Form 1099 requires each check to have exactly one vendor.')  
}

function noMatchesStopLooking(message, parsed, txns, thread) {
  removeLabel(thread, 'Multiple Matches')     
  removeLabel(thread, 'Awaiting Match')
  removeLabel(thread, 'Successful Match')
  reply(message, 'Thanks for submitting your receipt. Unfortunately, I could not find a matching expense/deposit with '+txns.query.split('WHERE')[1]+'. Please check the date and amount of your receipt below and resubmit:<br><pre>'+prettyJson(parsed)+'</pre>')
  addLabel(thread, 'No Matches') 
}

function noMatchesKeepLooking(message, parsed, txns, thread) {
  removeLabel(thread, 'Multiple Matches')     
  removeLabel(thread, 'No Matches')
  removeLabel(thread, 'Successful Match')
  reply(message, 'Thanks for submitting your receipt. Reply back with "cancel" if I misunderstand it:<br><pre>'+prettyJson(parsed)+'</pre>Unfortunately, I could not find a matching expense/deposit with '+txns.query.split('WHERE')[1]+' at the moment, but will keep checking until '+txns.endDate.toJSON().slice(0, 10)+', I find one, or you reply telling me to "cancel"<br>')
  addLabel(thread, 'Awaiting Match') 
}

function multipleMatches(message, parsed, txns, thread) {
  removeLabel(thread, 'Awaiting Match')
  removeLabel(thread, 'Successful Match')
  reply(message, 'Thanks for submitting your receipt. Reply back with "cancel" if I misunderstand it:<br><pre>'+prettyJson(parsed)+'</pre>I found '+txns.length+' matching expenses/deposits below, could you reply back with the ID of the match you want me to use?<br><pre>'+prettyJson(txns)+'</pre>', undefined, parsed.submitted) //Added parsed.submitted because of error on 2019-01-29.  George used special email address, but because multiple matches, George had to reply with the exact expense code, but in doing so the special email was lost so Book Keeper tried to use the generic Amazon email subject.  Fix this by updatting the subject of the reply with the one Book Keeper needs.
  addLabel(thread, 'Multiple Matches')
}

function successfulMatch(message, parsed, txns, thread) {

  var mins = 33 - new Date().getMinutes()
  
  if (mins <= 0) mins += 60
  
  removeLabel(thread, 'Multiple Matches')     
  removeLabel(thread, 'Awaiting Match')     
  reply(message, 'Thanks for submitting your receipt. Reply back with "cancel" if I misunderstand it:<br><pre>'+prettyJson(parsed)+'</pre>I found the matching expense/deposit below and will add it to QuickBooks in '+mins+' minutes unless you reply telling me to "cancel":<br><pre>'+prettyJson(txns)+'</pre>')
  addLabel(thread, 'Successful Match')
}

// if any field contains a split then all other fields must either have no splits or have the same number of splits
//
// Example of a Result:
// {
//  "submitted":<The orginial user supplied string>,
//  "errors":[],
//  "date":YYYY-MM-DD,
//  "total":13.80,
//  "programs":["US", "CA"],
//  "classes":["On-Going", "One-Time"],
//  "accounts":["Rx:Shipping", "Organizations"],
//  "amts":[6.50, 7.30],
// }

function parseTo(message) {
  
  var to   = getTo(message)
  var from = getFrom(message)
  
  var comments = to.match(/(.*)</) //e.g, US Boxes ('US Boxes' <receipts+us+boxes@sirum.org)
  var subject  = to.match(/receipts(.+?)@/) //e.g., +us+boxes

  if ( ! subject) return //debugEmail('parseTo has no matching subject', 'Date: '+date, 'From: '+from, 'To: '+to, 'Subject: '+subject, 'Comments: '+comments, message.getRawContent())
    
  Log('DEBUG parseTo', from, to, subject, comments, message.getRawContent())
      
  var submitted = subject[1].replace(/\+/g, ' ').trim()
    
  if (comments) //() get turned into the Emailer's name so we need to convert back to comments
    submitted += ' ('+comments[1].replace(/\+/g, ' ').trim()+')'
    
  //debugEmail('parseTo SUCCESS', 'Date: '+date, 'From: '+from, 'To: '+to, 'Subject: '+subject, 'Comments: '+comments, 'Submitted: '+submitted, message.getRawContent())
  
  return submitted  
}

function parseBody(subject, body, regex) {
  return subject.slice(0, 3) == 'Re:' && body.match(regex)
}

function testParseSubject() {
  importCategories()
  //var parsed = parseSubject("2018-05-01 One-time 501c3 Rent Rent Rent 50% 50%", "")
  //var parsed = parseSubject("2018-04-19 Good Pill Operations Pharmacist $602.00 $504.00 (James and Junushi)", "")
  //2018-05-02 SIRUM CA:Operations $14.89 split SIRUM US:Operations $14.89, Office:Supplies, Amazon Tape $29.77 total
  //var parsed = parseSubject("2018-06-05 501c3 Investment Fees Ongoing $140 (Human Interest)", "")
  //var parsed = parseSubject("Re: 2018-06-28 Fundraising Fees 501c3 Fundraising Check #90019 $1000 (Windfall Data Subscription)", "")
  //var toParse = "2018-07-12 Operations, Office Supplies, SIRUM US $62.14, SIRUM CA $62.14, Amazon total: $124.28"
  //var toParse = "Re: 2018-07-12 Operations, Office Supplies, SIRUM US 50%, SIRUM CA 50%, Amazon total: $124.28"
  var toParse = "Re: 2018-07-12 Ragini Operations, Office Supplies, SIRUM US 50%, SIRUM CA 50%, Amazon total: $124.28"
  //var expenses = searchExpenses(parsed.total, parsed.date)
  debugEmail('testParseSubject', parseSubject(toParse, ""))
}

function parseSubject(submitted, message) {
  
  submitted   = submitted.trim()
  var body    = message.getPlainBody()
  var subject = removeComments(submitted.trim())
  var parsed  = { submitted:submitted,  errors:[], invoiceNos:getInvoiceNos(subject), date:null, amts:[], percents:[], total:null, inEmail:[], attachments:message.getAttachments().length, programs:[], classes:[], accounts:[], vendors:[], subject:subject}
  
  findDate(parsed, message)
  findTotal(parsed, body)
  findAmts(parsed, body)
  findPercents(parsed, body)

  findVendors(parsed, vendors, body)

  findPrograms(parsed, programs)
  findClasses(parsed, classes)
  findAccounts(parsed, accounts) //Do this last because it's the most greedy regex and can inadvertently mess up the others
      
  if ( ! parsed.date)
    parsed.errors.push("Is the date in a YYYY-MM-DD or MM-DD-YYYY format?")
      
  if ( ! parsed.total)
    parsed.errors.push("Can you please specify the total for this receipt?")
  
  var amtSum   = Math.abs(sum(parsed.amts).toFixed(2))
  var total    = Math.abs((+parsed.total).toFixed(2))
  var emailSum = Math.abs(sum(parsed.inEmail).toFixed(2))
      
  //if (parsed.amts.length == 1 && parsed.amts[0] != parsed.total)
  //  parsed.errors.push("Could you please add the specified amount to the body of the email and briefly explain why it is not present in the receipt?")
  if (amtSum != total) { //Caused by 1) Amts don't add to a specified total, 2) percents don't add to 100%, 3) a mix of $s and %s
    if (amtSum == emailSum) //Maybe the receipt is itemized but does not include a total
      parsed.total = emailSum
    else if (parsed.total) //don't duplicate the warning: "Can you please specify the total for this receipt?"
      parsed.errors.push("Are you sure you the receipt amount(s) or percent(s) add up to the total?")
  } else if ( ! parsed.invoiceNos) { //these are checked against amt.length so no point in checking if we know because of the prior condition that amt.length is likely wrong
    checkLength(parsed, 'programs')
    checkLength(parsed, 'classes')
    checkLength(parsed, 'accounts')
  }
 
  return parsed
}

//Don't search within parenthesis/brackets if more than 2 letter (still include 501(c)3) //Use & not "and"
function removeComments(subject) {
  return subject.replace(/\[..+?\]/g, '').replace(/\(..+?\)/g, '').replace(/ and /g, ' & ') 
}

function checkLength(parsed, type) {
  
  //debugEmail('testParseSubject checkLength', type, parsed.amts.length, parsed[type].length, parsed[type], parsed)
  
  if (parsed[type].length == 1)
    parsed[type] = fillDefaults(parsed[type], parsed.amts.length)
    
  else if ( ! parsed[type].length) 
    parsed.errors.push('Did you specify the '+getSheetLink(type, false)+' correctly? I could not find any specified.') 
     
  else if (parsed[type].length < parsed.amts.length)
    parsed.errors.push('Did you forget to add one of the '+getSheetLink(type)+'? I got '+parsed[type].length+' but was expecting '+parsed.amts.length)
   
  else if (parsed[type].length > parsed.amts.length)
    parsed.errors.push('Did you accidently use a keyword in the description? Because there are '+parsed[type].length+' '+getSheetLink(type)+' and only '+parsed.amts.length+' '+getSheetLink('amounts')+'. Tip: I skip words enclosed in () or [].')
}

function getSheetLink(type, singular) {
  var gid = {instructions:931274598, classes:2063886151, programs:683783800, vendors:1847271979, accounts:0}[type]
  
  if (singular != null) //Undefined/Null leaves alone, TRUE converts to singular, FALSE converts to ambivalent e.g. program(es)
    type = type.replace(/(e?s)$/, singular ? '' : '($1)')
    
  return '<a href="https://docs.google.com/spreadsheets/d/1klEQQ7u73D8y1UdPLu2C3xChQ1ZlLfEpGfhACe9WNXQ/edit#gid='+gid+'">'+type+'</a>'
}

//Get rid of line breaks and indentation around arrays in order to save vertical space
function prettyJson(obj) {
  return JSON.stringify(obj).replace(/("id"|"memo"|"invoiceNos"|"total"|"amt"|"bank"|"account"|"submitted"|"attachments"|"date"|"amts"|"inEmail"|"percents"|"programs"|"classes"|"accounts"|"vendors"|"exact")/g, '\n  $1').replace(/}/g, '\n}')  //.replace(/\[\n */g, '[ ').replace(/\n *\]/g, ' ]').replace(/([^\]],)\n */g, '$1')
}

function addLabel(thread, label) {
  Logger.log('label 1', label)
  label = GmailApp.getUserLabelByName('The Book Keeper/'+label)
  Logger.log('label 2', label)
  thread.addLabel(label)
  thread.moveToArchive() //Move Out of Inbox
}

function removeLabel(thread, label) {
  thread.removeLabel(GmailApp.getUserLabelByName('The Book Keeper/'+label))
  //if ( ! thread.getLabels().length) thread.moveToInbox() //I don't believe this happens 
}

function hasLabel(thread, label) {
  var labels = thread.getLabels()
  for (var i in labels) {
    if (labels[i].getName() == 'The Book Keeper/'+label)
      return true
  }
}

//Even if we specify "to:receipts@sirum.org" in the search, the messages in the thread may still be
//ones that the Book Keeper sent.  We want to find the last message *NOT FROM* us.
function getLastMessage(thread) {
  
  if ( ! thread)
    return debugEmail('Error: getLastMessage', 'Thread has no getMessages method', typeof thread, thread, new Error().stack)
  
  var messages = thread.getMessages()
  
  Log('getLastMessage', thread)
 
  for (var i = messages.length-1; i >= 0; i--) {
    var msg = messages[i]
    if ( ! ~ msg.getFrom().indexOf('receipts@sirum.org'))
      return msg
  }
  
  return messages[messages.length-1] //if nothing in thread it maybe because someone was in Book Keeper's inbox reforwarding messages to him.
}

function debugSearch(threads) {
  var subjects = []
  for (var i in threads)
    subjects.push(threads[i].getFirstMessageSubject())
    
  debugEmail('debugSearch', 'subjects', subjects)
}

var programs
var classes
var accounts
var vendors
function importCategories() {
  var ss = SpreadsheetApp.openById('1klEQQ7u73D8y1UdPLu2C3xChQ1ZlLfEpGfhACe9WNXQ')
  programs = ss.getSheetByName("Programs").getDataRange().getValues()
  classes  = ss.getSheetByName("Classes").getDataRange().getValues()
  accounts = ss.getSheetByName("Accounts").getDataRange().getValues()
  vendors  = ss.getSheetByName("Vendors").getDataRange().getValues()
}