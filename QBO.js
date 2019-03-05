//Make sure you enter the redirect_uri (currently https%3A%2F%2Fscript.google.com%2Fmacros%2Fd%2F1sWiQq8knZjsJMQO8vOL8Y8LMLixvPh4XmOgyd-DbvpjGd_h5oL2exKmW%2Fusercallback)
//into QBO on the KEYs page
//--------------------------------------------------------

function testReceivable() {
   var res = getReceivable(30)
   Log(res)
}

function testSearchExpenses() {
   var res = searchExpenses(140, "2018-06-05")
   Log(res)
}

function testSearchDeposits() {
   var res = searchDeposits("250", "2018-06-05")
   Log(res)
}

function testUpload() {
   var res = uploadPDF(10021)
   Log(res)
}

function testNewRecent() {
   var service = getService();

  if ( ! service.hasAccess())
    return debugEmail('Open the following URL and re-run the script:', service.getAuthorizationUrl())

  var expenses = []

  //MetaData.CreateTime >= '2017-12-12T12:50:30Z'
  expenses.query = "SELECT * FROM Deposit WHERE TxnDate >= '2017-12-12' MAXRESULTS 100" //MetaData.LastUpdatedTime"
  var res = queryQBO(expenses.query, service)
  Log(expenses.query, res.QueryResponse.Deposit.length, res.QueryResponse.Deposit.map(function(expense) { return expense.Line[0].DepositLineDetail.AccountRef.name }))
}

function testClassify() {
  //https://qbo.intuit.com/app/expense?txnId=10021
  /*classifyExpense(10021, {
    amts:['50', '100', '150'],
    classes:['0 Program:On-Going', '0 Program:One-Time', '3 Lobbying'],
    programs:['501c3', 'Charitable Returns', 'SIRUM US'],
    accounts:['Accounting:Accounting, Auditing, & Tax', 'Other Service Fees:Destruction', 'Office:Banking & Investment Fees'],
    submitted:"This is Adam's user submitted line"
  })*/

  //https://qbo.intuit.com/app/expense?txnId=10021
  classifyTxn('E10021', {
    total:300,
    amts:['300'],
    classes:['0 Program:One-Time'],
    programs:['SIRUM US'],
    accounts:['Office:Registrations & Fees'],
    submitted:'Point Of Sale Withdrawal 7000202 Point Of Sale Withdrawal 700020201 DELAWARE DIV OF COR'
    //Point Of Sale Withdrawal 7000202 Point Of Sale Withdrawal 700020201 DELAWARE DIV OF COR
  })
}

function lastYearsExpenses() {
  return getRecentExpenses(365)
}

function getRecentExpenses(days){

  days = days || 7

  var service = getService();

  if ( ! service.hasAccess())
    return debugEmail('Open the following URL and re-run the script:', service.getAuthorizationUrl())

  //I shouldn't have to do it this way but GAS is doing weird things with dates right now
  var startDate = new Date()
  startDate.setDate(startDate.getDate()- days)

  var expenses = []

  expenses.query = "SELECT * FROM Purchase WHERE TxnDate > '" + startDate.toJSON().slice(0, 10) + "' ORDERBY TxnDate DESC MAXRESULTS 1000"
  var res = queryQBO(expenses.query, service)
 //debugEmail(expenses.query, days, res.QueryResponse.Purchase && res.QueryResponse.Purchase.reverse().slice(990))
  if(res.Fault)
    return Log("Error: There was an error is your query", expenses.query, res)

  Log(expenses.query, res.QueryResponse.Purchase && res.QueryResponse.Purchase.length, res)

  for(var i in res.QueryResponse.Purchase) {
    var expense = res.QueryResponse.Purchase[i]
    //Couldn't figure out how to do include account in query's WHERE, so filter the results here instead
    if ( ! expense.Line[0].AccountBasedExpenseLineDetail.AccountRef.name || expense.Line[0].AccountBasedExpenseLineDetail.AccountRef.name == "Uncategorized Expense")
      expenses.push({
        id:getTxnId(expense, i),
        date:expense.TxnDate,
        memo:expense.Line[0].Description+' | '+(expense.Line[0].Description == expense.PrivateNote ? '' : ' | '+expense.PrivateNote), //+' | '+JSON.stringify(expense),
        amt:(+expense.TotalAmt).toFixed(2),
        bank:expense.AccountRef.name,
      })
  }

  return expenses
}

function getRecentDeposits(days){

  days = days || 7

  var service = getService();

  if ( ! service.hasAccess())
    return debugEmail('Open the following URL and re-run the script:', service.getAuthorizationUrl())

  //I shouldn't have to do it this way but GAS is doing weird things with dates right now
  var startDate = new Date()
  startDate.setDate(startDate.getDate()- days)

  var deposits = []

  deposits.query = "SELECT * FROM Deposit WHERE TxnDate > '" + startDate.toJSON().slice(0, 10) + "' ORDERBY TxnDate DESC MAXRESULTS 1000"
  var res = queryQBO(deposits.query, service)

  if(res.Fault)
    return Log("Error: There was an error is your query", deposits.query, res)

  for(var i in res.QueryResponse.Deposit) {
    var deposit = res.QueryResponse.Deposit[i]
    //Couldn't figure out how to do include account in query's WHERE, so filter the results here instead
    //First Part (before the ||) is for Quickbooks Credit Card Payments of Invoices.  They don't have an AccountRef but do have a linked transaction
    if (isUncategorizedDeposit(deposit, "Accounts Receivable:Uncategorized Deposits"))
      deposits.push({
        id:getTxnId(deposit),
        date:deposit.TxnDate,
        memo:deposit.Line[0].Description+(deposit.Line[0].Description == deposit.PrivateNote ? '' : ' | '+deposit.PrivateNote), //+' | '+JSON.stringify(deposit),
        amt:(+deposit.TotalAmt).toFixed(2),
        bank:deposit.DepositToAccountRef.name,
      })
  }

  Log(deposits.query, res.QueryResponse.Deposit && res.QueryResponse.Deposit.length, deposits.length, deposits, res)


  return deposits
}


//https://developer.intuit.com/docs/00_quickbooks_online/2_build/20_explore_the_quickbooks_online_api/50_data_queries
//Select Statement = SELECT * | count(*)
//        FROM IntuitEntity
//        [WHERE WhereClause]
//        [ORDERBY OrderByClause]
//        [STARTPOSITION  Number] [MAXRESULTS  Number]
//AND
//TotalAmt = '" + amount.toString() + "'"
//
//

function searchExpenses(amt, date, account){

  var service = getService();

  if ( ! service.hasAccess())
    return debugEmail('Open the following URL and re-run the script:', service.getAuthorizationUrl())

  //I shouldn't have to do it this way but GAS is doing weird things with dates right now
  var startDate = new Date(date)
  var endDate   = new Date(date)
  startDate.setDate(startDate.getDate()-3)
  endDate.setDate(endDate.getDate()+5)

  var lowAmt  = +amt - 1
  var highAmt = +amt + 1

  var account = account || "Uncategorized Expense"
  var expenses = []

  expenses.query = "SELECT * FROM Purchase WHERE TotalAmt > '"+lowAmt.toFixed(2)+"' and TotalAmt < '"+highAmt.toFixed(2)+"' and TxnDate > '" + startDate.toJSON().slice(0, 10) + "' and TxnDate < '" + endDate.toJSON().slice(0, 10) + "'"
  var res = queryQBO(expenses.query, service)

  if(res.Fault)
    return Log("Error: There was an error is your query", expenses.query, res)

  Log(expenses.query, res.QueryResponse.Purchase && res.QueryResponse.Purchase.length, res)

  for(var i in res.QueryResponse.Purchase){
    var expense = res.QueryResponse.Purchase[i]
    //Couldn't figure out how to do include account in query's WHERE, so filter the results here instead
    if ( ! expense.Line[0].AccountBasedExpenseLineDetail.AccountRef.name || expense.Line[0].AccountBasedExpenseLineDetail.AccountRef.name == account)
      expenses.push({
        id:getTxnId(expense),
        date:expense.TxnDate,
        memo:expense.PrivateNote,
        amt:(+expense.TotalAmt).toFixed(2),
        exact:expense.TotalAmt == amt, //Should this be based on DATE too?
        account:account,
        bank:expense.AccountRef.name,
      })
  }

  return exactMatchesOrAll(expenses)
}

//Search for a matching Deposit. If no match, add label "awaiting match"
//Search invoice by invoice#. If no match, add label "parse error"
//If either is not found, email user with all errors
//Create a payment that references the invoice's id
//Update the deposit with a linked transaction containing the payment id

function searchDeposits(amt, date, account){

  account = account || "Accounts Receivable:Uncategorized Deposits"

  var service = getService();

  if ( ! service.hasAccess())
    return debugEmail('Open the following URL and re-run the script:', service.getAuthorizationUrl())

  //I shouldn't have to do it this way but GAS is doing weird things with dates right now
  var startDate = new Date(date)
  var endDate   = new Date(date)
  startDate.setDate(startDate.getDate()-3)
  endDate.setDate(endDate.getDate()+5)

  var account = account
  var deposits = []

  deposits.query = "SELECT * FROM Deposit WHERE TxnDate > '" + startDate.toJSON().slice(0, 10) + "' and TxnDate < '" + endDate.toJSON().slice(0, 10) + "'"
  var res = queryQBO(deposits.query, service)

  if(res.Fault)
    return Log("Error: There was an error is your query", deposits.query, res)

  for(var i in res.QueryResponse.Deposit){
    var deposit = res.QueryResponse.Deposit[i]
    //Couldn't figure out how to do include account or TotalAmt in query's WHERE, so filter the results here instead
    //This matches
    if (deposit.TotalAmt > +amt - 1 && deposit.TotalAmt < +amt + 1 && isUncategorizedDeposit(deposit, account)) {
      Log("DEBUG DEPOSIT FILTER:", deposit.TotalAmt, +amt - 1, +amt + 1)
      deposits.push({
        id:getTxnId(deposit),
        date:deposit.TxnDate,
        memo:deposit.PrivateNote,
        amt:(+deposit.TotalAmt).toFixed(2),
        exact:deposit.TotalAmt == amt,  //Should this be based on DATE too?
        account:account,
        bank:deposit.DepositToAccountRef.name,
      })
    }
  }

  Log(deposits.query, res.QueryResponse.Deposit && res.QueryResponse.Deposit.length, deposits.length, deposits, res)

  return exactMatchesOrAll(deposits)
}

//2018-10-09     $75.00 deposit, Cash:SVB 501c3 6746, Paid Labs could not find a match for this deposit | | {"DepositToAccountRef":{"value":"465","name":"Cash:SVB 501c3 6746"},"TotalAmt":75,"domain":"QBO","sparse":false,"Id":"12380","SyncToken":"1","MetaData":{"CreateTime":"2018-10-22T15:21:15-07:00","LastUpdatedTime":"2018-12-21T22:04:47-08:00"},"TxnDate":"2018-10-09","CurrencyRef":{"value":"USD","name":"United States Dollar"},"PrivateNote":"Paid Labs could not find a match for this deposit","Line":[{"Id":"1","LineNum":1,"Description":"Paid Labs could not find a match for this deposit","Amount":75,"DetailType":"DepositLineDetail","DepositLineDetail":{"Entity":{"value":"416","name":"1 Donor Unrestricted:Unrestricted Internally","type":"CUSTOMER"},"ClassRef":{"value":"3000000000000889887","name":"0 Program:On-Going:501c3"},"AccountRef":{"value":"432","name":"Accounts Receivable:Uncategorized Deposits"}}}]}
function getPayment(Id) {

  Id = Id || 12781

  var service = getService();

  if ( ! service.hasAccess())
    return debugEmail('Open the following URL and re-run the script:', service.getAuthorizationUrl())

  query = "SELECT * FROM Payment WHERE Id = '" +Id + "'"
  var res = queryQBO(query, service)

  if(res.Fault)
    return debugEmail("Error: There was an error is your query", query, res)

  debugEmail('getPayment', Id, res.QueryResponse.Payment.length, res.QueryResponse.Payment)
}

function getInvoices(docNumbers) {

  var service = getService();

  if ( ! service.hasAccess())
    return debugEmail('Open the following URL and re-run the script:', service.getAuthorizationUrl())

  query = "SELECT * FROM Invoice WHERE DocNumber IN ("+addQuotes(docNumbers)+")"
  var res = queryQBO(query, service)

  if(res.Fault)
    return debugEmail("Error: There was an error is your query", query, res)

  //debugEmail('getInvoice', docNumbers, query, res.QueryResponse)

  return res.QueryResponse.Invoice
}

//Quickbooks Credit Card Payments of Invoices don't have an AccountRef but do have a linked transaction
function isUncategorizedDeposit(deposit, accountName) {
  var account = deposit.Line[0].DepositLineDetail.AccountRef
  return account ? account.name == accountName : !deposit.Line[0].LinkedTxn
}

//Return exact matches if there are any, otherwise return all matches
function exactMatchesOrAll(matches) {
  var exactMatches = matches.filter(function(match) { return match.exact })
  return exactMatches.length >= 1 ? exactMatches : matches
}

function matchDeposit2Invoices(txn, parsed, thread, label) {
  try {
    removeLabel(thread, label)

    addInvoicePayment(txn, parsed)

    var attachments = thread2attachments(thread)
    for (var i in attachments)
      uploadPDF(txn.id, attachments[i])

    addLabel(thread, 'Added to QBO')

    infoEmail('matchDeposit2Invoices', txn.id, parsed, label)
  } catch (e) {
    removeLabel(thread, 'Multiple Matches')
    removeLabel(thread, 'Awaiting Match')
    removeLabel(thread, 'Successful Match')
    addLabel(thread, 'Other Error')
    debugEmail(e)
  }
}

//Couldn't figure out how to make a $0.00 payment that would link invoices with deposits.  Instead create a payment
//worth the full deposit amount and then set the deposit amount to 0
function addInvoicePayment(txn, parsed) {

  //For debugging
  txn = txn || {
    "id": "D12378",
    "date": "2018-10-15",
    "memo": "Paid Labs could not find a match for this deposit",
    "amt": "175.00",
    "exact": true,
    "account": "Accounts Receivable:Uncategorized Deposits",
    "bank": "Cash:SVB 501c3 6746"
  }

  //For debugging
  parsed.invoiceNos = parsed.invoiceNos || [
    "1643",
    "1644",
    "1645"
  ]

  var service = getService();

  if ( ! service.hasAccess())
    return debugEmail('Open the following URL and re-run the script:', service.getAuthorizationUrl())


  var oldTxn = UrlFetchApp.fetch("https://quickbooks.api.intuit.com/v3/company/" + COMPANY_ID + "/deposit/"+txn.id.slice(1),  {
    headers: {Authorization: 'Bearer ' + service.getAccessToken(), 'Accept': 'application/json'},
    method: 'get',
    muteHttpExceptions: true
  })

  oldTxn = JSON.parse(oldTxn)

  var invoices = getInvoices(parsed.invoiceNos) || []

  var lines = [{
    Amount:txn.amt,
    LinkedTxn: [
      {
        TxnId:txn.id.slice(1),
        TxnType:"Deposit"
      }
    ]
  }]

  for (var i in invoices) {
    lines.push({
     Amount: invoices[i].TotalAmt,
     LinkedTxn: [
      {
       //TxnLineId:i,
       TxnId:invoices[i].Id,
       TxnType:"Invoice"
      }
     ]
    })
  }

  var payment = {
   CustomerRef: {
    value: "416",
    name: "Unrestricted Internally"
   },
   //DepositToAccountRef: oldTxn.Deposit.DepositToAccountRef,
   PaymentMethodRef: {
    value: "12"  //Check
   },
   TotalAmt:txn.amt,
   TxnDate:txn.date,
   Line:lines,
    PrivateNote:("Book Keeper matched to Invoices "+parsed.invoiceNos+" on "+new Date().toJSON().slice(0, -8)+"\n"+oldTxn.Deposit.PrivateNote+"\n"+prettyJson(parsed)).slice(0, 4000)
  }

  var url     = "https://quickbooks.api.intuit.com/v3/company/" + COMPANY_ID + "/payment"
  var mime    = 'application/json'
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {Authorization: 'Bearer '+service.getAccessToken(), Accept:mime , "Content-Type":mime},
    payload: JSON.stringify(payment),
    muteHttpExceptions:true
  })

  res = JSON.parse(res)

  if ( ! res.Payment || ! res.Payment.Line.length) {
    debugEmail('Error linking invoices to deposit!', payment, res)
    throw res
  }

  oldTxn.Deposit.PrivateNote  = payment.PrivateNote
  oldTxn.Deposit.Line[0].Amount = 0
  oldTxn.Deposit.Line[0].DepositLineDetail.AccountRef = {
    value:"106",
    name:"Accounts Receivable:Earned Income Receivable"
  }

  oldTxn.Deposit.Line.push({
    Amount:txn.amt,
    LinkedTxn:[
      {
        TxnLineId:0,
        TxnId:res.Payment.Id,
        TxnType:"Payment"
      }
    ]
  })

  Log("classifyTxn: Updated Txn", txn)

  var url  = "https://quickbooks.api.intuit.com/v3/company/" + COMPANY_ID + "/deposit?operation=update"
  var mime = 'application/json'
  var newTxn = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {Authorization: 'Bearer '+service.getAccessToken(), Accept:mime , "Content-Type":mime},
    payload: JSON.stringify(oldTxn.Deposit),
    muteHttpExceptions:true
  })

  debugEmail('Invoices Linked to Desposit', oldTxn.Deposit,JSON.parse(newTxn))
}

function addToQBO(id, parsed, thread, label) {

  try {
    removeLabel(thread, label)

    classifyTxn(id, parsed)

    var attachments = thread2attachments(thread)
    for (var i in attachments)
      uploadPDF(id, attachments[i])

    addLabel(thread, 'Added to QBO')
    infoEmail('addToQBO', id, parsed, label)
  } catch (e) {
    removeLabel(thread, 'Multiple Matches')
    removeLabel(thread, 'Awaiting Match')
    removeLabel(thread, 'Successful Match')
    addLabel(thread, 'Other Error')
    debugEmail(e)
  }
}

// Example of a Parsed Object:
// {
//  "submitted":<The orginial user supplied string>,
//  "date":YYYY-MM-DD,
//  "programs":["US", "CA"],
//  "classes":["On-Going", "One-Time"],
//  "accounts":["Rx:Shipping", "Organizations"],
//  "amounts":[6.5, 7.3]
// }
function classifyTxn(id, parsed){
  var service = getService();

  if ( ! service.hasAccess())
    return debugEmail('Open the following URL and re-run the script:', service.getAuthorizationUrl())

  var entityType   = getEntityType(id)
  var lineName     = getLineName(id)
  var AccountRefs  = getAccountRefs(parsed.accounts)
  var CustomerRefs = getCustomerRefs(id, parsed.classes)
  var VendorRefs   = getVendorRefs(parsed.vendors)
  var ClassRefs    = getClassRefs(parsed.programs, parsed.classes)

  if ( ! AccountRefs || ! CustomerRefs || ! ClassRefs)
    return debugEmail('Aborting classifyExpense because missing Refs', AccountRefs, CustomerRefs, ClassRefs, parsed)

  //Update needs to full purchase/deposit object, so we retrieve it first.
  var oldTxn = UrlFetchApp.fetch("https://quickbooks.api.intuit.com/v3/company/" + COMPANY_ID + "/"+entityType.toLowerCase()+"/"+id.slice(1),  {
    headers: {Authorization: 'Bearer ' + service.getAccessToken(), 'Accept': 'application/json'},
    method: 'get',
    muteHttpExceptions: true
  })

  Log("classifyTxn: Retrieve Txn", id, parsed, oldTxn)

  infoEmail('classifyTxn: Retrieve Txn', id, parsed, JSON.stringify(JSON.parse(oldTxn), null, "  ")) //data must be in proto, have to parse and stringify for it to be visible

  var txn = JSON.parse(oldTxn)[entityType]
  var Line = txn.Line[0] //Additional lines should use the first line as a template
  txn.Line = [] //Get rid of any existing lines

  //Get Receipt Line Items to sum up to Purchase Amt
  parsed.amts[0] = +parsed.amts[0] + txn.TotalAmt - parsed.total

  for (var i in parsed.amts) {
    txn.Line[i] = JSON.parse(JSON.stringify(Line))
    txn.Line[i].Id = +i+1
    //txn.Line[i].Description = parsed.submitted
    txn.Line[i].Amount = Math.abs(parsed.amts[i]) //transaction amounts must always be > 0 or you will get "Business Validation Error: You must specify a transaction amount that is 0 or greater." this seems to be working for refunds/returns although not sure how: Maybe {"Name":"TxnType","Value":"54"} means a credit/debit?

    if (id[0] == 'D' && CustomerRefs[i]) {
       txn.Line[i][lineName].Entity = CustomerRefs[i]
       txn.Line[i][lineName].Entity.type = 'CUSTOMER'
    }
    else
      txn.Line[i][lineName].CustomerRef = CustomerRefs[i]

    if (id[0] == 'C' && VendorRefs[i]) {
      txn.EntityRef = VendorRefs[i]
      txn.EntityRef.type   = 'VENDOR'
    }

    txn.Line[i][lineName].AccountRef = AccountRefs[i]
    txn.Line[i][lineName].ClassRef = ClassRefs[i]
  }

  txn.PrivateNote += "\n\nBook Keeper matched on "+new Date().toJSON().slice(0, -8)+": "+prettyJson(parsed)

  Log("classifyTxn: Updated Txn", txn)

  var url  = "https://quickbooks.api.intuit.com/v3/company/" + COMPANY_ID + "/"+entityType.toLowerCase()+"?operation=update"
  var mime = 'application/json'
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {Authorization: 'Bearer '+service.getAccessToken(), Accept:mime , "Content-Type":mime},
    payload: JSON.stringify(txn),
    muteHttpExceptions:true
  })

  res = JSON.parse(res)

  Log("Classify Txn Response", res)

  if(res.Fault)
    throw ["Error: There was an error while classifying your TXN", 'RES', res, 'PARSED', parsed, 'TXN', txn, 'OLD TXN', JSON.stringify(JSON.parse(oldTxn), null, "  ")] //data must be in proto, have to parse and stringify for it to be visible
}

//Upload a pdf and then attach it to the relavant item
function uploadPDF(id, blob){

  var service = getService();

  if ( ! service.hasAccess())
    return debugEmail('Open the following URL and re-run the script:', service.getAuthorizationUrl())

  blob = blob || DriveApp.getFileById(DriveApp.getFilesByName("123856-9452.pdf").next().getId()).getBlob();

  var entityType = getEntityType(id)
  var boundary   = "gfkjndfgdg7e8dg"
  var attachment = Utilities.base64Encode(blob.getBytes())

  var json = {
    "AttachableRef": [{"EntityRef": {"type":entityType, "value":id.slice(1) }}],
    "FileName":blob.getName(),
    "ContentType":blob.getContentType()
  }

  //https://developer.intuit.com/docs/api/accounting/attachable
  //https://developer.intuit.com/docs/00_quickbooks_online/2_build/60_tutorials/0050_attach_images_and_notes#/Uploading_and_linking_new_attachments
  //https://help.developer.intuit.com/s/question/0D50f000050U3vVCAS/help-with-uploading-a-file-using-the-api
  var body = [
    '--'+boundary,
    'Content-Disposition: form-data; name="file_metadata_01"; filename="attachment.json"',
    'Content-Type: application/json',
    '', //Double line break is important
    JSON.stringify(json, null, '  '),
    '--'+boundary,
    'Content-Disposition: form-data; name="file_content_01"; filename="'+blob.getName()+'"',
    "Content-Type: "+blob.getContentType(),
    "Content-Transfer-Encoding: base64",
    '', //Double line break is important
    attachment,
    '--'+boundary+'--'
  ]

  var url = "https://quickbooks.api.intuit.com/v3/company/" + COMPANY_ID + "/upload"
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: "multipart/form-data; boundary="+boundary,
    headers: {Authorization: 'Bearer '+service.getAccessToken(), Accept:'application/json'},
    payload: body.join("\r\n"),
    muteHttpExceptions:true
  })

  res = JSON.parse(res)

  Log("Upload PDF Response", res)

  if(res.Fault)
    throw ["Error: There was an error is your attachment", body, res]
}

function getClassRefs(programs, classes){

  //If Class starts with 0 Program then we append Entity
  classes = classes.map(function(class, i) {
    return class[0] === '0' ? class+':'+programs[i] : class
  })

  var service = getService();

  if ( ! service.hasAccess())
    return debugEmail('Open the following URL and re-run the script:', service.getAuthorizationUrl())

  var query = 'SELECT * FROM Class WHERE FullyQualifiedName IN ('+addQuotes(classes)+')'

  var res = queryQBO(query, service)

  if(res.Fault || ! res.QueryResponse.Class)
    return debugEmail("Error: Couldn't match all Classes!", query, res)

  if(res.QueryResponse.Class.length > classes.length)
    return debugEmail("Error: Multiple Class Matches!", query, res)

  //Log(query, res.QueryResponse.Class.length, res)

  return classes.map(function(class) {
    return {name:class, value:findMatch(class, res.QueryResponse.Class)}
  })
}

function getCustomerRefs(id, classes) {

  return classes.map(function(class) {

    if ((id[0] == 'D' || id[0] == 'd') && ~ class.indexOf('One-Time'))
      return //One-Time Deposits are some type of Donor:Restricted Grant TODO: Do emails have enough info to figure this out?

    if ( ~ class.indexOf('Lobbying'))
      return {value:"416", name:"0 Donor Unrestricted:Internally Unrestricted"}

    return {value:"104", name:"0 Donor Unrestricted:Internally Restricted (No Lobbying)"}
  })
}

function getVendorRefs(vendors) {

  if ( ! vendors.length) return []

  //If Class starts with 0 Prorgam then we append Entity
  var service = getService();

  if ( ! service.hasAccess())
    return debugEmail('Open the following URL and re-run the script:', service.getAuthorizationUrl())

  var query = 'SELECT * FROM Vendor WHERE DisplayName IN ('+addQuotes(vendors)+')'

  var res = queryQBO(query, service)

  Log('getVendorRefs', query, res)

  if(res.Fault || ! res.QueryResponse.Vendor)
    return debugEmail("Error: Couldn't match all Vendors!", query, res)

  if(res.QueryResponse.Vendor.length > vendors.length)
    return debugEmail("Error: Multiple Vendor Matches!", query, res)

  //Log(query, res.QueryResponse.Account.length, res)

  return vendors.map(function(vendor) {
    return {name:vendor, value:findMatch(vendor, res.QueryResponse.Vendor)}
  })
}

function getAccountRefs(accounts) {

  //If Class starts with 0 Prorgam then we append Entity
  var service = getService();

  if ( ! service.hasAccess())
    return debugEmail('Open the following URL and re-run the script:', service.getAuthorizationUrl())

  var query = 'SELECT * FROM Account WHERE FullyQualifiedName IN ('+addQuotes(accounts)+')'

  var res = queryQBO(query, service)

  Log('getAccountRefs', query, res)

  if(res.Fault || ! res.QueryResponse.Account)
    return debugEmail("Error: Couldn't match all Accounts!", query, res)

  if(res.QueryResponse.Account.length > accounts.length)
    return debugEmail("Error: Multiple Account Matches!", query, res)

  //Log(query, res.QueryResponse.Account.length, res)

  return accounts.map(function(account) {
    return {name:account, value:findMatch(account, res.QueryResponse.Account)}
  })
}

function getTxnId(txn, i) {
  if ( ! txn.PurchaseEx) return 'D'+txn.Id
  return (txn.PrivateNote && txn.PrivateNote.match(/\bcheck\b/i)) ? 'C'+txn.Id : 'E'+txn.Id
}

function getEntityName(id) {
  return id[0] == 'D' || id[0] == 'd' ? 'deposit' : 'expense'
}

function getEntityType(id) {
  return id[0] == 'D' || id[0] == 'd' ? 'Deposit' : 'Purchase'
}

function getLineName(id) {
  return id[0] == 'D' || id[0] == 'd' ? 'DepositLineDetail' : 'AccountBasedExpenseLineDetail'
}

//Using IN() in queries saves us from separate queries but does NOT preserve ordering
function findMatch(needle, haystack) {
  var match = haystack.reduce(function(res, toCheck) {
    //Log('findMatch', res, match.FullyQualifiedName, needle, match.Id, match.FullyQualifiedName == needle && match.Id)
    return res || ((toCheck.FullyQualifiedName || toCheck.DisplayName) == needle && toCheck.Id)
  }, false)

  if (match) return match

  debugEmail('Could not findMatch()', needle, haystack)
}

//Double quotes didn't seem to work
function addQuotes(arr) {
  return JSON.stringify(arr).slice(1, -1).replace(/"/g, "'")
}

function queryQBO(query, options) {
  var url = "https://quickbooks.api.intuit.com/v3/company/" + COMPANY_ID + "/query?query=" + encodeURIComponent(query)

  options = options.headers ? options : {
    headers: {Authorization: 'Bearer ' + options.getAccessToken(), 'Accept': 'application/json'},
    method: 'get',
    muteHttpExceptions: true
  }

  var res = UrlFetchApp.fetch(url, options)

  //Log('queryQBO', url)

  try {
    return JSON.parse(res)
  } catch (e) {
    debugEmail('queryQBO JSON.parse ERROR', e, res)
  }
}

/**
 * Configures the service.
 */
function getService() {
  return OAuth2.createService('Dropbox')
      // Set the endpoint URLs.
      .setAuthorizationBaseUrl('https://appcenter.intuit.com/connect/oauth2?')
      .setTokenUrl('https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer')

      // Set the client ID and secret.
      .setClientId(CLIENT_ID)
      .setClientSecret(CLIENT_SECRET)

      // Set the name of the callback function that should be invoked to
      // complete the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())

      .setScope('com.intuit.quickbooks.accounting')
     // .setParam('state','security_token123456789101112314151617181920')

      // Set the response type to code (required).
      .setParam('response_type', 'code');
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function resetService() {
  var service = getService();
  service.reset();
}


/**
 * Handles the OAuth callback.
 */
function authCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied');
  }
}

/**
 * Logs the redict URI to register in the Quickbooks application settings.
 */
function logRedirectUri() {
  var service = getService();
  Log(service.getRedirectUri());
}

/*
//Add a note to an item (for a recipt)
function addNote(id, note){
  var service = getService();

  if ( ! service.hasAccess())
    return Log('Open the following URL and re-run the script:', service.getAuthorizationUrl())

  //adding a note to an invocie
  var mime = "application/json"
  var url  = "https://quickbooks.api.intuit.com/v3/company/" + COMPANY_ID + "/attachable?operation=update&minorversion=4"

  var body = {
    "AttachableRef": [{"EntityRef": {"type": "Purchase","value": id}}],
    "Note": note
  }

  var res = queryQBO(query, {
    method: 'post',
    headers: {Authorization: 'Bearer '+service.getAccessToken(), Accept:mime , "Content-Type":mime},
    payload: JSON.stringify(body),
    muteHttpExceptions:true
  })

  if( ~ response3.getContentText().indexOf("Fault"))
    return Log("Add_note failed")

  Log("Add_note succeeded")
  return true
}
*/
