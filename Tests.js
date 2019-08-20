function testPayment() {
  var payment = {
    "CustomerRef": {
      "value": "416",
      "name": "Unrestricted Internally"
    },
    "TotalAmt": 0,
    "UnappliedAmt": 0,
    "ProcessPayment": false,
    "domain": "QBO",
    "sparse": false,
    "Id": "10841",
    "SyncToken": "1",
    "MetaData": {
      "CreateTime": "2018-06-03T15:20:35-07:00",
      "LastUpdatedTime": "2018-06-03T15:24:09-07:00"
    },
    "TxnDate": "2018-06-03",
    "CurrencyRef": {
      "value": "USD",
      "name": "United States Dollar"
    },
    "Line": [
      {
        "Amount": 75,
        "LinkedTxn": [
          {
            "TxnId": "9280",
            "TxnType": "Invoice"
          }
        ],
        "LineEx": {
          "any": [
            {
              "name": "{http://schema.intuit.com/finance/v3}NameValue",
              "declaredType": "com.intuit.schema.finance.v3.NameValue",
              "scope": "javax.xml.bind.JAXBElement$GlobalScope",
              "value": {
                "Name": "txnId",
                "Value": "9280"
              },
              "nil": false,
              "globalScope": true,
              "typeSubstituted": false
            },
            {
              "name": "{http://schema.intuit.com/finance/v3}NameValue",
              "declaredType": "com.intuit.schema.finance.v3.NameValue",
              "scope": "javax.xml.bind.JAXBElement$GlobalScope",
              "value": {
                "Name": "txnOpenBalance",
                "Value": "75.00"
              },
              "nil": false,
              "globalScope": true,
              "typeSubstituted": false
            },
            {
              "name": "{http://schema.intuit.com/finance/v3}NameValue",
              "declaredType": "com.intuit.schema.finance.v3.NameValue",
              "scope": "javax.xml.bind.JAXBElement$GlobalScope",
              "value": {
                "Name": "txnReferenceNumber",
                "Value": "1459"
              },
              "nil": false,
              "globalScope": true,
              "typeSubstituted": false
            }
          ]
        }
      },
      {
        "Amount": 50,
        "LinkedTxn": [
          {
            "TxnId": "9276",
            "TxnType": "Invoice"
          }
        ],
        "LineEx": {
          "any": [
            {
              "name": "{http://schema.intuit.com/finance/v3}NameValue",
              "declaredType": "com.intuit.schema.finance.v3.NameValue",
              "scope": "javax.xml.bind.JAXBElement$GlobalScope",
              "value": {
                "Name": "txnId",
                "Value": "9276"
              },
              "nil": false,
              "globalScope": true,
              "typeSubstituted": false
            },
            {
              "name": "{http://schema.intuit.com/finance/v3}NameValue",
              "declaredType": "com.intuit.schema.finance.v3.NameValue",
              "scope": "javax.xml.bind.JAXBElement$GlobalScope",
              "value": {
                "Name": "txnOpenBalance",
                "Value": "50.00"
              },
              "nil": false,
              "globalScope": true,
              "typeSubstituted": false
            },
            {
              "name": "{http://schema.intuit.com/finance/v3}NameValue",
              "declaredType": "com.intuit.schema.finance.v3.NameValue",
              "scope": "javax.xml.bind.JAXBElement$GlobalScope",
              "value": {
                "Name": "txnReferenceNumber",
                "Value": "1456"
              },
              "nil": false,
              "globalScope": true,
              "typeSubstituted": false
            }
          ]
        }
      },
      {
        "Amount": 125,
        "LinkedTxn": [
          {
            "TxnId": "10270",
            "TxnType": "Deposit"
          }
        ],
        "LineEx": {
          "any": [
            {
              "name": "{http://schema.intuit.com/finance/v3}NameValue",
              "declaredType": "com.intuit.schema.finance.v3.NameValue",
              "scope": "javax.xml.bind.JAXBElement$GlobalScope",
              "value": {
                "Name": "txnId",
                "Value": "10270"
              },
              "nil": false,
              "globalScope": true,
              "typeSubstituted": false
            },
            {
              "name": "{http://schema.intuit.com/finance/v3}NameValue",
              "declaredType": "com.intuit.schema.finance.v3.NameValue",
              "scope": "javax.xml.bind.JAXBElement$GlobalScope",
              "value": {
                "Name": "txnOpenBalance",
                "Value": "125.00"
              },
              "nil": false,
              "globalScope": true,
              "typeSubstituted": false
            },
            {
              "name": "{http://schema.intuit.com/finance/v3}NameValue",
              "declaredType": "com.intuit.schema.finance.v3.NameValue",
              "scope": "javax.xml.bind.JAXBElement$GlobalScope",
              "value": {
                "Name": "txnReferenceNumber"
              },
              "nil": false,
              "globalScope": true,
              "typeSubstituted": false
            }
          ]
        }
      }
    ]
  }


  var service = getService();

  if ( ! service.hasAccess())
    return debugEmail('Open the following URL and re-run the script:', service.getAuthorizationUrl())

  var url = "https://quickbooks.api.intuit.com/v3/company/" + COMPANY_ID + "/payment"
  var mime = 'application/json'
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {Authorization: 'Bearer '+service.getAccessToken(), Accept:mime , "Content-Type":mime},
    payload: JSON.stringify(payment),
    muteHttpExceptions:true
  })

  debugEmail('testPayment', url, JSON.parse(res))


}
function getPayment() {

  var payment = 13049
  var service = getService();

  if ( ! service.hasAccess())
    return debugEmail('Open the following URL and re-run the script:', service.getAuthorizationUrl())

  var url = "https://quickbooks.api.intuit.com/v3/company/" + COMPANY_ID + "/payment/"+payment
  var oldTxn = UrlFetchApp.fetch(url,  {
    headers: {Authorization: 'Bearer ' + service.getAccessToken(), 'Accept': 'application/json'},
    method: 'get',
    muteHttpExceptions: true
  })

  debugEmail('getPayment', url, JSON.parse(oldTxn))

}

function testDeposit() {

  var service = getService();

  if ( ! service.hasAccess())
    return debugEmail('Open the following URL and re-run the script:', service.getAuthorizationUrl())

  var oldTxn = UrlFetchApp.fetch("https://quickbooks.api.intuit.com/v3/company/" + COMPANY_ID + "/deposit/12378",  {
    headers: {Authorization: 'Bearer ' + service.getAccessToken(), 'Accept': 'application/json'},
    method: 'get',
    muteHttpExceptions: true
  })

  oldTxn = JSON.parse(oldTxn)

  oldTxn.Deposit.PrivateNote += " Adam Changed $175 -> $0 for a test!!!"


  oldTxn.Deposit.Line.push({
    Amount:-oldTxn.Deposit.TotalAmt,
    LinkedTxn: [
      {
        TxnLineId:0,
        TxnId: "13049",
        TxnType: "payment"
      }
    ]
  })

  oldTxn.Deposit.TotalAmt = 0
  //oldTxn.Deposit.AccountRef = {value:"106", name:"1120 Receivable - Program Revenue"}


  var url  = "https://quickbooks.api.intuit.com/v3/company/" + COMPANY_ID + "/deposit?operation=update"
  var mime = 'application/json'
  var newTxn = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {Authorization: 'Bearer '+service.getAccessToken(), Accept:mime , "Content-Type":mime},
    payload: JSON.stringify(oldTxn.Deposit),
    muteHttpExceptions:true
  })

  debugEmail('testDeposit', oldTxn.Deposit, JSON.parse(newTxn))

}

function testing() {

  var search = GmailApp.search('after:2018/10/10 subject:Your Amazon.com order of "Medium Moving Boxes"')
  //var msg = GmailApp.getMessageById('WhctKJVBBbrgSckrGfNXxCsGxljlbnHDSphSvntCZWFPqwDqCXGrndXPcVPhdWdMPgBhZBq')
  Log('Threads', search)
  var thread = search[5]
  var msgs   = thread.getMessages()

  for (var i in msgs)
    Log(parseTo(msgs[i]),  msgs[i].getTo(), msgs[i].getRawContent())
}
