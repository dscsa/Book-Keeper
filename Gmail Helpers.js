//Trouble with Email Forwarding using the last message in a thread because to address may have changed.  For example, George autoforwards Amazon boxes
//but the boxes have not appeared on bank or QBO.  So book keeper then emails back saying its trying to find a match, which becomes the last email in the thread
//so if you do paseTo on last email it looks like it is to George or Amazon (depending on how the reply-to works).
//Instead lets look through all messages in this thread to see if any match an extended receipts email.
//TODO: Test whether this can be simplified by just looking at "to" of first message rather than first match from looping in reverse


function getDate(message) {

  var to = getTo(message)
  var subject = message.getSubject ? message.getSubject() : message

  return getYYYYMMDD(to)
    || getMMDDYYYY(to)
    || getYYYYMMDD(subject)
    || getMMDDYYYY(subject)
    || getForwardedDate(message.getBody ? message.getBody() : message)
    || getSentDate(message)
}

function getYYYYMMDD(string) {
  var yyyymmdd = string.match(/(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})\b/)
  //if (yyyymmdd) debugEmail('getYYYYMMDD',string, yyyymmdd)
  if (yyyymmdd) return yyyymmdd[1]+'-'+('0'+yyyymmdd[2]).slice(-2)+'-'+('0'+yyyymmdd[3]).slice(-2)
}

function getMMDDYYYY(string) {
  var mmddyyyy = string.match(/(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})\b/)
  //if (mmddyyyy) debugEmail('getMMDDYYYY',string, mmddyyyy)
  if (mmddyyyy) return mmddyyyy[3]+'-'+('0'+mmddyyyy[1]).slice(-2)+'-'+('0'+mmddyyyy[2]).slice(-2)
}

function getForwardedDate(string) {
  var date = string.match(/Date:(.*?)($|<br)/mi)



  if (date) {
    //Mon, Apr 2, 2018 at 1:06 PM
    //April 11, 2018 at 9:42:37 PM EDT
    //Tue, Apr 3, 2018 at 2:31 PM
    //Fri, Apr 6, 2018 at 3:06 PM
    //Thu, Mar 15, 2018 at 1:20 PM
    //Wed, Apr 4, 2018 at 12:41 AM
    //April 23, 2018 at 12:57:27 AM EDT
    var year  = date[1].match(/\b\d{4}\b/)
    var day   = date[1].match(/\b\d{1,2}\b/)
    var month = getMonth(date[1])
    //string = string.replace(/</g, '[').replace(/>/g, ']') //sanitize
    //debugEmail('getForwardedDate', year+'-'+month+'-'+('0'+day).slice(-2), year, month, day, date.index, string.slice(date.index-200, date.index+200), '@@@@@@@@', string)

    return year+'-'+month+'-'+('0'+day).slice(-2)
  }
}

function getSentDate(message) {
  return message.getDate ? message.getDate().toJSON().slice(0, 10) : new Date()
}

function getTo(message) {
  return getForward(message)[1] || (message.getTo ? message.getTo() : message)
}

function getFrom(message) {
  return getForward(message)[0] || (message.getFrom ? message.getFrom() : message)
}

function getName(message) {
  var from = getFrom(message)
  var name = from.replace(/<.*?>/, '').split(/ |@|_|\./) //Remove any hyperlinks and try to get the first name only

  if (name[0].length < 2 || ! name[1] || name[1].length < 2)
    return {
      first:from.replace(/"/g, '').trim(),
      last:''
    }

  return {
    first:name[0][0].toUpperCase()+name[0].slice(1).toLowerCase(),
    last:name[1][0] == '@' ? '' : name[1][0].toUpperCase()+name[1].slice(1).toLowerCase()
  }
}

//X-Forwarded-For is used by gmail forwarding, not sure if its a universal standard though
//Example https://mail.google.com/mail/u/0?ik=dc59420667&view=om&permmsgid=msg-f%3A1614438721172240768
function getForward(message) {
  var match = message.getRawContent ? message.getRawContent().match(/^X-Forwarded-For:\s*([^\s]+)\s*([^\s]+)/mi) : message
  return match ? match.slice(1, 3) : []
}

function findDate(parsed, message) {
   parsed.date = getDate(message)
   parsed.subject = parsed.subject.replace(parsed.date, '')
}

function reply(message, body, attachments, subject) {
  var name = getName(message)
  return message.forward(getFrom(message), {
    name:'Book Keeper',
    bcc:'adam.kircher@gmail.com',
    attachments:attachments, //optional
    subject:subject,         //optional
    htmlBody:'Hello '+name.first+' '+name.last+',<br><br>'+body+'<br>Thanks,<br><a href="https://docs.google.com/spreadsheets/d/1klEQQ7u73D8y1UdPLu2C3xChQ1ZlLfEpGfhACe9WNXQ/edit">Mr. Keeper</a><br><br>Saving Medicine : Saving Lives<br><br>-----<br><br>'+message.getBody()+'<br>'
  })
}

//Smarter way to do this? https://stackoverflow.com/questions/13566552/easiest-way-to-convert-month-name-to-month-number-in-js-jan-01
function getMonth(date) {

    if ( ~ date.indexOf('Jan'))
      return '01'

    if ( ~ date.indexOf('Feb'))
      return '02'

    if ( ~ date.indexOf('Mar'))
      return '03'

    if ( ~ date.indexOf('Apr'))
      return '04'

    if ( ~ date.indexOf('May'))
      return '05'

    if ( ~ date.indexOf('Jun'))
      return '06'

    if ( ~ date.indexOf('Jul'))
      return '07'

    if ( ~ date.indexOf('Aug'))
      return '08'

    if ( ~ date.indexOf('Sep'))
      return '09'

    if ( ~ date.indexOf('Oct'))
      return '10'

    if ( ~ date.indexOf('Nov'))
      return '11'

    if ( ~ date.indexOf('Dec'))
      return '12'
}
