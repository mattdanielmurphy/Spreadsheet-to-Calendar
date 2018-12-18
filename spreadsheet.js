const XLSX = require('xlsx')

let inspect = require('util').inspect
let fs = require('fs')
let base64 = require('base64-stream')
let Imap = require('imap')

let imap = new Imap({
  user: 'matt@mattmurphy.ca',
  password: 'we%Hu6TFs3;F{n)GB6Hr9yAt8EQg2X)nw2Gz{CV4^%o&NbWh4vcWaYdiF^PxQ6WK',
  host: 'imap.zoho.com',
  port: 993,
  tls: true
  //,debug: (msg){console.log('imap:', msg);=>  }
})

const toUpper = x => (x && x.toUpperCase ? x.toUpperCase() : x)

function findAttachmentParts(struct, attachments) {
  attachments = attachments || []
  for (let i = 0, len = struct.length, r; i < len; ++i) {
    if (Array.isArray(struct[i])) {
      findAttachmentParts(struct[i], attachments)
    } else if (
      struct[i].disposition &&
      ['INLINE', 'ATTACHMENT'].indexOf(toUpper(struct[i].disposition.type)) > -1
    )
      attachments.push(struct[i])
  }
  return attachments
}

function saveFile(stream, filename, encoding, prefix) {
    if (/\.xlsx$/.test(filename)) {
        console.log('writing...')
        filename = `${attachmentsDir}/${filename}`
        let writeStream = fs.createWriteStream(filename)

        writeStream.on('finish', () => {
        console.log(prefix + 'Done writing to file %s', filename)
        })

        if (toUpper(encoding) === 'BASE64') {
        base64.decode(writeStream)
        stream.pipe(base64.decode()).pipe(writeStream)
        } else {
        stream.pipe(writeStream)
        }
    }
}

function buildAttMessageFunction(attachment) {
  let filename = attachment.params.name
  let encoding = attachment.encoding

  return (msg, seqno) => {
    let prefix = '(#' + seqno + ') '
    msg.on('body', (stream, info) => {
      // only save if spreadsheet
      console.log('going to save')
      saveFile(stream, filename, encoding, prefix)
    })
    msg.once('end', () => {
      console.log(prefix + 'Finished attachment %s', filename)
    })
  }
}

function getSanagansMessages() {
  imap.search([['FROM', 'info@sanagansmeatlocker.com'], ['SUBJECT', 'Schedule']], (err, results) => {
    if (err) throw err

    let fetch = imap.fetch(results, {
      bodies: ['HEADER.FIELDS (FROM TO SUBJECT DATE)'],
      struct: true
    })
    fetch.on('message', (msg, seqno) => {
      let prefix = '(#' + seqno + ') '
      msg.on('body', (stream, info) => {
        let buffer = ''
        stream.on('data', chunk => {
          buffer += chunk.toString('utf8')
        })
        stream.once('end', () => {
          let header = Imap.parseHeader(buffer)
          Object.keys(header).forEach(key => {
          })
        })
      })
      msg.once('attributes', attrs => {
        let attachments = findAttachmentParts(attrs.struct)
        for (let i = 0, len = attachments.length; i < len; ++i) {
          let attachment = attachments[i]
          let fetch = imap.fetch(attrs.uid, {
            //do not use imap.seq.fetch here
            bodies: [attachment.partID],
            struct: true
          })
          //build function to process attachment message
          fetch.on('message', buildAttMessageFunction(attachment))
        }
      })
      msg.once('end', () => {
      })
    })
    fetch.once('error', err => {
      console.log('Fetch error: ' + err)
    })
    fetch.once('end', () => {
      console.log('Done fetching all messages!')
      imap.end()
    })
  })
}

function openMailbox(mailboxName, cb) {
  imap.openBox(mailboxName, true, cb)
}

function getSpreadsheets() {
  imap.once('ready', () => {
    openMailbox('Inbox', (err, box) => {
      if (err) throw err
      getSanagansMessages()
    })
    openMailbox('Archive', (err, box) => {
      if (err) throw err
      getSanagansMessages()
    })
  })

  imap.once('error', err => {
    console.log(err)
  })

  imap.once('end', () => {
    console.log('Connection ended')
  })

  imap.connect()
}

let attachmentsDir = './attachments'

module.exports = getSpreadsheets()