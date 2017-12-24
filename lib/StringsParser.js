const {EventEmitter} = require('events')

module.exports = class StringsParser extends EventEmitter {
  constructor (createParserFn) {
    super()
    this.createParserFn = createParserFn
  }

  parse (entry) {
    const parser = this.createParserFn()
    let stringsCollect = false
    let sl = []
    let s = ''

    const startElement = (name, attrs) => {
      if (name === 'si') {
        sl = []
      }
      if (name === 't') {
        stringsCollect = true
        s = ''
      }
    }

    const endElement = (name, attrs) => {
      if (name === 't') {
        sl.push(s)
        stringsCollect = false
      }
      if (name === 'si') {
        const string = sl.join('')
        this.emit('string', string, () => {
          entry.unpipe(parser)
          parser.pause()
          parser.end()
          entry.autodrain && entry.autodrain()
        })
      }
    }

    parser.on('startElement', startElement)
    parser.on('opentag', startElement)
    parser.on('endElement', endElement)
    parser.on('closetag', endElement)

    parser.on('text', txt => {
      if (stringsCollect) {
        s = s + txt
      }
    })

    parser.on('error', err => {
      this.emit('error', err)
    })
    parser.on('close', () => {
      this.emit('close')
    })
    parser.on('finish', () => {
      this.emit('finish')
    })

    entry.pipe(parser)
  }
}