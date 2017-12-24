const {EventEmitter} = require('events')

module.exports = class StringsParser extends EventEmitter {
  constructor (createParserFn) {
    super()
    this.createParserFn = createParserFn
  }

  parse (entry, onClose, onError) {
    const strings = this.strings
    const parser = this.createParserFn()
    let stringsCollect = false
    let sl = []
    let s = ''

    parser.on('startElement', name => {
      if (name === 'si') {
        sl = []
      }
      if (name === 't') {
        stringsCollect = true
        s = ''
      }
    })

    parser.on('endElement', name => {
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
          entry.autodrain()
        })
      }
    })

    parser.on('text', txt => {
      if (stringsCollect) {
        s = s + txt
      }
    })

    parser.on('error', err => {
      onError && onError(err)
      this.emit('error', err)
    })
    parser.on('close', () => {
      onClose && onClose()
      this.emit('finish')
    })
    entry.pipe(parser)
  }
}