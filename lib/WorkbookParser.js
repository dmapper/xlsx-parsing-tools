const {EventEmitter} = require('events')

module.exports = class WorkbookParser extends EventEmitter{
  constructor (createParserFn) {
    super()
    this.createParserFn = createParserFn
    this.sheets = []
  }

  parse (entry, onClose, onError) {
    const parser = this.createParserFn()
    parser.on('startElement', (name, attrs) => {
      if (name === 'sheet') {
        this.emit('sheet', attrs.name, attrs.sheetId, attrs['r:id'])
      }
    })
    parser.on('error', err => {
      onError && onError(err)
    })
    parser.on('close', () => {
      onClose && onClose()
      this.emit('finish')
    })
    entry.pipe(parser)
  }
}