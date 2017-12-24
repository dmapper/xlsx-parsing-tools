const {EventEmitter} = require('events')

module.exports = class WorkbookRelationsParser extends EventEmitter{
  constructor (createParserFn) {
    super()
    this.createParserFn = createParserFn
    this.sheets = []
  }

  parse (entry, onClose, onError) {
    const parser = this.createParserFn()
    parser.on('startElement', (name, attrs) => {
      if (name === 'Relationship') {
        this.emit('rel', attrs['Id'], attrs['Type'], attrs['Target'])
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