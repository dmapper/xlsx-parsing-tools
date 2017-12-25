const {EventEmitter} = require('events')

module.exports = class WorkbookRelationsParser extends EventEmitter{
  constructor (createParserFn) {
    super()
    this.createParserFn = createParserFn
    this.sheets = []
  }

  parse (entry) {
    const startElement = (name, attrs) => {
      if (name === 'Relationship') {
        this.emit('rel', attrs['Id'], attrs['Type'], attrs['Target'])
      }
    }

    const parser = this.createParserFn(startElement)

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