const {EventEmitter} = require('events')

module.exports = class WorkbookParser extends EventEmitter{
  constructor (createParserFn) {
    super()
    this.createParserFn = createParserFn
    this.sheets = []
  }

  parse (entry) {

    const startElement = (name, attrs) => {
      if (name === 'sheet') {
        this.emit('sheet', attrs.name, attrs.sheetId, attrs['r:id'])
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
    // TODO fixme
    parser.on('end', () => {
      this.emit('finish')
    })

    entry.pipe(parser)
  }
}