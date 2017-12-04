const	expat = require('node-expat')
const {EventEmitter} = require('events')

module.exports = class WorkbookParser extends EventEmitter{
  constructor (options) {
    super()
    this.sheets = []
  }

  parse (entry, onClose, onError) {
    const parser = expat.createParser()
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