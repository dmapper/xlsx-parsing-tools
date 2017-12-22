const	expat = require('node-expat')
const {EventEmitter} = require('events')

module.exports = class WorkbookRelationsParser extends EventEmitter{
  constructor (options) {
    super()
    this.sheets = []
  }

  parse (entry, onClose, onError) {
    const parser = expat.createParser()
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