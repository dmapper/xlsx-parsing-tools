const {splitFormats} = require('./utils')
const {EventEmitter} = require('events')

const fmts = {
  0: null, //General
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',

  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',

  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',

  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@'
}

module.exports = class StylesParser extends EventEmitter {
  constructor (createParserFn) {
    super()
    this.createParserFn = createParserFn
  }

  parse (entry) {
    const numFmts = {}
    const cellXfs = []
    let cellXfsCollect = false

    const startElement = (name, attrs) => {
      if (name === 'numFmt') {
        numFmts[attrs.numFmtId] = attrs.formatCode
      } else if (name === 'cellXfs') {
        cellXfsCollect = true
      } else if ((cellXfsCollect) && (name === 'xf')) {
        const fmtnr = parseInt(attrs.numFmtId)
        cellXfs.push(fmtnr)
        const fmt = numFmts[fmtnr] || fmts[fmtnr]

        const style = {
          fmt,
          fmtnr,
          fmts: (fmt ? splitFormats(fmt) : []),
          def: attrs
        }

        this.emit('style', style, () => {
          entry.unpipe(parser)
          parser.pause()
          parser.end()
          entry.autodrain && entry.autodrain()
        })
      }
    }

    const endElement = (name, attrs) => {
      if (name === 'cellXfs') {
        cellXfsCollect = false
      }
    }

    const parser = this.createParserFn(startElement, endElement)

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