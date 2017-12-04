/*
 converts a raw xlsx-date to js date
*/
  
function XLSXDate(value, date1904) {
    let date = Math.floor(value);
    let time = Math.round(86400 * (value - date));
    let d;
    if (date1904) {
		date += 1462
	}
    // Open XML stores dates as the number of days from 1 Jan 1900. Well, skipping the incorrect 29 Feb 1900 as a valid day.
    if (date === 60) {
		d = new Date(1900, 1, 29)
	} else {
		if (date > 60) {
			--date
		}
		/* 1 = Jan 1 1900 */
		d = new Date(1900, 0, 1, 0, 0, 0)
		d.setDate(d.getDate() + date - 1)
	}
    d.setSeconds(time % 60)
    time = Math.floor(time / 60)
    d.setMinutes(time % 60)
    time = Math.floor(time / 60)
    d.setHours(time)
    return d
}


function splitFormats(s) {
		/*
		 http://office.microsoft.com/en-gb/excel-help/create-or-delete-a-custom-number-format-HP005199500.aspx?redir=0
		 _-* #,##0\ _€_-;\-* #,##0\ _€_-;_-* "-"??\ _€_-;_-@_-
		 positiv value ; negativ value ; zero; string
		 */
	const fmts = s.split(/(?!\\);/);
	let nr = 0;
	let lastff = {t: 'x'};
	const result = [];
	for (let i = 0; i < fmts.length; i++) {
		let ff = parseFmtType(fmts[i]);
		ff = (ff.t === 'l' ? lastff : ff)
		lastff = ff
		const format = {fmt: fmts[i], fmt_type: ff.t};
		if (ff.f) {
			format.digits = ff.f
		}
		result.push(format)
		nr++
	}
	return result
}

function containsOnlyChars(value, chars) {
	for (let i = 0; i < value.length; i++) {
		if (!chars.includes(value[i])) {
			return false
		}
	}
	return (value.length > 0)
}

function parseFmtType(fmt) {
	// messy hack for extracting some infos from the number format (type and float-digits}
	let s = fmt;
	let b = '';
	while (s.length > 0) {
		const c = s[0];
		s = s.slice(1)
		if ((c === '_') || (c === '\\') || (c === '*')) {
			s = s.slice(1)
		} else if (c === '[') {
			s = s.slice(s.indexOf(']') + 1)
		} else if (c === '"') {
			s = s.slice(s.indexOf('"') + 1)
		} else if ((c === '(') || (c === ')')) {
			// nop
		} else {
			b += c
		}
	}
	b = b.replace(/#/g, '0').replace(/%/g, '')
	// deal with thousands separator 12000 -> 12 -> formatCode	'#,'
	let sp = b.split(',');
	b = sp[sp.length - 1]
	if (!isNaN(b)) {
		if (b.includes('.')) {
			let di = sp[sp.length - 1].split('.')[1].trim().length;
			if (b.includes('E+')) {
				di += 14
			}
			return {t: 'f', f: di}
		} else {
			return {t: 'i'}
		}
	} else if (b === '@') {
		return {t: 's'}
	}
	// '-'?? zero value
	if (b === '??') {
		return {t: 'l'} // last fmt should by used
	}
	sp = b.split(' ')
	// test # ??/??
	if ((sp.length > 1) && (containsOnlyChars(sp[sp.length - 1], '?/'))) {
		// '# ?/?', '# ??/??',
		const digits = sp[sp.length - 1].split('/')[0].trim().length + 1;
		return {t: 'f', f: digits}
	}
	// date format?
	if (containsOnlyChars(b, 'tmdyhseAPTMH:/-.0 ')) {
		return {t: 'd'}
	}
	return {t: 'x'}
}

function getEffectiveNumFormat (val, fmt) {
	if ((!fmt) || (fmt.fmts.length === 0)) {
		return null
	}
	if (fmt.fmts.length === 1) {
		return fmt.fmts[0]
	}
	if (isNaN(val)) {
		return fmt.fmts[3]
	}
	if (val < 0) {
		return fmt.fmts[1]
	}
	if (val > 0) {
		return fmt.fmts[0]
	}
	return fmt.fmts[(fmt.fmts.length > 2) ? 2 : 0]
}

function applyNumFormat (val, fmt, options) {
	const usefmt = getEffectiveNumFormat(val, fmt)
	
	if (usefmt) {
		switch (usefmt.fmt_type) {
			case 'd':
				if (options.convertValues.dates) {
					val = XLSXDate(val, options.date1904)
				}
				break
			case 'i':
				if (options.convertValues.numbers) {
					let i = null
					if (fmt && fmt.fmt === '0\\ %') {
						i = Math.round(parseFloat(val * 100))
					} else {
						i = parseInt(val, 10)
					}
					if (!isNaN(i)) {
						val = i
					}
				}
				break
			case 'f':
				if ((usefmt.digits > 0) && options.convertValues.floats) {
					if (options.roundFloats) {
						val = val.toFixed(usefmt.digits)
					}

					const v = parseFloat(val)
					if (!isNaN(v)) {
						val = v
					}
				}
				break
			default:
									// nop
				break
		}
	}
	return val
}

function getCellValue (cell, strings = [], styles = [], options = {}) {
	options = options || {
		date1904: false,
		rawValues: false,
		roundFloats: true,
		convertValues: {
			numbers: true,
			floats: true,
			dates: true,
			booleans: true
		}
	}

	cell.fmt = cell.styleIndex ? styles[cell.styleIndex] : null
	let val = cell.val

	if (cell.col >= 0) {
		if (cell.type === 's') {
			val = strings[parseInt(val, 10)]
		}
	}

	if (options.rawValues) return val

	if (val !== null) {
		switch (cell.type) {
			case 'n':
				const v = parseFloat(val)
				if (!isNaN(v)) {
					val = v
				}
				if ((cell.fmt) && (options.convertValues)) {
					val = applyNumFormat(val, cell.fmt, options)
				}
				break
			case 's':
			case 'str':
			case 'inlineStr':
				break // string, do nothing
			case 'b':
				if (options.convertValues && options.convertValues.booleans) {
					if (['0', 'FALSE', 'false'].includes(val)) {
						val = false
					} else if (['1', 'TRUE', 'true'].includes(val)) {
						val = true
					} else {
						console.log('Unknown boolean:', val)
					}
				}
				break
			case 'e':
				console.log('Error cell type: Value will be invalid ("#REF!", "#NAME?", "#VALUE!" or similar).')
				break
			default:
				console.log('Unknown cell type: "%s"', cell.type)
		}
	}

	return val
}
  
module.exports = {
	splitFormats: splitFormats,
	getCellValue: getCellValue
}