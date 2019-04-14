
class RootOp {
  constructor (op) {
    this.op = op
  }
  getTop () {
    // return Number
    throw Error('SHOULD OVERRIDE getTop')
  }
  getBottom () {
    // return Number
    throw Error('SHOULD OVERRIDE getBottom')
  }
}

const numFormat = /^[+-]?\d+$/
// &+2b+3
// Full match 0-6 &+2b+3
// Group 1. 0-4 &+2b
// Group 2. 1-3 +2
// Group 3. 3-4 b
// Group 4. 4-6 +3
const relativeRefFormat = /^(&([+-]\d+)?([A-Za-z]))([+-]\d+)?$/

// #5b+1
// Full match 0-5 #5b+1
// Group 1. 0-3 #5b
// Group 2. 1-2 5
// Group 3. 2-3 b
// Group 4. 3-5 +1
const absoluteRefFormat = /^(#(\d+)([A-Za-z]))([+-]\d+)?$/

// H#0b+0
// Full match 0-6 H#0b+0
// Group 1. 0-1 H
// Group 2. 1-6 #0b+0
// Group 4. 1-4 #0b
// Group 5. 2-3 0
// Group 6. 3-4 b
// Group 7. 4-6 +0
const coRefFormat = /^([A-Z]+)((\d+)|(#(\d+)([A-Za-z]))([+-]\d+)|(&([+-]\d+)?([A-Za-z]))([+-]\d+))$/

function getRowFromCo (co) {
  return co.match(coRefFormat)[2]
}

function getColumnFromCo (co) {
  return co.match(coRefFormat)[1]
}

function rowPlus (a, b) {
  if ((a.includes('&') || a.includes('#')) && (b.includes('&') || b.includes('#'))) {
    throw Error('ONLY ALLOW 1 VARIABLE CONTAINS REF')
  }

  if (a.match(numFormat) != null && b.match(numFormat) != null) {
    return (parseInt(a) + parseInt(b)).toString()
  }

  [[a, b], [b, a]].forEach(it => {
    const matchRelative = it[0].match(relativeRefFormat)
    if (matchRelative != null) {
      if (it[1].match(numFormat) == null) {
        throw Error('FORMAT ERROR ' + it[1])
      }

      const sum = parseInt(matchRelative[4] || 0) + parseInt(it[1])
      return matchRelative[1] + (sum >= 0 ? '+' : '') + sum
    }

    const matchAbsolute = it[0].match(absoluteRefFormat)
    if (matchAbsolute != null) {
      if (it[1].match(numFormat) == null) {
        throw Error('FORMAT ERROR ' + it[1])
      }

      const sum = parseInt(matchAbsolute[4] || 0) + parseInt(it[1])
      return matchAbsolute[1] + (sum >= 0 ? '+' : '') + sum
    }
  })
  throw Error('FORMAT ERROR')
}

// 這邊不處理一個有 REF 一個沒 REF 的 case
function rowMinus (a, b) {
  if (a.match(numFormat) != null && b.match(numFormat) != null) {
    return (parseInt(a) - parseInt(b)).toString()
  }

  const matchRelativeA = a.match(relativeRefFormat)
  const matchRelativeB = b.match(relativeRefFormat)
  if (matchRelativeA != null && matchRelativeB != null) {
    if (matchRelativeA[1] !== matchRelativeB[1]) {
      throw Error('CROSS REF ERROR ' + matchRelativeA[1] + ' != ' + matchRelativeB[1])
    }
    return (parseInt(matchRelativeA[4] || 0) - parseInt(matchRelativeB[4] || 0)).toString()
  }

  const matchAbsoluteA = a.match(absoluteRefFormat)
  const matchAbsoluteB = b.match(absoluteRefFormat)
  if (matchAbsoluteA != null && matchAbsoluteB != null) {
    if (matchAbsoluteA[1] !== matchAbsoluteB[1]) {
      throw Error('CROSS REF ERROR ' + matchAbsoluteA[1] + ' != ' + matchAbsoluteB[1])
    }
    return (parseInt(matchAbsoluteA[4] || 0) - parseInt(matchAbsoluteB[4] || 0)).toString()
  }

  throw Error('FORMAT ERROR')
}

class Section extends RootOp {
}

class AddSheet extends Section {
  constructor (name, order) {
    super('ADD_SHEET')
    this.name = name
    this.order = order
  }
}

class DeleteSheet extends Section {
  constructor (name) {
    super('DELETE_SHEET')
    this.name = name
  }
}

class RenameSheet extends Section {
  constructor (oldName, newName) {
    super('RENAME_SHEET')
    this.oldName = oldName
    this.newName = newName
  }
}

class CopyRows extends Section {
  constructor (srcSheet, srcRowRange, dstSheet, dstRow, extra) {
    super('COPY_ROWS')
    // TODO valid input
    this.srcSheet = srcSheet
    this.srcRowRange = srcRowRange
    this.dstSheet = dstSheet
    this.dstRow = dstRow
    this.extra = extra
  }
  getTop () {
    return this.dstRow
  }
  // should be called after value set
  getBottom () {
    return rowPlus((parseInt(rowMinus(this.srcRowRange.split('~')[1], this.srcRowRange.split('~')[0]))).toString(), this.dstRow)
  }
}

class Fill extends Section {
  constructor (sheet, co, title, extra) {
    super('FILL')
    this.sheet = sheet
    this.co = co
    this.title = title
    this.extra = extra
    this.value = ''
  }
  getTop () {
    return getRowFromCo(this.co)
  }

  // we only know co, but don't know if the dst is a MergedRegion, so we return the getTop()
  getBottom () {
    return this.getTop()
  }
}

class PreProcess extends RootOp {
  process () {
    // return new Section()
    throw Error('SHOULD OVERRIDE process')
  }
}

// Will be replaced
class CopyList extends PreProcess {
  // @param resolveSingle : 0 -> head, 1 -> body, 2 -> foot
  // @param sections : Currently we only support Fill
  constructor (srcSheet, headRowRange, bodyRowRange, footRowRange, dstSheet, dstRow, sections, resolveSingle, title, extra) {
    super('COPY_LIST')
    this.srcSheet = srcSheet
    this.headRowRange = headRowRange
    this.bodyRowRange = bodyRowRange
    this.footRowRange = footRowRange
    this.dstSheet = dstSheet
    this.dstRow = dstRow
    this.sections = sections
    this.resolveSingle = resolveSingle
    this.title = title
    this.extra = extra
    this.value = []
  }

  process () {
    if (this.dstRow.match(numFormat) == null) {
      throw Error('this function should be called after dstRow got resolved')
    }

    const result = []
    let dstRow = this.dstRow

    if (this.value.length > 0) {
      const srcRowRange = (this.value.length === 1) ? [this.headRowRange, this.bodyRowRange, this.footRowRange][this.resolveSingle] : this.headRowRange

      result.push(new CopyRows(this.srcSheet, srcRowRange, this.dstSheet, dstRow, ''))
      this.sections.map(section => {
        // currently we only support Fill
        // section.co would be relative from the base of CopyList & currently we don't support REF inside CopyList
        const newFill = new Fill(
          this.dstSheet,
          getColumnFromCo(section.co) + (parseInt(getRowFromCo(section.co)) - 1 + parseInt(dstRow)).toString(),
          section.title,
          section.extra
        )
        newFill.value = this.value[0][section.title]
        return newFill
      })
        .forEach(it => result.push(it))

      dstRow = rowPlus(dstRow, rowPlus(rowMinus(srcRowRange.split('~')[1], srcRowRange.split('~')[0]), '1'))
    }

    this.value.slice(1, -1).forEach(it => {
      const srcRowRange = this.bodyRowRange

      result.push(new CopyRows(this.srcSheet, srcRowRange, this.dstSheet, dstRow, ''))
      this.sections.map(section => {
        // currently we only support Fill
        // it.co would be relative from the base of CopyList & currently we don't support REF inside CopyList
        const newFill = new Fill(
          this.dstSheet,
          getColumnFromCo(section.co) + (parseInt(getRowFromCo(section.co)) - 1 + parseInt(dstRow)).toString(),
          section.title,
          section.extra
        )
        newFill.value = it[section.title]
        return newFill
      })
        .forEach(result.push)

      dstRow = rowPlus(dstRow, rowPlus(rowMinus(srcRowRange.split('~')[1], srcRowRange.split('~')[0]), '1'))
    })

    if (this.value.length > 1) {
      result.push(new CopyRows(this.srcSheet, this.footRowRange, this.dstSheet, dstRow, ''))
      this.sections.map(section => {
        // currently we only support Fill
        // it.co would be relative from the base of CopyList & currently we don't support REF inside CopyList
        const newFill = new Fill(
          this.dstSheet,
          getColumnFromCo(section.co) + (parseInt(getRowFromCo(section.co)) - 1 + parseInt(dstRow)).toString(),
          section.title,
          section.extra
        )
        newFill.value = this.value[this.value.length - 1][section.title]
        return newFill
      })
        .forEach(it => result.push(it))
    }
    return result
  }

  // TODO handle when value isEmpty
  // Should be called after value set
  getTop () {
    return this.dstRow
  }

  // TODO handle when value isEmpty
  // Should be called after value set
  getBottom () {
    const b = this.process().reduce((bottomInt, section) => {
      if (section instanceof CopyRows) {
        return Math.max(parseInt(section.getBottom()), bottomInt)
      } else {
        return bottomInt
      }
    }, parseInt(this.dstRow)).toString()
    return b
  }
}

function resolveRow (rootOps, it, index) {
  const matchRelative = it.match(relativeRefFormat)
  const matchAbsolute = it.match(absoluteRefFormat)

  if (it.match(numFormat) != null) {
    return it
  } else if (matchRelative != null) {
    // &+2b+3
    // Full match 0-6 &+2b+3
    // Group 1. 0-4 &+2b
    // Group 2. 1-3 +2
    // Group 3. 3-4 b
    // Group 4. 4-6 +3
    const refObj = rootOps[index + parseInt(matchRelative[2])]
    const refRow = getDstRow(refObj, matchRelative[3])
    return rowPlus(refRow, (matchRelative[4] || '0'))
  } else if (matchAbsolute != null) {
    // #5b+1
    // Full match 0-5 #5b+1
    // Group 1. 0-3 #5b
    // Group 2. 1-2 5
    // Group 3. 2-3 b
    // Group 4. 3-5 +1

    // index is 1based
    const refObj = rootOps[parseInt(matchAbsolute[2]) - 1]
    const refRow = getDstRow(refObj, matchAbsolute[3])
    return rowPlus(refRow, (matchAbsolute[4] || '0'))
  }
  throw Error('BAD FORMAT')
}

function getDstRow (refObj, side) {
  switch (side) {
    case 'b':
      return refObj.getBottom()
    case 't':
      return refObj.getTop()
    default:
      throw Error('BAD FORMAT ' + side)
  }
}

/**
* @param {[RootOp]} rootOps - A list of RootOp
* @param {string} data - json
* @returns {[Section]} - you may use this response as the payload for requesting xlsxmanipulator with your template
*/
function compile (rootOps, data) {
  // 1. fill value
  Object.keys(data).forEach(it => {
    const target = rootOps.find(rootOp => {
      return rootOp.title === it
    })
    if (target != null) {
      target.value = data[it]
    }
  })

  console.log(rootOps)

  const sections = []

  // 2. resolve topology order 3. generate full sections
  // 這邊假設所有 REF 都只會往前指，我們就照順序來解 REF
  rootOps.forEach((it, index) => {
    if (it instanceof CopyRows) {
      it.dstRow = resolveRow(rootOps, it.dstRow, index)
      sections.push(it)
    } else if (it instanceof Fill) {
      it.co = getColumnFromCo(it.co) + resolveRow(rootOps, getRowFromCo(it.co), index)
      sections.push(it)
    } else if (it instanceof CopyList) {
      it.dstRow = resolveRow(rootOps, it.dstRow, index)
      it.process().forEach(it => sections.push(it))
    } else {
      sections.push(it)
    }
  })

  return sections
}

const rootOpFormat = /^[A-Za-z_]+$/

function deserializeRootOps (objList) {
  return objList.map(it => {
    // validate it.op
    const op = it.op.match(rootOpFormat)
    if (op == null) {
      throw Error(it.op + ' is not a valid RootOp')
    }
    const pascalClass = toPascalCase(op[0])
    if (pascalClass === CopyList.op) {
      it.value = deserializeRootOps(it.value)
    }
    // eslint-disable-next-line no-eval
    Object.setPrototypeOf(it, eval(pascalClass).prototype)
    return it
  })
}

function toPascalCase (str) {
  return [...str].reduce((array, c) => {
    if (array[0]) {
      array[0] = false
      array[1] += c
    } else if (c === '_') {
      array[0] = true
    } else {
      array[1] += c.toLowerCase()
    }
    return array
  }, [true, ''])[1]
}

export { AddSheet, DeleteSheet, RenameSheet, CopyRows, Fill, CopyList, compile, deserializeRootOps }
