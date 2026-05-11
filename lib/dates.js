function startOfWeek(offset = 0) {
  const d = new Date()
  d.setHours(0, 0, 0, 0)
  const day = d.getDay()
  const diff = d.getDate() - day + (day === 0 ? -6 : 1)
  d.setDate(diff - (offset * 7))
  return d
}

function startOfMonth(offset = 0) {
  const d = new Date()
  d.setHours(0, 0, 0, 0)
  d.setDate(1)
  d.setMonth(d.getMonth() - offset)
  return d
}

function startOfQuarter(offset = 0) {
  const d = new Date()
  d.setHours(0, 0, 0, 0)
  d.setDate(1)
  const quarterMonth = Math.floor(d.getMonth() / 3) * 3
  d.setMonth(quarterMonth - (offset * 3))
  return d
}

function startOfHalfyear(offset = 0) {
  const d = new Date()
  d.setHours(0, 0, 0, 0)
  d.setDate(1)
  const halfyearMonth = Math.floor(d.getMonth() / 6) * 6
  d.setMonth(halfyearMonth - (offset * 6))
  return d
}

function startOfYear(offset = 0) {
  const d = new Date()
  d.setHours(0, 0, 0, 0)
  d.setDate(1)
  d.setMonth(0)
  d.setFullYear(d.getFullYear() - offset)
  return d
}
