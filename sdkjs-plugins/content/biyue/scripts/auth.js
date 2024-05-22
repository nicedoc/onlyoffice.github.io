// 设置token
let token = ''
let uuid = ''

function setXToken(v) {
  console.log('setXToken', v)
  token = v
}

function getToken() {
  console.log('getToken', token)
  return token
}

function getPaperUuid() {
  return uuid
}

function setPaperUuid(v) {
  uuid = v
}

export { setXToken, getToken, getPaperUuid, setPaperUuid };