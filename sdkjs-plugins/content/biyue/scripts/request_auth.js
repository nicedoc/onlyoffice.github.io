// eslint-disable-next-line no-unused-vars
import { getToken } from './auth.js'

let baseURL = 'https://edusys.xmdas-link.com'

// 创建axios实例
const service = axios.create({
  baseURL: baseURL, // api的base_url
  timeout: 300000, // 请求超时时间
  headers: {
    'Content-Type': 'application/x-www-form-urlencoded'
  }
})

// Token失效，需要重新获取token的StatusCode状态码
const tokenFailStatusCodes = [
  21 // token无效
]

// request拦截器
service.interceptors.request.use(
  config => {
    config.headers['X-Token'] = getToken()
    if (config.method === 'post' && !config.isFormData && !config.isFileData && !config.isJsonData) {
      let list = []
      Object.keys(config.data).forEach(key => {
        if (config.data[key] != undefined) {
          list.push(`${key}=${config.data[key]}`)
        }
      })
      config.data = list.join('&')
    }
    if (config.isFileData) {
      config.headers['Content-Type'] = 'multipart/form-data'
      if (config.org_code) {
        config.data.append('org_code', config.org_code)
      }
    }
    return config
  },
  error => {
    console.log(error) // for debug
    return Promise.reject(error)
  }
)

// respone拦截器
service.interceptors.response.use(
  response => {
    console.log(response)
    const code = response.data.code
    if (code === 1) {
      // 成功处理
      return response.data
    } else if (tokenFailStatusCodes.indexOf(code) !== -1) {
      // 表示token为失效状态，需要重新登录
      console.log('表示token为失效状态，需要重新登录 fedLogOut')
      alert(response.data.message)
      // 返回null表示不需要catch之后处理
      return Promise.reject(null)
    } else if (response.config.transmitError !== true) {
      // 返回null表示不需要catch之后处理
      return Promise.reject(response.data)
    } else {
      // 失败处理
      return Promise.reject(response.data)
    }
  },
  error => {
    return Promise.reject(error)
  }
)

export default service
