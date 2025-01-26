import { getToken } from './auth.js'
import { VUE_APP_GO_API } from '../apiConfig.js'
// Token失效，需要重新获取token的StatusCode状态码
const tokenFailStatusCodes = [
	21, // token无效
]

let baseURL = VUE_APP_GO_API

// 创建axios实例
const service = axios.create({
	baseURL: baseURL, // api的base_url
	timeout: 300000, // 请求超时时间
	headers: {
		'Content-Type': 'application/x-www-form-urlencoded',
	},
})

// request拦截器
service.interceptors.request.use(
	(config) => {
		// console.log(process.env.VUE_APP_TOKEN)
		config.headers['X-Token'] = getToken()
		if (
			config.method === 'post' &&
			!config.isFormData &&
			!config.isFileData &&
			!config.isJsonData
		) {
			let list = []
			Object.keys(config.data).forEach((key) => {
				if (config.data[key] != undefined) {
					list.push(`${key}=${config.data[key]}`)
				}
			})
			config.data = list.join('&')
		}
		if (config.isFileData) {
			config.headers['Content-Type'] = 'multipart/form-data'
			config.timeout = 1800000
			if (config.org_code) {
				config.data.append('org_code', config.org_code)
			}
		}
		return config
	},
	(error) => {
		console.log(error) // for debug
		return Promise.reject(error)
	}
)

// respone拦截器
service.interceptors.response.use(
	(response) => {
		const code = response.data.code
		const state = response.data.state || ''
		if (code === 1 || (state === 'SUCCESS' && response.config.isEditorImage)) {
			// 成功处理
			return response.data
		} else if (tokenFailStatusCodes.indexOf(code) !== -1) {
			// 表示token为失效状态，需要重新登录 todo..
			alert(response.data.message)
			return Promise.reject(null)
		} else if (
			response.config.transmitError !== true &&
			!response.config.isEditorImage
		) {
			// 返回null表示不需要catch之后处理
			return Promise.reject(response.data)
		} else {
			// 失败处理
			let errData = response.data
			if (response.config.isEditorImage) {
				errData = state
			}
			return Promise.reject(errData)
		}
	},
	(error) => {
		if (error && error.response) {
			return Promise.reject(error.response)
		} else {
			return Promise.reject(error)
		}
	}
)
function setBaseURL(url) {
    baseURL = url;
    service.defaults.baseURL = baseURL;
}

export { setBaseURL };
export default service
