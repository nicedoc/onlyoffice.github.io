
// base api
var VUE_APP_GO_API = 'https://eduques.xmdas-link.com'
var VUE_APP_AUTH_API = 'https://edusys.xmdas-link.com'
var VUE_APP_VER_PREFIX = '' // 版本前缀
var VUE_APP_DEBUG = true // 是否可配置调试，正式服要配置为false

function setApiConfig(apiConfig) {
	if (!apiConfig) return
	if (apiConfig.go_api) {
		VUE_APP_GO_API = apiConfig.go_api
	}
	if (apiConfig.auth_api) {
		VUE_APP_AUTH_API = apiConfig.auth_api
	}
	if (apiConfig.label) {
		VUE_APP_VER_PREFIX = apiConfig.label
	}
}

export {
	VUE_APP_GO_API,
	VUE_APP_AUTH_API,
	VUE_APP_VER_PREFIX,
    VUE_APP_DEBUG,
    setApiConfig
}