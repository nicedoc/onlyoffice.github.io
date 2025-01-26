import ComponentSelect from '../../components/Select.js'
import { addClickEvent, showCom } from '../model/util.js'
import { VUE_APP_DEBUG, VUE_APP_VER_PREFIX, setApiConfig } from '../../apiConfig.js'
import { setXToken } from '../auth.js'
import { setBaseURL } from '../request.js'
import { setAuthBaseURL } from '../request_auth.js'
var ev_list = [{
	value: 'master',
	label: 'master',
	go_api: 'https://eduques.xmdas-link.com',
	auth_api: 'https://edusys.xmdas-link.com'
}, {
	value: 'qa',
	label: 'qa',
	go_api: 'https://qaques.biyue.tech',
	auth_api: 'https://qasys.biyue.tech'
}, {
	value: 'prod',
	label: 'prod',
	go_api: 'https://ques.biyue.tech',
	auth_api: 'https://sys.biyue.tech'
}, {
	value: 'bonlyoffice',
	label: 'bonlyoffice',
	go_api: 'https://eduques.xmdas-link.com/bonlyoffice',
	auth_api: 'https://edusys.xmdas-link.com/bonlyoffice'
}]
var select_ev = null
function initSetEv() {
	showCom('#setConfigEv', VUE_APP_DEBUG && VUE_APP_VER_PREFIX != 'prod')
	showCom('#ev-set-box', false)
	addClickEvent('#setConfigEv', function() {
		showCom('#ev-set-box', true)
		render()
	})
}

function render() {
	var com = $('#ev-set-box')
	if (!com) {
		return
	}
	var html = `<div id="ev-set-box">
				<div class="row-align-center" style="margin-bottom: 8px">
					<div class="keepall" style="margin-right: 12px;">环境</div>
					<div id="evSelect"></div>
				</div>
				<input placeholder="token" id="token" autocomplete="off" />
				<input placeholder="paper_id" id="paperId" style="margin: 8px 0;" autocomplete="off" />
				<div class="row-arround">
					<div id="confirmSetEv" class="btn-text-default" style="padding: 2px 16px;">确定 </div>
					<div id="cancelSetEv" class="btn-text-default"  style="padding: 2px 16px;">取消</div>
				</div>
			</div>`
	com.html(html)
	select_ev = new ComponentSelect({
		id: 'evSelect',
		options: ev_list,
		value_select: '',
		width: '70%',
		enabled: true
	})
	addClickEvent('#confirmSetEv', onConfirmSet)
	addClickEvent('#cancelSetEv', onCancelSet)
}

function onConfirmSet() {
	var token = $('#token').val()
	var paperId = $('#paperId').val()
	var ev = select_ev.getValue()
	if (!token || !ev || !paperId) {
		alert('token, paperId和环境不能为空')
		return
	}
	var item = ev_list.find(e => {
		return e.value == ev
	})
	setApiConfig(item)
	setXToken(token)
	setBaseURL(item.go_api)
	setAuthBaseURL(item.auth_api)
	window.BiyueCustomData.xtoken = token
	window.BiyueCustomData.paper_uuid = paperId
	window.biyue.handleInit()
	showCom('#ev-set-box', false)
}

function onCancelSet() {
	showCom('#ev-set-box', false)
}

export {
	initSetEv
}