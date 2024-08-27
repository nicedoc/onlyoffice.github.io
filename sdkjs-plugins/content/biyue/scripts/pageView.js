// 用于管理主界面

import {
	initFeature,
} from './panelFeature.js'

import { showQuesData, initListener } from './panelQuestionDetail.js'
import {
	reqGetQuestionType,
	handleAllWrite,
	showAskCells,
	g_click_value,
	clearRepeatControl,
	tidyNodes
} from './QuesManager.js'
import { showCom, updateText, addClickEvent, getInfoForServerSave } from './model/util.js'
import { reqSaveInfo } from './api/paper.js'
import { resetStack } from './command.js'
function initView() {
	showCom('#initloading', true)
	showCom('#hint1', false)
	showCom('.tabs', false)
	showCom('#panelList', false)
	initListener()
	window.tab_select = 'tabList'
	$('.tabitem').on('click', changeTab)
	document.addEventListener('clickSingleQues', function (event) {
		console.log('clickSingleQues', event)
		if (window.tab_select != 'tabQues') {
			changeTabPanel('tabQues')
		}
	})
	addClickEvent('#retry', () => {
		window.biyue.handleInit()
	})
	if ($('#writeSelect')) {
		$('#writeSelect').on('change', function() {
			var selectedValue = $('#writeSelect').val()
			console.log('writeSelect', selectedValue)
			handleAllWrite(selectedValue).then(() => {
				if (selectedValue != 'del') {
					showAskCells(selectedValue)
				}
			})
		})
	}
	addClickEvent('#reSplitQuestionBtn', () => {
		window.biyue.showMessageBox({
			title: '提示',
			content: '确定要重新切题吗？',
			extra_data: {
				func: 'reSplitQustion'
			}
		})
	})
	addClickEvent('#uploadTree', () => {
		var qmap = window.BiyueCustomData.question_map
		if (!qmap || Object.keys(qmap).length == 0) {
			window.biyue.showMessageBox({
				content: '未找到可更新的题目，请检查题目列表',
				showCancel: false
			})
		} else {
			window.biyue.showMessageBox({
				content: '确定要全量更新吗？',
				extra_data: {
					func: 'reqUploadTree'
				}
			})
		}
	})
	addClickEvent('#getQuesType', reqGetQuestionType)
	addClickEvent('#viewQuesType', onViewQuesType)
	addClickEvent('#saveData', onSaveData)
	addClickEvent('#clearRepeatControl', clearRepeatControl)
	addClickEvent('#extroSwitch', showExtroButtons)
	addClickEvent('#printData', () => {
		console.log('BiyueCustomData', window.BiyueCustomData)
	})
	showCom('#extro_buttons', false)
	addClickEvent('#printStack', () => {
		if (window.commandStack) {
			console.log('commandStack len:', window.commandStack.length)
		}
	})
	addClickEvent('#clearStack', resetStack)
	addClickEvent('#tidyNodes', tidyNodes)
}

function handlePaperInfoResult(success, res) {
	showCom('#initloading', false)
	if (success) {
		showCom('.tabs', true)
		changeTabPanel('tabList')
	} else {
		showCom('#hint1', true)
		if (res && res.status) {
			updateText('#message', `获取信息失败, status:${res.status}`)
		} else {
			updateText('#message', '获取课时信息失败')
		}
	}
}

function changeTab(e) {
	var id
	if (e.target && e.target.id && e.target.id != '') {
		id = e.target.id
	} else if (e.currentTarget) {
		id = e.currentTarget.id
	}
	changeTabPanel(id)
}

function changeTabPanel(id) {
	var tabs = ['tabList', 'tabQues', 'tabFeature']
	tabs.forEach((tab) => {
		if (tab == window.tab_select && tab != id) {
			$('#' + tab).removeClass('selected')
		} else if (tab == id) {
			$('#' + tab).addClass('selected')
		}
	})
	window.tab_select = id
	var targetPanel = id.replace('tab', 'panel')
	var panels = ['panelList', 'panelQues', 'panelFeature']
	panels.forEach((panel) => {
		if (panel == targetPanel) {
			$('#' + panel).show()
		} else {
			$('#' + panel).hide()
		}
	})
	if (id == 'tabFeature') {
		initFeature()
	} else if (id == 'tabQues') {
		if (g_click_value) {
			showQuesData(Object.assign({}, g_click_value.Tag, {
				InternalId: g_click_value.InternalId,
				Appearance: g_click_value.Appearance,
			}))
		} else {
			showQuesData()
		}
	}
}

function onViewQuesType() {
	// console.log('展示题型') todo..
}

function onSaveData(print = true) {
	var str = getInfoForServerSave()
	return new Promise((resolve, reject) => {
		reqSaveInfo(window.BiyueCustomData.paper_uuid, str).then(res => {
			console.log('保存数据到后端成功', str)
			if (print) {
				window.biyue.showMessageBox({
					content: '保存成功',
					showCancel: false
				})
			}
			resolve()
		}).catch(res => {
			console.log('保存数据到后端失败', res)
			if (print) {
				window.biyue.showMessageBox({
					content: '保存失败',
					showCancel: false
				})
			}
			resolve()
		})
	}) 
}

function showExtroButtons() {
	var com = $('#extro_buttons')
	if (!com) {
		return
	}
	com.toggle()
}

export {
	initView,
	handlePaperInfoResult,
	clearRepeatControl,
	onSaveData
}