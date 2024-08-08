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
	showLevelSetDialog
} from './QuesManager.js'
import { showCom, updateText, addClickEvent } from './model/util.js'
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
		window.biyue.showMessageBox({
			content: '确定要全量更新吗？',
			extra_data: {
				func: 'reqUploadTree'
			}
		})
	})
	addClickEvent('#getQuesType', reqGetQuestionType)
	addClickEvent('#viewQuesType', onViewQuesType)
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

export {
	initView,
	handlePaperInfoResult
}