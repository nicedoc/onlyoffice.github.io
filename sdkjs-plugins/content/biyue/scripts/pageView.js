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
import { reqSaveInfo, onLatexToImg} from './api/paper.js'
import { biyueCallCommand, resetStack } from './command.js'
import ComponentSelect from '../components/Select.js'
var select_ask_shortcut = null
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
	//addClickEvent('#mathpix', onMathpix)
	addClickEvent('#paste', onPaste)
	var vShortcut = '0'
	if (window.BiyueCustomData && window.BiyueCustomData.ask_shortcut) {
		vShortcut = window.BiyueCustomData.ask_shortcut
	}
	select_ask_shortcut = new ComponentSelect({
		id: 'shortcutKeyDiv',
		options: [
			{ value: '0', label: '未定义' },
			{ value: 'ctrl', label: '双击 + ctrl' },
			{ value: 'alt', label: '双击 + alt' },
			{ value: 'shift', label: '双击 + shift' }
		],
		value_select: vShortcut,
		callback_item: (data) => {
			changeAskShortcut(data)
		},
		width: '60%',
	})
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

function changeAskShortcut(data) {
	window.BiyueCustomData.ask_shortcut = data.value
}

async function getClipboardContents() {
	try {
		const clipboardItems = await navigator.clipboard.read()
		for (const clipboardItem of clipboardItems) {
			for (const type of clipboardItem.types) {
				const blob = await clipboardItem.getType(type)
				console.log('blob', blob)
			}
		}
	} catch(err) {
		console.log(err)
	}
}

function onPaste() {
	console.log('黏贴')
	setTimeout(async() => {
        try {
            const text = await navigator.clipboard.readText()
            console.log('text', text)
        } catch (error) {
            console.log('Failed to read clipboard contents:', error)
        }
    }, 2000)
}

function insertImage(src, width, height) {
	console.log('insertImage', src, width, height)
	Asc.scope.insert_data = {
		type: 'image',
		scr: src,
		width: width,
		height: height
	}
	window.Asc.plugin.executeMethod("GetSelectionRange", [], function(range) {
        if (range) {
            // 插入图片
            window.Asc.plugin.executeMethod("InsertImage", [base64Image, width, height]);
        } else {
            console.error("Failed to get the selection range or cursor position.");
        }
    });
	return biyueCallCommand(window, function() {
		var insertData = Asc.scope.insert_data
		var image = Api.CreateImage(insertData.src, insertData.width * 36e3, insertData.height * 36e3)


	})
}

function onMathpix() {
	var text = "$\frac{\sqrt{145}}{5}$"
	onLatexToImg(text).then(res => {
		console.log('onLatexToImg success', res)
		var content = res.data.new_content
		const srcMatch = content.match(/src="([^"]*)"/);
		// const widthMatch = content.match(/width:(\d+px)/);
		// const heightMatch = content.match(/height:(\d+px)/);
		const widthMatch = content.match(/width:(\d+)px/);
		const heightMatch = content.match(/height:(\d+)px/);

		const src = srcMatch ? srcMatch[1] : null;
		const width = widthMatch ? widthMatch[1] : null;
		const height = heightMatch ? heightMatch[1] : null;
		if (src) {
			insertImage(src, width, height)
		}	

	}).catch(res => {
		console.log('onLatexToImg fail', res)
	})
	// var contentInput = $('#paste-target')
	// if (!contentInput) {
	// 	return
	// }
	// contentInput.select()
	// contentInput.focus()
	// setTimeout(async() => {
	// 	try {
	// 		const text = await navigator.clipboard.readText()
	// 		console.log('text', text)
	// 	} catch (error) {
	// 		console.log('Failed to read clipboard contents:', error)
	// 	}
	// }, 2000)
}

export {
	initView,
	handlePaperInfoResult,
	clearRepeatControl,
	onSaveData
}