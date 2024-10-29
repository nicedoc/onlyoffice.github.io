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
	tidyNodes,
	layoutDetect
} from './QuesManager.js'
import {
	imageAutoLink,
	onAllCheck
} from './linkHandler.js'
import { showCom, updateText, addClickEvent, getInfoForServerSave } from './model/util.js'
import { reqSaveInfo, onLatexToImg} from './api/paper.js'
import { biyueCallCommand, resetStack } from './command.js'
import ComponentSelect from '../components/Select.js'
var select_ask_shortcut = null
var timeout_paste_hint = null
var select_image_link = null
function initView() {
	showCom('#initloading', true)
	showCom('#hint1', false)
	showCom('.tabs', false)
	showCom('#panelList', false)
	initListener()
	window.tab_select = 'tabList'
	$('.tabitem').on('click', changeTab)
	document.addEventListener('clickSingleQues', function (event) {
		if (window.tab_select != 'tabQues') {
			changeTabPanel('tabQues', event)
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
	addClickEvent('#extroSwitch', (e) => {
		showExtroButtons()
	})
	addClickEvent('#tools', (e) => {
		showTools()
	})
	addClickEvent('#printData', () => {
		console.log('BiyueCustomData', window.BiyueCustomData)
	})
	showCom('#extro_buttons', false)
	showCom('#tool_buttons', false)
	addClickEvent('#printStack', () => {
		if (window.commandStack) {
			console.log('commandStack len:', window.commandStack.length)
		}
	})
	addClickEvent('#clearStack', resetStack)
	addClickEvent('#tidyNodes', tidyNodes)
	addClickEvent('#mathpix', onMathpix)
	addClickEvent('#paste', onPaste)
	addClickEvent('#pasteclear', onPasteInputClear)
	addClickEvent('#showShortcutKey', () => {
		showButton('func_key_container')
	})
	showCom('#func_key_container', false)
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
		pop_width: '100%'
	})
	addClickEvent('#showLinkPop', () => {
		showButton('link_container')
	})
	showCom('#link_container', false)
	select_image_link = new ComponentSelect({
		id: 'imageLinkSelect',
		options: [
			{ value: '0', label: '关闭就近关联' },
			{ value: '1', label: '开启就近关联' }
		],
		value_select: '0',
		callback_item: (data) => {
			changeImageLink(data)
		},
		width: '60%',
		pop_width: '100%'
	})
	enableBtnImageLink(false)
	showCom('#imageLinkTip')
	addClickEvent('#btnImageLink', onImageLink)
	addClickEvent('#imageLinkCheck', onImageLinkCheck)
	addClickEvent('#addWord', openSymbol)
	addClickEvent('#fixLayout', () => {
		layoutDetect(true)
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

function changeTabPanel(id, event) {
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
		if (event.detail.parentTag) {
			showQuesData(event.detail)
		} else if (g_click_value) {
			showQuesData(Object.assign({}, g_click_value.Tag, {
				InternalId: g_click_value.InternalId,
				Appearance: g_click_value.Appearance,
			}))
		} else {
			showQuesData()
		}
	}
	scroll(false)
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

function showExtroButtons(bshow) {
	showButton('extro_buttons', bshow)
	showButton('tool_buttons', false)
	scroll(true)
}

function showTools(bshow) {
	showButton('tool_buttons', bshow)
	showButton('extro_buttons', false)
	scroll(true)
}

function scroll(isBottom) {
	var com = $('#iframe_parent')
	if (!com) {
		return
	}
	com.animate({
		scrollTop: isBottom ? com[0].scrollHeight : 0
	}, 200)
}

function showButton(id, bshow) {
	var com = $(`#${id}`)
	if (!com) {
		return
	}
	if (bshow == undefined) {
		com.toggle()
	} else {
		bshow ? com.show() : com.hide()
	}
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
	setTimeout(() => {
		$('#paste-target').focus()
	}, 1000)
	setTimeout(async() => {
        try {
            const text = await navigator.clipboard.readText()
            console.log('text', text)
        } catch (error) {
            console.log('Failed to read clipboard contents:', error)
        }
    }, 2000)
}

function insertContent(str) {
	Asc.scope.insert_str = str
	return biyueCallCommand(window, function() {
		var content = Asc.scope.insert_str		
		const srcMatch = content.match(/src="([^"]*)"/);
		var src = srcMatch ? srcMatch[1] : null;
		var oDocument = Api.GetDocument()
		var pos = oDocument.Document.Get_CursorLogicPosition()
		if (src) {
			const widthMatch = content.match(/width:(\d+)px/);
			const heightMatch = content.match(/height:(\d+)px/);
			const width = widthMatch ? widthMatch[1] : null;
			const height = heightMatch ? heightMatch[1] : null;
			src = src.replace('img.xmdas-link.com', 'img2.xmdas-link.com')
			var scale = 0.25 * 36000
			var oDrawing = Api.CreateImage(src, width * scale, height * scale)
			if (pos && pos.length) {
				var lastElement = pos[pos.length - 1].Class
				if (lastElement.Add_ToContent) {
					lastElement.Add_ToContent(
						pos[pos.length - 1].Position,
						oDrawing.Drawing
					)
				}
			}
		} else { // 不是图片
			if (pos) {
				var lastElement = pos[pos.length - 1].Class
				if (lastElement.AddText) {
					lastElement.AddText(content, pos[pos.length - 1].Position)
				}
			}
		}
	}, false, true)
}

function updateHintById(id, message, color) {
	var tooltip = document.getElementById(id);
	if (!tooltip) {
		return
	}
	tooltip.textContent = message;
	tooltip.style.color = color || '#999';
	tooltip.style.display = 'block';
	clearTimeout(timeout_paste_hint)
	timeout_paste_hint = setTimeout(function() {
		tooltip.style.display = 'none';
	}, 1500);
}

function updatePasteHint(message, color) {
	updateHintById('pastehint', message, color)
}

function onMathpix() {
	var com = $('#paste-target')
	if (!com) {
		updatePasteHint('未找到输入框', '#ff0000')
		return
	}
	var text1 = $('#paste-target').val()
	if (!text1 || !text1.length) {
		updatePasteHint('输入框内容为空', '#ff0000')
		return
	}
	onLatexToImg(text1).then(res => {
		insertContent(res.data.new_content).then(() => {
			console.log('插入成功')
			updatePasteHint('插入成功', '#0080ff')
			onPasteInputClear()
		})
	}).catch(res => {
		if (res && res.message) {
			updatePasteHint(res.message, '#ff0000')
		} else if (!res) {
			updatePasteHint('网络不稳定', '#ff0000')
		}
		console.log('onLatexToImg fail', res)
	})
}

function onPasteInputClear() {
	var com = $('#paste-target')
	if (!com) {
		return
	}
	com.val('')
}

function openSymbol() {
	window.biyue.showDialog('addSymbolWindow', '插入符号', 'addSymbol.html', 600, 400, false)
}

function changeImageLink(data) {
	enableBtnImageLink(data.value * 1)
}

function enableBtnImageLink(v) {
	var btnlink = $('#btnImageLink')
	if (!btnlink) {
		return
	}
	if (v) {
		btnlink.removeClass('btn-unable')
	} else {
		btnlink.addClass('btn-unable')
	}
}
// 图片就近关联
function onImageLink() {
	imageAutoLink().then(res => {
		if (res) {
			window.BiyueCustomData.client_node_id = res.client_node_id
			updateHintById('imageLinkTip', '就近关联完成', '#4EAB6D')
			onImageLinkCheck()
		}
	})
}
// 图片关联检查
function onImageLinkCheck() {
	onAllCheck()
}

export {
	initView,
	handlePaperInfoResult,
	clearRepeatControl,
	onSaveData
}