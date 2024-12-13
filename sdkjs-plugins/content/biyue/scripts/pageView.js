// 用于管理主界面

import {
	initFeature,
} from './panelFeature.js'

import { showQuesData, initListener } from './panelQuestionDetail.js'
import {
	handleAllWrite,
	showAskCells,
	g_click_value,
	clearRepeatControl,
	tidyNodes,
	preGetExamTree,
	getQuestionHtml
} from './QuesManager.js'
import {
	imageAutoLink,
	onAllCheck
} from './linkHandler.js'
import { showCom, updateText, addClickEvent, getInfoForServerSave, setBtnLoading, isLoading, getYYMMDDHHMMSS } from './model/util.js'
import { reqSaveInfo, onLatexToImg, logOnlyOffice} from './api/paper.js'
import { biyueCallCommand } from './command.js'
import { generateTree, updateTreeSelect, clickTreeLock, initTreeListener } from './panelTree.js'
import ComponentSelect from '../components/Select.js'
var timeout_paste_hint = null
var select_image_link = null
let check_text_ques = true
let check_level = false
let inner_pop_list = ['link_container', 'func_key_container', 'tree_container']
const CLR_SUCCESS = '#4EAB6D'
const CLR_FAIL = '#ff0000'
function initView() {
	showCom('#initloading', true)
	showCom('#hint1', false)
	showCom('.tabs', false)
	showCom('#panelTree', false)
	initListener()
	initTreeListener()
	window.tab_select = 'tabTree'
	$('.tabitem').on('click', changeTab)
	document.addEventListener('clickSingleQues', function (event) {
		if (window.tab_select == 'tabTree' && window.tree_lock) {
			if (event && event.detail) {
				updateTreeSelect(event.detail)
			}
			return
		}
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
	addClickEvent('#reSplitQues', () => {
		window.biyue.reSplitQustion()
	})
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
	addClickEvent('#tidyNodes', tidyNodes)
	addClickEvent('#mathpix', onMathpix)
	addClickEvent('#paste', onPaste)
	addClickEvent('#pasteclear', onPasteInputClear)
	select_image_link = new ComponentSelect({
		id: 'imageLinkSelect',
		options: [
			{ value: '0', label: '关闭就近关联' },
			{ value: '1', label: '开启就近关联' }
		],
		value_select: '1',
		callback_item: (data) => {
			changeImageLink(data)
		},
		width: '60%',
		pop_width: '100%'
	})
	enableBtnImageLink(true)
	showCom('#imageLinkTip')
	addClickEvent('#btnImageLink', onImageLink)
	addClickEvent('#imageLinkCheck', onImageLinkCheck)
	if ($('#check_level')) {
		$('#check_level').prop('checked', check_level)
		$('#check_level').on('click', () => {
			check_level = !check_level
			renderQuesTypeTree()
		})
	}
	if ($('#check_text_ques')) {
		$('#check_text_ques').prop('checked', check_text_ques)
		$('#check_text_ques').on('click', () => {
			check_text_ques = !check_text_ques
			renderQuesTypeTree()
		})
	}
	addClickEvent('#uploadTypeError', onUploadTypeError)
	showCom('#uploadHint', false)
	showCom('#panelFeature', false)
	showCom('#panelLink', false)
	showCom('#panelTypeErrorUpload', false)
	addClickEvent('#tabFeature', () => {
		$('#panelFeature').show()
		initFeature()
	})
	addClickEvent('.panelclose', (e) => {
		var target = e.currentTarget || e.target
		if (target && target.dataset && target.dataset.panel) {
			showCom(`#${target.dataset.panel}`, false)
		}
	})
	addClickEvent('#extrolclose', () => {
		showButton('extro_buttons', false)
	})
	document.addEventListener('contextmenu', e=> {
		e.preventDefault()	
	})
	// 点击其他地方隐藏所有菜单
	$(document).on('click', function() {
		if (window.show_dynamic_menu) {
			$('.customContextMenu').hide();
		}
	});
	$('#panelTree #tree').on('scroll', function() {
		if (window.show_dynamic_menu) {
			$('.customContextMenu').hide();
		}
	})
	addClickEvent('#panelTree #lock', clickTreeLock)
	addClickEvent('#downloadExamHtml', clickDownloadExamHtml)
}

function handlePaperInfoResult(success, res) {
	showCom('#initloading', false)
	if (success) {
		showCom('.tabs', true)
		// changeTabPanel('tabTree')
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
	var tabs = ['tabList', 'tabQues', 'tabTree']
	tabs.forEach((tab) => {
		if (tab == window.tab_select && tab != id) {
			$('#' + tab).removeClass('selected')
		} else if (tab == id) {
			$('#' + tab).addClass('selected')
		}
	})
	window.tab_select = id
	$('.customContextMenu').hide();
	var targetPanel = id.replace('tab', 'panel')
	var panels = ['panelList', 'panelQues', 'panelTree']
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
		if (event && event.detail && event.detail.parentTag) {
			showQuesData(event.detail)
		} else if (g_click_value) {
			showQuesData(Object.assign({}, g_click_value.Tag, {
				InternalId: g_click_value.InternalId,
				Appearance: g_click_value.Appearance,
			}))
		} else {
			showQuesData(event ? event.detail : null)
		}
	} else if (id == 'tabTree') {
		generateTree()
	}
	scroll(false)
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
						oDrawing.getParaDrawing()
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

function updateHintById(id, message, color, duration = 1500) {
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
	}, duration);
}

function updatePasteHint(message, color) {
	updateHintById('pastehint', message, color)
}

function onMathpix() {
	var com = $('#paste-target')
	if (!com) {
		updatePasteHint('未找到输入框', CLR_FAIL)
		return
	}
	var text1 = $('#paste-target').val()
	if (!text1 || !text1.length) {
		updatePasteHint('输入框内容为空', CLR_FAIL)
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
			updatePasteHint(res.message, CLR_FAIL)
		} else if (!res) {
			updatePasteHint('网络不稳定', CLR_FAIL)
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
	window.auto_image_link = v
}
// 图片就近关联
function onImageLink() {
	imageAutoLink(null, true).then(res => {
		if (res) {
			updateHintById('imageLinkTip', '就近关联完成', CLR_SUCCESS)
			onImageLinkCheck()
		}
	})
}
// 图片关联检查
function onImageLinkCheck() {
	onAllCheck()
}

function renderTreeNode(parent, item, identation = 0) {
	if (!parent) {
		return
	}
	var question_map = window.BiyueCustomData.question_map || {}
	var quesData = question_map[item.id]
	if (!quesData) {
		return
	}
	var  types = window.BiyueCustomData.paper_options.question_type || []
	var html = ''
	if (item.level_type == 'struct') {
		if (check_level) {
			html += `<div class=item id="group-id-${item.id}" style="padding-left:${identation}px"><div class="text-over-ellipsis flex-1 ml-4" title="${quesData.text}">${quesData.text}</div></div>`
		}
	} else if (item.level_type == 'question') {
		if (quesData.question_type != 6 || check_text_ques) {
			var quesType = types.find(e => {
				return e.value == quesData.question_type
			})
			var typeOptions = types.map(e => {
				return `<option value="${e.value}">${e.label}</option>`
			})
			html += `<div class="item" style="padding-left:${identation}px">
        				<div class="content" title="${quesData.text}">
          					<input type="checkbox" id="check_${item.id}">
          					<div class="text-over-ellipsis flex-1 ml-4">${quesData.text}</div>
        				</div>
        				<div class="ques-type-name" id="ques-type-${item.id}">${quesType ? quesType.label : '未定义'}</div>
						<select class="target-select" id="target_type_${item.id}">
							<option value="" style="display:none;"></option>
							${typeOptions.join('')}
						</select>
      				</div>`;
		}
	}
	parent.append(html)
	if (item.children && item.children.length > 0) {
		identation += 20
		for (var child of item.children) {
			renderTreeNode(parent, child, check_level ? identation : 0)
		}
	}
}

function renderQuesTypeTree() {
	var tree_info = Asc.scope.tree_info || {}
	var rootElement = $('#typeErrorDiv')
	rootElement.empty()
	if (tree_info.tree && tree_info.tree.length) {
		tree_info.tree.forEach(item => {
			renderTreeNode(rootElement, item, 0)
		})
	} else {
		$('#typeErrorDiv').html('<div class="ques-none">暂无题目，请先切题</div>')
	} 
}

function onUploadTypeError() {
	if (isLoading('uploadTypeError')) {
		return
	}
	var list = []
	var types = window.BiyueCustomData.paper_options.question_type || []
	var typeMaps = {}
	types.forEach(e => {
		typeMaps[e.value + ''] = e.label
	})
	for (var item of Asc.scope.tree_info.list) {
		var quesData = window.BiyueCustomData.question_map[item.id]
		if (item.level_type == 'question' && quesData) {
			var check = $('#check_' + item.id)
			if (check) {
				if (check.prop('checked')) {
					var target_type = $('#target_type_' + item.id).val()
					if (!target_type) {
						updateHintById('uploadHint', '请设置目标题型', CLR_FAIL)
						return
					}
					list.push({
						paper_uuid: window.BiyueCustomData.paper_uuid,
						id: item.id,
						question_type: quesData.question_type,
						question_type_name: typeMaps[quesData.question_type + ''],
						target_type: target_type,
						target_type_name: typeMaps[target_type + '']
					})
				}
			}
		}
	}
	if (list.length == 0) {
		updateHintById('uploadHint', '请勾选需要上传的题目', CLR_FAIL)
		return
	}
	if (!list || list.length == 0) {
		return
	}
	getQuestionHtml(list.map(e => e.id), 2).then(html_list => {
		if (html_list) {
			list.forEach(item => {
				var find = html_list.find(e => {
					return e.id == item.id
				})
				if (find) {
					item.content_type = find.content_type
					item.content_html = find.content_html
					if (find.context_list && find.context_list.length) {
						item.context_html = find.context_list.map(e => {
							return e.content_html
						}).join('')
					}
				}
			})
			logQuesTypeError(list, 0)
		}
	})
}
function logQuesTypeError(list, index) {
	if (index >= list.length) {
		updateHintById('uploadHint', '上传成功', CLR_SUCCESS)
		setTimeout(() => {
			setBtnLoading('uploadTypeError', false)
			showCom('#panelTypeErrorUpload', false)
		}, 2000)
		return
	}
	logOnlyOffice(JSON.stringify(list[index])).then(res => {
		var com = document.getElementById('ques-type-' + list[index].id);
		if (com) {
			com.textContent = '已上传'
			com.style.color = CLR_SUCCESS
		}
		logQuesTypeError(list, index + 1)
	}).catch(res => {
		updateHintById('uploadHint', res ? res.message || '上传失败' : '网络不稳定', CLR_FAIL)
		setBtnLoading('uploadTypeError', false)
	})
}

function hidePops(type, name) {
	if (inner_pop_list) {
		for (var item of inner_pop_list) {
			if (type == 'all') {
				showButton(item, false)
			} else if (type == 'except') { // 除了name以外的，全部关闭
				if (item == name) {
					continue
				}
				showButton(item, false)
			} else if (type == 'only') { // 只关闭name
				if (item == name) {
					showButton(item, false)
					break
				}
			}
		}
	}
}

function clickDownloadExamHtml() {
	return getQuestionHtml().then(htmlList => {
		if (htmlList) {
			var types = window.BiyueCustomData.paper_options.question_type || []
			var typeMaps = {}
			types.forEach(e => {
				typeMaps[e.value + ''] = e.label
			})
			var obj = {
				paper_uuid: window.BiyueCustomData.paper_uuid,
				exam_title: window.BiyueCustomData.exam_title,
				content_list: []
			}
			for (var item of htmlList) {
				if (item.content_type == 'question') {
					var quesData = window.BiyueCustomData.question_map[item.id]
					if (quesData) {
						obj.content_list.push({
							id: item.id,
							content_type: item.content_type,
							question_type: quesData.question_type,
							question_type_name: typeMaps[quesData.question_type + ''],
							content_html: item.content_html
						})
					}
				} else {
					obj.content_list.push(item)
				}
			}
			// 创建一个 Blob 对象
			var textToDownload = JSON.stringify(obj)
			var blob = new Blob([textToDownload], { type: 'text/plain' });

			// 创建一个指向该 Blob 对象的 URL
			var url = URL.createObjectURL(blob);

			// 创建一个临时的 <a> 元素，用于触发下载
			var a = document.createElement('a');
			a.href = url;
			a.download = `${getYYMMDDHHMMSS()}_题型识别错误` ; // 指定下载的文件名

			// 触发下载
			document.body.appendChild(a);
			a.click();

			// 移除临时 <a> 元素
			document.body.removeChild(a);

			// 释放这个 URL 对象
			URL.revokeObjectURL(url);
		}
	})
}

function clickSplitQues() {
	window.biyue.showMessageBox({
		title: '提示',
		content: '确定要重新切题吗？',
		extra_data: {
			func: 'reSplitQustion'
		}
	})
}

function clickUploadTree() {
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
}

function showTypeErrorPanel() {
	showCom('#panelTypeErrorUpload', true)
	preGetExamTree().then(res => {
		Asc.scope.tree_info = res
		renderQuesTypeTree()
	})
}
export {
	initView,
	handlePaperInfoResult,
	clearRepeatControl,
	onSaveData,
	clickSplitQues,
	clickUploadTree,
	showTypeErrorPanel,
	changeTabPanel
}