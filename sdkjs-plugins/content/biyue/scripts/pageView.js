// 用于管理主界面
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
	onAllCheck,
	getPictureList
} from './linkHandler.js'
import { showCom, updateText, addClickEvent, getInfoForServerSave, setBtnLoading, isLoading, getYYMMDDHHMMSS, updateHintById } from './model/util.js'
import { reqSaveInfo, onLatexToImg, logOnlyOffice} from './api/paper.js'
import { biyueCallCommand } from './command.js'
import { generateTree, updateTreeSelect, clickTreeLock, initTreeListener } from './panelTree.js'
import ComponentSelect from '../components/Select.js'
import NumberInput from '../components/NumberInput.js'
import { initSetEv } from './debugging/evSet.js'
var select_image_link = null
var select_link_type = null
var input_coverage_percent = null
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
		if (window.error_list) {
			var len = window.error_list.length
			updateText('#printStack', 'error数量：' + len)
			if (len == 0) {
				setTimeout(() => {
					updateText('#printStack', '查看error数量')
				}, 1000);
			} else {
				$('#printStack').css('color', "#ff0000")
			}
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
		width: '50%',
		pop_width: '100%'
	})
	select_link_type = new ComponentSelect({
		id: 'linkTypeSelect',
		options: [
			{ value: 'all', label: '全包关联' },
			{ value: 'area', label: '面积关联' }
		],
		value_select: 'all',
		callback_item: (data) => {
			changeLinkType(data)
		},
		width: '40%',
		pop_width: '100%'
	})
	input_coverage_percent = new NumberInput('CoveragePercentInput', {
		min: 60,
		max: 100,
		change: (id, data) => {
			changeCoveragePrecent(data)
		},
		width: '40%',
	})
	if (input_coverage_percent) {
		input_coverage_percent.setValue(80)
	}
	enableBtnImageLink(true)
	showCom('#imageLinkTip')
	addClickEvent('#btnImageLink', onImageLink)
	addClickEvent('#imageLinkCheck', onImageLinkCheck)
	showCom('#uploadHint', false)
	showCom('#panelFeature', false)
	showCom('#panelLink', false)
	addClickEvent('#tabFeature', onFeature)
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
	initSetEv()
}

function onFeature() {
	window.biyue.refreshDialog({
		winName:'featureDialog',
		name:'智批元素',
		url:'featureDialog.html',
		width:400,
		height:800,
		isModal:false,
		type:'panel',
		icons:['resources/light/check.png']
	})
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
	if (id == 'tabQues') {
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
		return generateTree()
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
			// console.log('[insertContent] begin')
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
	}, false, true, {name: 'insertContent'})
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
function showPanelLink() {
	return getPictureList().then(res => {
		if (!res) {
			return
		}
		if (res.picture_id) {
			window.BiyueCustomData.picture_id = res.picture_id
		}
		if (res.table_id) {
			window.BiyueCustomData.table_id = res.table_id
		}
		Asc.scope.list_picture = res.list
		Asc.scope.list_ignore = res.list_ignore
		window.biyue.refreshDialog({
			winName:'pictureIndex',
			name:'图片关联',
			url:'pictureIndex.html',
			width:400,
			height:800,
			isModal:false,
			type:'panelRight',
			icons:['resources/light/img.png']
		}, 'pictureIndexMessage', {
			list: res.list
		})
	})
	// showCom('#panelLink', true)
	// var link_type = window.BiyueCustomData.link_type || 'all'
	// if (select_link_type) {
	// 	select_link_type.setSelect(link_type)
	// }
	// showCom('#CoveragePercentInput', link_type == 'area')
	// if (input_coverage_percent && link_type == 'area') {
	// 	var percent = window.BiyueCustomData.link_coverage_percent || 80
	// 	input_coverage_percent.setValue(percent + '')
	// }
}
function changeImageLink(data) {
	enableBtnImageLink(data.value * 1)
	showCom('#linkwrapper', data.value * 1)
}

function changeLinkType(data) {
	window.BiyueCustomData.link_type = data.value
	showCom('#CoveragePercentInput', data.value == 'area')
}

function changeCoveragePrecent(data) {
	window.BiyueCustomData.link_coverage_percent = data
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
	if ($('#btnImageLink').hasClass('btn-unable')) {
		return
	}
	window.biyue.showMessageBox({
		content: '图片关联将会清除旧有数据，确定继续操作吗？',
		extra_data: {
			func: 'onImageAutoLink'
		}
	})
}
function onImageAutoLink() {
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

function onUploadTypeErrorList(list) {
	if (!list) {
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
			window.biyue.sendMessageToWindow('quesTypeErrorReport', 'uploadLoading', false)
			setBtnLoading('uploadTypeError', false)
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
		window.biyue.sendMessageToWindow('quesTypeErrorReport', 'uploadLoading', false)
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
	preGetExamTree().then(res => {
		Asc.scope.tree_info = res
		window.biyue.refreshDialog({
			winName:'quesTypeErrorReport',
			name:'题型错误上报',
			url:'quesTypeErrorReport.html',
			width:400,
			height:800,
			isModal:false,
			type:'panel',
			icons:['resources/light/error.png']
		}, 'quesTypeErrorReportMessage', {
			tree_info: res,
			BiyueCustomData: window.BiyueCustomData
		})
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
	changeTabPanel,
	onFeature,
	showPanelLink,
	onImageAutoLink,
	insertContent,
	onUploadTypeErrorList,
	clickDownloadExamHtml
}