import { showCom, addClickEvent, isLoading, setBtnLoading } from '../model/util.js'
;(function (window, undefined) {
	var BiyueCustomData = {}
	var tree_info = {}
	let check_text_ques = true
	let check_level = true
	var _init = false
	window.Asc.plugin.init = function () {
		console.log('quesTypeErrorReport init')
		window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'initDialog', initmsg: 'quesTypeErrorReportMessage' })
	}
	function init() {
		showCom('#uploadHint', false)
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
		addClickEvent('#downloadExamHtml', clickDownloadExamHtml)
		addClickEvent('#uploadTypeError', onUploadTypeError)
		addClickEvent('#refresh', onRefresh)
		renderQuesTypeTree()
		_init = true
	}

	function renderTreeNode(parent, item, identation = 0) {
		if (!parent) {
			return
		}
		var question_map = BiyueCustomData.question_map || {}
		var quesData = question_map[item.id]
		if (!quesData) {
			return
		}
		var  types = BiyueCustomData.paper_options.question_type || []
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
		var types = BiyueCustomData.paper_options.question_type || []
		var typeMaps = {}
		types.forEach(e => {
			typeMaps[e.value + ''] = e.label
		})
		for (var item of tree_info.list) {
			var quesData = BiyueCustomData.question_map[item.id]
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
							paper_uuid: BiyueCustomData.paper_uuid,
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
		setBtnLoading('uploadTypeError', true)
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'quesTypeErrorReportMessage',
			cmd: 'uploadTypeError',
			data: list 
		})
	}

	function onRefresh() {
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'quesTypeErrorReportMessage',
			cmd: 'refreshTypeError'
		})
	}
	function clickDownloadExamHtml() {
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'quesTypeErrorReportMessage',
			cmd: 'downloadExamHtml'
		})
	}
	window.Asc.plugin.attachEvent('quesTypeErrorReportMessage', function (message) {
		console.log('接收的消息', message)
		if (message && message.BiyueCustomData) {
			BiyueCustomData = message.BiyueCustomData
			tree_info = message.tree_info || []
			if (_init) {
				renderQuesTypeTree()
			} else {
				init()
			}
		}
	})
	window.Asc.plugin.attachEvent('uploadLoading', function (message) {
		setBtnLoading('uploadTypeError', message)
	})
})(window, undefined)
