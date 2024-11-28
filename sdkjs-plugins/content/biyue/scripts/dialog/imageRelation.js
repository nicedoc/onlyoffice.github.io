;(function (window, undefined) {
	let biyueCustomData = {}
	let question_map = {}
	let questionList = []
	let tree_map = {}
	let type_options = {
		0: '未关联',
		1: '已关联'
	}
	let ques_use = []
	let hidden_empty_struct = false
	let target_id = 0
	let target_type = 'drawing'
	let tree_info = {}
	let display_tree = true
	window.Asc.plugin.init = function () {
	  console.log('imageRelation init')
	  window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'imageRelationMessage' })
	}

	function init() {
		window.Asc.plugin.callCommand(function() {
			var oDocument = Api.GetDocument()
			var selectDrawings = oDocument.GetSelectedDrawings() || []
			var selectionInfo = oDocument.Document.getSelectionInfo()
			if (selectDrawings.length && selectDrawings[0].Drawing) {
				var drawing = selectDrawings[0].getParaDrawing()
				if (!drawing.GetRun) {
					return
				}
				var parentRun = drawing.GetRun()
				if (parentRun) {
					var oRun = Api.LookupObject(parentRun.Id)
					var oRange = oRun.GetRange()
					oRange.private_SetRangePos(selectionInfo.selectionStart, selectionInfo.selectionEnd)
					oRange.Select()
					let text_data = {
						data:     "",
						// 返回的数据中class属性里面有binary格式的dom信息，需要删除掉
						pushData: function (format, value) {
							this.data = value ? value.replace(/class="[a-zA-Z0-9-:;+"\/=]*/g, "") : "";
						}
					};
					Api.asc_CheckCopy(text_data, 2);
					return {
						html: text_data.data,
						title: selectDrawings[0].GetTitle(),
						type: 'drawing',
						target_id: selectDrawings[0].Drawing.Id
					}
				}
			} else if (selectionInfo.curPos) {
				var oTable = null
				for (var i = 0; i < selectionInfo.curPos.length; ++i) {
					if (selectionInfo.curPos[i].Class) {
						var oElement = Api.LookupObject(selectionInfo.curPos[i].Class.Id)
						if (oElement && oElement.GetClassType && oElement.GetClassType() == 'table') {
							oTable = oElement
							break
						}
					}
				}
				if (oTable) {
					var oRange = oTable.GetRange()
					oRange.Select()
					let text_data = {
						data:     "",
						// 返回的数据中class属性里面有binary格式的dom信息，需要删除掉
						pushData: function (format, value) {
							this.data = value ? value.replace(/class="[a-zA-Z0-9-:;+"\/=]*/g, "") : "";
						}
					};
					Api.asc_CheckCopy(text_data, 2);
					return {
						html: text_data.data,
						title: oTable.GetTableTitle(),
						type: 'table',
						target_id: oTable.Table.Id
					}
				}
			}
		}, false, false, function(res) {
			if (res) {
				initData(res)
			}
		})
	}

	function getJsonData(str) {
		if (!str || str == '' || typeof str != 'string') {
			return {}
		}
		try {
			return JSON.parse(str)
		} catch (error) {
			console.log('json parse error', error)
			return {}
		}
	}

	function initData(res) {
		if (!res) {
			return
		}
		target_id = res.target_id
		target_type = res.type
		var titleObj = getJsonData(res.title)
		if (target_type == 'table') {
			ques_use = titleObj.ques_use ? titleObj.ques_use.split('_') : []
		} else {
			if (titleObj.feature && titleObj.feature.ques_use) {
				ques_use = titleObj.feature.ques_use.split('_')
			} else {
				ques_use = []
			}
		}
		getOptions()
		$('#header').html(target_type == 'table' ? '当前表格：' : '当前图片：')
		$('#imageContainer').html(res.html)
		$('#confirm').on('click', onConfirm)
		$('#hidden_empty_struct').on('click', onSwitchStruct)
		$('#switch_tree').prop('checked', display_tree)
		$('#switch_tree').on('click', onSwitchTree)
	}

	function getOptions() {
	  renderData()
	}

	function handleQuesUse(id, val) {
		if (val == 1) {
			ques_use.push(id + '')
		} else {
			var index = ques_use.indexOf(id + '')
			ques_use.splice(index, 1)
		}
	}

	function renderTree() {
		question_map = biyueCustomData.question_map || {}
		var rootElement = $('.batch-setting-info')
		rootElement.empty()
		questionList = []
		tree_map = {}
		if (!tree_info || !tree_info.tree || tree_info.tree.length == 0) {
			$('.batch-setting-info').html('<div class="ques-none">暂无题目，请先切题</div>')
			showBottomBtns(false)
		} else {
			showBottomBtns(true)
			tree_info.tree.forEach(item => {
				renderTreeNode(rootElement, item, 0)
			})
			addEvents()
			updateHideEmptyStruct()
		}
	  }

	  function showBottomBtns(v) {
		if (v) {
			$('.hidden_empty_struct').show()
			$('.switch-tree').show()
		} else {
			$('.hidden_empty_struct').hide()
			$('.switch-tree').hide()
		}
	  }
	
	  function getSelectHtml(id, quesData) {
		var html = ''
		if (quesData && quesData.level_type == 'struct') {
			html += `<select id="bat-group-${id}" class="type-item">`
			html += `<option value="" style="display: none;"></option>`
			for (const key in type_options) {
				html += `<option value="${key}">${type_options[key]}</option>`
			}
			html += "</select>"
		} else {
			var used = ques_use.find(item => item == id)
			html += `<select class="type-item ques-${ id }" style="color:${used ? '#4CAF50' : ''}">`
			for (const key in type_options) {
				let selected = ''
				if (used && key == 1) {
					selected = 'selected'
				} else if (!used && key == 0) {
					selected = 'selected'
				}
				html += `<option value="${key}" ${selected}>${type_options[key]}</option>`
			}
			html += `</select>`
		}
		return html
	  }
	
	  function renderTreeNode(parent, item, identation = 0) {
		if (!parent || !item) {
			return
		}
		var quesData = question_map[item.id]
		if (!quesData) {
			return
		}
		const div = $('<div></div>')
		var html = ''
		if (item.level_type == 'struct') {
			html += `<div class=group id="group-id-${item.id}" style="padding-left:${identation}px"><span title="${quesData.text}">${quesData.text.split('\r\n')[0]}</span>`
			if (item.children && item.children.length) {
				html += '<div class="bat-set">设置为'
				html += getSelectHtml(item.id, quesData)
				html += `<span class="bat-set-btn" id="bat-set-btn-${ item.id }" data-id="${ item.id }">设置</span></div></div>`
			}
			html += '</div>'
			tree_map[item.id] = []
		} else if (item.level_type == 'question') {
			html += `<span class="question" title="${quesData.text}" style="padding-left:${identation}px">${quesData.ques_name || quesData.ques_default_name || ''}`
			html += getSelectHtml(item.id, quesData)
			html += '</span>'
			questionList.push(item.id)
		}
		if (display_tree || item.level_type == 'struct') {
			div.html(html)
			parent.append(div)
		} else {
			parent.append(html)
		}
		if (item.children && item.children.length > 0) {
			identation += 20
			for (var child of item.children) {
				if (item.level_type == 'struct') {
					if (child.level_type == 'question') {
						if (tree_map[item.id]) {
							tree_map[item.id].push(child.id)
						} else {
							tree_map[item.id] = [child.id]
						}
					}
				}
				renderTreeNode(parent, child, display_tree ? identation : 0)
			}
		}
	}

	function addEvents() {
		// 处理题目的下拉选项事件
		for (const key in questionList) {
			let id = questionList[key]
			let doms = document.querySelectorAll('.ques-' + id) || []
			doms.forEach(function(dom) {
				function changeHandler() {
					handleQuesUse(id, dom.value || '')
					dom.style.color = dom.value > 0 ? '#4CAF50' : ''
				}
				dom.removeEventListener('change', changeHandler)
				dom.addEventListener('change', changeHandler)
				dom.style.color = dom.value > 0 ? '#4CAF50' : ''
			})
		}
	
		for (const key in tree_map) {
			if (tree_map[key].length > 0) {
				// 对有题的结构增加批量设置题型的下拉框
				let btnDom = document.querySelector(`#bat-set-btn-${ key }`)
				if (btnDom) {
					function clickHandler() {
						let id = btnDom.dataset.id || ''
						let inputDom = document.querySelector(`#bat-group-${ id }`)
						batchSetStructType(id, inputDom.value || 0)
					}
					btnDom.removeEventListener('click', clickHandler)
					btnDom.addEventListener('click', clickHandler)
				}
			}
		}
	}

	function renderData() {
		renderTree()
	}

	function batchSetStructType(struct_id, value) {
	  // 根据结构批量写入题目类型
	  	let arr = tree_map[struct_id] || []
	  	for (const key in arr) {
			handleQuesUse(arr[key], value || '')
	  	}
	  renderData()
	}

	function onSwitchStruct() {
	  	//  开启/隐藏 无题目的结构
	  	hidden_empty_struct = !hidden_empty_struct
		updateHideEmptyStruct()
	}

	function updateHideEmptyStruct() {
		$('#hidden_empty_struct').prop('checked', hidden_empty_struct)

	  	for (const key in tree_map) {
		 	let arr = tree_map[key] || []
		  	if (arr.length === 0) {
			  	if (hidden_empty_struct) {
				  	$('#group-id-'+ key).hide()
			  	} else {
				  	$('#group-id-'+ key).show()
			  	}
		  	}
	  	}
	}

	function onSwitchTree() {
		display_tree = !display_tree
		$('#switch_tree').prop('checked', display_tree)
		renderData()
	}

	function onConfirm() {
		// 将窗口的信息传递出去
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'imageRelationMessage',
			data: {
				target_id: target_id,
				target_type: target_type,
				ques_use: ques_use
			}
		})
	}

	window.Asc.plugin.attachEvent('initInfo', function (message) {
	  	biyueCustomData = message ? (message.biyueCustomData || {}) : {}
		tree_info = message.tree_info || {}
	  	init()
	})
  })(window, undefined)
