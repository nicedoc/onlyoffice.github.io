;(function (window, undefined) {
	let biyueCustomData = {}
	let question_map = {}
	let questionList = []
	let linked_list = []
	let tree_map = {}
	let type_options = {
		0: '未关联',
		1: '已关联'
	}
	let ques_use = []
	let hidden_empty_struct = false
	let target_id = 0
	let target_index = 0
	let tree_info = {}
	let display_tree = true
	window.Asc.plugin.init = function () {
	  console.log('elementLinks init')
	  window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'elementLinkedMessage' })
	}
	function initData(res) {
		if (!res) {
			return
		}
		$('#confirm').on('click', onConfirm)
		if (res.length == 0) {
			$('.window-body').hide()
			$('.empty').show()
			$('.hidden_empty_struct').hide()
			return
		}
		$('.empty').hide()
		var strhtml = ''
		res.forEach((e, index) => {
			var item = `<div class="link-item-box linkItem${ e.target_id}"><div class="link-item">${e.html}</div><div class="link-index">${e.type == 'table' ? '表格' : '图片'} ${index + 1}</div></div>`
			strhtml += item
		})
		$('#linksContainer').html(strhtml)
		$('#hidden_empty_struct').on('click', onSwitchStruct)
		res.forEach((e, index) => {
			let doms = document.querySelectorAll('.linkItem' + e.target_id) || []
			doms.forEach(function(dom) {
				dom.addEventListener('click', function() {
					selectItem(index)
				})
			})
		})
		$('#preview').on('click', locateItem)
		selectItem(0)
		$('#switch_tree').prop('checked', display_tree)
		$('#switch_tree').on('click', onSwitchTree)
	}

	function locateItem() {
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'elementLinkedMessage',
			cmd: 'locate',
			data: {
				target_id: target_id,
				type: linked_list[target_index].type
			}
		})
	}

	function selectItem(index) {
		if (target_index != index) {
			var pre = $(`.linkItem${linked_list[target_index].target_id}`)
			if (pre) {
				pre.removeClass('active')
			}
		}
		
		target_index = index
		if (index < linked_list.length) {
			$(`#preview`).html(linked_list[index].html)
			$(`.linkItem${linked_list[target_index].target_id}`).addClass('active')
			target_id = linked_list[index].target_id
			var use = linked_list[index].ques_use
			if (typeof use == 'string') {
				ques_use =  use.split('_')	
			} else {
				ques_use = use
			}
		}
		renderData()
	}

	function handleQuesUse(id, val) {
		if (val == 1) {
			ques_use.push(id)
		} else {
			var index = ques_use.indexOf(id)
			if (index == -1) {
				if (typeof id == 'string') {
					index = ques_use.indexOf(id * 1)	
				} else {
					index = ques_use.indexOf(id + '')
				}
			}
			if (index >= 0) {
				ques_use.splice(index, 1)
			}
		}
		linked_list[target_index].ques_use = ques_use.join('_')
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
				renderTreeNode(rootElement, item, 0, 0)
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
		if (quesData.level_type == 'struct') {
			html += `<select id="bat-group-${id}" class="type-item">`
			html += `<option value="" style="display: none;"></option>`
			for (const key in type_options) {
				html += `<option value="${key}">${type_options[key]}</option>`
			}
			html += "</select>"
		} else {
			var used = ques_use && ques_use.find(item => item == id)
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
	
	  function renderTreeNode(parent, item, structId, identation = 0) {
		if (!parent) {
			return
		}
		var quesData = question_map[item.id]
		if (!quesData) {
			return
		}
		const div = $('<div></div>')
		var html = ''
		if (item.level_type == 'struct') {
			html += `<div class=group id="group-id-${item.id}" style="padding-left:${identation}px"><span title="${quesData.text}">${quesData.text.split('\r\n')[0] }</span>`
			if (item.children && item.children.length) {
				html += '<div class="bat-set">设置为'
				html += getSelectHtml(item.id, quesData)
				html += `<span class="bat-set-btn" id="bat-set-btn-${ item.id }" data-id="${ item.id }">设置</span></div>`
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
				var sId = item.level_type == 'struct' ? item.id : structId
				if (child.level_type == 'question') {
					if (tree_map[sId]) {
						tree_map[sId].push(child.id)
					} else {
						tree_map[sId] = [child.id]
					}
				}
				renderTreeNode(parent, child, sId,  display_tree ? identation : 0)
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
		if (display_tree) {
			renderTree()
			return
		}
	  let node_list = biyueCustomData.node_list || []
	  question_map = biyueCustomData.question_map || {}

	  let html = ''
	  let question_list = []
	  let tree = {}
	  let pre_struct = ''
	  for (const key in node_list) {
		let item = question_map[node_list[key].id] || ''
		if (node_list[key].level_type == 'question') {
		  if (item) {
			html += `<span class="question" title="${ item.text }">${(item.ques_name || item.ques_default_name || '')}`
			var used = ques_use && ques_use.find(item => item == node_list[key].id)
			html += `<select class="type-item ques-${ node_list[key].id }" style="color:${used ? '#4CAF50' : ''}">`
			// html += `<option value="" style="display:none;"></option>`
			for (const key in type_options) {
			  let selected = ''
			  if (used && key == 1) {
				  selected = 'selected'
			  } else if (!used && key == 0) {
				  selected = 'selected'
			  }
			  html += `<option value="${key}" ${selected}>${type_options[key]}</option>`
			}

			html += `</select></span>`

			if (!question_list.includes(node_list[key].id)) {
			  question_list.push(node_list[key].id)
			}
			if (!pre_struct && !tree[pre_struct]) {
			  tree[pre_struct] = []
			}
			if (pre_struct && tree[pre_struct]) {
			  tree[pre_struct].push(node_list[key].id)
			}
		  }
		} else if (node_list[key].level_type == 'struct' && item){
		  pre_struct = node_list[key].id

		  html += `<div class="group" id="group-id-${ pre_struct }"><span>${ item.text.split('\r\n')[0] }</span></div>`
		  if (!tree[pre_struct]) {
			  tree[pre_struct] = []
		  }
		}
	  }
	  tree_map = tree
	  questionList = question_list || []
	  $('.batch-setting-info').html(html)
	  if (html == '') {
		$('.batch-setting-info').html('<div class="ques-none">暂无题目，请先切题</div>')
		$('.hidden_empty_struct').hide()
	  }

	  // 处理题目的下拉选项事件
	  for (const key in questionList) {
		let id = questionList[key]
		let doms = document.querySelectorAll('.ques-' + id) || []
		doms.forEach(function(dom) {
			dom.addEventListener('change', function() {
				handleQuesUse(id, dom.value || '')
				if (dom.value > 0) {
				dom.style.color = '#4CAF50'
				} else {
				dom.style.color = ''
				}
			})
			if (dom.value > 0) {
				dom.style.color = '#4CAF50'
			} else {
				dom.style.color = ''
			}
		})
	  }

	  for (const key in tree_map) {
		  if (tree_map[key].length > 0) {
			  // 对有题的结构增加批量设置题型的下拉框
			  let dom = document.querySelector('#group-id-' + key)
			  if (dom) {
				  let html = dom.innerHTML // 取出当前的题组内容
				  let selectHtml = `<select id="bat-group-${key}" class="type-item">`
				  selectHtml += `<option value="" style="display: none;"></option>`
				  for (const key in type_options) {
					selectHtml += `<option value="${key}">${type_options[key]}</option>`
				  }
				  selectHtml += "</select>"

				  let inputHtml = `<div class="bat-set">设置为${selectHtml}<span class="bat-set-btn" id="bat-set-btn-${ key }" data-id="${ key }">设置</span></div>`
				  dom.innerHTML = html + inputHtml

				  let btnDom = document.querySelector(`#bat-set-btn-${ key }`)
				  if (btnDom) {
					btnDom.addEventListener('click', function() {
						let id = btnDom.dataset.id || ''
						let inputDom = document.querySelector(`#bat-group-${ id }`)
						batchSetStructType(id, inputDom.value || 0)
					})
				  }
			  }
		  }
	  }
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
		var data = linked_list.map(e => {
			return {
				ques_use: e.ques_use,
				target_id: e.target_id,
				type: e.type
			}
		})
		// 将窗口的信息传递出去
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'elementLinkedMessage',
			cmd: 'confirm',
			data: data
		})
	}

	window.Asc.plugin.attachEvent('initLinkedInfo', function (message) {
	  	// biyueCustomData = message || {}
		  console.log('initLinkedInfo  message', message)
		if (message) {
			biyueCustomData = message.biyueCustomData
			linked_list = message.linkedList || []
			tree_info = message.tree_info || {}
			initData(message.linkedList)
		}
	})
  })(window, undefined)
