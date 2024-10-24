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
	window.Asc.plugin.init = function () {
	  console.log('elementLinks init')
	  window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'elementLinkedMessage' })
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
		var strhtml = ''
		res.forEach(e => {
			var item = `<div class="link-item linkItem${ e.target_id}">${e.html}</div>`
			strhtml += item
		})
		$('#linksContainer').html(strhtml)
		$('#confirm').on('click', onConfirm)
		$('#hidden_empty_struct').on('click', onSwitchStruct)
		res.forEach((e, index) => {
			let doms = document.querySelectorAll('.linkItem' + e.target_id) || []
			doms.forEach(function(dom) {
				dom.addEventListener('click', function() {
					selectItem(index)
				})
			})
		})
		selectItem(0)
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

	function renderData() {
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
	  $('.batch-setting-type-info').html(html)

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
				  let selectHtml = `<select id="bat-type-group-${key}" class="type-item">`
				  selectHtml += `<option value="" style="display: none;"></option>`
				  for (const key in type_options) {
					selectHtml += `<option value="${key}">${type_options[key]}</option>`
				  }
				  selectHtml += "</select>"

				  let inputHtml = `<div class="bat-type-set">设置为${selectHtml}<span class="bat-type-set-btn" id="bat-type-set-btn-${ key }" data-id="${ key }">设置</span></div>`
				  dom.innerHTML = html + inputHtml

				  let btnDom = document.querySelector(`#bat-type-set-btn-${ key }`)
				  if (btnDom) {
					btnDom.addEventListener('click', function() {
						let id = btnDom.dataset.id || ''
						let inputDom = document.querySelector(`#bat-type-group-${ id }`)
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
			data: data
		})
	}

	window.Asc.plugin.attachEvent('initLinkedInfo', function (message) {
	  	// biyueCustomData = message || {}
		  console.log('initLinkedInfo  message', message)
		if (message) {
			biyueCustomData = message.biyueCustomData
			linked_list = message.linkedList || []
			initData(message.linkedList)
		}
	})
  })(window, undefined)
