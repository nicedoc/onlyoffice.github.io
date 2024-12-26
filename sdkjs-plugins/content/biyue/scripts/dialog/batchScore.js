import { setBtnLoading, addClickEvent, addTypeEvent, showCom, updateText } from '../model/util.js'
var _ = window._;
;(function (window, undefined) {
	let biyueCustomData = {}
	let question_map = {}
	let questionList = []
	let tree_map = {}
	let hidden_empty_struct = false
	let tree_info = {}
	let display_tree = true
	let isLock = false
	let oldQuestionMap;
	let oldTreeInfo;
	window.Asc.plugin.init = function () {
	  	console.log('batch score init')
	  	window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'PaperMessage' })
	}
  
	function init() {
	  	renderData()
		addClickEvent('#refresh', onRefresh)
		addClickEvent('#confirm', onConfirm)
		addClickEvent('#hidden_empty_struct', onSwitchStruct)
		$('#switch_tree').prop('checked', display_tree)
		addClickEvent('#switch_tree', onSwitchTree)
	}

	let eventHandlers = {
		inputHandler: function(e) {
			let inputDom = e.target;
			let inputValue = inputDom.value
			inputDom.value = checkInputValue(inputValue, 100)
		},
		clickHandler: function(e) {
			let btnDom = e.target;
			let id = btnDom.dataset.id || ''
			let inputDom = document.querySelector(`#bat-group-${ id }`)
			batchSetStructScore(id, inputDom.value || 0)
			getScoreSum()
		},
		changeHandler: function(e) {
			getScoreSum()
		},
		anotherInputHandler: function(e) {
			let dom = e.target;
			let inputValue = dom.value
			dom.value = checkInputValue(inputValue, 100)
		},
		focusHandler: function(e) {
			focusInput(e)
		}
	};
  
	function renderTree() {
		question_map = biyueCustomData.question_map || {}
		var rootElement = $('.batch-setting-info')
		rootElement.empty()
		questionList = []
		tree_map = {}
		if (!tree_info || !tree_info.tree || tree_info.tree.length == 0) {
			updateText('.batch-setting-info', '<div class="ques-none">暂无题目，请先切题</div>')
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
		showCom('.hidden_empty_struct', v)
		showCom('.switch-tree', v)
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
			var structName = quesData.ques_name || quesData.ques_default_name || quesData.text.split('\r\n')[0]
			html += `<div class="group row-between" id="${item.id}"><div class="text-over-ellipsis" style="display: block;width: 100px;" title="${quesData.text}">${structName}</div>`
			if (item.children && item.children.length) {
				html += `<div class="bat-set"><input class="input" type="text" id="bat-group-${item.id}"><span class="bat-set-btn" id="bat-set-btn-${item.id}" data-id=${item.id}>设置</span></div>`
			}
			html += '</div>'
		} else if (item.level_type == 'question' && quesData.question_type != 6) {
			html += `<span class="question" title="${quesData.text}">${quesData.ques_name || quesData.ques_default_name || ''}`
			var showAsks = quesData.ask_list && quesData.ask_list.length > 0
			if (showAsks) {
				let choice_display = biyueCustomData.choice_display || {}
				let show_choice_region = choice_display.style == 'show_choice_region' // 判断是否为开启集中作答区
				if (show_choice_region) {
					var nodeId = quesData.ids && quesData.ids.length ? quesData.ids[0] : item.id 
					var nodeData = biyueCustomData.node_list.find(e => {
						return e.id == nodeId
					})
					if (nodeData && nodeData.use_gather) {
						showAsks = false
					}
				}
			}
			if (showAsks) {
				if (quesData.ask_list && quesData.ask_list.length > 0) {
					quesData.ask_list.forEach((ask, index) => {
						if (index == 0) {
							html += `<input type="text" data-qid=${item.id} data-aidx=${index} class="score ques-${item.id} ask-index-${index} ask-${ask.id}" value=${ask.score || 0}>`
						} else {
							html += `<div style="padding-left: 12px"><input type="text" data-qid=${item.id} data-aidx=${index} class="score ques-${item.id} ask-index-${index} ask-${ask.id}" value=${ask.score || 0}></div>`
						}
					})
				}
			} else {
				html += `<input type="text" data-qid=${item.id} class="score ques-${ item.id }" value="${ quesData.score || 0 }">`
			}
			html += `</span>`
			questionList.push(item.id)
		}
		if (display_tree) {
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
  
	function renderData() {
	  	renderTree()
	}

	function onAllInput() {
		let inputValue = $('#bat-group-all').val()
		$('#bat-group-all').val(checkInputValue(inputValue, 100))
	}

	function clickAllSet() {
		let inputValue = $('#bat-group-all').val()
		batchSetEmptyScore(inputValue)
		$('#bat-group-all').val('')
	}
  
	function addEvents() {
		// 处理批量设置所有未设置分数的输入框事件
		addTypeEvent('#bat-group-all', 'input', onAllInput)
		// 处理批量设置所有未设置分数的输入框事件
		addClickEvent('#bat-set-btn-all', clickAllSet)
		for (var item of tree_info.list) {
			if (item.level_type == 'struct') {
				let inputDom = document.querySelector(`#bat-group-${ item.id }`)
				if (inputDom) {
					inputDom.addEventListener('input', eventHandlers.inputHandler);
				}
	
				let btnDom = document.querySelector(`#bat-set-btn-${ item.id }`)
				if (btnDom) {
					btnDom.addEventListener('click', eventHandlers.clickHandler);
				}
			} else {
				let doms = document.querySelectorAll('.ques-' + item.id) || []
				doms.forEach(function(dom) {
					dom.addEventListener('change', eventHandlers.changeHandler);
					dom.addEventListener('input', eventHandlers.anotherInputHandler);
					dom.addEventListener('focus', eventHandlers.focusHandler);
				})
			}
		}
		getScoreSum()
	}
	// 移除事件监听器
	function removeEvents() {
		if (!tree_info || !tree_info.list) {
			return
		}
		for (var item of tree_info.list) {
			if (item.level_type == 'struct') {
				let inputDom = document.querySelector(`#bat-group-${ item.id }`)
				if (inputDom) {
					inputDom.removeEventListener('input', eventHandlers.inputHandler);
				}
	
				let btnDom = document.querySelector(`#bat-set-btn-${ item.id }`)
				if (btnDom) {
					btnDom.removeEventListener('click', eventHandlers.clickHandler);
				}
			} else {
				let doms = document.querySelectorAll('.ques-' + item.id) || []
				doms.forEach(function(dom) {
					dom.removeEventListener('change', eventHandlers.changeHandler);
					dom.removeEventListener('input', eventHandlers.anotherInputHandler);
					dom.removeEventListener('focus', eventHandlers.focusHandler);
				})
			}
		}
	}

	function focusInput(e) {
		var target = e.target || e.srcElement
		if (!target || !target.dataset) {
			return
		}
		if (target.dataset.qid) {
			window.Asc.plugin.sendToPlugin('onWindowMessage', {
				type: 'focusQuestion',
				data: {
				  	ques_id: target.dataset.qid,
				  	ask_index: target.dataset.aidx
				},
			})
		}
	}
  
	function checkInputValue(val = '', max) {
		// 只允许数字和一个小数点，并且小数点后最多一位数字
		var sanitizedValue = val.replace(/[^0-9.]/g, ''); // 移除非数字和小数点的字符
	
		// 处理多个小数点的情况
		var parts = sanitizedValue.split('.');
		if (parts.length > 2) {
			sanitizedValue = parts[0] + '.' + parts.slice(1).join('');
		}
	
		// 处理最多一位小数的情况
		if (parts.length === 2) {
			sanitizedValue = parts[0] + '.' + (parts[1].length > 1 ? parts[1].substring(0, 1) : parts[1]);
		}
		if (max) {
			sanitizedValue = sanitizedValue > max ? max : sanitizedValue
		}
		return sanitizedValue
	}
  
	function batchSetStructScore(struct_id, value) {
		saveDataBeforeRender()
		// 根据结构批量写入分数
		let arr = tree_map[struct_id] || []
		for (const key in arr) {
			let id = arr[key]
			let ask_list = question_map[id].ask_list || []
			if (ask_list.length > 0) {
				let sum = 0
				for (const k in ask_list) {
					ask_list[k].score = value
					sum += parseFloat(value) || 0
				}
				question_map[id].score = sum
			} else {
				question_map[id].score = value
			}
		}
		renderData()
	}
  
	function batchSetEmptyScore(value) {
	  // 根据找到所有没有设置分数的空写入分数
	  for (const key in tree_map) {
		if (tree_map[key].length > 0) {
		  for (const k in tree_map[key]) {
			// 题目的id
			let id = tree_map[key][k]
			var quesData = question_map[id] || question_map[id + '']
			if (quesData && quesData.ask_list && quesData.ask_list.length > 0) {
			  let ask_list = quesData.ask_list || []
			  let sum = 0
			  let hasChange = false
			  for (const k in ask_list) {
				let score = parseFloat(ask_list[k].score) || 0
				if (!score || score * 1 == 0) {
				  ask_list[k].score = value
				  hasChange = true
				}
				sum += parseFloat(ask_list[k].score) || 0
			  }
			  if (hasChange) {
				  quesData.score = sum
			  }
			} else if (!(quesData.score * 1)) {
			  quesData.score = value
			}
		  }
		}
	  }
	  renderData()
	}
  
	function getScoreSum() {
	  let score_sum = 0
	  for (const key in questionList) {
		let id = questionList[key]
		let doms = document.querySelectorAll('.ques-' + id) || []
		doms.forEach(function(dom) {
		  let val = parseFloat(dom.value) || 0
		  if (val == 0) {
			dom.value = 0
			dom.style.color='#ff0000'
		  } else {
			dom.style.color=''
			dom.value = val * 1
		  }
		  score_sum += val
		})
	  }
  
	  $('#score_sum').html(score_sum)
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
	function saveDataBeforeRender() {
		let choice_display = biyueCustomData.choice_display || {}
		let show_choice_region = choice_display.style == 'show_choice_region' // 判断是否为开启集中作答区
		let node_list = biyueCustomData.node_list || []
		for (const key in questionList) {
			let id = questionList[key] || ''
			let dom = $('.ques-'+id)
			if (id && question_map[id] && dom) {
			let ask_list = question_map[id].ask_list || []
			var nodeData = node_list.find(e => {
				return e.id == id
			})
			if (ask_list.length > 0 && (!show_choice_region || show_choice_region && (!nodeData || !nodeData.use_gather))) {
				// 有小问区的题
				let sumScore = 0
				for (const k in ask_list) {
				let ask_dom = $(`.ques-${ id }.ask-index-${ k }`)
				if (ask_dom) {
					let params = {
					id: id,
					type: 'ask',
					index: k,
					score: ask_dom.val()
					}
					sumScore += params.score * 1
					changeScore(params)
				}
				}
				// 题目的分数设置为小问的分数的和
				let params = {
				id: id,
				type: 'question',
				score: sumScore
				}
				changeScore(params)
			} else {
				// 没有小问区的题目
				let params = {
				id: id,
				type: 'question',
				score: dom.val()
				}
				changeScore(params)
			}
			}
		}
	}
	function onConfirm() {
		if (isLock) {
			return
		}
		setBtnLoading('confirm', true)
		isLock = true
		saveDataBeforeRender()
		// 将窗口的信息传递出去
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'updateScore',
			data: {
				question_map: question_map
			},
		})
		setBtnLoading('confirm', false)
		$('#confirm').html('保存成功')
		setTimeout(() => {
			$('#confirm').html('保存')
			isLock = false
		}, 1500)
	}
  
	function changeScore({id, type, index, score}) {
	  if (type == 'ask') {
		question_map[id].ask_list[index].score = score
	  } else {
		question_map[id].score = score
	  }
	}

	function onRefresh(e) {
		if (e) {
			e.cancelBubble = true
			e.preventDefault()
		}
		if (isLock) {
			return
		}
		isLock = true
		window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'RefreshPaper' })
	}
  
	window.Asc.plugin.attachEvent('initPaper', function (message) {
	  console.log('batchScore 接收的消息', message)
	  biyueCustomData = message.biyueCustomData || {}
	  if (tree_info) {
		removeEvents()
	  }
	  tree_info = message.tree_info || {}
	  init()
	  isLock = false
	})

	window.Asc.plugin.attachEvent('quesMapUpdate', function (message) {
		console.log('quesMapUpdate 接收的消息', message)
		var changed = false
		if (message.question_map != undefined) {
			if (!oldQuestionMap || isQuestionMapChanged(message.question_map)) {
				biyueCustomData.question_map = message.question_map;
				oldQuestionMap = _.cloneDeep(message.question_map);
				changed = true
			}
		}
		if (message.node_list != undefined) {
			biyueCustomData.node_list = message.node_list
		}
		if (message.tree_info !== undefined) {
			if (!oldTreeInfo || isTreeInfoChanged(message.tree_info)) {
				tree_info = message.tree_info;
				oldTreeInfo = _.cloneDeep(message.tree_info);
				changed = true
			}
		}
		if (changed) {
			console.log('====== 变化了')
			removeEvents()
			renderData()
		} else {
			console.log('====== 没有变化')
		}
		isLock = false
	})

	function isQuestionMapChanged(newQuestionMap) {
		const fields = ['ques_name', 'ques_default_name', 'score', 'ask_list', 'ques_mode']
		for (let key in newQuestionMap) {
			for (let field of fields) {
				if (!oldQuestionMap[key] || !_.isEqual(newQuestionMap[key][field], oldQuestionMap[key][field])) {
					return true
				}
			}
		}
		return false
	}
	function isTreeInfoChanged(newTreeInfo) {
		return !_.isEqual(newTreeInfo.list, oldTreeInfo.list)
	}
  })(window, undefined)
  
  