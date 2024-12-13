import { setBtnLoading } from '../model/util.js'
var _ = window._;
;(function (window, undefined) {
	let biyueCustomData = {}
	let mark_type_info = {}
	let question_map = {}
	let questionList = []
	let tree_map = {}
	let question_type_options = []
	let hidden_empty_struct = false
	let needUpdateInteraction = false
	let needUpdateChoice = false
	let tree_info = {}
	let display_tree = true
	let isLock = false
	let oldQuestionMap;
	let oldTreeInfo;
	window.Asc.plugin.init = function () {
	  window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'PaperMessage' })
	}
  
	function init() {
	  needUpdateChoice = false
	  needUpdateInteraction = false
	  getOptions()
	  $('#refresh').on('click', onRefresh)
	  $('#confirm').on('click', onConfirm)
	  $('#hidden_empty_struct').on('click', onSwitchStruct)
	  $('#switch_tree').prop('checked', display_tree)
	  $('#switch_tree').on('click', onSwitchTree)
	}
  
	function getOptions() {
	  question_type_options = []
	  if (biyueCustomData.paper_options) {
		question_type_options = biyueCustomData.paper_options.question_type || []
	  }
	  renderData()
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
	  getQuestionTypeNum()
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
		  html += `<select class="type-item" style="width: 80px" id="bat-group-${ id }" data-qid=${id}>`
	  } else {
		  html += `<select class="type-item ques-${ id }" data-qid=${id}>`
	  }
	  html += `<option value="" style="display:none;"></option>`
	  for (const key in question_type_options) {
		  if (quesData.level_type == 'struct') {
			  html += `<option value="${question_type_options[key].value}" title="${question_type_options[key].label}">${question_type_options[key].label}</option>`
		  } else {
			  let selected = (quesData.level_type == 'question' && quesData.question_type * 1 === question_type_options[key].value * 1) ? 'selected' : ''
			  html += `<option value="${question_type_options[key].value}" ${selected}>${question_type_options[key].label}</option>`
		  }
	  }
	  html += `</select>`
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
	  const div = $(`<div data-qid=${item.id}></div>`)
	  var html = ''
	  if (item.level_type == 'struct') {
		var structName = quesData.ques_name || quesData.ques_default_name || quesData.text.split('\r\n')[0]
		  html += `<div class="group row-between" id="group-id-${item.id}"><div class="text-over-ellipsis" style="display: block;max-width: 60px;" title="${quesData.text}">${structName}</div>`
		  if (item.children && item.children.length) {
			  html += '<div class="bat-set">'
			  html += getSelectHtml(item.id, quesData)
			  html += `<span class="bat-set-btn" id="bat-set-btn-${ item.id }" data-id="${ item.id }">设置</span></div></div>`
		  }
		  html += '</div>'
		  tree_map[item.id] = []
	  } else if (item.level_type == 'question') {
		div.addClass('row-align-center question')
		  html += `<div class="text-over-ellipsis" style="display: block;max-width: 90px;" title="${quesData.text}">${quesData.ques_name || quesData.ques_default_name || ''}</div>`
		  html += getSelectHtml(item.id, quesData)
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
  
	function renderData() {
	  renderTree()
	}
  
	function addEvents() {
	  // 处理题目的下拉选项事件
	  for (const key in questionList) {
		  let id = questionList[key]
		  let doms = document.querySelectorAll('.ques-' + id) || []
		  doms.forEach(function(dom) {
			  // 提前定义事件处理函数
			  const changeHandler = function() {
				  getQuestionTypeNum();
			  };
			  dom.removeEventListener('change', changeHandler);
			  dom.addEventListener('change', changeHandler)
			  dom.addEventListener('click', function(e) {
				focusItem(e)
			  })
		  })
	  }
	  for (const key in tree_map) {
		  if (tree_map[key].length > 0) {
			  // 对有题的结构增加批量设置题型的下拉框
			  let dom = document.querySelector('#group-id-' + key)
			  if (dom) {
				  let btnDom = document.querySelector(`#bat-set-btn-${ key }`)
				  if (btnDom) {
					  const clickHandler = function() {
						  let id = btnDom.dataset.id || ''
						  let inputDom = document.querySelector(`#bat-group-${ id }`)
						  if (inputDom.value > 0) {
							  batchSetStructQuestionType(id, inputDom.value || 0)
							  getQuestionTypeNum()
						  }
					  };
					  btnDom.removeEventListener('click', clickHandler)
					  btnDom.addEventListener('click', clickHandler)
				  }
				  
			  }
		  }
	  }
	  getQuestionTypeNum()
	}
  
	function batchSetStructQuestionType(struct_id, value) {
	  // 根据结构批量写入题目类型
	  let arr = tree_map[struct_id] || []
	  for (const key in arr) {
		  let id = arr[key]
		  var oldMode = question_map[id].ques_mode
		  var ques_mode = getQuesMode(parseFloat(value) || 0)
		  question_map[id].ques_mode = ques_mode
		  question_map[id].question_type = value
		  if (!needUpdateInteraction) {
			  needUpdateInteraction = question_map[id].interaction != 'none' && (oldMode == 6 || ques_mode == 6) 
		  }
		  if (!needUpdateChoice) {
			  needUpdateChoice = oldMode == 1 || oldMode == 5 || ques_mode == 1 || ques_mode == 5
		  }
	  }
  
	  renderData()
	}
  
	function getQuestionTypeNum() {
	  let question_type_map = {}
	  for (const key in question_type_options) {
		if (question_type_options[key]) {
		  let item = question_type_options[key] || {}
		  question_type_map[item.value] = {
			label: item.label,
			number: 0
		  }
		}
		question_type_map[null] = {
		  label: '未知题型',
		  number: 0
		}
	  }
	  for (const key in questionList) {
		let id = questionList[key]
		let doms = document.querySelectorAll('.ques-' + id) || []
		doms.forEach(function(dom) {
		  let val = dom.value || ''
		  if (question_type_map[val]) {
			question_type_map[val].number += 1
		  } else {
			question_type_map[null].number += 1
		  }
		})
	  }
	  let html = ''
	  for (const key in question_type_map) {
		  if (question_type_map[key].number > 0) {
			html += `<span style="margin: 0 8px;">${question_type_map[key].label}${question_type_map[key].number}题</span>`
		  }
	  }
	  $('#getQuestionTypeNum').html(html)
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
  
	// 获取作答模式
	function getQuesMode(question_type) {
	  question_type = question_type * 1
	  var ques_mode = 0
	  if (question_type == 6) {
		  return 6
	  }
	  if (question_type > 0) {
		if (mark_type_info && mark_type_info.list) {
		  var find = mark_type_info.list.find(e => {
			return e.question_type_id == question_type
		  })
		  return find ? find.ques_mode : 3
		}
	  }
	  return ques_mode
	}
  
	function onConfirm() {
		if (isLock) {
			return
		}
		setBtnLoading('confirm', true)
		isLock = true
	  for (const key in questionList) {
		let id = questionList[key] || ''
		let dom = $('.ques-'+id)
		if (id && question_map[id] && dom) {
		  let params = {
			id: id,
			question_type: dom.val()
		  }
		  changeQuestionType(params)
		}
	  }
	  // 将窗口的信息传递出去
	  window.Asc.plugin.sendToPlugin('onWindowMessage', {
		type: 'changeQuestionMap',
		data: {
		  question_map: question_map,
		  needUpdateChoice: needUpdateChoice,
		  needUpdateInteraction: needUpdateInteraction
		},
	  })
	  setBtnLoading('confirm', false)
		$('#confirm').html('保存成功')
		setTimeout(() => {
			$('#confirm').html('保存')
			isLock = false
		}, 1500)
	}
  
	function changeQuestionType({id, question_type}) {
	  var oldMode = question_map[id].ques_mode
	  question_map[id].question_type = parseFloat(question_type) || 0
	  var ques_mode = getQuesMode(parseFloat(question_type) || 0)
	  question_map[id].ques_mode = ques_mode
	  if (!needUpdateInteraction) {
		  needUpdateInteraction = question_map[id].interaction != 'none' && (oldMode == 6 || ques_mode == 6) 
	  }
	  if (!needUpdateChoice) {
		  needUpdateChoice = oldMode == 1 || oldMode == 5 || ques_mode == 1 || ques_mode == 5
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

	function focusItem(e) {
		if (!e) {
			return
		}
		e.cancelBubble = true
		e.preventDefault()
		var target = e.target || e.srcElement
		if (!target || !target.dataset) {
			return
		}
		if (target.dataset.qid) {
			window.Asc.plugin.sendToPlugin('onWindowMessage', {
				type: 'focusQuestion',
				data: {
				  	ques_id: target.dataset.qid
				},
			})
		}
	}
  
	window.Asc.plugin.attachEvent('initPaper', function (message) {
	  console.log('batchSettingQuestionType 接收的消息', message)
	  biyueCustomData = message.biyueCustomData || {}
	  mark_type_info = message.subject_mark_types || {}
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
		if (message.tree_info !== undefined) {
			if (!oldTreeInfo || isTreeInfoChanged(message.tree_info)) {
				tree_info = message.tree_info;
				oldTreeInfo = _.cloneDeep(message.tree_info);
				changed = true
			}
		}
		if (changed) {
			console.log('====== 变化了')
			renderData()
		} else {
			console.log('====== 没有变化')
		}
		isLock = false
	})

	function isQuestionMapChanged(newQuestionMap) {
		const fields = ['ques_name', 'ques_default_name', 'question_type']
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
  