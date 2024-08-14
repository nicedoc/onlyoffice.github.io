import ComponentSelect from '../components/Select.js'
import NumberInput from '../components/NumberInput.js'
import { reqSaveQuestion } from './api/paper.js'
import { setInteraction } from './featureManager.js'
import { changeProportion, deleteAsks, focusAsk, updateAllChoice, deleteChoiceOtherWrite, getQuesMode, updateQuesScore, splitControl } from './QuesManager.js'
import { addClickEvent, getListByMap, showCom } from '../scripts/model/util.js'
// 单题详情
var proportionTypes = [
	{ value: '1', label: '默认' },
	{ value: '2', label: '1/2' },
	{ value: '3', label: '1/3' },
	{ value: '4', label: '1/4' },
	{ value: '5', label: '1/5' },
	{ value: '6', label: '1/6' },
	{ value: '7', label: '1/7' },
	{ value: '8', label: '1/8' },
]
var interactionTypes = [
	{ 	value: 'none', label: '无互动'},
	{	value: 'simple', label: '简单互动'},
	{	value: 'accurate', label: '精准互动'}
]
var g_client_id
var g_ques_id
var select_type = null // 题型
var select_ques_mode = null // 作答模式
var select_proportion = null // 占比
var select_interaction = null // 互动
var select_mark_mode = null // 批改模式
var select_score_mode = null // 打分方式
var select_score_layout = null // 分数布局
var score_list = [] // 分数选项
var input_score = null // 分数/权重
var list_ask = [] // 小问
var inited = false
var timeout_save = null
var workbook_id = 0
var select_ask_index = -1

function initElements() {
	console.log('====================================================== panelquestiondetail initElements')
	var content = ''
	content = `
  <div class="hint" style="text-align:center">请先选中一道题</div>
  <div id="panelQuesWrapper">
    <div>题目：<span id="ques_text"></span></div>
    <table style="width: 100%">
      <tbody>
	  	<tr>
          <td class="padding-small" colspan="2" width="100%">
            <label class="header">题号</label>
			<div id="ques_name" class="spinner" style="width: 100%">
      			<input type="text" class="form-control" spellcheck="false">
    		</div>
          </td>
        </tr>
        <tr>
          <td class="padding-small" colspan="2" width="100%">
            <label class="header">题型</label>
            <div id="questionType"></div>
          </td>
        </tr>
		<tr>
          <td class="padding-small" colspan="2" width="100%">
            <label class="header">作答模式</label>
            <div id="questionMode"></div>
          </td>
        </tr>
        <tr>
          <td class="padding-small" colspan="1" width="50%">
            <label class="header">占比</label>
            <div id="proportion"></div>
          </td>
		  <td class="padding-small" colspan="1" width="50%">
            <label class="header">互动</label>
            <div id="quesInteraction"></div>
          </td>
        </tr>
        <tr>
          <td class="padding-small" colspan="2" width="100%">
            <label class="header">权重/分数</label>
            <div id="ques_weight"></div>
          </td>
        </tr>
		<tr id="markModeTr">
          <td class="padding-small" colspan="2" width="100%">
            <label class="header">批改模式</label>
            <div id="markMode"></div>
          </td>
        </tr>
		<tr id="scoreTr">
          <td class="padding-small" colspan="1" width="50%">
            <label class="header">打分模式</label>
            <div id="scoreMode"></div>
          </td>
		  <td class="padding-small" colspan="1" width="50%">
            <label class="header">分数布局</label>
            <div id="scoreLayout"></div>
          </td>
        </tr>
      </tbody>
    </table>
	<div id="scores">
		<div class="row-between">
			<div>分数选项</div>
			<div class="clicked under" id="cancelAllScore">取消全选</div>
		</div>
		<div id="scorelist"></div>
	</div>
    <div id="panelQuesAsks"></div>
	<div id="resplitQues" class="under clicked">重新切题</div>
  </div>
  `
	$('#panelQues').html(content)
	var paper_options = window.BiyueCustomData.paper_options || {}
	var questionTypes = []
	if (paper_options.question_type) {
		questionTypes = questionTypes.concat(paper_options.question_type)
	}
	questionTypes.unshift({ value: '0', label: '未定义' })
	select_type = new ComponentSelect({
		id: 'questionType',
		options: questionTypes,
		value_select: '0',
		callback_item: (data) => {
			changeQuestionType(data)
		},
		width: '100%',
	})
	var mark_type_info = Asc.scope.subject_mark_types
	select_ques_mode = new ComponentSelect({
		id: 'questionMode',
		options: mark_type_info ? getListByMap(mark_type_info.ques_mode_map) : [],
		value_select: '0',
		width: '100%',
		enabled: false,
	})
	select_proportion = new ComponentSelect({
		id: 'proportion',
		options: proportionTypes,
		value_select: '0',
		callback_item: (data) => {
			onChangeProportion(data)
		},
		width: '110px',
	})
	input_score = new NumberInput('ques_weight', {
		width: '100%',
		min: 0,
		change: (id, data) => {
      		let val = data
		  	val = checkInputValue(data, 100)
			changeScore(id, val)
		},
	})
	$(`#ques_name input`).on('input', () => {
		changeQuesName($(`#ques_name input`).val())
	  })
	$(`#ques_weight input`).on('input', () => {
    	let val = $(`#ques_weight input`).val()
		val = checkInputValue(val, 100)
    	$(`#ques_weight input`).val(val)
  	})
	select_interaction = new ComponentSelect({
		id: 'quesInteraction',
		options: interactionTypes,
		value_select: 'none',
		callback_item: (data) => {
			changeInteraction(data)
		},
		width: '110px',
	})
	select_mark_mode = new ComponentSelect({
		id: 'markMode',
		options: mark_type_info ? getListByMap(mark_type_info.mark_type_map) : [],
		value_select: 'none',
		callback_item: (data) => {
			changeMarkMode(data)
		},
		width: '100%',
	})
	select_score_mode = new ComponentSelect({
		id: 'scoreMode',
		options: [{
			value: 0,
			label: '自动选择'
		}, {
			value: 1,
			label: '普通模式'
		}, {
			value: 2,
			label: '大分值模式'
		}],
		value_select: 0,
		callback_item: (data) => {
			changeScoreMode(data)
		},
		width: '110px',
	})
	select_score_layout = new ComponentSelect({
		id: 'scoreLayout',
		options: [{
			value: 1,
			label: '顶部浮动'
		}, {
			value: 2,
			label: '嵌入式'
		}],
		value_select: 1,
		callback_item: (data) => {
			changeScoreLayout(data)
		},
		width: '110px',
	})
	addClickEvent('#cancelAllScore', cancelAllScore)
	addClickEvent('#resplitQues', resplitQues)
	inited = true
	workbook_id = window.BiyueCustomData.workbook_info ? window.BiyueCustomData.workbook_info.id : 0
}

function resetEvent() {
	if (list_ask) {
		for (var i = 0; i < list_ask.length; ++i) {
			var btnDelete = $(`#ask${i}_delete`)
			if (btnDelete) {
				btnDelete.off('click')
			}
		}
	}
}

function updateElements(quesData, hint, ignore_ask_list) {
	resetEvent()
	if (!inited) {
		initElements()
	}
	if (!quesData) {
		$('#panelQues .hint').show()
		$('#panelQuesWrapper').hide()
		$('#panelQues .hint').html(`${hint ? (hint + '，') : ''}请先选中一道题`)
		return
	}
	$('#panelQues .hint').hide()
	$('#panelQuesWrapper').show()
	$(`#ques_name input`).val(quesData.ques_name || quesData.ques_default_name || '')
	if (select_type) {
		select_type.setSelect((quesData.question_type || 0) + '')
	}
	updateQuesMode(quesData.ques_mode)
	handleMark(quesData)
	if (select_proportion) {
		select_proportion.setSelect((quesData.proportion || 1) + '')
		if (!quesData.proportion) {
			window.BiyueCustomData.question_map[g_ques_id].proportion = 1
		}
	}
	if (select_interaction) {
		select_interaction.setSelect(quesData.interaction || 'none')
	}
	if (input_score) {
		input_score.setValue((quesData.score || 0) + '')
	}
	if (quesData.text) {
		$('#ques_text').html(quesData.text)
		$('#ques_text').attr('title', quesData.text)
	}
	if (quesData.ask_list && quesData.ask_list.length > 0 && !ignore_ask_list) {
		var content = ''
		content += '<div class="row-between"><label class="header">每空权重/分数</label><div class="clicked under" id="clearAllAsks">删除所有小问</div></div>'
		content += '<div class="asks">'
		quesData.ask_list.forEach((ask, index) => {
			content += `<div class="item"><span id="asklabel${index}" class="asklabel">(${
				index + 1
			})</span><div id="ask${index}"></div><i class="iconfont icon-shanchu clicked" id="ask${index}_delete"></i></div>`
		})
		content += '</div>'
		$('#panelQuesAsks').html(content)
		var askcount = quesData.ask_list.length
		quesData.ask_list.forEach((ask, index) => {
			if (list_ask && index < list_ask.length) {
				list_ask[index].render()
				list_ask[index].setValue((ask.score || 0) + '')
			} else {
				var askInput = new NumberInput(`ask${index}`, {
					width: '60px',
					min: 0,
					change: (id, data) => {
            			let val = data
            			val = checkInputValue(data, 100)
						changeScore(id, data)
					},
					focus: (id) => {
						onFocusAsk(id, index)
					}
				})
				list_ask.push(askInput)
				askInput.setValue((ask.score || 0) + '')
			}
			$(`#ask${index}_delete`).on('click', () => {
				deleteAsk(index)
			})
      $(`#ask${index} input`).on('input', () => {
        let val = $(`#ask${index} input`).val()
        val = checkInputValue(val, 100)
        $(`#ask${index} input`).val(val)
      })
		})
		if (list_ask && list_ask.length > askcount) {
			for (var i = list_ask.length - 1; i >= askcount; --i) {
				delete list_ask[i]
			}
			list_ask.splice(askcount, list_ask.length - askcount)
		}
	} else {
		$('#panelQuesAsks').html(
			'<div style="text-align:center;color:#bbb">暂无小问</div>'
		)
		for (var j = list_ask.length - 1; j >= 0; --j) {
			delete list_ask[j]
		}
		list_ask = []
	}
	$('#clearAllAsks').on('click', () => {
		onClearAllAsks()
	})
}

function showQuesData(params) {
	console.log('showQuesData', params)
	if (!params || !params.client_id ) {
		updateElements(null)
		return
	}
	g_client_id = params.client_id
	var question_map = window.BiyueCustomData.question_map || {}
  let choice_display = window.BiyueCustomData.choice_display || {}
  let ignore_ask_list = false
	var quesData = question_map ? question_map[g_client_id] : null
	if (g_client_id) {
		var node_list = window.BiyueCustomData.node_list || []
		var nodeData = node_list.find(e => {
			return e.id == g_client_id
		})
		console.log('nodeData', nodeData)
		if (nodeData) {
			if (nodeData.level_type == 'struct') {
				updateElements(null, `当前选中为题组`)
				return
			} else if (nodeData.level_type == 'text') {
				updateElements(null, `当前选中为待处理文本`)
				return
			}
		}
	}
	if (!question_map) {
		updateElements(null)
		return
	}
	var ques_client_id = 0
	if (params.regionType == 'write') {
		var keys = Object.keys(question_map)
		for (var i = 0; i < keys.length; ++i) {
			var ask_list = question_map[keys[i]].ask_list || []
			var findIndex = ask_list.findIndex(e => {
				return e.id == params.client_id
			})
			if (findIndex >= 0) {
				ques_client_id = keys[i]
				select_ask_index = findIndex
				break
			}
		}
	} else if (params.regionType == 'question') {
		ques_client_id = params.client_id
	}
	if (!ques_client_id) {
		updateElements(null)
		return
	}
	quesData = question_map[ques_client_id]
	if (!quesData) {
		updateElements(null)
		return
	}
	console.log('=========== showQuesData ques:', quesData)
	g_ques_id = ques_client_id
	if (quesData.level_type == 'question') {
    	if (choice_display.style && choice_display.style === 'show_choice_region') {
      		// 如果是开启了集中作答区状态，则需要忽略对应题目的小问
      		let node = node_list.find(e => {
        		return e.id == ques_client_id
      		})
      		if (node.use_gather) {
        		ignore_ask_list = true
      		}
    	}
		updateElements(quesData, null, ignore_ask_list)
		updateAskSelect(findIndex)
	} else if (quesData.level_type == 'struct') {
		updateElements(null, `当前选中为题组：${quesData ? quesData.ques_default_name : ''}`)
		return
	} else {
		updateElements(null)
	}
}

function changeQuestionType(data) {
	if (window.BiyueCustomData.question_map[g_ques_id]) {
		var oldvalue = window.BiyueCustomData.question_map[g_ques_id].question_type
		window.BiyueCustomData.question_map[g_ques_id].question_type = data.value * 1
		var oldMode = getQuesMode(oldvalue)
		var quesMode = getQuesMode(data.value)
		updateQuesMode(quesMode)
		if (window.BiyueCustomData.question_map && window.BiyueCustomData.question_map[g_ques_id]) {
			window.BiyueCustomData.question_map[g_ques_id].ques_mode = quesMode
		}
		if (quesMode == 1 || quesMode == 5) {
			deleteChoiceOtherWrite([g_ques_id], false).then(() => {
				updateAllChoice().then(() => {
					autoSave()
					showQuesData({
						client_id: g_client_id,
						regionType: 'question'
					})
				})
			})
		} else if (oldMode == 1 || oldMode == 5) {
			updateAllChoice().then(() => {
				autoSave()
			})
		} else {
			autoSave()
		}
	}
}
// 修改占比
function onChangeProportion(data) {
	changeProportion([g_ques_id], data.value).then(() => {
		window.biyue.StoreCustomData()
	})
}
// 修改互动
function changeInteraction(data) {
	if (window.BiyueCustomData.question_map[g_ques_id]) {
		setInteraction(data.value, [g_ques_id]).then(() => {
			window.BiyueCustomData.question_map[g_ques_id].interaction = data.value
			window.biyue.StoreCustomData()
		})
	}
}

function initListener() {
	document.addEventListener('clickSingleQues', (event) => {
		if (!event || !event.detail) return
		console.log('receive clickSingleQues', event.detail)
		showQuesData(event.detail)
	})
	document.addEventListener('updateQuesData', (event) => {
		if (!event) return
		if (event.detail && event.detail.field) {
			var quesData = window.BiyueCustomData.question_map[g_ques_id]
			// 来源为批量设置，只针对单个field更新
			switch(event.detail.field) {
				case 'question_type':
					if (select_type) {
						select_type.setSelect(quesData.question_type + '')
					}
					break
				case 'proportion':
					if (select_proportion) {
						select_proportion.setSelect(quesData.proportion + '')
					}
					break
				case 'score':
					if (input_score) {
						input_score.setValue((quesData.score || 0) + '')
					}
					break
				default:
					break
			}
			updateElements(quesData)
		} else {
			var id = event.detail && event.detail.client_id ? event.detail.client_id : g_client_id
			var nodeData = window.BiyueCustomData.node_list.find(e => {
				return e.id == id
			})
			if (nodeData) {
				showQuesData({
					client_id: nodeData.id,
					regionType: nodeData.regionType
				})
			} else {
				updateElements(null)
			}
		}
	})
}

function changeQuesName(data) {
	window.BiyueCustomData.question_map[g_ques_id].ques_name = data
	autoSave()
}

function changeScore(id, data) {
	var v = data * 1
	if (isNaN(v)) {
		return
	}
	var ask_list = window.BiyueCustomData.question_map[g_ques_id].ask_list
	var sum = 0
	if (id == 'ques_weight') {
		window.BiyueCustomData.question_map[g_ques_id].score = v
		if (ask_list && ask_list.length) {
			var avg = (v / ask_list.length).toFixed(1) * 1
			for (var i = 0; i < ask_list.length; ++i) {
				if (i == ask_list.length - 1) {
					ask_list[i].score = v - sum
				} else {
					ask_list[i].score = avg
					sum += avg
          sum = sum.toFixed(1) * 1
				}
				if (list_ask && list_ask[i]) {
          let score = ask_list[i].score || 0
          score = score.toFixed(1) * 1
					list_ask[i].setValue(score + '')
				}
			}
		}
	} else {
		var index = id.replace('ask', '') * 1
		ask_list[index].score = v
		for (var i = 0; i < ask_list.length; ++i) {
			var askscore = ask_list[i].score * 1
			if (!isNaN(askscore)) {
				sum += askscore
        		sum = sum.toFixed(1) * 1
			}
		}
		window.BiyueCustomData.question_map[g_ques_id].score = sum
		if (input_score) {
			input_score.setValue(sum + '')
		}
	}
	window.BiyueCustomData.question_map[g_ques_id].ask_list = ask_list
	updateScores()
	autoSave(true)
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

function autoSave(updatescore) {
	var quesData = window.BiyueCustomData.question_map[g_ques_id]
	if (!quesData) {
		return
	}
	clearTimeout(timeout_save)
	timeout_save = setTimeout(() => {
		window.biyue.StoreCustomData()
		if (updatescore && !workbook_id) {
			updateQuesScore([g_ques_id]).then(() => {
				setInteraction('useself', [g_ques_id])
			})
		}
	}, 500)
	if (!quesData.uuid) {
		return
	}
	var scores = []
	var qname = quesData.ques_name || quesData.ques_default_name
	if (quesData.question_type != 6) {
		if (qname == '') {
			return
		}
		if (quesData.ask_list && quesData.ask_list.length > 1) {
			scores = quesData.ask_list.map(e => {
				return e.score * 1
			})
		} else {
			scores = [quesData.score * 1]
		}
		for (var i = 0; i < scores.length; ++i) {
			if (isNaN(scores[i]) || !scores[i] || scores[i] < 0) {
				return
			}
		}
	}
	reqSaveQuestion(window.BiyueCustomData.paper_uuid, quesData.uuid, quesData.question_type, qname, scores.join(',')).then(res => {
		console.log('保存单题成功')
	}).catch(err => {
		console.warn(err)
	})
}

function deleteAsk(index) {
	var quesData = window.BiyueCustomData.question_map[g_ques_id]
	deleteAsks([{
		ques_id: g_ques_id,
		ask_id: quesData.ask_list[index].id
	}])
}

function onFocusAsk(id, idx) {
	var index = id.replace('ask', '') * 1
	var quesData = window.BiyueCustomData.question_map[g_ques_id]
	if (!quesData || !quesData.ask_list || index >= quesData.ask_list.length) {
		return
	}
	var nodeData = window.BiyueCustomData.node_list.find(e => {
		return e.id == g_ques_id
	})
	if (nodeData && nodeData.write_list) {
		var writeData = nodeData.write_list.find(e => {
			return e.id == quesData.ask_list[index].id
		})
		focusAsk(writeData)
		updateAskSelect(idx)
	}
}

function updateQuesMode(ques_mode) {
	if (select_ques_mode) {
		select_ques_mode.setSelect(ques_mode + '')
	}
	return ques_mode
}

function changeMarkMode(data) {
	window.BiyueCustomData.question_map[g_ques_id].mark_mode = data.value
	updateScoreComponent()
	autoSave(true)
}

function updateScoreComponent() {
	var mark_mode = window.BiyueCustomData.question_map[g_ques_id].mark_mode
	if (!mark_mode || mark_mode == 1) {
		showCom('#scoreTr', false)
		showCom('#scores', false)
		showCom('#panelQuesAsks', true)
	} else if (mark_mode == 2) {
		var score = window.BiyueCustomData.question_map[g_ques_id].score * 1
		if (!score || isNaN(score)) {
			showCom('#scoreTr', false)
			showCom('#scores', false)
		} else {
			showCom('#scoreTr', true)
			showCom('#scores', true)
		}
		showCom('#panelQuesAsks', false)
		updateScores()
	}
	autoSave(true)
}

function changeScoreMode(data) {
	window.BiyueCustomData.question_map[g_ques_id].score_mode = data.value
	updateScores()
	autoSave(true)
}

function changeScoreLayout(data) {
	window.BiyueCustomData.question_map[g_ques_id].score_layout = data.value
	autoSave(true)
}

function removeScoreEvent() {
	var doms = document.querySelectorAll('.score-item') || []
	doms.forEach(function(dom) {
		dom.removeEventListener('click', clickScore)
	})
}

function clickScore(e) {
	if (e && e.target && e.target.dataset && e.target.dataset.idx) {
		e.target.classList.toggle('selected')
		window.BiyueCustomData.question_map[g_ques_id].score_list[e.target.dataset.idx].use = !window.BiyueCustomData.question_map[g_ques_id].score_list[e.target.dataset.idx].use
	}
	autoSave(true)
}

function updateScores() {
	removeScoreEvent()
	var mark_mode = window.BiyueCustomData.question_map[g_ques_id].mark_mode
	if (!mark_mode || mark_mode == 1) {
		showCom('#scores', false)
		return
	}
	var score = window.BiyueCustomData.question_map[g_ques_id].score * 1
	var score_mode = window.BiyueCustomData.question_map[g_ques_id].score_mode || 0
	var score_mode_use = score_mode
	if (score_mode == 0) {
		score_mode_use = score >= 15 ? 2 : 1
	}
	if (score_mode_use == 2) {
		window.BiyueCustomData.question_map[g_ques_id].score_list = []
		showCom('#scores', false)
		return
	} else {
		showCom('#scores', true)
	}
	// 分数选项
	var list = window.BiyueCustomData.question_map[g_ques_id].score_list || []
	var count = 1
	if (score >= 1) {
		count += Math.floor(score - 1)
	}
	if (list.length == 0 || list.length != count) {
		list = []
		for (var i = 1; i < score; ++i) {
			list.push({
				value: i,
				use: true
			})
		}
		list.push({
			value: 0.5,
			use: true
		})
	}
	var com = $('#scorelist')
	var str = ''
	for (var i = 0; i < list.length; ++i) {
		str += `<div class="clicked score-item ${list[i].use ? 'selected' : ''}" id="s${i}" data-value="${list[i].value}" data-idx="${i}">${list[i].value}</div>`
	}
	window.BiyueCustomData.question_map[g_ques_id].score_list = list
	com.html(str)
	var doms = document.querySelectorAll('.score-item') || []
	doms.forEach(function(dom) {
		dom.addEventListener('click', clickScore)
	})
}

function cancelAllScore() {
	var coms = document.querySelectorAll('.score-item')
	if (coms && coms.length) {
		for (var i = 0; i < coms.length; ++i) {
			coms[i].classList.remove('selected')
		}
		var list = window.BiyueCustomData.question_map[g_ques_id].score_list
		score_list.forEach(e => {
			e.use = false
		})
		window.BiyueCustomData.question_map[g_ques_id].score_list = list
		autoSave(true)
	}
}

function handleMark(quesData) {
	if (workbook_id) {
		showCom('#markModeTr', false)
		showCom('#scoreTr', false)
		showCom('#scores', false)
		return
	}
	if (!quesData) {
		showCom('#markModeTr', false)
		showCom('#scoreTr', false)
		showCom('#scores', false)
		return
	}
	if (quesData.ques_mode == 6) { // 文本
		showCom('#markModeTr', false)
		showCom('#scoreTr', false)
		showCom('#scores', false)
		showCom('#panelQuesAsks', false)
		return
	}
	showCom('#markModeTr', true)
	var markMode = quesData.mark_mode || 0
	if (markMode == 0) {
		if (quesData.ques_mode == 3 || quesData.ques_mode == 8) {
			markMode = 2
		} else {
			markMode = 1
		}
	}
	if (select_mark_mode) {
		select_mark_mode.setSelect(markMode + '')
		quesData.mark_mode = markMode
	}
	updateScoreComponent()
}

function updateAskSelect(index) {
	if (select_ask_index >= 0) {
		var old = $(`#asklabel${select_ask_index}`)
		if (old) {
			old.removeClass('ask-selected')
		}
	}
	select_ask_index = index
	if (select_ask_index >= 0) {
		var newel = $(`#asklabel${select_ask_index}`)
		if (newel) {
			newel.addClass('ask-selected')
		}
	}
}

function onClearAllAsks() {
	var quesData = window.BiyueCustomData.question_map[g_ques_id]
	if (quesData.ask_list) {
		var list = quesData.ask_list.map(e => {
			return {
				ques_id: g_ques_id,
				ask_id: e.id
			}
		})
		deleteAsks(list)
	}
}

function resplitQues() {
	splitControl(g_ques_id)
}

export { showQuesData, initListener }
