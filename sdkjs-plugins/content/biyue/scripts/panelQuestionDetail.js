import ComponentSelect from '../components/Select.js'
import NumberInput from '../components/NumberInput.js'
import { reqSaveQuestion } from './api/paper.js'
import { setInteraction } from './featureManager.js'
import { changeProportion, deleteAsks, focusAsk, updateAllChoice, deleteChoiceOtherWrite, getQuesMode, updateQuesScore, splitControl, setChoiceOptionLayout } from './QuesManager.js'
import { addClickEvent, getListByMap, showCom } from '../scripts/model/util.js'
// 单题详情
var proportionTypes = [
	{ value: 1, label: '默认' },
	{ value: 2, label: '1/2' },
	{ value: 3, label: '1/3' },
	{ value: 4, label: '1/4' },
	{ value: 5, label: '1/5' },
	{ value: 6, label: '1/6' },
	{ value: 7, label: '1/7' },
	{ value: 8, label: '1/8' },
]
var interactionTypes = [
	{ value: 'none', label: '无互动'},
	{ value: 'simple', label: '简单互动'},
	{ value: 'accurate', label: '精准互动'}
]
var scoreModes = [
	{ value: 0, label: '自动选择'},
	{ value: 1, label: '普通模式'},
	{ value: 2, label: '大分值模式'}
]
var scoreLayoutTypes = [
	{ value: 1,	label: '顶部浮动'},
	{ value: 2,	label: '嵌入式'}
]
var choiceAlignTypes = [
	{ value: 0, label: '未设置'},
	{ value: 4, label: '一行4个'},
	{ value: 2, label: '一行2个'},
	{ value: 1, label: '一行1个'},
	{ value: 3, label: '一行3个'},
	{ value: 5, label: '一行5个'},
	{ value: 6, label: '一行6个'},
]
var choiceIndLeftTypes = [
	{ value: -1, label: '未设置'},
	{ value: 0, label: '无缩进'},
	{ value: 5, label: '0.5厘米'},
	{ value: 10, label: '1厘米'},
	{ value: 15, label: '1.5厘米'},
	{ value: 20, label: '2厘米'},
]
var choiceBracketTypes = [
	{ value: 'none', label: '未设置'},
	{ value: 'eng', label: '英文括号'},
	{ value: 'ch', label: '中文括号'},
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
var select_choice_align = null // 选择题对齐选项
var select_choice_ind_left = null // 选择题自动缩进
var select_choice_bracket = null // 选择题括号样式
var input_choice_space = null // 选择题括号空格数量
var score_list = [] // 分数选项
var input_score = null // 分数/权重
var list_ask = [] // 小问
var inited = false
var timeout_save = null
var workbook_id = 0
var select_ask_index = -1
var choice_check = false

function getTdText(colspan, width, title, id) {
	return `<td class="padding-small" colspan=${colspan} width=${width}>
				<label class="header">${title}</label>
				<div id=${id}></div>
			</td>`
}
function initElements() {
	console.log('====================================================== panelquestiondetail initElements')
	var content = ''
	content = `
  	<div class="hint" style="text-align:center">请先选中一道题</div>
  	<div id="panelQuesWrapper">
    	<div><span id="texttitle">题目：</span><span id="ques_text"></span></div>
    	<table style="width: 100%">
      		<tbody>
	  			<tr>
          			<td class="padding-small" colspan="2" width="100%">
            			<label class="header" id="nametitle">题号</label>
						<div id="ques_name" class="spinner" style="width: 100%">
      						<input type="text" class="form-control" spellcheck="false">
    					</div>
          			</td>
        		</tr>
        		<tr id="quesTypeTr">
					${getTdText("2", "100%", '题型', "questionType")}
        		</tr>
				<tr id="quesModeTr">
					${getTdText("2", "100%", '作答模式', "questionMode")}
        		</tr>
				<tr id="proportionTr">
					${getTdText("1", "50%", '占比', "proportion")}
					${getTdText("1", "50%", '互动', "quesInteraction")}
				</tr>
				<tr id="weightTr">
					${getTdText("2", "100%", '权重/分数', "ques_weight")}
				</tr>
				<tr id="markModeTr">
					${getTdText("2", "100%", '批改模式', "markMode")}
				</tr>
				<tr id="scoreTr">
					${getTdText("1", "50%", '打分模式', "scoreMode")}
					${getTdText("1", "50%", '分数布局', "scoreLayout")}
				</tr>
				<tr class="choicetr" style="border-top:1px solid #bbb;">
					<td colspan="2" style="padding-top:8px">
						<label>选择题设置项</label>
					</td>
				</tr>
				<tr class="choicetr">
					${getTdText("1", "50%", '自动对齐', "choiceAlign")}
					${getTdText("1", "50%", '自动缩进', "choiceIndLeft")}
				</tr>
				<tr class="choicetr">
					${getTdText("1", "50%", '括号样式', "choiceBracket")}
					${getTdText("1", "50%", '括号空格个数', "bracketSpace")}
				</tr>
				<tr class="choicetr">
					<td colspan="2" style="padding-bottom:8px">
						<div class="row-arround" style="padding-bottom:8px;border-bottom:1px solid #bbb;">
							<div id="applyToQues" class="btn2 clicked">只应用到本题</div>
							<div id="applyToAllQues" class="btn2 clicked">应用到同模式题目</div>
						</div>
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
	select_type = createSelect('questionType', questionTypes, '0', changeQuestionType, '100%')
	var mark_type_info = Asc.scope.subject_mark_types
	select_ques_mode = new ComponentSelect({
		id: 'questionMode',
		options: mark_type_info ? getListByMap(mark_type_info.ques_mode_map) : [],
		value_select: '0',
		width: '100%',
		enabled: false,
	})
	select_proportion = createSelect('proportion', proportionTypes, '0', onChangeProportion)
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
	select_interaction = createSelect('quesInteraction', interactionTypes, 'none', changeInteraction)
	var markModes = mark_type_info ? getListByMap(mark_type_info.mark_type_map) : []
	select_mark_mode = createSelect('markMode', markModes, 'none', changeMarkMode, '100%')
	select_score_mode = createSelect('scoreMode', scoreModes, 0, changeScoreMode)
	select_score_layout = createSelect('scoreLayout', scoreLayoutTypes, 1, changeScoreLayout)
	select_choice_align = createSelect('choiceAlign', choiceAlignTypes, '0', null)
	select_choice_ind_left = createSelect('choiceIndLeft', choiceIndLeftTypes, '-1', null)
	select_choice_bracket = createSelect('choiceBracket', choiceBracketTypes, 'none', null)
	input_choice_space = new NumberInput('bracketSpace', {
		width: '110px',
		min: 0,
		change: (id, data) => {
      		let val = data
		  	val = checkInputValueNumber(data, 20)
			changeBracketSpace(id, val)
		},
	})
	addClickEvent('#cancelAllScore', cancelAllScore)
	addClickEvent('#resplitQues', resplitQues)
	addClickEvent('#applyToQues', () => {onApplyAllQues(false)})
	addClickEvent('#applyToAllQues', () => {onApplyAllQues(true)})
	inited = true
	workbook_id = window.BiyueCustomData.workbook_info ? window.BiyueCustomData.workbook_info.id : 0
}

function createSelect(id, options, vSelect, callback, width = '110px') {
	return new ComponentSelect({
		id: id,
		options: options,
		value_select: vSelect,
		callback_item: callback,
		width: width,
	})
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

function showQuesCom(v) {
	showCom('#quesTypeTr', v)
	showCom('#quesModeTr', v)
	showCom('#proportionTr', v)
	showCom('#weightTr', v)
	showCom('#markModeTr', v)
	showCom('#scoreTr', v)
	showCom('#scores', v)
	showCom('#panelQuesAsks', v)
	showCom('#resplitQues', v)
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
	if (quesData.level_type == 'struct') {
		$('#panelQues .hint').hide()
		$('#panelQuesWrapper').show()
		showQuesCom(false)
		showCom('.choicetr', false)
		if (quesData.text) {
			$('#texttitle').html('题组：')
			$('#nametitle').html('题组名')
			$('#ques_text').html(quesData.text)
			$('#ques_text').attr('title', quesData.text)
		}
		$(`#ques_name input`).val(quesData.ques_name || quesData.ques_default_name || '')
		return
	}
	$('#texttitle').html('题目：')
	$('#nametitle').html('题号')
	$('#panelQues .hint').hide()
	$('#panelQuesWrapper').show()
	showQuesCom(true)
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
		content += '<div class="row-between"><label class="header" style="margin-bottom:0px">每空权重/分数</label><div class="clicked under" id="clearAllAsks">删除所有小问</div></div>'
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
				setQueId(g_client_id)
				updateElements(question_map[nodeData.id], `当前选中为题组`)
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
				if (e.other_fileds && e.other_fileds.includes(g_client_id)) {
					return true
				}
				return e.id == g_client_id
			})
			if (findIndex >= 0) {
				ques_client_id = keys[i]
				select_ask_index = findIndex
				break
			}
		}
	} else if (quesData && quesData.level_type == 'question') {
		ques_client_id = g_client_id
	} else if (params.regionType == 'question') {
		ques_client_id = g_client_id
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
	setQueId(ques_client_id)
	if (quesData.level_type == 'question') {
    	if ((quesData.ques_mode == 1 || quesData.ques_mode == 5) && choice_display.style && choice_display.style === 'show_choice_region') {
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
		setQueId(ques_client_id)
		updateElements(quesData, `当前选中为题组：${quesData ? quesData.ques_default_name : ''}`)
		return
	} else {
		updateElements(null)
	}
}

function setQueId(id) {
	if (id != g_ques_id) {
		choice_check = false
	}
	g_ques_id = id
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
		var interaction = window.BiyueCustomData.question_map[g_ques_id].interaction
		var need_update_interaction = false
		if (interaction != 'none') {
			if (oldMode == 6 || quesMode == 6) {
				need_update_interaction = true
			}
		}
		if (need_update_interaction) {
			setInteraction('useself', [g_ques_id]).then(() => {
				updateQuesType(quesMode, oldMode)
			})
		} else {
			updateQuesType(quesMode, oldMode)
		}
	}
}

function updateQuesType(quesMode, oldMode) {
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
	if (window.BiyueCustomData.question_map[g_ques_id]) {
		window.BiyueCustomData.question_map[g_ques_id].ques_name = data
	}
	autoSave()
}

function changeScore(id, data) {
	var v = data * 1
	if (isNaN(v) || !window.BiyueCustomData.question_map[g_ques_id]) {
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

function checkInputValueNumber(val = '', max) {
	// 只允许数字，且不允许点
	var sanitizedValue = val.replace(/[^\d]/g, '');
	if (max) {
		sanitizedValue = sanitizedValue > max ? max : sanitizedValue
	}
	input_choice_space.setValue(sanitizedValue)
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
	if (!quesData || !quesData.ask_list || index >= quesData.ask_list.length) {
		return
	}
	var askdata = quesData.ask_list[index]
	var list = [{
		ques_id: g_ques_id,
		ask_id: askdata.id
	}]
	if (askdata.other_fileds) {
		askdata.other_fileds.forEach(e => {
			list.push({
				ques_id: g_ques_id,
				ask_id: e
			})
		})
	}
	deleteAsks(list)
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
		if (writeData) {
			focusAsk(writeData)
			updateAskSelect(idx)
		} else {
			console.log('找不到小问数据,node_list和question_map里的数据不一致')
		}
		
	}
}

function updateQuesMode(ques_mode) {
	if (select_ques_mode) {
		select_ques_mode.setSelect(ques_mode + '')
	}
	showCom('.choicetr', ques_mode == 1 || ques_mode == 5)
	return ques_mode
}

function changeMarkMode(data) {
	if (!window.BiyueCustomData.question_map[g_ques_id]) {
		return
	}
	window.BiyueCustomData.question_map[g_ques_id].mark_mode = data.value
	updateScoreComponent()
	autoSave(true)
}

function updateScoreComponent() {
	if (!window.BiyueCustomData.question_map[g_ques_id]) {
		return
	}
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
	if (!window.BiyueCustomData.question_map[g_ques_id]) {
		return
	}
	window.BiyueCustomData.question_map[g_ques_id].score_mode = data.value
	updateScores()
	autoSave(true)
}

function changeScoreLayout(data) {
	if (!window.BiyueCustomData.question_map[g_ques_id]) {
		return
	}
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
	if (!window.BiyueCustomData.question_map[g_ques_id] || !window.BiyueCustomData.question_map[g_ques_id].score_list) {
		return
	}
	if (e && e.target && e.target.dataset && e.target.dataset.idx) {
		e.target.classList.toggle('selected')
		window.BiyueCustomData.question_map[g_ques_id].score_list[e.target.dataset.idx].use = !window.BiyueCustomData.question_map[g_ques_id].score_list[e.target.dataset.idx].use
	}
	autoSave(true)
}

function updateScores() {
	removeScoreEvent()
	if (!window.BiyueCustomData.question_map[g_ques_id]) {
		return
	}
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
	if (!window.BiyueCustomData.question_map[g_ques_id]) {
		return
	}
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
	if (quesData) {
		deleteAsks([{
			ques_id: g_ques_id,
			ask_id: 0
		}])
	}
}

function resplitQues() {
	deleteAsks([{
		ques_id: g_ques_id,
		ask_id: 0
	}], false, false).then(() => {
		splitControl(g_ques_id).then(res => {
			if (window.BiyueCustomData.question_map[g_ques_id] && window.BiyueCustomData.question_map[g_ques_id].interaction == 'accurate') {
				setInteraction(window.BiyueCustomData.question_map[g_ques_id].interaction, [g_ques_id])
			}
		})
	})
}

function onApplyAllQues(forAll) {
	choice_check = forAll
	notifyChoiceAlign('all')
}

function notifyChoiceAlign(from) {
	var question_map = window.BiyueCustomData.question_map || {}
	var list = []
	if (choice_check && question_map[g_ques_id]) {
		var qmode = question_map[g_ques_id].ques_mode
		Object.keys(question_map).forEach(id => {
			if (question_map[id].ques_mode == qmode) {
				list.push(id * 1)
			}
		})
	} else {
		list = [g_ques_id * 1]
	}
	var spaceNum = 0
	if (input_choice_space) {
		spaceNum = input_choice_space.getValue()
		if (spaceNum == '') {
			spaceNum = 0
		} else {
			spaceNum *= 1
		}
	}
	setChoiceOptionLayout({
		list: list,
		part: select_choice_align ? select_choice_align.getValue() * 1 : 0,
		indLeft: select_choice_ind_left ? select_choice_ind_left.getValue() * 1 : -1,
		bracket: select_choice_bracket ? select_choice_bracket.getValue() : 'none',
		spaceNum: spaceNum,
		from: from
	})
}

function changeBracketSpace(id, val) {
	console.log('changeBracketSpace', id, val)
}

export { showQuesData, initListener }
