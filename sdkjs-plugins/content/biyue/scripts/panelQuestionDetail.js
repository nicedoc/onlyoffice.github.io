import ComponentSelect from '../components/Select.js'
import NumberInput from '../components/NumberInput.js'
import { changeProportion } from './ExamTree.js'
import { reqSaveQuestion } from './api/paper.js'
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
var ques_control_id
var select_type = null // 题型
var select_proportion = null // 占比
var input_score = null // 分数/权重
var list_ask = [] // 小问
var inited = false

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
          <td class="padding-small" width="100%">
            <label class="header">题号</label>
			<div id="ques_name" class="spinner" style="width: 100%">
      			<input type="text" class="form-control" spellcheck="false">
    		</div>
          </td>
        </tr>
        <tr>
          <td class="padding-small" colspan="1" width="100%">
            <label class="header">题型</label>
            <div id="questionType"></div>
          </td>
        </tr>
        <tr>
          <td class="padding-small" colspan="1" width="100%">
            <label class="header">占比</label>
            <div id="proportion"></div>
          </td>
        </tr>
        <tr>
          <td class="padding-small" width="100%">
            <label class="header">权重/分数</label>
            <div id="ques_weight"></div>
          </td>
        </tr>
      </tbody>
    </table>
    <div id="panelQuesAsks"></div>
  </div>
  `
	$('#panelQues').html(content)
	var paper_options = window.BiyueCustomData.paper_options || {}
	var questionTypes = []
	if (paper_options.question_type) {
		questionTypes = questionTypes.concat(paper_options.question_type)
	}
	questionTypes.unshift({ value: '0', label: '未定义' })
	console.log('============ questionTypes', questionTypes)
	select_type = new ComponentSelect({
		id: 'questionType',
		options: questionTypes,
		value_select: '0',
		callback_item: (data) => {
			changeQuestionType(data)
		},
		width: '100%',
	})
	select_proportion = new ComponentSelect({
		id: 'proportion',
		options: proportionTypes,
		value_select: '0',
		callback_item: (data) => {
			chagneProportion(data)
		},
		width: '100%',
	})
	input_score = new NumberInput('ques_weight', {
		width: '100%',
		change: (id, data) => {
			changeScore(id, data)
		},
	})
	$(`#ques_name input`).on('input', () => {
		changeQuesName($(`#ques_name input`).val())
	  })
	inited = true
}

function updateElements(quesData) {
	if (!inited) {
		initElements()
	}
	if (!quesData) {
		$('#panelQues .hint').show()
		$('#panelQuesWrapper').hide()
		return
	}
	$('#panelQues .hint').hide()
	$('#panelQuesWrapper').show()
	$(`#ques_name input`).val(quesData.ques_name || '')
	if (select_type) {
		select_type.setSelect((quesData.question_type || 0) + '')
	}
	if (select_proportion) {
		select_proportion.setSelect((quesData.proportion || 0) + '')
	}
	if (input_score) {
		input_score.setValue((quesData.score || 0) + '')
	}
	if (quesData.text) {
		$('#ques_text').html(quesData.text)
	}
	if (quesData.ask_list && quesData.ask_list.length > 0) {
		var content = ''
		content += '<label class="header">每空权重/分数</label><div class="asks">'
		quesData.ask_list.forEach((ask, index) => {
			content += `<div class="item"><span class="asklabel">(${
				index + 1
			})</span><div id="ask${index}"></div></div>`
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
					change: (id, data) => {
						changeScore(id, data)
					},
				})
				list_ask.push(askInput)
				askInput.setValue((ask.score || 0) + '')
			}
		})
		if (list_ask && list_ask.length > askcount) {
			for (var i = askcount; i < list_ask.length; i++) {
				delete list_ask[i]
			}
			list_ask.splice(askcount, list_ask.length - askcount)
		}
	} else {
		$('#panelQuesAsks').html(
			'<div style="text-align:center;color:#bbb">暂无小问</div>'
		)
		list_ask.forEach((e, index) => {
			delete list_ask[index]
		})
		list_ask = []
	}
}

function showQuesData(control_id, reginType) {
	console.log('showQuesData', control_id, reginType)
	ques_control_id = control_id || Asc.scope.controlId || Asc.scope.click_id
	if (!ques_control_id) {
		updateElements(null)
		return
	}
	var question_map = window.BiyueCustomData.question_map || {}
	if (!question_map) {
		updateElements(null)
		return
	}
	if (!question_map[ques_control_id]) {
		var keys = Object.keys(question_map)
		for (var i = 0; i < keys.length; ++i) {
			var ask_list = question_map[keys[i]].ask_list || []
			var find = ask_list.find(e => {
				return e.control_id == control_id
			})
			if (find) {
				ques_control_id = keys[i]
				break
			}
		}
	}
	if (!question_map[ques_control_id]) {
		updateElements(null)
		return
	}
	var quesData = question_map[ques_control_id]
	console.log('=========== showQuesData ques:', quesData)
	updateElements(quesData)
}

function changeQuestionType(data) {
	if (window.BiyueCustomData.question_map[ques_control_id]) {
		window.BiyueCustomData.question_map[ques_control_id].question_type = data.value * 1
		autoSave()
	}
}
// 修改占比
function chagneProportion(data) {
	console.log('chagneProportion', data)
	changeProportion(ques_control_id, data.value)
}

function initListener() {
	document.addEventListener('clickSingleQues', (event) => {
		if (!event || !event.detail) return
		console.log('receive clickSingleQues', event.detail)
		showQuesData(event.detail.control_id)
	})
	document.addEventListener('updateQuesData', (event) => {
		if (!event) return
		var quesData = window.BiyueCustomData.question_map[ques_control_id]
		if (event.detail && event.detail.field) {
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
		} else {
			updateElements(quesData)
		}
	})
}

function changeQuesName(data) {
	window.BiyueCustomData.question_map[ques_control_id].ques_name = data
	autoSave()
}

function changeScore(id, data) {
	if (id == 'ques_weight') {
		window.BiyueCustomData.question_map[ques_control_id].score = data
	} else {
		var index = id.replace('ask', '') * 1
		window.BiyueCustomData.question_map[ques_control_id].ask_list[index].score = data
	}
	autoSave()
}

function autoSave() {
	var quesData = window.BiyueCustomData.question_map[ques_control_id]
	if (!quesData || !quesData.uuid) {
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

export { showQuesData, initListener }
