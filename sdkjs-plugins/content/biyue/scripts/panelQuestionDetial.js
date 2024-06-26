import ComponentSelect from '../components/Select.js'
import NumberInput from '../components/NumberInput.js'
// 单题详情
var questionTypes = [
	{ value: '0', label: '未定义' },
	{ value: '1', label: '单选' },
	{ value: '2', label: '填空' },
	{ value: '3', label: '作答' },
	{ value: '4', label: '判断' },
	{ value: '5', label: '多选' },
	{ value: '6', label: '文本' },
	{ value: '7', label: '单选组合' },
	{ value: '8', label: '作文' },
]
var proportionTypes = [
	{ value: '0', label: '默认' },
	{ value: '1', label: '1/2' },
	{ value: '2', label: '1/3' },
	{ value: '3', label: '1/4' },
	{ value: '4', label: '1/5' },
	{ value: '5', label: '1/6' },
	{ value: '6', label: '1/7' },
	{ value: '7', label: '1/8' },
]
var ques_control_id
var ques_control_index
var select_type = null // 题型
var select_proportion = null // 占比
var input_score = null // 分数/权重
var list_ask = [] // 小问
var inited = false

function initElements() {
	var content = ''
	content = `
  <div class="hint" style="text-align:center">请先选中一道题</div>
  <div id="panelQuesWrapper">
    <div>题目：<span id="ques_name"></span></div>
    <table>
      <tbody>
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
	inited = true
}

function updateElements(controlData) {
	if (!inited) {
		initElements()
	}
	if (!controlData) {
		$('#panelQues .hint').show()
		$('#panelQuesWrapper').hide()
		return
	}
	$('#panelQues .hint').hide()
	$('#panelQuesWrapper').show()
	if (select_type) {
		select_type.setSelect((controlData.ques_type || 0) + '')
	}
	if (select_proportion) {
		select_proportion.setSelect((controlData.proportion || 0) + '')
	}
	if (input_score) {
		input_score.setValue((controlData.score || 0) + '')
	}
	if (controlData.ques_name) {
		$('#ques_name').html(controlData.ques_name)
	}
	if (controlData.ask_controls && controlData.ask_controls.length > 0) {
		var content = ''
		content += '<label class="header">每空权重/分数</label><div class="asks">'
		controlData.ask_controls.forEach((ask, index) => {
			content += `<div class="item"><span class="asklabel">(${
				index + 1
			})</span><div id="ask${index}"></div></div>`
		})
		content += '</div>'
		$('#panelQuesAsks').html(content)
		var askcount = controlData.ask_controls.length
		controlData.ask_controls.forEach((ask, index) => {
			if (list_ask && index < list_ask.length) {
				list_ask[index].setValue(ask.score || 0)
				list_ask[index].render()
			} else {
				var askInput = new NumberInput(`ask${index}`, {
					width: '60px',
					change: (id, data) => {
						changeScore(id, data)
					},
				})
				list_ask.push(askInput)
				askInput.setValue(ask.score || 0)
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
	ques_control_id = control_id
	var control_list = window.BiyueCustomData.control_list || []
	var control_index = -1
	if (control_id) {
		var control_index = control_list.findIndex((e) => {
			return e.control_id == control_id
		})
	}
	if (control_index >= 0) {
		if (reginType == 'sub-question' || reginType == 'write') {
			var quesCtrlId = control_list[control_index].parent_ques_control_id
			control_index = control_list.findIndex((e) => {
				return e.control_id == quesCtrlId
			})
		}
	}
	if (control_index < 0) {
		updateElements(null)
		return
	}
	ques_control_index = control_index
	var controlData = control_list[control_index]
	updateElements(controlData)
}

function changeQuestionType(data) {
	var controlData = window.BiyueCustomData.control_list[ques_control_index]
	controlData.ques_type = data.value * 1
}

function chagneProportion(data) {
	console.log('chagneProportion', data)
}

function initListener() {
	document.addEventListener('clickSingleQues', (event) => {
		if (!event || !event.detail) return
		console.log('receive clickSingleQues', event.detail)
		showQuesData(event.detail.control_id, event.detail.regionType)
	})
	document.addEventListener('updateQuesData', (event) => {
		if (!event || !event.detail) return
		console.log('receive updateQuesData', event.detail)
		var detail = event.detail
		if (ques_control_index >= 0) {
			var control_data = window.BiyueCustomData.control_list[ques_control_index]
			if (detail.list && detail.list.indexOf(control_data.control_id) >= 0) {
				if (detail.field == 'ques_type') {
					if (select_type) {
						select_type.setSelect(detail.value + '')
					}
				} else if (detail.field == 'proportion') {
					if (select_proportion) {
						select_proportion.setSelect(detail.value + '')
					}
				} else if (detail.field == 'score') {
					if (input_score) {
						input_score.setValue((detail.field || 0) + '')
					}
					// toto ask
				}
			}
		}
	})
}

function changeScore(id, data) {
	var controlData = window.BiyueCustomData.control_list[ques_control_index]
	if (id == 'ques_weight') {
		controlData.score = data
	} else {
		var index = id.replace('ask', '') * 1
		controlData.ask_controls[index].score = data
	}
}

export { showQuesData, initListener }
