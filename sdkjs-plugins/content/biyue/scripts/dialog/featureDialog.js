import ComponentSelect from '../../components/Select.js'
import NumberInput from '../../components/NumberInput.js'
import { ZONE_TYPE, ZONE_TYPE_NAME, getInteractionTypes } from '../model/feature.js';
import { showCom } from '../model/util.js'
;(function (window, undefined) {
	var select_choice_style = null
	var select_choice_area = null
	var input_choice_num = null
	var timeout_change_choice_num = null
	var feature_map = {}
	const choiceStyles = [
		{
			value: 'brackets_choice_region',
			label: '括号识别'
		}, {
			value: 'show_choice_region',
			label: '集中作答区'
		}
	]
	const choiceAreas = [
		{
		  value: false,
		  label: '作答区前置'
		},
		{
		  value: true,
		  label: '作答区后置'
		}
	]
	const types = [
		{
			value: 'close',
			label: '关',
		},
		{
			value: 'open',
			label: '开',
		},
	]
	var list_feature = []
	var BiyueCustomData = {}
	window.Asc.plugin.init = function () {
		window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'initDialog', initmsg: 'featureMessage' })
	}
	function init() {
		list_feature = generateList()
		render()
	}
	function getSelectValue(id, vDefault = 'open') {
		if (feature_map && feature_map[id]) {
			return feature_map[id].sel || vDefault
		} else {
			return vDefault
		}
	}
	function generateList() {
		var list = []
		var workbook = BiyueCustomData.workbook_info
		if (!workbook) {
			return list
		}
		var extra_info = workbook.parse_extra_data
		var page_type = BiyueCustomData.page_type
		if (page_type == 0) { // 课时
			if (extra_info.workbook_qr_code_show) {
				list.push({
					zone_type: ZONE_TYPE.QRCODE,
					id: ZONE_TYPE_NAME[ZONE_TYPE.QRCODE],
					label: '二维码',
					value_select: getSelectValue(ZONE_TYPE_NAME[ZONE_TYPE.QRCODE])
				})
			}
			if (extra_info.practise_again && extra_info.practise_again.switch) {
				list.push({
					zone_type: ZONE_TYPE.AGAIN,
					id: ZONE_TYPE_NAME[ZONE_TYPE.AGAIN],
					label: '再练',
					value_select: getSelectValue(ZONE_TYPE_NAME[ZONE_TYPE.AGAIN]),
					separator: true
				})
			}
			if (extra_info.custom_evaluate) {
				list.push({
					zone_type: ZONE_TYPE.SELF_EVALUATION,
					id: ZONE_TYPE_NAME[ZONE_TYPE.SELF_EVALUATION],
					label: extra_info.self_evaluate || '自我评价',
					value_select: getSelectValue(ZONE_TYPE_NAME[ZONE_TYPE.SELF_EVALUATION]),
				})
				list.push({
					zone_type: ZONE_TYPE.THER_EVALUATION,
					id: ZONE_TYPE_NAME[ZONE_TYPE.THER_EVALUATION],
					label: '教师评价',
					value_select: getSelectValue(ZONE_TYPE_NAME[ZONE_TYPE.THER_EVALUATION]),
				})
				list.push({
					zone_type: ZONE_TYPE.PASS,
					id: ZONE_TYPE_NAME[ZONE_TYPE.PASS],
					label: '通过',
					value_select: getSelectValue(ZONE_TYPE_NAME[ZONE_TYPE.PASS]),
					separator: true
				})
			} else if (extra_info.hiddenComplete && extra_info.hiddenComplete.checked === false) {
				list.push({
					zone_type: ZONE_TYPE.END,
					id: ZONE_TYPE_NAME[ZONE_TYPE.END],
					label: '完成',
					value_select: getSelectValue(ZONE_TYPE_NAME[ZONE_TYPE.END]),
					separator: true
				})
			}
			list.push({
				zone_type: ZONE_TYPE.STATISTICS,
				id: ZONE_TYPE_NAME[ZONE_TYPE.STATISTICS],
				label: '统计',
				value_select: getSelectValue(ZONE_TYPE_NAME[ZONE_TYPE.STATISTICS]),
			})
			list.push({
				zone_type: ZONE_TYPE.IGNORE,
				id: ZONE_TYPE_NAME[ZONE_TYPE.IGNORE],
				label: '日期/评语',
				value_select: getSelectValue(ZONE_TYPE_NAME[ZONE_TYPE.IGNORE]),
				separator: true
			})
			list.push({
				id: 'interaction',
				label: '互动模式',
				value_select: getSelectValue('interaction') || 'none'
			})
		}
		return list
	}

	function render() {
		var interactionTypes = getInteractionTypes(BiyueCustomData)
		$('#wrapperFeature').empty()
		var content = '<table style="width: 100%"><tbody>'
		content += `<tr><td colspan="2"><label class="header">全部</label></td></tr><tr><td class="padding-small" width="100%" colspan="2"><div id='all'></div></td></tr>`
		content += '<tr><td class="padding-small" colspan="2"><div class="separator horizontal"></div></td></tr>'
		var list = list_feature
		var flag = 0
		list.forEach((e, index) => {
			if (!e.hidden) {
				var str = ''
				if (flag == 0) {
					str += '<tr>'
				}
				str += `<td class="padding-small" width="50%">
							<label class="input-label">${e.label}</label>
							<div id=${e.id}></div>
						</td>`
				flag++
				if (e.id == 'interaction') {
					str += `<td class="padding-small" width="50%" id="tdSimple">
								<label class="input-label">简单使用方式</label>
								<div id="simpleMode"></div>
							</td>`
					flag++
				}
				content += str
				if (e.separator) {
					content += '</tr>'
					flag = 0
				} else {
					if (flag % 2 == 0) {
						content += '</tr>'
						flag = 0
					}
				}
				if (index != list.length - 1 && e.separator) {
					content +=
						'<tr><td class="padding-small" colspan="2"><div class="separator horizontal"></div></td></tr>'
				}
			}
		})
		// 添加选择题集中作答区
		var choice_display = BiyueCustomData.choice_display || {}
		content += '<tr><td class="padding-small" colspan="2"><div class="separator horizontal"></div></td></tr>'
		content += `<tr><td colspan="2"><label class="header">选择题作答区</label></td></tr><tr><td class="padding-small" width="100%" colspan="2"><div id='select_choice_style'></div></td></tr>`
		content += `<tr id="choiceGather"><td class="padding-small" width="40%"><label class="input-label">每行数量</label><div id="input_choice_num"></div></td><td class="padding-small" width="60%"><label class="input-label">作答区位置</label><div id="select_choice_area"></div></td></tr>`
		content += '</tbody></table>'
		$('#wrapperFeature').html(content)
		var allComSelect = new ComponentSelect({
			id: 'all',
			options: types,
			value_select: getSelectValue('all', types[1].value),
			width: '100%',
			force_click_notify: true,
			callback_item: (data) => {
				changeAll(data)
			},
		})
		list.forEach((e, index) => {
			var optionsTypes = e.id == 'interaction' ? interactionTypes : types
			if (!e.hidden) {
				e.comSelect = new ComponentSelect({
					id: e.id,
					options: optionsTypes,
					value_select: e.value_select || optionsTypes[0].value,
					width: '100%',
					force_click_notify: true,
					callback_item: (data) => {
						changeItem(e.zone_type, data, e.id)
					},
				})
				if (e.id == 'interaction') {
					var vInteraction = e.value_select || optionsTypes[0].value
					e.comSelect2 = new ComponentSelect({
						id: 'simpleMode',
						options: [
							{ value: '1', label: '题号' },
							{ value: '2', label: '浮动图片' }
						],
						value_select: (BiyueCustomData.simple_interaction || 1) + '',
						callback_item: (data) => {
							changeSimpleInteraction(data.value)
						},
						width: '100%',
						pop_width: '100%',
						enabled: vInteraction != 'none'
					})
				}
			}
		})
		if (choice_display) {
			select_choice_style = new ComponentSelect({
				id: 'select_choice_style',
				options: choiceStyles,
				value_select: choice_display.style,
				callback_item: (data) => {
					changeChoiceStyle(data)
				},
				width: '100%',
			})
			select_choice_area = new ComponentSelect({
				id: 'select_choice_area',
				options: choiceAreas,
				value_select: choice_display.area,
				callback_item: (data) => {
					changeChoiceArea(data)
				},
				width: '100%',
			})
			input_choice_num = new NumberInput('input_choice_num', {
				width: '100%',
				change: (id, data) => {
					changeChoiceNum(id, data)
				},
			})
			if (input_choice_num) {
				input_choice_num.setValue(choice_display.num_row + '')
			}
		}
		showCom('#choiceGather', choice_display && choice_display.style != 'brackets_choice_region')
	}

	function updateInfo() {
		if (!list_feature) {
			return
		}
		for (var item of list_feature) {
			var targetValue = getSelectValue(item.id, 'open')
			if (item.comSelect) {
				if (item.comSelect.getValue() != targetValue) {
					item.comSelect.setSelect(targetValue)
				}
			}
			if (item.id == 'interaction' && item.comSelect2) {
				item.comSelect2.setEnable(targetValue != 'none')
			}
		}
	}

	function changeAll(data) {
		if (!BiyueCustomData.workbook_info) {
			return
		}
		send('all', data)
	}

	function changeItem(zoneType, data, id) {
		send('zoneType', {
			zone_type: zoneType,
			id: id,
			data: data
		})
	}

	function changeSimpleInteraction(data) {
		send('simple_interaction', data)
	}

	// 切换选择题作答区样式
	function changeChoiceStyle(data) {
		if (data) {
			showCom('#choiceGather', data.value != 'brackets_choice_region')
			send('choiceStyle', data)
		}
	}

	function changeChoiceArea(data) {
		if (data) {
			send('choiceArea', data)
		}
	}

	function changeChoiceNum(id, data) {
		clearTimeout(timeout_change_choice_num)
		timeout_change_choice_num = setTimeout(() => {
			send('choice_num_row', data)
		}, 400);
	}

	function send(cmd, data) {
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'featureMessage',
			cmd: cmd,
			data: data
		})	
	}

	window.Asc.plugin.attachEvent('featureMessage', function (message) {
		console.log('featureMessage 接收的消息', message)
		if (message) {
			if (message.BiyueCustomData) {
				BiyueCustomData = message.BiyueCustomData
			}
 			if (message.feature_map) {
				feature_map = message.feature_map
			}
 		}
		init()
	})

	window.Asc.plugin.attachEvent('featureUpdate', function (message) {
		console.log('featureUpdate 接收的消息', message)
		if (message) {
			if (message.BiyueCustomData) {
				BiyueCustomData = message.BiyueCustomData
			} else if (message.field) {
				BiyueCustomData[message.field] = message.data
			}
			if (message.feature_map) {
				feature_map = message.feature_map
			}
			updateInfo()
		}
	})
})(window, undefined)
