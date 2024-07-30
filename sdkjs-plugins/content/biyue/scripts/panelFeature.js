import ComponentSelect from '../components/Select.js'
import NumberInput from '../components/NumberInput.js'
import { ZONE_SIZE, ZONE_TYPE, ZONE_TYPE_NAME } from './model/feature.js'
import { handleFeature, handleHeader, drawExtroInfo, setLoading, deleteAllFeatures, setInteraction } from './featureManager.js'
import { biyueCallCommand, dispatchCommandResult } from "./command.js";
var list_feature = []
var timeout_pos = null

function initExtroInfo() {
	list_feature = getList()
	return initPositions1()
}

function getValue(v1, v2) {
	if (v1 == '' || isNaN(v1 * 1)) {
		return v2
	} else {
		return v1
	}
}

function getList() {
	var workbook = window.BiyueCustomData.workbook_info
	if (!workbook) {
		return
	}
	var extra_info = {}
	if (workbook.extra_info && workbook.extra_info.length > 0) {
		extra_info = JSON.parse(workbook.extra_info)
		window.BiyueCustomData.workbook_info.parse_extra_data = extra_info
	}
	if (workbook.practise_again) {
		extra_info.practise_again = workbook.practise_again
	}
	var list = []
	var scale = 0.2647058823529412
	if (extra_info.workbook_qr_code_show) {
		var size = getValue(extra_info.workbook_qr_code_size, 45 * scale)
		list.push({
			zone_type: ZONE_TYPE.QRCODE,
		  	id: ZONE_TYPE_NAME[ZONE_TYPE.QRCODE],
		  	label: '二维码',
			ox: getValue(extra_info.workbook_qr_code_x, 700 * scale),
			oy: getValue(extra_info.workbook_qr_code_y, 20 * scale),
			ow: size,
			oh: size,
			url: 'https://by-qa-image-cdn.biyue.tech/qrCodeUnset.png',
			value_select: 'open'
		})
	}
	if (extra_info.practise_again && extra_info.practise_again.switch) {
		list.push({
			zone_type: ZONE_TYPE.AGAIN,
			id: ZONE_TYPE_NAME[ZONE_TYPE.AGAIN],
			label: '再练',
			ox: extra_info.practise_again.x,
			oy: extra_info.practise_again.y,
			value_select: 'open'
		})
	}
	if (extra_info.custom_evaluate) {
		list.push({
			zone_type: ZONE_TYPE.SELF_EVALUATION,
			id: ZONE_TYPE_NAME[ZONE_TYPE.SELF_EVALUATION],
			label: extra_info.self_evaluate || '自我评价',
			icon_url: 'https://by-base-cdn.biyue.tech/xiaoyue.png',
			flowers: Object.values(extra_info.self_filling_imgs || {}),
			value_select: 'open'
		})
		list.push({
			zone_type: ZONE_TYPE.THER_EVALUATION,
			id: ZONE_TYPE_NAME[ZONE_TYPE.THER_EVALUATION],
			label: '教师评价',
			icon_url: 'https://by-base-cdn.biyue.tech/xiaotao.png',
			flowers: Object.values(extra_info.teacher_filling_imgs || {}),
			value_select: 'open'
		})
		list.push({
			zone_type: ZONE_TYPE.PASS,
			id: ZONE_TYPE_NAME[ZONE_TYPE.PASS],
			label: '通过',
			value_select: 'open'
		})
	} else {
		list.push({
			zone_type: ZONE_TYPE.END,
			id: ZONE_TYPE_NAME[ZONE_TYPE.END],
			label: '完成',
			value_select: 'open'
		})
	}

	list.push({
		zone_type: ZONE_TYPE.IGNORE,
		id: ZONE_TYPE_NAME[ZONE_TYPE.IGNORE],
		label: '日期/评语',
		value_select: 'open'
	})

	list.push({
		zone_type: ZONE_TYPE.STATISTICS,
		id: ZONE_TYPE_NAME[ZONE_TYPE.STATISTICS],
		label: '统计',
		p: 0,
		v: 1,
		// hidden: true,
		url: 'https://by-qa-image-cdn.biyue.tech/statistics.png',
		value_select: 'open'
	})
	if (!extra_info.hidden_correct_region.checked) {
		var value_select = extra_info.start_interaction.checked ? 'accurate' : 'simple'
		window.BiyueCustomData.interaction = value_select
		list.push({
			id: 'interaction',
			label: '互动模式',
			value_select: value_select
		})
	} else {
		window.BiyueCustomData.interaction = 'none'
	}
	return list
}

function initFeature() {
	var types = [
		{
			value: 'close',
			label: '关',
		},
		{
			value: 'open',
			label: '开',
		},
	]
	var interactionTypes = [
		{
			value: 'none',
			label: '无互动'
		},
		{
			value: 'simple',
			label: '简单互动'
		}, {
			value: 'accurate',
			label: '精准互动'
		}
	]
	$('#panelFeature').empty()
	var content = '<table><tbody>'
	content += `<tr><td colspan="2"><label class="header">全部</label></td></tr><tr><td class="padding-small" width="100%"><div id='all'></div></td></tr>`
	content += '<tr><td class="padding-small" colspan="2"><div class="separator horizontal"></div></td></tr>'
	var list = getList()
	list.forEach((e, index) => {
		if (!e.hidden) {
			var str = `<tr><td colspan="2"><label class="header">${e.label}</label></td></tr><tr><td class="padding-small" width="100%"><div id=${e.id}></div></td></tr>`
			if (e.id != 'header') {
				str += `<tr id="${e.id}Pos"><td class="padding-small" width="50%"><label class="input-label">X坐标</label><div id="${e.id}X"></div></td><td class="padding-small" width="50%"><label class="input-label">Y坐标</label><div id="${e.id}Y"></div></td></tr>`
			}
			content += str
			if (index != list.length - 1) {
				content +=
					'<tr><td class="padding-small" colspan="2"><div class="separator horizontal"></div></td></tr>'
			}
		}
	})
	content += '</tbody></table>'
	$('#panelFeature').html(content)
	var allComSelect = new ComponentSelect({
		id: 'all',
		options: types,
		value_select: types[1].value,
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
			e.inputX = new NumberInput(`${e.id}X`, {
				change: (id, data) => {
					changePos(e.zone_type, 'x', data)
				},
			})
			e.inputY = new NumberInput(`${e.id}Y`, {
				change: (id, data) => {
					changePos(e.zone_type, 'y', data)
				},
			})
		}
	})
	list_feature = list
	initPositions2()
}

function changeAll(data) {
	var vinteraction = 'none'
	var extra_info = window.BiyueCustomData.workbook_info.parse_extra_data
	if (extra_info.hidden_correct_region.checked == false) {
		vinteraction = extra_info.start_interaction.checked ? 'accurate' : 'simple'
	}
	window.BiyueCustomData.interaction = vinteraction
	if (data.value == 'close') {
		deleteAllFeatures()
	} else {
		drawExtroInfo([].concat(list_feature)).then(res => {
			setLoading(false)
			setInteraction(vinteraction)
		})
	}
	if (list_feature) {
		list_feature.forEach(e => {
			if (e.id == 'interaction') {
				e.comSelect.setSelect(vinteraction)
				updateAllInteraction(vinteraction)
			} else {
				if (e.comSelect) {
					e.comSelect.setSelect(data.value)
				}
			}
		})
	}
} 

function changeItem(type, data, id) {
	var poscom = $(`#${id}Pos`)
	if (poscom) {
		if (data.value == 'open') {
			poscom.show()
		} else {
			poscom.hide()
		}
	}
	var fdata = list_feature.find(e => {
		return e.id == id
	})
	if (fdata) {
		fdata.value_select = data.value
	}
	if (id == 'header') {
		handleHeader(data.value, window.BiyueCustomData.exam_title || '试卷标题')
	} else if (id == 'interaction') {
		window.BiyueCustomData.interaction = data.value
		if (data.value == 'none') {
			deleteAllFeatures(null, ['ques_interaction'])
		} else {
			setInteraction(data.value)	
		}
		updateAllInteraction(data.value)
	} else {
		if (id == ZONE_TYPE_NAME[ZONE_TYPE.STATISTICS]) {
			drawExtroInfo([Object.assign({}, fdata, {
				cmd:data.value
			})]).then(() => {
				setLoading(false)
			})
		} else {
			handleFeature(Object.assign({}, fdata, {
				cmd: data.value,
			}))
		}
	}

}
// 重新切题后同步互动情况
function syncInteractionWhenReSplit() {
	if (window.BiyueCustomData.interaction != 'none') {
		setInteraction(window.BiyueCustomData.interaction)
		updateAllInteraction(window.BiyueCustomData.interaction)
	}
}

function changePos(zone_type, id, data) {
	var index = list_feature.findIndex((e) => e.zone_type == zone_type)
	if (index == -1) {
		return
	}
	list_feature[index][id] = data * 1
	handleFeature({
		zone_type: list_feature[index].zone_type,
		cmd: 'open',
		x: list_feature[index].x,
		y: list_feature[index].y,
		p: list_feature[index].p,
	})
}

function setXY(index, p, x, y, size) {
	if (!list_feature) {
		return
	}
	if (index < 0 || index >= list_feature.length) {
		return
	}
	list_feature[index].p = p
	list_feature[index].x = x
	if (list_feature[index].inputX) {
		list_feature[index].inputX.setValue(x)
	}
	list_feature[index].y = y
	if (list_feature[index].inputY) {
		list_feature[index].inputY.setValue(y)
	}
	if (size) {
		list_feature[index].size = size
	}
}

function getPageData() {
	Asc.scope.workbook = window.BiyueCustomData.workbook_info
	return biyueCallCommand(window, function () {
		var workbook = Asc.scope.workbook
		var oDocument = Api.GetDocument()
		var sections = oDocument.GetSections()
		function MM2Twips(mm) {
			return mm / (25.4 / 72 / 20)
		}
		function get2(v) {
			return MM2Twips(v * (workbook.page_size.width / 816)) 
		}
		if (sections && sections.length > 0) {
			var oSection = sections[0]
			var pageNum = oDocument.Document.Pages.length
			var hasHeader = !!oSection.GetHeader('title', false)
			return {
				Num: oSection.Section.Columns.Num,
				PageSize: oSection.Section.PageSize,
				PageMargins: oSection.Section.PageMargins,
				pageNum: pageNum,
				hasHeader: hasHeader,
			}
		}
		return null
	},false,false)
}

function updateFeatureList(res) {
	if (res) {
		var Num = res.Num
		var PageSize = res.PageSize
		var PageMargins = res.PageMargins
		var bottom = 6
		// 后端给的坐标是基于页面尺寸 816*1100 的
		var lastleft = ((Num - 1) * PageSize.W) / Num
		var evaluationX = Num > 1 ? lastleft : PageMargins.Left
		console.log('evaluationX', evaluationX)
		list_feature.forEach((e, index) => {
			var x = e.ox != undefined ? e.ox : 0
			var y = e.oy != undefined ? e.oy : 0
			var size = {}
			if (e.ow != undefined && e.oh != undefined) {
				size.w = e.ow
				size.h = e.oh
			}
			var page_num = res.pageNum - 1
			if (e.zone_type == ZONE_TYPE.QRCODE) {
				size.imgSize = size.w - 3
				page_num = 0
			} else if (e.zone_type == ZONE_TYPE.AGAIN) {
				page_num = 0
			} else if (e.zone_type == ZONE_TYPE.SELF_EVALUATION) {
				x = evaluationX
				y = PageSize.H - PageMargins.Bottom
			} else if (e.zone_type == ZONE_TYPE.THER_EVALUATION) {
				x = evaluationX + 60
				y = PageSize.H - PageMargins.Bottom
			} else if (e.zone_type == ZONE_TYPE.PASS || e.zone_type == ZONE_TYPE.END) {
				x = PageSize.W - PageMargins.Right - ZONE_SIZE[ZONE_TYPE.IGNORE].w - 4 - ZONE_SIZE[ZONE_TYPE.PASS].w
				y = PageSize.H - PageMargins.Bottom
			} else if (e.zone_type == ZONE_TYPE.IGNORE) {
				x = PageSize.W - PageMargins.Right - ZONE_SIZE[ZONE_TYPE.IGNORE].w
				y = PageSize.H - PageMargins.Bottom
			} else if (e.zone_type == ZONE_TYPE.STATISTICS) {
				x = PageSize.W - PageMargins.Right
				y = PageSize.H - PageMargins.Bottom
				page_num = e.p
				size.pagination_ver_pos = window.BiyueCustomData.workbook_info.parse_extra_data.pagination_margin_bottom * 1 - PageMargins.Bottom
			}
			setXY(
				index,
				page_num,
				x,
				y,
				size
			)
		})
		if (res.hasHeader) {
			var headerIndex = list_feature.findIndex((e) => e.id == 'header')
			if (headerIndex >= 0) {
				list_feature[headerIndex].comSelect.setSelect('open')
			}
		}
	}
	list_feature.forEach((e) => {
		var pos = $(`#${e.id}Pos`)
		if (pos && e.comSelect) {
			if (e.comSelect.getValue() == 'open') {
				pos.show()
			} else {
				pos.hide()
			}
		}
	})
}

function initPositions1() {
	return getPageData().then(res => {
		updateFeatureList(res)
		deleteAllFeatures([], list_feature)
	}).then(res => {
		var vinteraction = window.BiyueCustomData.interaction
		updateAllInteraction(vinteraction)
		if (vinteraction != 'none') {
			return setInteraction(vinteraction)
		} else {
			return new Promise((resolve, reject) => {
				resolve()
			})
		}
	}).then(() => {
		return drawExtroInfo(list_feature)
	}).then(() => {
		setLoading(false)
		return MoveCursor()
	})		
}

function initPositions2() {
	return getPageData().then(res => {
		updateFeatureList(res)
	})
}

function MoveCursor() {
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		if (controls && controls.length) {
			oDocument.Document.MoveCursorToContentControl(controls[0].Sdt.GetId(), true)
		} else {
			oDocument.Document.MoveCursorToPageEnd()
		}
	}, false, false)
}

function updateAllInteraction(vinteraction) {
	var question_map = window.BiyueCustomData.question_map || {}
	Object.keys(question_map).forEach(e => {
		question_map[e].interaction = vinteraction
	})
}

export { initFeature, initExtroInfo, syncInteractionWhenReSplit }
