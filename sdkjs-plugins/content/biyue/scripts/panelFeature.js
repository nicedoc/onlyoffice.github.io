import ComponentSelect from '../components/Select.js'
import NumberInput from '../components/NumberInput.js'
import { ZONE_SIZE, ZONE_TYPE, ZONE_TYPE_NAME } from './model/feature.js'
import { handleFeature, handleHeader, drawExtroInfo, deleteAllFeatures, setInteraction } from './featureManager.js'
import { biyueCallCommand, dispatchCommandResult } from "./command.js";
var list_feature = []
var timeout_pos = null

function initExtroInfo() {
	list_feature = getList()
	updatePageSizeMargins().then(() => {
		initPositions(true)
	})
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
	var extra_info = {}
	if (workbook.extra_info && workbook.extra_info.length > 0) {
		extra_info = JSON.parse(workbook.extra_info)
	}
	if (workbook.practise_again) {
		extra_info.practise_again = workbook.practise_again
	}
	var list = []
	var scale = 0.2647058823529412
	if (extra_info.workbook_qr_code_show) {
		list.push({
			zone_type: ZONE_TYPE.QRCODE,
		  	id: ZONE_TYPE_NAME[ZONE_TYPE.QRCODE],
		  	label: '二维码',
			ox: getValue(extra_info.workbook_qr_code_x, 700 * scale) ,
			oy: getValue(extra_info.workbook_qr_code_y, 20 * scale),
			ow: 45 * scale, // extra_info.workbook_qr_code_size,
			oh: 45 * scale, //  extra_info.workbook_qr_code_size,
			url: 'https://by-qa-image-cdn.biyue.tech/qrCodeUnset.png'
		})
	}
	if (extra_info.practise_again && extra_info.practise_again.switch) {
		list.push({
			zone_type: ZONE_TYPE.AGAIN,
			id: ZONE_TYPE_NAME[ZONE_TYPE.AGAIN],
			label: '再练',
			ox: extra_info.practise_again.x,
			oy: extra_info.practise_again.y
		})
	}
	if (extra_info.custom_evaluate) {
		list.push({
			zone_type: ZONE_TYPE.SELF_EVALUATION,
			id: ZONE_TYPE_NAME[ZONE_TYPE.SELF_EVALUATION],
			label: extra_info.self_evaluate || '自我评价',
			icon_url: 'https://by-base-cdn.biyue.tech/xiaoyue.png',
			flowers: ['https://by-base-cdn.biyue.tech/flower.png', 'https://by-base-cdn.biyue.tech/flower.png', 'https://by-base-cdn.biyue.tech/flower.png', 'https://by-base-cdn.biyue.tech/flower.png']
		})
		list.push({
			zone_type: ZONE_TYPE.THER_EVALUATION,
			id: ZONE_TYPE_NAME[ZONE_TYPE.THER_EVALUATION],
			label: '教师评价',
			icon_url: 'https://by-base-cdn.biyue.tech/xiaotao.png',
			flowers: ['https://by-base-cdn.biyue.tech/flower.png', 'https://by-base-cdn.biyue.tech/flower.png', 'https://by-base-cdn.biyue.tech/flower.png', 'https://by-base-cdn.biyue.tech/flower.png']
		})
		list.push({
			zone_type: ZONE_TYPE.PASS,
			id: ZONE_TYPE_NAME[ZONE_TYPE.PASS],
			label: '通过'
		})
	} else {
		list.push({
			zone_type: ZONE_TYPE.END,
			id: ZONE_TYPE_NAME[ZONE_TYPE.END],
			label: '完成'
		})
	}

	list.push({
		zone_type: ZONE_TYPE.IGNORE,
		id: ZONE_TYPE_NAME[ZONE_TYPE.IGNORE],
		label: '日期/评语'
	})

	list.push({
		zone_type: ZONE_TYPE.STATISTICS,
		id: ZONE_TYPE_NAME[ZONE_TYPE.STATISTICS],
		label: '统计',
		p: 0,
		v: 1,
		// hidden: true,
		url: 'https://by-qa-image-cdn.biyue.tech/statistics.png'
	})

	// if (extra_info.start_interaction && extra_info.start_interaction.checked) {
		list.push({
			id: 'interaction',
			label: '互动模式'
		})
	// }
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
				value_select: optionsTypes[0].value,
				width: '100%',
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
	initPositions()
}

function changeAll(data) {
	if (data.value == 'close') {
		deleteAllFeatures()
	} else {
		drawExtroInfo([].concat(list_feature))
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
	if (id == 'header') {
		handleHeader(data.value, window.BiyueCustomData.exam_title || '试卷标题')
	} else if (id == 'interaction') {
		if (data.value == 'none') {
			deleteAllFeatures(null, ['ques_interaction'])
		} else {
			setInteraction(data.value)
		}
	} else {
		if (id == ZONE_TYPE_NAME[ZONE_TYPE.STATISTICS]) {
			var list = list_feature.filter(e => {
				return e.zone_type == ZONE_TYPE.STATISTICS
			})
			if (list) {
				drawExtroInfo(list.map(e => {
					return Object.assign({}, e, {
						cmd: data.value,
					})
				}))
			}
		} else {
			var fdata = list_feature.find(e => {
				return e.id == id
			})
			if (!fdata) {
				return
			}
			handleFeature(Object.assign({}, fdata, {
				cmd: data.value,
			}))
		}
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

function updatePageSizeMargins() {
	Asc.scope.workbook = window.BiyueCustomData.workbook_info
	return biyueCallCommand(window, function () {
		var workbook = Asc.scope.workbook
		var oDocument = Api.GetDocument()
		var sections = oDocument.GetSections()
		function MM2Twips(mm) {
			var m = Math.max(mm, 10)
			return m / (25.4 / 72 / 20)
		}
		function get2(v) {
			return MM2Twips(v * (workbook.page_size.width / 816)) 
		}
		if (sections && sections.length > 0) {
			sections.forEach(oSection => {
				oSection.SetPageSize(MM2Twips(workbook.page_size.width), MM2Twips(workbook.page_size.height))
				oSection.SetPageMargins(MM2Twips(workbook.margin.left), MM2Twips(workbook.margin.top), MM2Twips(workbook.margin.right), MM2Twips(workbook.margin.bottom))
				oSection.SetFooterDistance(MM2Twips(workbook.margin.bottom));
				oSection.SetHeaderDistance(MM2Twips(workbook.margin.top))
				// oSection.SetPageMargins(get2(workbook.margin.left), get2(workbook.margin.top), get2(workbook.margin.right), get2(workbook.margin.bottom))
			})
		}
		return null
	}, false, false).then()
}

function initPositions(draw) {
	Asc.scope.workbook = window.BiyueCustomData.workbook_info
	biyueCallCommand(window, function () {
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
			// sections.forEach(oSection => {
			// 	oSection.SetPageSize(MM2Twips(workbook.page_size.width), MM2Twips(workbook.page_size.height))
			// 	// oSection.SetPageMargins(MM2Twips(workbook.margin.left), MM2Twips(workbook.margin.top), MM2Twips(workbook.margin.right), MM2Twips(workbook.margin.bottom))
			// 	oSection.SetPageMargins(get2(workbook.margin.left), get2(workbook.margin.top), get2(workbook.margin.right), get2(workbook.margin.bottom))
			// })
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
	},
	false,
	false
	).then(res => {
		console.log('initPositions result:', res, window.BiyueCustomData.workbook_extra_info)
		if (res) {
			var Num = res.Num
			var PageSize = res.PageSize
			var PageMargins = res.PageMargins
			var bottom = 8
			// 后端给的坐标是基于页面尺寸 816*1100 的
			var scale = PageSize.W / 816
			console.log('ssss ', scale)
			var lastleft = ((Num - 1) * PageSize.W) / Num
			var evaluationX = Num > 1 ? lastleft : PageMargins.Left
			var statsIndex = list_feature.findIndex(e => {
				return e.zone_type == ZONE_TYPE.STATISTICS
			})
			if (statsIndex >= 0 && res.pageNum > 1) {
				var addStats = []
				for (var j = 1; j < res.pageNum; j++) {
					addStats.push(Object.assign({}, list_feature[statsIndex], {
						p: j,
						hidden: true,
						v: j + 1
					}))
				}
				list_feature.splice(statsIndex + 1, 0, ...addStats)
			}
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
					size.imgSize = scale * (45 - 10)
					page_num = 0
				} else if (e.zone_type == ZONE_TYPE.AGAIN) {
					page_num = 0
				} else if (e.zone_type == ZONE_TYPE.SELF_EVALUATION) {
					x = evaluationX
					y = PageSize.H - PageMargins.Bottom - bottom
				} else if (e.zone_type == ZONE_TYPE.THER_EVALUATION) {
					x = evaluationX + 60
					y = PageSize.H - PageMargins.Bottom - bottom
				} else if (e.zone_type == ZONE_TYPE.PASS || e.zone_type == ZONE_TYPE.END) {
					x = evaluationX + 120
					y = PageSize.H - PageMargins.Bottom - bottom
				} else if (e.zone_type == ZONE_TYPE.IGNORE) {
					x = PageSize.W - 15 - ZONE_SIZE[ZONE_TYPE.IGNORE].w
					y = PageSize.H - PageMargins.Bottom - bottom	
				} else if (e.zone_type == ZONE_TYPE.STATISTICS) {
					x = PageSize.W - PageMargins.Right
					y = PageSize.H - PageMargins.Bottom
					page_num = e.p
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
		console.log('list_feature', list_feature)
		if (draw) {
			deleteAllFeatures(list_feature).then(() => {
				drawExtroInfo(list_feature)
			})
			// drawExtroInfo([].concat(list_feature.slice(6, 9)))
		}
	})
}

export { initFeature, initExtroInfo }
