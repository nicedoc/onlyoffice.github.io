import ComponentSelect from '../components/Select.js'
import NumberInput from '../components/NumberInput.js'
import { ZONE_SIZE, ZONE_TYPE } from './model/feature.js'
import { handleFeature, handleHeader } from './featureManager.js'
import { biyueCallCommand, dispatchCommandResult } from "./command.js";
var list_feature = []
var timeout_pos = null

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
	$('#panelFeature').empty()
	var content = '<table><tbody>'
	var list = [
		// {
		//   zone_type: ZONE_TYPE.QRCODE,
		//   id: 'qrCode',
		//   label: '二维码'
		// },
		{
			zone_type: ZONE_TYPE.THER_EVALUATION,
			id: 'evaluation',
			label: '评价图标',
		},
		// {
		//   zone_type: ZONE_TYPE.AGAIN,
		//   id: 'again',
		//   label: '再练区域'
		// },
		{
			zone_type: ZONE_TYPE.PASS,
			id: 'pass',
			label: '通过',
		},
		{
			zone_type: ZONE_TYPE.IGNORE,
			id: 'ignore',
			label: '日期/评语',
		},
		{
			id: 'header',
			label: '页眉',
		},
	]
	list.forEach((e, index) => {
		var str = `<tr><td colspan="2"><label class="header">${e.label}</label></td></tr><tr><td class="padding-small" width="100%"><div id=${e.id}></div></td></tr>`
		if (e.id != 'header') {
			str += `<tr id="${e.id}Pos"><td class="padding-small" width="50%"><label class="input-label">X坐标</label><div id="${e.id}X"></div></td><td class="padding-small" width="50%"><label class="input-label">Y坐标</label><div id="${e.id}Y"></div></td></tr>`
		}
		content += str
		if (index != list.length - 1) {
			content +=
				'<tr><td class="padding-small" colspan="2"><div class="separator horizontal"></div></td></tr>'
		}
	})
	content += '</tbody></table>'
	$('#panelFeature').html(content)
	list.forEach((e, index) => {
		e.comSelect = new ComponentSelect({
			id: e.id,
			options: types,
			value_select: types[0].value,
			width: '100%',
			callback_item: (data) => {
				changeItem(e.zone_type, data, index)
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
	})
	list_feature = list
	initPositions()
}

function changeItem(type, data, index) {
	var fdata = list_feature[index]
	var poscom = $(`#${fdata.id}Pos`)
	if (poscom) {
		if (data.value == 'open') {
			poscom.show()
		} else {
			poscom.hide()
		}
	}
	if (fdata.id == 'header') {
		handleHeader(data.value, window.BiyueCustomData.exam_title || '试卷标题')
	} else {
		handleFeature({
			zone_type: fdata.zone_type,
			cmd: data.value,
			x: fdata.x,
			y: fdata.y,
			p: fdata.p,
		})
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

function setXY(zone_type, p, x, y) {
	if (!list_feature) {
		return
	}
	var index = list_feature.findIndex((e) => e.zone_type == zone_type)
	if (index == -1) {
		return
	}
	list_feature[index].p = p
	list_feature[index].x = x
	list_feature[index].inputX.setValue(x)
	list_feature[index].y = y
	list_feature[index].inputY.setValue(y)
}

function initPositions() {
	biyueCallCommand(window, function () {
			var oDocument = Api.GetDocument()
			var sections = oDocument.GetSections()
			if (sections && sections.length > 0) {
				var oSection = sections[0]
				console.log('oSection', oSection, oSection.Section)
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
		console.log('initPositions result:', res)
		if (res) {
			var Num = res.Num
			var PageSize = res.PageSize
			var PageMargins = res.PageMargins
			var bottom = 0
			setXY(
				ZONE_TYPE.QRCODE,
				0,
				Num > 1 ? PageSize.W / Num - 40 : PageSize.W - PageMargins.Right - 20,
				PageMargins.Top
			)
			setXY(ZONE_TYPE.AGAIN, 0, PageMargins.Left, PageMargins.Top)
			var lastleft = ((Num - 1) * PageSize.W) / Num
			var evaluationX = Num > 1 ? lastleft : PageMargins.Left
			setXY(
				ZONE_TYPE.THER_EVALUATION,
				res.pageNum - 1,
				evaluationX,
				PageSize.H - PageMargins.Bottom - bottom
			)
			setXY(
				ZONE_TYPE.PASS,
				res.pageNum - 1,
				evaluationX + 150,
				PageSize.H - PageMargins.Bottom - bottom
			)
			setXY(
				ZONE_TYPE.IGNORE,
				res.pageNum - 1,
				PageSize.W - 15 - ZONE_SIZE[ZONE_TYPE.IGNORE].w,
				PageSize.H - PageMargins.Bottom - bottom
			)
			if (res.hasHeader) {
				var headerIndex = list_feature.findIndex((e) => e.id == 'header')
				if (headerIndex >= 0) {
					list_feature[headerIndex].comSelect.setSelect('open')
				}
			}
		}
		list_feature.forEach((e) => {
			var pos = $(`#${e.id}Pos`)
			if (pos) {
				if (e.comSelect.getValue() == 'open') {
					pos.show()
				} else {
					pos.hide()
				}
			}
		})
	})
}

export { initFeature }
