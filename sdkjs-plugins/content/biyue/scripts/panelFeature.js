import ComponentSelect from '../components/Select.js'
import NumberInput from '../components/NumberInput.js'
import { ZONE_SIZE, ZONE_TYPE, ZONE_TYPE_NAME, getInteractionTypes } from './model/feature.js'
import { handleFeature, handleHeader, drawExtroInfo, setLoading, deleteAllFeatures, setInteraction, updateChoice, handleChoiceUpdateResult, drawHeaderFooter, drawStatistics } from './featureManager.js'
import { biyueCallCommand } from "./command.js";
import { showCom } from './model/util.js'
var list_feature = []
var choiceStyles = [
	{
		value: 'brackets_choice_region',
		label: '括号识别'
	}, {
		value: 'show_choice_region',
		label: '集中作答区'
	}
]
var choiceAreas = [
	{
	  value: false,
	  label: '作答区前置'
	},
	{
	  value: true,
	  label: '作答区后置'
	}
]
var select_choice_style = null
var select_choice_area = null
var input_choice_num = null
var timeout_change_choice_num = null
var imageDimensionsCache = {} // 图片的宽高比缓存
var STAT_URL = 'https://by-qa-image-cdn.biyue.tech/statistics.png'
function initExtroInfo() {
	list_feature = getList()
	if (window.BiyueCustomData.page_type == 1) {
		return initPositions1_intro()
	}
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
		try {
			extra_info = JSON.parse(workbook.extra_info)
		} catch (error) {
			console.log('json parse error', error)
			return
		}
		
		var choice_display = Object.assign({
			num_row: 10,
			area: false,
			style: 'brackets_choice_region'
		}, window.BiyueCustomData.choice_display || {})
		if (extra_info.show_choice_region_bottom != undefined) {
			choice_display.area = extra_info.show_choice_region_bottom.checked
		}
		if (extra_info.brackets_choice_region) {
			if (extra_info.brackets_choice_region.checked) {
				choice_display.style = 'brackets_choice_region'
			} else {
				choice_display.style = 'show_choice_region'
				choice_display.num_row = extra_info.show_choice_region.content
			}
		}
		window.BiyueCustomData.choice_display = choice_display
	}
	if (workbook.practise_again) {
		extra_info.practise_again = workbook.practise_again
	}
	if (workbook.page_layout) {
		extra_info.page_header_text = workbook.page_layout.page_header || ''
		extra_info.page_footer_text = workbook.page_layout.page_footer || ''
		extra_info.page_image_url = workbook.page_layout.page_top_left || ''
	}
	window.BiyueCustomData.workbook_info.parse_extra_data = extra_info
	console.log('【extra_info】', extra_info)
	var list = []
	var scale = 0.2647058823529412
	var page_type = window.BiyueCustomData.page_type
	if (page_type == 0) {
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
				icon_url: workbook.self_evaluate_img, // || 'https://by-base-cdn.biyue.tech/xiaoyue_s.png',
				flowers: Object.values(extra_info.self_filling_imgs || {}),
				value_select: 'open'
			})
			list.push({
				zone_type: ZONE_TYPE.THER_EVALUATION,
				id: ZONE_TYPE_NAME[ZONE_TYPE.THER_EVALUATION],
				label: '教师评价',
				icon_url: workbook.teacher_evaluate_img, // || 'https://by-base-cdn.biyue.tech/xiaotao_s.png',
				flowers: Object.values(extra_info.teacher_filling_imgs || {}),
				value_select: 'open'
			})
			list.push({
				zone_type: ZONE_TYPE.PASS,
				id: ZONE_TYPE_NAME[ZONE_TYPE.PASS],
				label: '通过',
				value_select: 'open'
			})
		} else if (extra_info.hiddenComplete && extra_info.hiddenComplete.checked === false) {
			list.push({
				zone_type: ZONE_TYPE.END,
				id: ZONE_TYPE_NAME[ZONE_TYPE.END],
				label: '完成',
				value_select: 'open'
			})
		}
		list.push({
			zone_type: ZONE_TYPE.STATISTICS,
			id: ZONE_TYPE_NAME[ZONE_TYPE.STATISTICS],
			label: '统计',
			value_select: 'open'
		})
		list.push({
			zone_type: ZONE_TYPE.IGNORE,
			id: ZONE_TYPE_NAME[ZONE_TYPE.IGNORE],
			label: '日期/评语',
			value_select: 'open'
		})
	} 
	if (page_type == 0) {
		if (!extra_info.hidden_correct_region.checked) {
			var value_select = extra_info.start_interaction.checked ? 'accurate' : 'simple'
			window.BiyueCustomData.interaction = value_select
		} else {
			window.BiyueCustomData.interaction = 'none'
		}
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
	var interactionTypes = getInteractionTypes()
	console.log('=============== initFeature interactionTypes', interactionTypes)
	$('#wrapperFeature').empty()
	var content = '<table style="width: 100%"><tbody>'
	content += `<tr><td colspan="2"><label class="header">全部</label></td></tr><tr><td class="padding-small" width="100%" colspan="2"><div id='all'></div></td></tr>`
	content += '<tr><td class="padding-small" colspan="2"><div class="separator horizontal"></div></td></tr>'
	var list = list_feature || getList()
	list.forEach((e, index) => {
		if (!e.hidden) {
			var str = `<tr><td colspan="2"><label class="header">${e.label}</label></td></tr><tr><td class="padding-small" width="100%" colspan="2"><div id=${e.id}></div></td></tr>`
			if (e.id != 'header' && e.id != 'statistics') {
				str += `<tr id="${e.id}Pos"><td class="padding-small" width="50%"><label class="input-label">X坐标</label><div id="${e.id}X"></div></td><td class="padding-small" width="50%"><label class="input-label">Y坐标</label><div id="${e.id}Y"></div></td></tr>`
			}
			content += str
			if (index != list.length - 1) {
				content +=
					'<tr><td class="padding-small" colspan="2"><div class="separator horizontal"></div></td></tr>'
			}
		}
	})
	// 添加选择题集中作答区
	var choice_display = window.BiyueCustomData.choice_display || {}
	content += '<tr><td class="padding-small" colspan="2"><div class="separator horizontal"></div></td></tr>'
	content += `<tr><td colspan="2"><label class="header">选择题作答区</label></td></tr><tr><td class="padding-small" width="100%" colspan="2"><div id='select_choice_style'></div></td></tr>`
	content += `<tr id="choiceGather"><td class="padding-small" width="40%"><label class="input-label">每行数量</label><div id="input_choice_num"></div></td><td class="padding-small" width="60%"><label class="input-label">作答区位置</label><div id="select_choice_area"></div></td></tr>`
	content += '</tbody></table>'
	$('#wrapperFeature').html(content)
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
				width: '100%',
			})
			e.inputY = new NumberInput(`${e.id}Y`, {
				change: (id, data) => {
					changePos(e.zone_type, 'y', data)
				},
				width: '100%',
			})
		}
	})
	list_feature = list
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
	initPositions2()
}

function changeAll(data) {
	if (!window.BiyueCustomData.workbook_info) {
		return
	}
	var vinteraction = 'none'
	var extra_info = window.BiyueCustomData.workbook_info.parse_extra_data
	if (!extra_info) {
		return
	}
	if (data.value != 'close' && (!extra_info.hidden_correct_region || extra_info.hidden_correct_region.checked == false)) {
		vinteraction = extra_info.start_interaction && extra_info.start_interaction.checked ? 'accurate' : 'simple'
	}
	window.BiyueCustomData.interaction = vinteraction
	if (data.value == 'close') {
		deleteAllFeatures(['pagination'])
	} else {
		drawExtroInfo([].concat(list_feature), false)
		.then(() => {
			return drawPageHeaderFooter(true)
		})
		.then(res => {
			setLoading(false)
			setInteraction(vinteraction)
		})
	}
	if (list_feature) {
		list_feature.forEach(e => {
			if (e.id == 'interaction') {
				if (e.comSelect) {
					e.comSelect.setSelect(vinteraction)
				}
				e.value_select = vinteraction
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
			var extra_info = window.BiyueCustomData.workbook_info.parse_extra_data
			drawStatistics({
				cmd: data.value,
				stat: Object.assign({}, extra_info.onlyoffice_options.statis, {
					width: 4.76,
					height: 4.76,
					url: STAT_URL
				}),
				page_type: window.BiyueCustomData.page_type
			}, true).then(() => {
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
		try {
			console.log('[getPageData] begin')
			var workbook = Asc.scope.workbook || {}
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
		} catch (error) {
			console.error('[getPageData]', error)
		}
	},false,false)
}

function updateFeatureList(res) {
	if (res) {
		var Num = res.Num
		var PageSize = res.PageSize
		var PageMargins = res.PageMargins
		var bottom = 6
		// 后端给的坐标是基于页面尺寸 816*1100 的
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
				x = PageMargins.Left
				y = PageSize.H - PageMargins.Bottom
			} else if (e.zone_type == ZONE_TYPE.THER_EVALUATION) {
				x = PageMargins.Left + 60
				y = PageSize.H - PageMargins.Bottom
			} else if (e.zone_type == ZONE_TYPE.PASS || e.zone_type == ZONE_TYPE.END) {
				x = PageSize.W - PageMargins.Right - ZONE_SIZE[ZONE_TYPE.IGNORE].w - 4 - ZONE_SIZE[ZONE_TYPE.PASS].w
				y = PageSize.H - PageMargins.Bottom
			} else if (e.zone_type == ZONE_TYPE.IGNORE) {
				x = PageSize.W - PageMargins.Right - ZONE_SIZE[ZONE_TYPE.IGNORE].w
				y = PageSize.H - PageMargins.Bottom
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
			if (headerIndex >= 0 && list_feature[headerIndex].comSelect) {
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

function initPositions1_intro() {
	return getPageData().then(res => {
		updateFeatureList(res)
		var specifyFeatures = [
			ZONE_TYPE_NAME[ZONE_TYPE.AGAIN],
			ZONE_TYPE_NAME[ZONE_TYPE.IGNORE],
			ZONE_TYPE_NAME[ZONE_TYPE.QRCODE],
			ZONE_TYPE_NAME[ZONE_TYPE.SELF_EVALUATION],
			ZONE_TYPE_NAME[ZONE_TYPE.THER_EVALUATION],
			ZONE_TYPE_NAME[ZONE_TYPE.STATISTICS],
			ZONE_TYPE_NAME[ZONE_TYPE.PASS],
			ZONE_TYPE_NAME[ZONE_TYPE.END]
		]
		return deleteAllFeatures([], specifyFeatures)
	}).then(() => { // 插入获取图片宽高比的步骤
		return loadImages()
  	})
  	.then(() => {
		return drawPageHeaderFooter(true)
  	})
	.then(() => {
		setLoading(false)
		return MoveCursor()
	})
}

function loadImages() {
	const imageUrls = [];
	if (window.BiyueCustomData.page_type == 0) {
		list_feature.forEach(e => {
			if (e.url) {
				imageUrls.push(e.url)
			}
			if (e.icon_url) {
				imageUrls.push(e.icon_url)
			}
			if (e.flowers) {
				imageUrls.concat(e.flowers)
			}
		})
	}
	var extra_info = window.BiyueCustomData.workbook_info.parse_extra_data
	if (extra_info && extra_info.page_image_url) {
		imageUrls.push(extra_info.page_image_url)
	}
	return Promise.all(imageUrls.map(getImageDimensions));
}
function initPositions1() {
	return getPageData().then(res => {
		updateFeatureList(res)
		// var specifyFeatures = list_feature.map(e => {
		// 	return ZONE_TYPE_NAME[e.zone_type]
		// })
		// if (!specifyFeatures.includes(ZONE_TYPE_NAME[ZONE_TYPE.STATISTICS])) {
		// 	specifyFeatures.push(ZONE_TYPE_NAME[ZONE_TYPE.STATISTICS])
		// }
		// if (!specifyFeatures.includes(ZONE_TYPE_NAME[ZONE_TYPE.PAGINATION])) {
		// 	specifyFeatures.push(ZONE_TYPE_NAME[ZONE_TYPE.PAGINATION])
		// }
		return deleteAllFeatures([])
	})
	.then(() => {
		return updateChoice(false)
	})
	.then(res => {
		if (res) {
			window.BiyueCustomData.node_list = res
		}
		var vinteraction = window.BiyueCustomData.interaction
		updateAllInteraction(vinteraction, false)
		return setInteraction('useself')
	})
	.then(() => { // 插入获取图片宽高比的步骤
		return loadImages()
	})
	.then(() => {
		return drawExtroInfo(list_feature, imageDimensionsCache, false)
	})
	.then(() => {
		return drawPageHeaderFooter(true)
	})
	.then(() => {
		setLoading(false)
		return MoveCursor()
	})
}

function initPositions2() {
	console.log('=============== initPositions2')
	return getPageData().then(res => {
		updateFeatureList(res)
	})
}

function MoveCursor() {
	return biyueCallCommand(window, function() {
		try {
			console.log('[MoveCursor] begin')
			var oDocument = Api.GetDocument()
			var controls = oDocument.GetAllContentControls()
			if (controls && controls.length) {
				oDocument.Document.MoveCursorToContentControl(controls[0].Sdt.GetId(), true)
			} else {
				oDocument.Document.MoveCursorToPageEnd()
			}
		} catch (error) {
			console.error('[MoveCursor]', error)
		}
	}, false, false)
}

function updateAllInteraction(vinteraction, isForce = true) {
	var question_map = window.BiyueCustomData.question_map || {}
	Object.keys(question_map).forEach(e => {
		if (isForce || question_map[e].interaction == undefined) {
			question_map[e].interaction = vinteraction
		}
	})
}

// 切换选择题作答区样式
function changeChoiceStyle(data) {
	if (data) {
		window.BiyueCustomData.choice_display.style = data.value
		showCom('#choiceGather', data.value != 'brackets_choice_region')
	}
	return updateChoice(true).then(res => {
		return handleChoiceUpdateResult(res)
	})
}
// 切换选择题作答区位置
function changeChoiceArea(data) {
	window.BiyueCustomData.choice_display.area = data.value
	return updateChoice(true).then(res => {
		return handleChoiceUpdateResult(res)
	})
}

function changeChoiceNum(id, data) {
	if (!window.BiyueCustomData.choice_display) {
		return
	}
	clearTimeout(timeout_change_choice_num)
	timeout_change_choice_num = setTimeout(() => {
		window.BiyueCustomData.choice_display.num_row = data
		updateChoice(true).then(res => {
			return handleChoiceUpdateResult(res)
		})
	}, 400);
}
// 获取图片宽高比
function getImageDimensions(url) {
	if (imageDimensionsCache[url]) {
		return Promise.resolve(imageDimensionsCache[url]);
	}
	return new Promise((resolve, reject) => {
	  	const img = new Image();
		img.crossOrigin = "anonymous";  // 设置跨域属性
	  	img.onload = function() {
			const dimensions = { width: img.naturalWidth, height: img.naturalHeight, aspectRatio: img.naturalWidth / img.naturalHeight };
			// 缓存获取的宽高比信息
			imageDimensionsCache[url] = dimensions;
			resolve(dimensions);
		};	  
		img.onerror = function() {
			reject(`Failed to load image: ${url}`);
		};
		img.src = url;
	});
}
function drawPageHeaderFooter(recalc) {
	var extra_info = window.BiyueCustomData.workbook_info.parse_extra_data || {}
	var page_logo_width = extra_info.page_logo_width
	if (!page_logo_width && extra_info.page_image_url) {
		var imageData = imageDimensionsCache[extra_info.page_image_url]
		if (!extra_info.page_logo_height) {
			extra_info.page_logo_height = 10
		}
		if (imageData) {
			extra_info.page_logo_width = imageData.aspectRatio * extra_info.page_logo_height
		}
	}
	var logo_absolute_position = extra_info.logo_absolute_position && extra_info.logo_absolute_position.checked
	var hidden_page_header_border = extra_info.hidden_page_header_border && extra_info.hidden_page_header_border.checked
	var hidden_page_footer_border = extra_info.hidden_page_footer_border && extra_info.hidden_page_footer_border.checked
	var options = {
		header: {
			text: extra_info.page_header_text,
			font_family: extra_info.page_header_family,
			font_bold: extra_info.page_header_bold_font,
			font_size: extra_info.page_header_font_size || 3.71,
			align: extra_info.page_header_position,
			line_visible: hidden_page_header_border === true ? false : true,
			image_url: extra_info.page_image_url,
			image_height: extra_info.page_logo_height,
			image_width: extra_info.page_logo_width,
			image_x: logo_absolute_position ? extra_info.logo_x : null,
			image_y: logo_absolute_position ? extra_info.logo_y : null
		},
		footer: {
			text: extra_info.page_footer_text,
			font_family: extra_info.page_footer_family,
			font_bold: extra_info.page_footer_bold_font,
			font_size: extra_info.page_footer_font_size || 3.71,
			align: extra_info.page_footer_position,
			line_visible: hidden_page_footer_border === true ? false : true,
		},
		pagination: extra_info.onlyoffice_options.pagination,
		page_type: window.BiyueCustomData.page_type,
		imageDimensionsCache: imageDimensionsCache
	}
	if (options.page_type == 0) {
		options.stat = Object.assign({}, extra_info.onlyoffice_options.statis, {
			width: 4.76,
			height: 4.76,
			url: STAT_URL
		})
	}
	return drawHeaderFooter(options, recalc)
}
export { initFeature, initExtroInfo, syncInteractionWhenReSplit }
