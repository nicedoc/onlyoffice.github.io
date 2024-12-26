var ZONE_TYPE = {
	AGAIN: 15, // 再练区域
	STATISTICS: 14, // 统计
	PASS: 11, // 通过
	SELF_EVALUATION: 9, // 自定义评价
	THER_EVALUATION: 10, // 教师评价
	END: 16, // 完成批改
	IGNORE: 18, // 忽略区
	CHECK: 28, // 检查区
	SEALING_LINE: 100, // 密封线
	QRCODE: 200, // 二维码
	PAGINATION: 201, // 页码
}

var ZONE_TYPE_NAME = {
	[`${ZONE_TYPE.AGAIN}`]: 'again',
	[`${ZONE_TYPE.END}`]: 'end',
	[`${ZONE_TYPE.PASS}`]: 'pass',
	[`${ZONE_TYPE.SELF_EVALUATION}`]: 'self_evaluation',
	[`${ZONE_TYPE.THER_EVALUATION}`]: 'teacher_evaluation',
	[`${ZONE_TYPE.STATISTICS}`]: 'statistics',
	[`${ZONE_TYPE.IGNORE}`]: 'ignore',
	[`${ZONE_TYPE.CHECK}`]: 'check',
	[`${ZONE_TYPE.SEALING_LINE}`]: 'sealing_line',
	[`${ZONE_TYPE.QRCODE}`]: 'qr_code',
	[`${ZONE_TYPE.PAGINATION}`]: 'pagination',
}

var ZONE_SIZE = {
	[`${ZONE_TYPE.AGAIN}`]: {
		w: 11.011764705882353,
		h: 6.247058823529412, 
		font_size: 18,
		shape_type: 'rect',
		stroke_width: 0.05,
	},
	[`${ZONE_TYPE.END}`]: {
		w: 15.236470588235296,
		h: 7.9411764705882355,
		font_size: 18,
		shape_type: 'rect',
		stroke_width: 0.05,
	},
	[`${ZONE_TYPE.PASS}`]: {
		w: 15.236470588235296,
		h: 7.9411764705882355,
		font_size: 18,
		shape_type: 'rect',
		stroke_width: 0.05,
	},
	[`${ZONE_TYPE.SELF_EVALUATION}`]: {
		w: 8,
		h: 9,
		font_size: 20,
		flower_size: 24 * 0.25,
		icon_size: 21.33 * 0.25,
	},
	[`${ZONE_TYPE.THER_EVALUATION}`]: {
		w: 8,
		h: 9,
		font_size: 20,
		flower_size: 24 * 0.25,
		icon_size: 21.33 * 0.25,
	},
	[`${ZONE_TYPE.STATISTICS}`]: {
		w: 4.76,
		h: 4.76,
	},
	[`${ZONE_TYPE.CHECK}`]: {
		w: 15.236470588235296,
		h: 7.9411764705882355,
		font_size: 18,
		shape_type: 'rect',
		stroke_width: 0.05,
	},
	[`${ZONE_TYPE.IGNORE}`]: {
		w: 40,
		h: 8,
		font_size: 16,
		jc: 'left',
		shape_type: 'roundRect',
		stroke_width: 0.12,
		sr: 153,
		sg: 153,
		sb: 153,
	},
	[`${ZONE_TYPE.QRCODE}`]: {
		w: 45 * 0.25,
		h: 45 * 0.25,
		font_size: 14,
		imgSize: 35 * 0.25
	},
}
// 是否有互动
function hasInteraction() {
	if (window.BiyueCustomData && 
		window.BiyueCustomData.workbook_info && 
		window.BiyueCustomData.workbook_info.parse_extra_data &&
		window.BiyueCustomData.workbook_info.parse_extra_data.hidden_correct_region
	) {
		if (window.BiyueCustomData.workbook_info.parse_extra_data.hidden_correct_region.checked === false) {
			return true
		}
	}
	return false
}
// 获取互动类型列表
function getInteractionTypes() {
	var interactionTypes = [{
		value: 'none',
		label: '无互动'
	}]
	var extra_info = window.BiyueCustomData.workbook_info.parse_extra_data
	if (extra_info && extra_info.hidden_correct_region) {
		if (!extra_info.hidden_correct_region.checked) {
			interactionTypes.push({
				value: 'simple',
				label: '简单互动'
			})
			if (extra_info.start_interaction.checked) {
				interactionTypes.push({
					value: 'accurate',
					label: '精准互动'
				})
			}
		}
	}
	return interactionTypes
}

export { ZONE_TYPE, ZONE_SIZE, ZONE_TYPE_NAME, hasInteraction, getInteractionTypes }
