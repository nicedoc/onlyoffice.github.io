import { getPaperInfo } from '../business.js'
import { ZONE_TYPE, ZONE_TYPE_NAME } from '../model/feature.js'
import { isTextMode } from '../model/ques.js'
import { MARK_METHOD_TYPE, UPLOAD_VALIDATE_RESULT } from '../model/uploadValidateEnum.js'
let numericFields = ['x', 'y', 'w', 'h', 'page']
// 这个文件用于处理上传前的校验
function onValidate() {
	var paper_info = getPaperInfo() || {}
	if (paper_info && paper_info.workbook && paper_info.workbook.id > 0) {
		// 练习册课时
		return onCheckPositions()
	} else {
		// todo..散页
	}
	return true
}

function notifyError(result) {
	if (result) {
		result.exam_info = getExamInfo()
	}
	Asc.scope.upload_validate = result
	window.biyue.refreshDialog({
		winName:'uploadValidation',
		name:'上传检查',
		url:'uploadValidation.html',
		width:400,
		height:800,
		isModal:false,
		type:'panelRight',
		icons:['resources/light/check.png']
	}, 'uploadValidationMessage', {validate_info: result})
}

function onCheckPositions() {
	var page_size = ''
	var paper_info = getPaperInfo() || {}
	if (paper_info && paper_info.workbook && paper_info.workbook.layout) {
		page_size = paper_info.workbook.layout
	} else {
		notifyError({
			code: UPLOAD_VALIDATE_RESULT.NOT_WORKBOOK_LAYOUT
		})
		return false
	}
	var preQuestionPositions = getPositions()
	if (window.BiyueCustomData.page_type == 0) {
		if (Object.keys(preQuestionPositions).length === 0) {
			notifyError({
				code: UPLOAD_VALIDATE_RESULT.NOT_QUESTION
			})
			return false
		}
		var find = Object.keys(preQuestionPositions).find(e => {
			return !preQuestionPositions[e].ref_id
		})
		if (find) {
			notifyError({
				code: UPLOAD_VALIDATE_RESULT.MISS_QUESTION_UUID
			})
			return false
		}
		var res_check_postion = checkPositions(preQuestionPositions) || {}
		var featureResult = checkEvaluationPosition()
		if (Object.keys(res_check_postion).length || featureResult.length) {
			notifyError({
				code: UPLOAD_VALIDATE_RESULT.QUESTION_ERROR,
				detail: res_check_postion,
				feature: featureResult
			})
			return false
		}
	}
	var preEvaluationPosition = getEvaluationPosition(page_size)
	if (window.BiyueCustomData.page_type == 0) {
		for (const key in preQuestionPositions) {
			let item = preQuestionPositions[key]
			if (item.content) {
				delete item.content
			}
			if (item.text) {
				delete item.text
			}
		}
		// 过滤题目选区里的字段类型,如果是数值类型则还会去除小数位
		const need_formatted = [
			'title_region',
			'write_ask_region',
			'mark_ask_region',
			'correct_region',
			'correct_ask_region',
		]
		formattedPositions(preQuestionPositions, need_formatted)
	}
	console.log('本地检查通过')
	window.biyue.closeDialog('uploadValidation')
	Asc.scope.preQuestionPositions = preQuestionPositions
	Asc.scope.preEvaluationPosition = preEvaluationPosition
	// showExportDialog()
	return true
}

function formattedPositions(positions, fieldsToCheck) {
    for (const key in positions) {
      const item = positions[key]
      fieldsToCheck.forEach(field => {
        if (field in item) {
          positions[key][field] = convertDataToNumber(positions[key][field], numericFields)
        }
      })
    }
  }
function getExamInfo() {
	let info = {
		struct_count: 0,
		ques_count: 0
	}
	var typeMaps = {}
	var types = window.BiyueCustomData.paper_options.question_type || []
	var types1 = {}
	types.forEach(e => {
		types1[e.value + ''] = e.label
	})
	var question_map = window.BiyueCustomData.question_map || {}
	for (const key in question_map) {
		if (question_map[key].level_type == 'struct') {
			info.struct_count++
		} else if (question_map[key].level_type == 'question') {
			info.ques_count++
			if (!typeMaps[question_map[key].question_type]) {
				typeMaps[question_map[key].question_type] = {
					count: 1,
					type_name: types1[question_map[key].question_type + '']
				}
			} else {
				typeMaps[question_map[key].question_type].count += 1
			}
		}
	}
	var typelist = Object.values(typeMaps).map(e => {
		return `${e.type_name}${e.count}题`
	})
	info.type_info = typelist.join('、')
	return info
}

function getPositions() {
	let positions = {}
	var questionPositions = Asc.scope.questionPositions || {}
	let ques_list = questionPositions.ques_list || []
	let index = 1
	for (const key in ques_list) {
		let item = ques_list[key] || ''
		if (item && item.ref_id && !isTextMode(item.ques_mode)) {
			if (!positions[item.ref_id]) {
				positions[item.ref_id] = {}
			}
			item.ques_no = index++
			positions[item.ref_id] = item
		}
	}
	return positions
}

function getQuesObject(obj) {
	return {
		ques_id: obj.id,
		ques_name: obj.ques_name,
		text: obj.text,
	}
}

function checkPositions(postions) {
	const title_region_err = []
	const write_ask_region_err = []
	const correct_ask_region_err = []
	const ques_score_err = []
	const ask_score_err = []
	const ask_score_empty_err = []
	const ques_type_err = []

	for (const key in postions) {
		let score = parseFloat(postions[key].score) || 0
		if ((!score || score <= 0) && !isTextMode(postions[key].ques_mode)) {
			// 文本题不需要分数判断
			ques_score_err.push(getQuesObject(postions[key]))
		}
		if (!postions[key].ques_type) {
			ques_type_err.push(getQuesObject(postions[key]))
		}
		if (checkRegionEmpty(postions[key], 'title_region')) {
			title_region_err.push(getQuesObject(postions[key]))
		}
		if (
			checkRegionEmpty(postions[key], 'write_ask_region', false) ||
			!postions[key]['write_ask_region'] ||
			postions[key]['write_ask_region'].length === 0
		) {
			write_ask_region_err.push(getQuesObject(postions[key]))
		}
		if (checkRegionEmpty(postions[key], 'correct_ask_region', false)) {
			correct_ask_region_err.push(getQuesObject(postions[key]))
		}
		postions[key].score = parseFloat(postions[key].score) || 0
		if (postions[key].mark_method == MARK_METHOD_TYPE.CHOICE_OPTION) {
			postions[key].ask_num = 1
		} else if (
			postions[key].mark_ask_region &&
			postions[key].mark_method != MARK_METHOD_TYPE.CHOICE_OPTION
		) {
			const mark_ask_region = postions[key].mark_ask_region || {}
			postions[key].ask_num = Object.keys(mark_ask_region).length

			if (postions[key].ask_num > 1) {
				const question_ask = {}
				if (postions[key].mark_method == MARK_METHOD_TYPE.SCORE) {
					// 主观题
				} else {
					const positionsObj = postions[key].mark_ask_region
					let ask_score_sum = 0
					let has_ask_score_empty_err = false
					for (const k in positionsObj) {
						if (positionsObj[k] && positionsObj[k][0]) {
							question_ask[positionsObj[k][0].order] = {
								order: positionsObj[k][0].order + '',
								score: positionsObj[k][0].v + '', // 因为以前的历史原因 小问的分数需要字符串类型
							}
							if (postions[key].score) {
								//  检查小问的分数和题目的分数是否一致
								let ask_score = parseFloat(positionsObj[k][0].v) || 0
								ask_score_sum += ask_score
								if (!ask_score || ask_score <= 0) {
									has_ask_score_empty_err = true
								}
							}
						}
					}
					if (
						ask_score_sum * 1 !== postions[key].score * 1 &&
						postions[key].score
					) {
						//  小问分数和题目分数不一致
						ask_score_err.push(getQuesObject(postions[key]))
					}
					if (has_ask_score_empty_err) {
						// 题目有分数的情况下，有小问的分数为空
						ask_score_empty_err.push(getQuesObject(postions[key]))
					}
					postions[key].question_ask = question_ask
				}
			}
		}
	}
	let obj = {}
	if (ques_score_err.length > 0) {
		obj.ques_score_err = ques_score_err
	}
	if (ask_score_err.length > 0) {
		obj.ask_score_err = ask_score_err
	}
	if (ask_score_empty_err.length > 0) {
		obj.ask_score_empty_err = ask_score_empty_err
	}
	if (ques_type_err.length > 0) {
		obj.ques_type_err = ques_type_err
	}
	if (title_region_err.length > 0) {
		obj.title_region_err = title_region_err
	}
	if (write_ask_region_err.length > 0) {
		obj.write_ask_region_err = write_ask_region_err
	}
	if (correct_ask_region_err.length > 0) {
		obj.correct_ask_region_err = correct_ask_region_err
	}
	return obj
}

function checkRegionEmpty(postions, name, checkValue) {
    const ret = false
    if (!postions[name]) {
      return true
    }
    const arr = postions[name] || []
    for (const key in arr) {
      const w = arr[key].w || 0
      const h = arr[key].h || 0
      const v = arr[key].v || 0
      if (w * 1 + h * 1 === 0 || checkValue && v * 1 === 0) {
        return true
      }
    }
    return ret
  }

function getFieldsByZoneType(zoneType) {
	var questionPositions = Asc.scope.questionPositions || {}
	// 遍历 feature_list 查找匹配的 zone_type
	let result = zoneType === ZONE_TYPE_NAME[ZONE_TYPE.PASS] ? {} : []

	for (const feature of questionPositions.feature_list) {
		if (
			zoneType === ZONE_TYPE_NAME[ZONE_TYPE.PASS] &&
			feature.zone_type === zoneType
		) {
			// pass区域需要传对象的结构
			return feature.fields && feature.fields.length > 0
				? feature.fields[0]
				: {}
		} else if (feature.zone_type === zoneType) {
			return feature.fields
		}
	}
	return result
}
// 功能区判断
function checkEvaluationPosition() {
	var failList = []
	if (window.BiyueCustomData.page_type == 0) {
		if (!window.BiyueCustomData.workbook_info) {
			return false
		}
		var parse_extra_data = window.BiyueCustomData.workbook_info.parse_extra_data
		if (!parse_extra_data) {
			return false
		}
		if (parse_extra_data.practise_again && 
			parse_extra_data.practise_again.switch && 
			!hasEvaluationPosition(ZONE_TYPE_NAME[ZONE_TYPE.AGAIN])) {
			failList.push('再练区')
		}
		if (parse_extra_data.workbook_qr_code_show && !hasEvaluationPosition(ZONE_TYPE_NAME[ZONE_TYPE.QRCODE])) {
			failList.push('二维码')
		}
		if (parse_extra_data.custom_evaluate) {
			if (!hasEvaluationPosition(ZONE_TYPE_NAME[ZONE_TYPE.SELF_EVALUATION])) {
				failList.push('学生评价')
			}
			if (!hasEvaluationPosition(ZONE_TYPE_NAME[ZONE_TYPE.THER_EVALUATION])) {
				failList.push('教师评价')
			}
			if (!hasEvaluationPosition(ZONE_TYPE_NAME[ZONE_TYPE.PASS])) {
				failList.push('通过区')
			}
		} else if (parse_extra_data.hiddenComplete && 
			parse_extra_data.hiddenComplete.checked === false &&
			!hasEvaluationPosition(ZONE_TYPE_NAME[ZONE_TYPE.END])
		) {
			failList.push('完成区')
		}
		if (parse_extra_data.onlyoffice_options && 
			parse_extra_data.onlyoffice_options.statis &&
			!hasEvaluationPosition(ZONE_TYPE_NAME[ZONE_TYPE.STATISTICS])) {
			failList.push('统计区')
		}
		if (parse_extra_data.onlyoffice_options && 
			parse_extra_data.onlyoffice_options.pagination &&
			!hasEvaluationPosition(ZONE_TYPE_NAME[ZONE_TYPE.PAGINATION])) {
			failList.push('页码区')
		}
		if (!hasEvaluationPosition(ZONE_TYPE_NAME[ZONE_TYPE.IGNORE])) {
			failList.push('日期评语')
		}
	}
	return failList
}
// 是否有功能区，这里先只针对功能区存在进行判断，也许以后还要进行位置判断
function hasEvaluationPosition(zone_type) {
	var questionPositions = Asc.scope.questionPositions || {}
	if (questionPositions.feature_list) {
		for (var feature of questionPositions.feature_list) {
			if (feature.zone_type == zone_type) {
				return true
			}
		}
	}
	return false
}
function getEvaluationPosition(page_size) {
	var questionPositions = Asc.scope.questionPositions || {}
	var evaluationPosition = {
		partial_no_dot_regional: questionPositions.partical_no_dot_list || [],
	}
	if (window.BiyueCustomData.page_type == 0) {
		evaluationPosition = Object.assign(evaluationPosition, {
			self_evaluation: getFieldsByZoneType(ZONE_TYPE_NAME[ZONE_TYPE.SELF_EVALUATION]),
			teacher_evaluation: getFieldsByZoneType(ZONE_TYPE_NAME[ZONE_TYPE.THER_EVALUATION]),
			pass_regional: getFieldsByZoneType(ZONE_TYPE_NAME[ZONE_TYPE.PASS]),
			ignore_region: getFieldsByZoneType(ZONE_TYPE_NAME[ZONE_TYPE.IGNORE]),
			end_regional: getFieldsByZoneType(ZONE_TYPE_NAME[ZONE_TYPE.END]),
			again_regional: getFieldsByZoneType(ZONE_TYPE_NAME[ZONE_TYPE.AGAIN]),
			stat_regional: getFieldsByZoneType(ZONE_TYPE_NAME[ZONE_TYPE.STATISTICS]),
		})
	}
	// 处理选区的数据类型
	for (const key in evaluationPosition) {
		evaluationPosition[key] = convertDataToNumber(
			evaluationPosition[key],
			numericFields
		)
	}

	evaluationPosition['page_size'] = page_size
	evaluationPosition['exam_type'] = 'exercise'
	return evaluationPosition
}

function convertDataToNumber(data, numericFields) {
	// 判断是否为数字的函数
	const isNumber = (value) => !isNaN(parseFloat(value)) && isFinite(value)

	// 转换对象数组中的数据类型
	const convertObjectArray = (arr) => {
		arr.forEach((item) => {
			for (const prop in item) {
				if (numericFields.includes(prop)) {
					item[prop] =
						isNumber(item[prop]) && item[prop] > 0 ? Math.floor(item[prop]) : 0
				} else {
					item[prop] = String(item[prop])
				}
			}
		})
	}
	// 遍历数据对象或数组
	const convert = (item) => {
		for (const key in item) {
			if (
				Array.isArray(item[key]) &&
				item[key].length > 0 &&
				typeof item[key][0] === 'object'
			) {
				convertObjectArray(item[key])
			} else {
				if (numericFields.includes(key)) {
					item[key] =
						isNumber(item[key]) && item[key] > 0 ? Math.floor(item[key]) : 0
				} else {
					item[key] = String(item[key])
				}
			}
		}
	}

	if (Array.isArray(data)) {
		data.forEach((item) => convert(item))
	} else {
		convert(data)
	}

	return data
}
function showExportDialog() {
	window.biyue.showDialog(
		'exportExamWindow',
		'上传试卷',
		'examExport.html',
		1000,
		800,
		true
	)
}

export default {
	onValidate,
}
