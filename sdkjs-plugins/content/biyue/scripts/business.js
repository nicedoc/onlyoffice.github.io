// 主要用于处理业务逻辑
import {
	reqPaperInfo,
	structAdd,
	questionCreate,
	questionDelete,
	questionUpdateContent,
	structDelete,
	structRename,
	reqSubjectMarkTypes
} from './api/paper.js'
import { getBase64, map_base64 } from '../resources/list_base64.js'
import { biyueCallCommand } from './command.js'
import { handlePaperInfoResult } from './pageView.js'
let paper_info = {} // 从后端返回的试卷信息
let select_ques_ids = []
const MM2EMU = 36000 // 1mm = 36000EMU
// 根据paper_uuid获取试卷信息
function initPaperInfo() {
	return new Promise((resolve, reject) => {
		reqPaperInfo(window.BiyueCustomData.paper_uuid)
			.then((res) => {
				handlePaperInfoResult(true, res)
				paper_info = res.data
				if (!paper_info) {
					return
				}
				const { paper, workbook } = res.data
				if (workbook) {
					$('#workbookName').text(
						`${workbook.school_id ? workbook.school_name : '内部练习册'}`
					)
					$('#workbookId').text(`#${workbook.id}`)
					window.BiyueCustomData.workbook_info = workbook
				}
				if (paper) {
					$('#exam_title').text(`《${paper.title}》`)
					window.BiyueCustomData.exam_title = paper.title
					$('#grade_data').text(
						`${paper.period_name}${paper.subject_name} / ${
							paper.edition_name || ''
						} / ${paper.phase_name || ''}`
					)
				}
				var options = res.data.options || {}
				var paper_options = {}
				Object.keys(options).forEach((key) => {
					var list = []
					Object.keys(options[key]).forEach((e) => {
						list.push({
							value: e,
							label: options[key][e],
						})
					})
					paper_options[key] = list
				})
				window.BiyueCustomData.paper_options = paper_options
				return reqSubjectMarkTypes(paper.subject_value).then(res3 => {
					Asc.scope.subject_mark_types = res3.data
					resolve(res)
				}).catch(res => {
					reject(res)
				})
			})
			.catch((res) => {
				handlePaperInfoResult(false, res)
				reject(res)
				console.log('catch', res)
			})
	})
}

function updatePageSizeMargins() {
	Asc.scope.workbook = window.BiyueCustomData.workbook_info
	Asc.scope.control_hightlight = true
	return biyueCallCommand(
		window,
		function () {
			var workbook = Asc.scope.workbook || {}
			var oDocument = Api.GetDocument()
			var sections = oDocument.GetSections()
			function MM2Twips(mm) {
				var m = Math.max(mm, 10)
				return m / (25.4 / 72 / 20)
			}
			if (sections && sections.length > 0) {
				sections.forEach((oSection) => {
					if (workbook.page_size) {
						oSection.SetPageSize(
							MM2Twips(workbook.page_size.width),
							MM2Twips(workbook.page_size.height)
						)
					}
					if (workbook.margin) {
						oSection.SetPageMargins(
							MM2Twips(workbook.margin.left),
							MM2Twips(workbook.margin.top),
							MM2Twips(workbook.margin.right),
							MM2Twips(workbook.margin.bottom)
						)
						oSection.SetFooterDistance(MM2Twips(workbook.margin.bottom))
						oSection.SetHeaderDistance(MM2Twips(workbook.margin.top))
					}
					oSection.RemoveFooter('default')
					oSection.RemoveFooter('even')
					oSection.RemoveHeader('default')
				})
			}
			var odrawings = oDocument.GetAllDrawingObjects() || []
			odrawings.forEach(oDrawing => {
				if (oDrawing.Drawing && oDrawing.Drawing.docPr && oDrawing.Drawing.docPr.descr && oDrawing.Drawing.docPr.descr.indexOf('biyue') == -1) {
					oDrawing.Drawing.Set_Props({
						description: ''
					})
				}
				if (oDrawing.GetParentTable()) {
					var oParagraph = oDrawing.GetParentParagraph()
					if (oParagraph) {
						var linespacing = oParagraph.GetSpacingLineValue()
						oParagraph.SetSpacingLine(linespacing, 'atLeast')
					}
				}
			})
			Api.asc_SetGlobalContentControlShowHighlight(true, 255, 191, 191)
			return null
		},
		false,
		true
	)
}

function getPaperInfo() {
	return paper_info
}

function p_Twips2MM(twips) {
	return (25.4 / 72 / 20) * twips
}
function p_MM2Twips(mm) {
	return mm / (25.4 / 72 / 20)
}
function p_EMU2MM(EMU) {
	return EMU / 36e3
}
function p_MM2EMU(mm) {
	return mm * 36e3
}

function updateQeusTree() {
	var listElement = document.getElementById('ques-list')
	if (!listElement) {
		return
	}
	var control_list = window.BiyueCustomData.control_list || []
	// var html = ''
	// control_list.forEach(e => {
	//   if (e.regionType == 'struct') {
	//     html += `<div class="quesitem" id=${e.control_id} draggable="true"><div class="tag-struct"></div><span class="text-struct">${e.name}</span></div>`
	//   } else if (e.regionType == 'question') {
	//     html += `<div class="quesitem" id=${e.control_id} draggable="true"><div class="tag-ques"></div><span class="text-ques">${e.text}</span></div></div>`
	//   } else if (e.regionType == 'sub-question') {
	//     html += `<div class="quesitem" id=${e.control_id} draggable="true"><div class="tag-sub-ques"></div><span class="text-ques-sub">${e.text}</span></div></div>`
	//   }
	// })
	// listElement.innerHTML = html
	// $('.quesitem').on('click', onQuesTreeClick)
	// $('.quesitem').on('dragstart', onTreeDragStart)
	// $('.quesitem').on('dragover', onDragOver)
	// $('.quesitem').on('drop', function(e) {
	//   console.log('drop', e)
	//   e.preventDefault();
	// })
	// listElement.addEventListener('drop', (e) => {
	//   console.log('list drop', e)
	//   e.preventDefault();
	// })
}

function onTreeDragStart(e) {
	console.log('onTreeDragStart', e)
	e.originalEvent.dataTransfer.setData('drag_data', e.target.id)
}

function onDragOver(e) {
	e.preventDefault()
	e.cancelBubble = true
	if (e.dataTransfer) {
		e.dataTransfer.dropEffect = 'move'
	} else {
		e.originalEvent.dataTransfer.dropEffect = 'move'
	}
}

function onQuesTreeClick(e) {
	console.log('onQuesTreeClick', e)
	var id
	if (e.target && e.target.id && e.target.id != '') {
		id = e.target.id
	} else if (e.currentTarget) {
		id = e.currentTarget.id
	}
	var control_list = window.BiyueCustomData.control_list || []
	var controlData = control_list.find((item) => {
		return item.control_id == id
	})
	if (!controlData) {
		console.log('onQuesTreeClick cannot find ', id)
		return
	}
	var ctrlKey = e.ctrlKey
	var newlist = []
	if (ctrlKey) {
		newlist = [].concat(select_ques_ids)
		var index = newlist.indexOf(id)
		if (index >= 0) {
			// 原本已存在，取消选中
			newlist.splice(index, 1)
		} else {
			newlist.push(id)
		}
	} else {
		newlist.push(id)
		var event = new CustomEvent('clickSingleQues', {
			detail: {
				control_id: id,
				regionType: 'question',
			},
		})
		document.dispatchEvent(event)
	}
	updateQuesStyle(newlist)
	Asc.scope.click_ids = newlist
	biyueCallCommand(
		window,
		function () {
			var ids = Asc.scope.click_ids
			var oDocument = Api.GetDocument()
			oDocument.RemoveSelection()
			var controls = oDocument.GetAllContentControls()
			var firstRange = null
			ids.forEach((id, index) => {
				var control = controls.find((e) => {
					return e.Sdt.GetId() == id
				})
				if (control) {
					if (index == 0) {
						firstRange = control.GetRange()
					} else {
						var oRange = control.GetRange()
						firstRange = firstRange.ExpandTo(oRange)
					}
				}
			})
			firstRange.Select()
		},
		false,
		false
	)
}
// 更新题目选中样式
function updateQuesStyle(idList) {
	console.log('updateQuesStyle', idList)
	for (var i = 0; i < select_ques_ids.length; ++i) {
		if (idList.indexOf(select_ques_ids[i]) == -1) {
			$('#' + select_ques_ids[i]).removeClass('selected')
		}
	}
	for (var j = 0; j < idList.length; ++j) {
		$('#' + idList[j]).addClass('selected')
	}
	select_ques_ids = idList
}

function updateControls() {
	Asc.scope.paper_info = paper_info
	return biyueCallCommand(
		window,
		function () {
			let paperinfo = Asc.scope.paper_info
			var oDocument = Api.GetDocument()
			let controls = oDocument.GetAllContentControls() || []
			let ques_no = 1
			let struct_index = 0
			var control_list = []
			function getJsonData(str) {
				if (!str || str == '' || typeof str != 'string') {
					return {}
				}
				try {
					return JSON.parse(str)
				} catch (error) {
					console.log('json parse error', error)
					return {}
				}
			}
			controls.forEach((control) => {
				var rect = Api.asc_GetContentControlBoundingRect(
					control.Sdt.GetId(),
					true
				)
				tagInfo = getJsonData(control.GetTag())
				var text = control.GetRange().GetText()
				let obj = {
					control_id: control.Sdt.GetId(),
					regionType: tagInfo.regionType,
					text: text,
				}
				if (tagInfo.regionType == 'struct') {
					++struct_index
					const pattern = /^[一二三四五六七八九十0-9]+.*?(?=[：:])/
					const result = pattern.exec(text)
					obj.name = result ? result[0] : null
					if (
						paperinfo.ques_struct_list &&
						struct_index - 1 < paperinfo.ques_struct_list.length
					) {
						obj.struct_id =
							paperinfo.ques_struct_list[struct_index - 1].struct_id
					}
				} else if (tagInfo.regionType == 'question') {
					obj.ques_no = ques_no
					const regex = /^([^.．、]*)/
					const match = obj.text.match(regex)
					obj.ques_name = match ? match[1] : ''
					if (
						paperinfo.info &&
						paperinfo.info.questions &&
						ques_no <= paperinfo.info.questions.length
					) {
						obj.score = paperinfo.info.questions[ques_no - 1].score
						obj.ques_uuid = paperinfo.info.questions[ques_no - 1].uuid
						obj.ask_controls = []
						if (
							paperinfo.ques_struct_list &&
							struct_index - 1 < paperinfo.ques_struct_list.length
						) {
							obj.struct_id =
								paperinfo.ques_struct_list[struct_index - 1].struct_id
						}
					}
					ques_no++
				} else if (
					tagInfo.regionType == 'write' ||
					tagInfo.regionType == 'sub-question'
				) {
					let parentContentControl = control.GetParentContentControl()
					if (!parentContentControl) {
						console.log('parentContentControl is null')
					}
					if (parentContentControl) {
						obj.parent_control_id = parentContentControl.Sdt.GetId()
					}
					function getParentQues(id) {
						for (var i = 0, imax = control_list.length; i < imax; ++i) {
							if (control_list[i].control_id == id) {
								if (control_list[i].regionType == 'question') {
									return control_list[i]
								} else {
									return getParentQues(control_list[i].parent_control_id)
								}
							}
						}
						return null
					}
					var parentQues = getParentQues(obj.parent_control_id)
					if (parentQues) {
						obj.ques_uuid = parentQues.ques_uuid
						obj.parent_ques_control_id = parentQues.control_id
						if (tagInfo.regionType == 'write') {
							if (!parentQues.ask_controls) {
								parentQues.ask_controls = []
							}
							parentQues.ask_controls.push({
								control_id: control.Sdt.GetId(),
								v: 0,
							})
						} else {
							if (!parentQues.sub_questions) {
								parentQues.sub_questions = []
							}
							parentQues.sub_questions.push({
								control_id: control.Sdt.GetId(),
							})
						}
					}
				} else if (tagInfo.regionType == 'feature') {
					obj.zone_type = tagInfo.zone_type
					obj.v = tagInfo.v
				}
				control_list.push(obj)
			})
			console.log(' updatecontrol           control_list', control_list)
			return control_list
		},
		false,
		false
	)
}

// 更新customData的control_list
function updateCustomControls() {
	Asc.scope.paper_info = paper_info
	return biyueCallCommand(
		window,
		function () {
			console.log('+++++++++++++++++++++++')
			let paperinfo = Asc.scope.paper_info
			var oDocument = Api.GetDocument()
			let controls = oDocument.GetAllContentControls() || []
			let ques_no = 1
			let struct_index = 0
			var control_list = []
			function getJsonData(str) {
				if (!str || str == '' || typeof str != 'string') {
					return {}
				}
				try {
					return JSON.parse(str)
				} catch (error) {
					console.log('json parse error', error)
					return {}
				}
			}
			controls.forEach((control) => {
				// var rect = Api.asc_GetContentControlBoundingRect(
				// 	control.Sdt.GetId(),
				// 	true
				// )
				let tagInfo = getJsonData(control.GetTag())
				var text = control.GetRange().GetText()
				let obj = {
					control_id: control.Sdt.GetId(),
					regionType: tagInfo.regionType,
					text: text,
				}
				if (tagInfo.regionType == 'struct') {
					++struct_index
					const pattern = /^[一二三四五六七八九十0-9]+.*?(?=[：:])/
					const result = pattern.exec(text)
					obj.name = result ? result[0] : null
					if (
						paperinfo.ques_struct_list &&
						struct_index - 1 < paperinfo.ques_struct_list.length
					) {
						obj.struct_id =
							paperinfo.ques_struct_list[struct_index - 1].struct_id
					}
				} else if (tagInfo.regionType == 'question') {
					obj.ques_no = ques_no
					const regex = /^([^.．、]*)/
					const match = obj.text.match(regex)
					obj.ques_name = match ? match[1] : ''
					if (
						paperinfo.info &&
						paperinfo.info.questions &&
						ques_no <= paperinfo.info.questions.length
					) {
						obj.score = paperinfo.info.questions[ques_no - 1].score
						obj.ques_uuid = paperinfo.info.questions[ques_no - 1].uuid
						obj.ask_controls = []
						if (
							paperinfo.ques_struct_list &&
							struct_index - 1 < paperinfo.ques_struct_list.length
						) {
							obj.struct_id =
								paperinfo.ques_struct_list[struct_index - 1].struct_id
						}
					}
					ques_no++
				} else if (
					tagInfo.regionType == 'write' ||
					tagInfo.regionType == 'sub-question'
				) {
					let parentContentControl = control.GetParentContentControl()
					if (!parentContentControl) {
						console.log('parentContentControl is null')
					}
					if (parentContentControl) {
						obj.parent_control_id = parentContentControl.Sdt.GetId()
					}
					function getParentQues(id) {
						for (var i = 0, imax = control_list.length; i < imax; ++i) {
							if (control_list[i].control_id == id) {
								if (control_list[i].regionType == 'question') {
									return control_list[i]
								} else {
									return getParentQues(control_list[i].parent_control_id)
								}
							}
						}
						return null
					}
					var parentQues = getParentQues(obj.parent_control_id)
					if (parentQues) {
						obj.ques_uuid = parentQues.ques_uuid
						obj.parent_ques_control_id = parentQues.control_id
						if (tagInfo.regionType == 'write') {
							if (!parentQues.ask_controls) {
								parentQues.ask_controls = []
							}
							parentQues.ask_controls.push({
								control_id: control.Sdt.GetId(),
								v: 0,
							})
						} else {
							if (!parentQues.sub_questions) {
								parentQues.sub_questions = []
							}
							parentQues.sub_questions.push({
								control_id: control.Sdt.GetId(),
							})
						}
					}
				} else if (tagInfo.regionType == 'feature') {
					obj.zone_type = tagInfo.zone_type
					obj.v = tagInfo.v
				}
				control_list.push(obj)
			})
			console.log(' updatecontrol           control_list', control_list)
			return control_list
		},
		false,
		false
	).then((res) => {
		window.BiyueCustomData.control_list = res
	})
}

// 清除试卷结构和所有题目
async function clearStruct() {
	console.log('开始清除结构')
	if (paper_info.info && paper_info.info.questions) {
		for (const e of paper_info.info.questions) {
			await questionDelete(window.BiyueCustomData.paper_uuid, e.uuid, 1)
		}
	}
	if (paper_info.ques_struct_list) {
		for (const estruct of paper_info.ques_struct_list) {
			await structDelete(window.BiyueCustomData.paper_uuid, estruct.struct_id)
		}
	}
	console.log('清除结构成功')
	initPaperInfo()
}

// 获取试卷结构
async function getStruct() {
	if (!window.BiyueCustomData.control_list) {
		return
	}
	console.log('开始获取试卷结构')
	var struct_index = 0
	for (var control of window.BiyueCustomData.control_list) {
		if (control.regionType == 'struct') {
			++struct_index
			if (
				paper_info.ques_struct_list &&
				struct_index - 1 < paper_info.ques_struct_list.length
			) {
				control.struct_id =
					paper_info.ques_struct_list[struct_index - 1].struct_id
			}
			if (!control.struct_id) {
				await structAdd({
					paper_uuid: window.BiyueCustomData.paper_uuid,
					name: control.name,
					rich_name: control.text,
				}).then((res) => {
					if (res.code == 1) {
						control.struct_id = res.data.struct_id
						// control.struct_name = control.name
					}
				})
			} else if (
				control.name !=
				paper_info.ques_struct_list[struct_index - 1].struct_name
			) {
				await structRename(
					window.BiyueCustomData.paper_uuid,
					control.struct_id,
					control.name
				)
			}
		}
	}
	await getQuesUuid()
}

async function getQuesUuid() {
	for (const e of window.BiyueCustomData.control_list) {
		if (e.regionType == 'question') {
			var uuid = ''
			if (paper_info.info && paper_info.info.questions) {
				var ques = paper_info.info.questions.find((item) => {
					return item.no == e.ques_no
				})
				if (ques) {
					uuid = ques.uuid
				}
			}
			if (uuid == '') {
				await questionCreate({
					paper_uuid: window.BiyueCustomData.paper_uuid,
					content: encodeURIComponent(e.text), // e.text,
					blank: '',
					type: 1,
					score: 0,
					no: e.ques_no,
					struct_id: e.struct_id,
				})
			} else {
				await questionUpdateContent({
					paper_uuid: window.BiyueCustomData.paper_uuid,
					question_uuid: uuid,
					content: encodeURIComponent(e.text), // e.text,
				})
			}
		}
	}
	console.log('所有结构和题目都更新完')
	initPaperInfo()
}

function showQuestionTree() {
	const questionList = $('#questionList')
	if (questionList) {
		if (
			questionList.children() &&
			questionList.children().length > 0 &&
			questionList.children().length < 2
		) {
			questionList.toggle()
		} else {
			initTree()
		}
	}
}

function initTree() {
	const questionList = $('#questionList')
	var needRefreshStruct = false
	if (!paper_info.info) {
		needRefreshStruct = true
	}
	if (
		!needRefreshStruct &&
		window.BiyueCustomData.control_list &&
		window.BiyueCustomData.control_list.length > 0
	) {
		var find = window.BiyueCustomData.control_list.find((e) => {
			return (
				(e.regionType == 'struct' && e.struct_id == 0) ||
				(e.regionType == 'question' && !e.ques_uuid)
			)
		})
		if (find) {
			needRefreshStruct = true
		}
	}
	if (needRefreshStruct) {
		questionList.append(
			'<div style="color:#ff0000">试卷结构有异，请先更新试卷结构</div>'
		)
		return
	}
	if (
		!window.BiyueCustomData.control_list ||
		window.BiyueCustomData.control_list.length == 0
	) {
		questionList.append('<div><button id="refreshTree">刷新树</button></div>')
		$('#refreshTree').on('click', () => {
			if (
				window.BiyueCustomData.control_list &&
				window.BiyueCustomData.control_list.length > 0
			) {
				initTree()
			}
		})
		return
	}
	questionList.append(
		'<div><button id="addfield">添加作答题分数框</button></div>'
	)
	var question_types = [
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
	var structItem
	var structList
	window.BiyueCustomData.control_list.forEach((control, index) => {
		if (control.regionType == 'struct') {
			if (structList) {
				structItem.append(structList)
				questionList.append(structItem)
			}
			structItem = $('<div>').text(control.name)
			structList = $('<ul>')
			const select = $('<select>')
			question_types.forEach((type) => {
				const option = $('<option>').val(type.value).text(type.label)
				select.append(option)
			})

			select.val('0')

			select.on('change', function () {
				changeStructQuesType(control.struct_id, parseInt(select.val()))
			})

			structItem.append(select)
		} else if (control.regionType == 'question') {
			const questionItem = $('<div>')
			const select = $('<select>')
			const input = $('<input>')
			const unittext = $('<span>')

			question_types.forEach((type) => {
				const option = $('<option>').val(type.value).text(type.label)
				select.append(option)
			})

			select.attr('id', `select${control.control_id}`)
			select.val('0')
			select.on('change', function () {
				control.ques_type = parseInt(select.val())
				console.log(
					'window.BiyueCustomData.control_list',
					window.BiyueCustomData.control_list
				)
			})
			input.attr('id', `score${control.control_id}`)
			input
				.attr('type', 'text')
				.val(control.score === '' ? '（未设置）' : control.score)
			input.on('input', function () {
				control.score = input.val()
				input.attr('placeholder', control.score === '' ? '（未设置）' : '')
			})

			questionItem.text(`${control.ques_name}.`)
			unittext.text('分')
			questionItem.append(select)
			questionItem.append(input)
			questionItem.append(unittext)
			structList.append(questionItem)
			let str = ''
			if (control.sub_questions && control.sub_questions.length) {
				str += `含${control.sub_questions.length}小题，`
			}
			if (control.ask_controls && control.ask_controls.length) {
				str += `${control.ask_controls.length}小问`
			}
			if (str != '') {
				structList.append(`<div style="color:#999">(${str})</div>`)
			}
		}
	})
	if (structList) {
		structItem.append(structList)
		questionList.append(structItem)
	}
	$('#addfield').on('click', addScoreField)
}

function changeStructQuesType(struct_id, v) {
	console.log('changeStructQuesType', struct_id, v)
	window.BiyueCustomData.control_list.forEach((control) => {
		if (control.regionType == 'question' && control.struct_id == struct_id) {
			control.ques_type = v
			$(`#select${control.control_id}`).val(v)
		}
	})
}
// 添加分数框
function addScoreField(score, mode, layout, posall) {
	Asc.scope.control_list = window.BiyueCustomData.control_list
	Asc.scope.params = {
		score: parseFloat(score) || 0,
		mode: mode,
		layout: layout,
		scores: getScores(score, mode),
		posall: posall,
	}
	console.log('setup post task for addScoreField')
	biyueCallCommand(
		window,
		function () {
			var control_list = Asc.scope.control_list
			var controls = Api.GetDocument().GetAllContentControls()
			var params = Asc.scope.params
			var score = params.score
			var res = {
				add: false,
			}
			function getJsonData(str) {
				if (!str || str == '' || typeof str != 'string') {
					return {}
				}
				try {
					return JSON.parse(str)
				} catch (error) {
					console.log('json parse error', error)
					return {}
				}
			}
			for (var i = 0; i < controls.length; ++i) {
				var control = controls[i]
				var tag = getJsonData(control.GetTag())
				if (tag.regionType == 'question') {
					var control_id = control.Sdt.GetId()
					var controlIndex = control_list.findIndex((item) => {
						return (
							item.regionType == 'question' && item.control_id == control_id
						)
					})
					var controldata = control_list.find((item) => {
						return (
							item.regionType == 'question' && item.control_id == control_id
						)
					})
					var rect = Api.asc_GetContentControlBoundingRect(control_id, true)
					console.log('rect', rect)
					var width = rect.X1 - rect.X0
					var trips_width = width / (25.4 / 72 / 20)
					if (!controldata) {
						console.log(control_id, control_list)
					}
					if (controldata) {
						if (
							controldata.score_options &&
							controldata.score_options.run_id &&
							controldata.score_options.paragraph_id
						) {
							var paragraph = new Api.private_CreateApiParagraph(
								AscCommon.g_oTableId.Get_ById(
									controldata.score_options.paragraph_id
								)
							)
							if (paragraph) {
								for (var i = 0; i < paragraph.Paragraph.Content.length; ++i) {
									if (
										paragraph.Paragraph.Content[i].Id ==
										controldata.score_options.run_id
									) {
										return {
											control_id: control_id,
											add: false,
										}
									}
								}
							}
						}
						var scores = params.scores
						var cellcount = scores.length
						var cell_width_mm = 8
						var cell_height_mm = 8
						var cellwidth = cell_width_mm / (25.4 / 72 / 20)
						var cellHeight = cell_height_mm / (25.4 / 72 / 20)
						var maxTableWidth = trips_width / params.layout
						var rowcount = 1
						var columncount = cellcount
						if (maxTableWidth < cellcount * cellwidth) {
							// 需要换行
							rowcount = Math.ceil((cellcount * cellwidth) / maxTableWidth)
							columncount = Math.ceil(maxTableWidth / cellwidth)
						}
						var oDocument = Api.GetDocument()
						var oTable = Api.CreateTable(columncount, rowcount)
						oTable.SetWidth('twips', columncount * cellwidth)
						var oTableStyle = oDocument.CreateStyle('CustomTableStyle', 'table')
						var oTableStylePr =
							oTableStyle.GetConditionalTableStyle('wholeTable')
						oTable.SetTableLook(true, true, true, true, true, true)
						oTableStylePr.GetTableRowPr().SetHeight('atLeast', cellHeight) // 高度至少多少trips
						var oTableCellPr = oTableStyle.GetTableCellPr()
						oTableCellPr.SetVerticalAlign('center')
						oTable.SetWrappingStyle(params.layout == 1 ? true : false)
						oTable.SetStyle(oTableStyle)
						var mergecount = rowcount * columncount - cellcount
						if (mergecount > 0) {
							var cells = []
							for (var k = 0; k < mergecount; ++k) {
								cells.push(oTable.GetRow(rowcount - 1).GetCell(k))
							}
							oTable.MergeCells(cells)
						}
						var scoreindex = -1
						// 设置单元格文本
						for (var irow = 0; irow < rowcount; ++irow) {
							var cbegin = 0
							var cend = columncount
							if (mergecount > 0 && irow == rowcount - 1) {
								// 最后一行
								cbegin = 1
								cend = columncount - mergecount + 1
							}
							for (var icolumn = cbegin; icolumn < cend; ++icolumn) {
								var cr = irow
								var cc = icolumn
								scoreindex++
								if (scoreindex >= scores.length) {
									break
								}
								var cell = oTable.GetCell(cr, cc)
								if (cell) {
									var cellcontent = cell.GetContent()
									if (cellcontent) {
										var oCellPara = cellcontent.GetElement(0)
										if (oCellPara) {
											oCellPara.AddText(scores[scoreindex].v)
											cell.SetWidth('twips', cellwidth)
											oCellPara.SetJc('center')
											oCellPara.SetColor(0, 0, 0, false)
											oCellPara.SetFontSize(16)
											scores[scoreindex].row = cr
											scores[scoreindex].column = cc
										} else {
											console.log('oCellPra is null')
										}
									} else {
										console.log('cellcontent is null')
									}
								} else {
									console.log('cannot get cell', cc, cr)
								}
							}
						}
						// 表格-高级设置 相关参数
						var Props = {
							CellSelect: true,
							Locked: false,
							PositionV: {
								Align: 1,
								RelativeFrom: 2,
								UseAlign: true,
								Value: 0,
							},
							PositionH: {
								Align: 4,
								RelativeFrom: 0,
								UseAlign: true,
								Value: 0,
							},
							TableDefaultMargins: {
								Bottom: 0,
								Left: 0,
								Right: 0,
								Top: 0,
							},
						}
						oTable.Table.Set_Props(Props)
						if (oTable.SetLockValue) {
							oTable.SetLockValue(true)
						} else {
							console.log('not function SetLockValue')
						}
						console.log('add table', oTable)
						var oFill = Api.CreateNoFill()
						var oStroke = Api.CreateStroke(3600, Api.CreateNoFill())
						var wline = params.layout == 2 ? 0.3 : 0.25
						var shapew = (cell_width_mm + wline) * columncount
						var shapeh = cell_height_mm * rowcount + 4
						var oDrawing = Api.CreateShape(
							'rect',
							shapew * 36e3,
							shapeh * 36e3,
							oFill,
							oStroke
						)

						// oDrawing.SetLockValue("noSelect", true) // 锁定，保证不可拖动，因为拖动会进行复制属性操作，导致ID改变，之后无法再追踪
						var drawDocument = oDrawing.GetContent()
						drawDocument.AddElement(0, oTable)
						if (params.layout == 2) {
							// 嵌入式
							oDrawing.SetWrappingStyle('square')
							oDrawing.SetHorPosition('column', (width - shapew) * 36e3)
						} else {
							// 顶部
							oDrawing.SetWrappingStyle('topAndBottom')
							oDrawing.SetHorPosition('column', (width - shapew) * 36e3)
							oDrawing.SetVerAlign('paragraph')
						}
						var titleobj = {
							type: 'qscore',
							ques_control_id: control_id,
						}
						oDrawing.Drawing.Set_Props({
							title: JSON.stringify(titleobj),
						})
						var paragraph
						if (params.posall) {
							// paragraph = Api.CreateParagraph();
							// paragraph.AddElement(oRun, 1);
							control.GetRange().Select()
							oDocument.AddDrawingToPage(
								oDrawing,
								0,
								(rect.X1 - shapew) * 36e3,
								rect.Y0 * 36e3
							) // 用这种方式加入的一定是相对页面的
							oDrawing.SetVerAlign('paragraph', rect.Y0 * 36e3)
							Api.asc_RemoveSelection()
						} else {
							console.log('++++++++++++++++')
							var oRun = Api.CreateRun()
							oRun.AddDrawing(oDrawing)
							var paragraphs = control.GetContent().GetAllParagraphs()
							paragraph = paragraphs[0]
							paragraph.AddElement(oRun, 1)
						}
						res = {
							control_id: control.Sdt.GetId(),
							add: true,
							score: parseFloat(score) || 0,
							score_options: {
								paragraph_id: paragraph ? paragraph.Paragraph.Id : 0,
								table_id: oTable.Table.Id,
								run_id: oRun ? oRun.Run.Id : 0,
								drawing_id: oDrawing.Drawing.Id,
								mode: params.mode,
								layout: params.layout,
								table_cells: scores,
								pos_all: params.posall,
							},
						}
						console.log('======== ', res.score_options)
						break
					}
				}
			}
			console.log('777777777777 ', res)
			// debugger
			return res
		},
		false,
		true
	).then((res) => {
		console.log(res)
		if (res && res.add) {
			for (var i = 0; i < window.BiyueCustomData.control_list.length; ++i) {
				if (
					window.BiyueCustomData.control_list[i].control_id == res.control_id
				) {
					window.BiyueCustomData.control_list[i].score = score
					window.BiyueCustomData.control_list[i].score_options =
						res.score_options
					break
				}
			}
		}
	})
}

function selectQues(treeInfo, index) {
	Asc.scope.temp_sel_index = index
	biyueCallCommand(
		window,
		function () {
			var res = Api.GetDocument().GetAllContentControls()
			var index = Asc.scope.temp_sel_index
			if (res && res[index]) {
				res[index].GetRange().Select()
			}
		},
		false,
		false
	)
}

function drawPosition2(data) {
	console.log('drawPosition', data)
	// 绘制区域，需要判断原本是否有这个区域，如果有，修改位置，如果没有，添加
	Asc.scope.pos = data
	Asc.scope.MM2EMU = MM2EMU
	console.log(
		'window.BiyueCustomData.control_list',
		window.BiyueCustomData.control_list
	)
	var find = window.BiyueCustomData.control_list.find((e) => {
		return (
			e.regionType == 'feature' &&
			e.zone_type == data.zone_type &&
			e.v == data.v
		)
	})
	if (find) {
		// 原本已经有
		Asc.scope.control_id = find.control_id
	} else {
		Asc.scope.control_id = null
	}
	biyueCallCommand(
		window,
		function () {
			var posdata = Asc.scope.pos
			var MM2EMU = Asc.scope.MM2EMU
			var control_id = Asc.scope.control_id
			var oDocument = Api.GetDocument()
			var oControls = oDocument.GetAllContentControls()
			if (control_id) {
				var tag = JSON.stringify({
					regionType: 'feature',
					zone_type: posdata.zone_type,
					mode: 6,
					v: posdata.v,
				})
				for (var i = 0; i < oControls.length; i++) {
					var oControl = oControls[i]
					if (oControl.Sdt.GetId() === control_id) {
						console.log('找到control 了', oControl)
						break
					}
				}
				var controls = oDocument.GetContentControlsByTag(tag)
				return {
					add: false,
				}
			} else {
				var oFill = Api.CreateNoFill()
				var oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(125, 125, 125))
				var oStroke = Api.CreateStroke(3600, oFill2)
				// 目前oStroke.Ln.Join == null, 需要拓展API才能修改Join，实现虚线效果
				var oDrawing = Api.CreateShape(
					'rect',
					MM2EMU * posdata.w,
					MM2EMU * posdata.h,
					oFill,
					oStroke
				)
				var drawDocument = oDrawing.GetContent()
				var oParagraph = Api.CreateParagraph()
				var text = ''
				if (posdata.zone_type == 15) {
					text = '再练'
				} else if (posdata.zone_type == 16) {
					text = '完成'
				}
				oParagraph.AddText(text)
				oParagraph.SetColor(125, 125, 125, false)
				oParagraph.SetFontSize(24)
				oParagraph.SetJc('center')
				drawDocument.AddElement(0, oParagraph)
				oDrawing.SetVerticalTextAlign('center')
				oDrawing.SetVerPosition('page', 914400)
				oDrawing.SetSize(914400, 914400)

				oDocument.AddDrawingToPage(
					oDrawing,
					0,
					MM2EMU * posdata.x,
					MM2EMU * posdata.y
				)
				var range = drawDocument.GetRange()
				range.Select()
				var tag = {
					regionType: 'feature',
					zone_type: posdata.zone_type,
					mode: 6,
					v: posdata.v,
				}
				var oResult = Api.asc_AddContentControl(range.controlType || 1, {
					Tag: JSON.stringify(tag),
				})
				Api.asc_RemoveSelection()
				if (oResult) {
					return {
						add: true,
						control: {
							control_id: oResult.InternalId,
							regionType: 'feature',
							v: posdata.v,
							zone_type: posdata.zone_type,
						},
					}
				} else {
					return {
						add: false,
					}
				}
			}
		},
		false,
		true
	).then((res) => {
		if (res && res.add) {
			window.BiyueCustomData.control_list.push(res.control)
		}
	})
}

function drawPositions(list) {
	Asc.scope.positions_list = list
	Asc.scope.pos_list = window.BiyueCustomData.pos_list
	Asc.scope.MM2EMU = MM2EMU
	Asc.scope.map_base64 = map_base64
	biyueCallCommand(
		window,
		function () {
			var positions_list = Asc.scope.positions_list || []
			var pos_list = Asc.scope.pos_list || []
			var MM2EMU = Asc.scope.MM2EMU
			var oDocument = Api.GetDocument()
			var objs = oDocument.GetAllDrawingObjects()
			var map_base64 = Asc.scope.map_base64
			positions_list.forEach((e) => {
				var posdata = pos_list.find((pos) => {
					return pos.zone_type == e.zone_type && pos.v == e.v
				})
				var oDrawing = null
				if (posdata && posdata.drawing_id) {
					// 已存在
					var index = objs.findIndex((obj) => {
						return obj.Drawing.Id == posdata.drawing_id
					})
					if (index >= 0) {
						oDrawing = objs[index]
					}
				}
				if (oDrawing) {
					console.log('已存在')
					oDrawing.SetSize(MM2EMU * e.w, MM2EMU * e.h)
					oDrawing.SetVerPosition('page', MM2EMU * e.x)
					oDrawing.SetHorPosition('page', MM2EMU * e.y)
				} else {
					console.log('不存在')
					if (e.draw_type == 'shape') {
						var oFill = Api.CreateNoFill()
						var oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(125, 125, 125))
						var oStroke = Api.CreateStroke(3600, oFill2)
						// 目前oStroke.Ln.Join == null, 需要拓展API才能修改Join，实现虚线效果
						oDrawing = Api.CreateShape(
							'rect',
							MM2EMU * e.w,
							MM2EMU * e.h,
							oFill,
							oStroke
						)

						var drawDocument = oDrawing.GetContent()
						var oParagraph = Api.CreateParagraph()
						var text = ''
						if (e.zone_type == 15) {
							text = '再练'
						} else if (e.zone_type == 16) {
							text = '完成'
						}
						oParagraph.AddText(text)
						oParagraph.SetFontSize(24)
						oParagraph.SetColor(125, 125, 125, false)
						oParagraph.SetJc('center')
						drawDocument.AddElement(0, oParagraph)
						oDrawing.SetVerticalTextAlign('center')

						// debugger
						oDrawing.SetVerPosition('page', 914400)
						oDrawing.SetSize(MM2EMU * e.w, MM2EMU * e.h)
						// oDrawing.SetLockValue("noSelect", true) // 锁定，保证不可拖动，因为拖动会进行复制属性操作，导致ID改变，之后无法再追踪
						oDocument.AddDrawingToPage(
							oDrawing,
							e.page_num,
							MM2EMU * e.x,
							MM2EMU * e.y
						)
					} else if (e.draw_type == 'image') {
						var imgname = ''
						var imgurl = ''
						if (e.zone_type == 16) {
							imgname = 'complete'
						} else if (e.zone_type == 28) {
							imgname = 'check'
						} else if (e.zone_type == 11) {
							imgname = 'pass'
						}
						imgurl = map_base64[imgname] //  getBase64(imgname)
						oDrawing = Api.CreateImage(
							imgurl,
							77 * 0.3 * 36000,
							28 * 0.3 * 36000
						)
						oDocument.AddDrawingToPage(
							oDrawing,
							e.page_num,
							MM2EMU * e.x,
							MM2EMU * e.y
						)
					}
				}
				if (!posdata) {
					console.log('add')
					pos_list.push({
						zone_type: e.zone_type,
						v: e.v,
						drawing_id: oDrawing.Drawing.Id,
					})
				} else {
					console.log('update')
					posdata.drawing_id = oDrawing.Drawing.Id
				}
			})
			return pos_list
		},
		false,
		true
	).then((res) => {
		console.log('drawPositions result:', res)
		window.BiyueCustomData.pos_list = res
	})
}

function drawPosition(data) {
	// 绘制区域，需要判断原本是否有这个区域，如果有，修改位置，如果没有，添加
	Asc.scope.pos = data
	Asc.scope.MM2EMU = MM2EMU
	var find = null
	if (window.BiyueCustomData.pos_list) {
		find = window.BiyueCustomData.pos_list.find((e) => {
			return e.zone_type == data.zone_type && e.v == data.v
		})
	}
	if (find) {
		// 原本已经有
		Asc.scope.drawing_id = find.id
	} else {
		Asc.scope.drawing_id = null
	}
	biyueCallCommand(
		window,
		function () {
			var posdata = Asc.scope.pos
			console.log('posdata', posdata)
			var MM2EMU = Asc.scope.MM2EMU
			var drawing_id = Asc.scope.drawing_id
			var oDocument = Api.GetDocument()
			if (drawing_id) {
				var objs = oDocument.GetAllDrawingObjects()
				for (var i = 0; i < objs.length; i++) {
					var oDrawing = objs[i]
					if (oDrawing.Drawing.Id == drawing_id) {
						oDrawing.SetSize(MM2EMU * posdata.w, MM2EMU * posdata.h)
						oDrawing.SetVerPosition('page', MM2EMU * posdata.x)
						oDrawing.SetHorPosition('page', MM2EMU * posdata.y)
						break
					}
				}
				return {
					add: false,
				}
			} else {
				var oFill = Api.CreateNoFill()
				var oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(125, 125, 125))
				var oStroke = Api.CreateStroke(3600, oFill2)
				// 目前oStroke.Ln.Join == null, 需要拓展API才能修改Join，实现虚线效果
				var oDrawing = Api.CreateShape(
					'rect',
					MM2EMU * posdata.w,
					MM2EMU * posdata.h,
					oFill,
					oStroke
				)

				var drawDocument = oDrawing.GetContent()
				var oParagraph = Api.CreateParagraph()
				var text = ''
				if (posdata.zone_type == 15) {
					text = '再练'
				} else if (posdata.zone_type == 16) {
					text = '完成'
				}
				oParagraph.AddText(text)
				oParagraph.SetFontSize(24)
				oParagraph.SetColor(125, 125, 125, false)
				oParagraph.SetJc('center')
				drawDocument.AddElement(0, oParagraph)
				oDrawing.SetVerticalTextAlign('center')

				// debugger
				oDrawing.SetVerPosition('page', 914400)
				oDrawing.SetSize(MM2EMU * posdata.w, MM2EMU * posdata.h)
				// oDrawing.SetLockValue("noSelect", true) // 锁定，保证不可拖动，因为拖动会进行复制属性操作，导致ID改变，之后无法再追踪
				oDocument.AddDrawingToPage(
					oDrawing,
					posdata.page_num,
					MM2EMU * posdata.x,
					MM2EMU * posdata.y
				)
				return {
					add: true,
					id: oDrawing.Drawing.Id,
					zone_type: posdata.zone_type,
					v: posdata.v,
				}
			}
		},
		false,
		true
	).then((res) => {
		console.log(res)
		if (res && res.add) {
			if (!window.BiyueCustomData.pos_list) {
				window.BiyueCustomData.pos_list = []
			}
			window.BiyueCustomData.pos_list.push({
				id: res.id,
				zone_type: res.zone_type,
				v: res.v,
			})
		}
	})
}

function getScores(score, mode) {
	var scores = []
	if (mode == 1) {
		// 普通模式
		for (var i = 0; i <= score; ++i) {
			scores.push({
				v: i + '',
			})
		}
	} else if (mode == 2) {
		// 大分值模式
		var ten = (score - (score % 10)) / 10
		for (var i = 0; i <= ten; ++i) {
			scores.push({
				v: i == 0 ? `${i}` : `${i}0+`,
			})
		}
		if (ten >= 1) {
			for (var j = 1; j < 10; ++j) {
				scores.push({
					v: j + '',
				})
			}
		}
	}
	scores.push({
		v: '0.5',
	})
	return scores
}
// 打分区用base64实现
function handleScoreField4(options) {
	if (!options) {
		return
	}
	var control_list = window.BiyueCustomData.control_list
	var list = []
	options.forEach((e) => {
		var index = control_list.findIndex((control) => {
			return e.ques_no == control.ques_no && control.regionType == 'question'
		})
		if (index >= 0) {
			var controlData = control_list[index]
			var needHandle = true
			if (controlData.score == e.score) {
				if (e.score) {
					var score_options = controlData.score_options
					if (
						score_options &&
						score_options.mode == e.mode &&
						score_options.layout == e.layout
					) {
						needHandle = false
					}
				}
			}
			if (needHandle) {
				var obj = Object.assign({}, e, {
					control_index: index,
				})
				if (e.score) {
					obj.scores = getScores(e.score, e.mode)
				}
				list.push(obj)
			}
		}
	})
	if (list.length == 0) {
		console.log('没有要处理的题目')
		return
	}
	Asc.scope.control_list = control_list
	Asc.scope.list = list
	Asc.scope.map_base64 = map_base64
	biyueCallCommand(
		window,
		function () {
			var oDocument = Api.GetDocument()
			var controls = oDocument.GetAllContentControls()
			var control_list = Asc.scope.control_list
			var list = Asc.scope.list
			var map_base64 = Asc.scope.map_base64
			console.log('list', list)
			var resList = []
			var shapes = oDocument.GetAllShapes()
			var drawings = oDocument.GetAllDrawingObjects()
			var cell_width_mm = 6
			var cell_height_mm = 6
			var MM2TWIPS = 25.4 / 72 / 20
			var cellWidth = cell_width_mm / MM2TWIPS
			var cellHeight = cell_height_mm / MM2TWIPS
			function getJsonData(str) {
				if (!str || str == '' || typeof str != 'string') {
					return {}
				}
				try {
					return JSON.parse(str)
				} catch (error) {
					console.log('json parse error', error)
					return {}
				}
			}
			for (var idx = 0, maxidx = list.length; idx < maxidx; ++idx) {
				var options = list[idx]
				var controlData = control_list[options.control_index]
				var control = controls.find((e) => {
					return e.Sdt.GetId() == controlData.control_id
				})
				if (!control) {
					resList.push({
						ques_no: options.ques_no,
						message: '题目控件不存在',
					})
					continue
				}
				var oShape = null
				var score_options = controlData.score_options
				if (score_options && score_options.run_id) {
					var scoreParagraph = new Api.private_CreateApiParagraph(
						AscCommon.g_oTableId.Get_ById(score_options.paragraph_id)
					)
					for (var i = 0; i < scoreParagraph.Paragraph.Content.length; ++i) {
						if (
							scoreParagraph.Paragraph.Content[i].Id == score_options.run_id
						) {
							console.log('remove run')
							scoreParagraph.RemoveElement(i)
							break
						}
					}
				} else if (score_options && score_options.drawing_id) {
					oShape = shapes.find((e) => {
						return e.Drawing.Id == controlData.score_options.drawing_id
					})
				} else {
					for (var i = 0, imax = shapes.length; i < imax; ++i) {
						var dtitle = shapes[i].Drawing.docPr.title
						if (dtitle && dtitle != '') {
							var titlejson = getJsonData(dtitle)
							if (
								titlejson.type == 'qscore' &&
								titlejson.ques_control_id == controlData.control_id
							) {
								oShape = shapes[i]
								break // 暂时假设打分区不会重复
							}
						}
					}
				}
				var oDrawing = null
				if (oShape) {
					oDrawing = drawings.find((e) => {
						return e.Drawing.Id == oShape.Drawing.Id
					})
				}
				if (oDrawing) {
					oDrawing.Delete()
				}
				if (!options.score || options.score == '') {
					// 删除
					resList.push({
						ques_no: options.ques_no,
						code: 1,
						options: options,
					})
					continue
				}
				var rect = Api.asc_GetContentControlBoundingRect(
					controlData.control_id,
					true
				)
				var newRect = {
					Left: rect.X0,
					Right: rect.X1,
					Top: rect.Y0,
					Bottom: rect.Y1,
				}
				var controlContent = control.GetContent()
				if (controlContent) {
					var pageIndex = 0
					if (
						controlContent.Document &&
						controlContent.Document.Pages &&
						controlContent.Document.Pages.length > 1
					) {
						for (var p = 0; p < controlContent.Document.Pages.length; ++p) {
							if (!control.Sdt.IsEmptyPage(p)) {
								pageIndex = p
								break
							}
						}
					}
					console.log('controlContent', controlContent)
					console.log('pageIndex', pageIndex)
					var pagebounds = controlContent.Document.Get_PageBounds(pageIndex)
					if (pagebounds) {
						newRect.Right = Math.max(pagebounds.Right, newRect.Right)
					}
				}
				console.log(
					controlData.ques_no,
					controlData.control_id,
					'rect',
					rect,
					'pagebounds',
					pagebounds
				)
				var width = newRect.Right - newRect.Left
				console.log('newRect', newRect, 'width', width)
				var trips_width = width / MM2TWIPS
				console.log('trips_width', trips_width, options)
				var scores = options.scores
				var maxWidth = trips_width / options.layout
				var rowcount = 1
				var cellcount = scores.length
				var columncount = cellcount
				if (maxWidth < cellcount * cellWidth) {
					// 需要换行
					rowcount = Math.ceil((cellcount * cellWidth) / maxWidth)
					columncount = Math.floor(maxWidth / cellWidth)
				}
				console.log('rowcount', rowcount, 'columncount', columncount)
				var oFill = Api.CreateNoFill()
				var oStroke = Api.CreateStroke(3600, Api.CreateNoFill())
				var shapew = columncount * cell_width_mm + 6
				var shapeh = rowcount * cell_height_mm + 3
				oDrawing = Api.CreateShape(
					'rect',
					shapew * 36e3,
					shapeh * 36e3,
					oFill,
					oStroke
				)
				var drawDocument = oDrawing.GetContent()
				var oParagraph = Api.CreateParagraph()
				for (var i = 0; i < scores.length; ++i) {
					var imgurl = map_base64['1']
					console.log(i, scores[i].v)
					var imgDrawing = Api.CreateImage(
						imgurl,
						cell_width_mm * 36e3,
						cell_height_mm * 36e3
					)
					console.log(
						'imgDrawing width height',
						imgDrawing.GetWidth(),
						imgDrawing.GetHeight()
					)
					oParagraph.AddDrawing(imgDrawing)
				}
				oParagraph.SetJc('right')
				oParagraph.SetColor(125, 125, 125, false)
				oParagraph.SetFontSize(24)
				oParagraph.SetFontFamily('黑体')
				console.log('oParagraph', oParagraph)
				drawDocument.AddElement(0, oParagraph)
				if (options.layout == 2) {
					// 嵌入式
					oDrawing.SetWrappingStyle('square')
					oDrawing.SetHorPosition('column', (width - shapew) * 36e3)
				} else {
					// 顶部
					oDrawing.SetWrappingStyle('topAndBottom')
					oDrawing.SetHorPosition('column', (width - shapew) * 36e3)
					oDrawing.SetVerPosition('paragraph', 1 * 36e3)
					// oDrawing.SetVerAlign("paragraph")
				}
				var titleobj = {
					type: 'qscore',
					ques_control_id: controlData.control_id,
				}
				oDrawing.Drawing.Set_Props({
					title: JSON.stringify(titleobj),
				})
				// 下面是全局插入的
				// control.GetRange().Select()
				// oDocument.AddDrawingToPage(oDrawing, 0, (newRect.Right - shapew) * 36E3, newRect.Top * 36E3); // 用这种方式加入的一定是相对页面的
				// oDrawing.SetVerAlign("paragraph", newRect.Top * 36E3)
				// Api.asc_RemoveSelection();
				// 在题目内插入
				var oRun = Api.CreateRun()
				oRun.AddDrawing(oDrawing)
				console.log('oDrawing', oDrawing)
				var paragraphs = controlContent.GetAllParagraphs()
				console.log('paragraphs', paragraphs)
				if (paragraphs && paragraphs.length > 0) {
					paragraphs[0].AddElement(oRun, 1)
					resList.push({
						code: 1,
						ques_no: options.ques_no,
						options: options,
						paragraph_id: paragraphs[0].Paragraph.Id,
						run_id: oRun.Run.Id,
						drawing_id: oDrawing.Drawing.Id,
						run_id: oRun.Run.Id,
					})
				}
			}
			return resList
		},
		false,
		true
	).then((res) => {
		console.log('callback for handleScoreField', res)
		if (!res) {
			return
		}
		res.forEach((e) => {
			if (e.options && e.options.control_index != undefined) {
				control_list[e.options.control_index].score = e.options.score
				if (e.options.score) {
					control_list[e.options.control_index].score_options = {
						paragraph_id: e.paragraph_id,
						run_id: e.run_id,
						drawing_id: e.drawing_id,
						table_id: e.table_id,
						mode: e.options.mode,
						layout: e.options.layout,
					}
				} else {
					control_list[e.options.control_index].score_options = null
				}
			}
		})
	})
}
// 打分区用添加表格单元格距离实现
function handleScoreField(options) {
	if (!options) {
		return
	}
	var control_list = window.BiyueCustomData.control_list
	var list = []
	options.forEach((e) => {
		var index = control_list.findIndex((control) => {
			return e.ques_no == control.ques_no && control.regionType == 'question'
		})
		if (index >= 0) {
			var controlData = control_list[index]
			var needHandle = true
			if (controlData.score == e.score) {
				if (e.score) {
					var score_options = controlData.score_options
					if (
						score_options &&
						score_options.mode == e.mode &&
						score_options.layout == e.layout
					) {
						needHandle = false
					}
				}
			}
			if (needHandle) {
				var obj = Object.assign({}, e, {
					control_index: index,
				})
				if (e.score) {
					obj.scores = getScores(e.score, e.mode)
				}
				list.push(obj)
			}
		}
	})
	if (list.length == 0) {
		console.log('没有要处理的题目')
		return
	}
	Asc.scope.control_list = control_list
	Asc.scope.list = list
	biyueCallCommand(
		window,
		function () {
			var oDocument = Api.GetDocument()
			var controls = oDocument.GetAllContentControls()
			var control_list = Asc.scope.control_list
			var list = Asc.scope.list
			console.log('list', list)
			var resList = []
			var shapes = oDocument.GetAllShapes()
			var drawings = oDocument.GetAllDrawingObjects()
			var cell_width_mm = 8 + 3
			var cell_height_mm = 8 + 3.5
			var MM2TWIPS = 25.4 / 72 / 20
			var cellWidth = cell_width_mm / MM2TWIPS
			var cellHeight = cell_height_mm / MM2TWIPS
			function getJsonData(str) {
				if (!str || str == '' || typeof str != 'string') {
					return {}
				}
				try {
					return JSON.parse(str)
				} catch (error) {
					console.log('json parse error', error)
					return {}
				}
			}
			for (var idx = 0, maxidx = list.length; idx < maxidx; ++idx) {
				var options = list[idx]
				var controlData = control_list[options.control_index]
				var control = controls.find((e) => {
					return e.Sdt.GetId() == controlData.control_id
				})
				if (!control) {
					resList.push({
						ques_no: options.ques_no,
						message: '题目控件不存在',
					})
					continue
				}
				var oShape = null
				var score_options = controlData.score_options
				if (score_options && score_options.run_id) {
					var scoreParagraph = new Api.private_CreateApiParagraph(
						AscCommon.g_oTableId.Get_ById(score_options.paragraph_id)
					)
					for (var i = 0; i < scoreParagraph.Paragraph.Content.length; ++i) {
						if (
							scoreParagraph.Paragraph.Content[i].Id == score_options.run_id
						) {
							console.log('remove run')
							scoreParagraph.RemoveElement(i)
							break
						}
					}
				} else if (score_options && score_options.drawing_id) {
					oShape = shapes.find((e) => {
						return e.Drawing.Id == controlData.score_options.drawing_id
					})
				} else {
					for (var i = 0, imax = shapes.length; i < imax; ++i) {
						var dtitle = shapes[i].Drawing.docPr.title
						if (dtitle && dtitle != '') {
							var titlejson = getJsonData(dtitle)
							if (
								titlejson.type == 'qscore' &&
								titlejson.ques_control_id == controlData.control_id
							) {
								oShape = shapes[i]
								break // 暂时假设打分区不会重复
							}
						}
					}
				}
				var oDrawing = null
				if (oShape) {
					oDrawing = drawings.find((e) => {
						return e.Drawing.Id == oShape.Drawing.Id
					})
				}
				if (oDrawing) {
					oDrawing.Delete()
				}
				if (!options.score || options.score == '') {
					// 删除
					resList.push({
						ques_no: options.ques_no,
						code: 1,
						options: options,
					})
					continue
				}
				var rect = Api.asc_GetContentControlBoundingRect(
					controlData.control_id,
					true
				)
				var newRect = {
					Left: rect.X0,
					Right: rect.X1,
					Top: rect.Y0,
					Bottom: rect.Y1,
				}
				var controlContent = control.GetContent()
				if (controlContent) {
					var pageIndex = 0
					if (
						controlContent.Document &&
						controlContent.Document.Pages &&
						controlContent.Document.Pages.length > 1
					) {
						for (var p = 0; p < controlContent.Document.Pages.length; ++p) {
							if (!control.Sdt.IsEmptyPage(p)) {
								pageIndex = p
								break
							}
						}
					}
					console.log('controlContent', controlContent)
					console.log('pageIndex', pageIndex)
					var pagebounds = controlContent.Document.Get_PageBounds(pageIndex)
					if (pagebounds) {
						newRect.Right = Math.max(pagebounds.Right, newRect.Right)
					}
				}
				console.log(
					controlData.ques_no,
					controlData.control_id,
					'rect',
					rect,
					'pagebounds',
					pagebounds
				)
				var width = newRect.Right - newRect.Left
				console.log('newRect', newRect, 'width', width)
				var trips_width = width / MM2TWIPS
				console.log('trips_width', trips_width, options)
				var maxTableWidth = trips_width / options.layout
				var scores = options.scores
				var cellcount = scores.length
				var rowcount = 1
				var columncount = cellcount
				var maxTableWidth = trips_width / options.layout
				console.log('maxTableWidth', maxTableWidth, cellcount, cellWidth)
				if (maxTableWidth < cellcount * cellWidth) {
					// 需要换行
					rowcount = Math.ceil((cellcount * cellWidth) / maxTableWidth)
					console.log(
						'rowCount',
						rowcount,
						'cellcount',
						cellcount,
						'cellWidth',
						cellWidth,
						'maxTableWidth',
						maxTableWidth
					)
					columncount = Math.floor(maxTableWidth / cellWidth)
				} else {
					console.log('无需换行')
				}
				var mergecount = rowcount * columncount - cellcount
				if (rowcount <= 0 || columncount <= 0) {
					console.log('行数或列数异常', list)
					return
				}
				var oTable = Api.CreateTable(columncount, rowcount)
				if (mergecount > 0) {
					var cells = []
					for (var k = 0; k < mergecount; ++k) {
						var cellrow = oTable.GetRow(rowcount - 1)
						if (cellrow) {
							var cell = cellrow.GetCell(k)
							if (cell) {
								cells.push(cell)
							} else {
								console.log(rowcount - 1, k, 'cell is null')
							}
						} else {
							console.log(
								'cellrow is null',
								rowcount - 1,
								oTable.GetRowsCount()
							)
						}
					}
					if (cells.length > 0) {
						oTable.MergeCells(cells)
					}
				}
				var oTableStyle = oDocument.CreateStyle('CustomTableStyle', 'table')
				var oTableStylePr = oTableStyle.GetConditionalTableStyle('wholeTable')
				oTable.SetTableLook(true, true, true, true, true, true)
				oTableStylePr.GetTableRowPr().SetHeight('atLeast', cellHeight) // 高度至少多少trips
				var oTableCellPr = oTableStyle.GetTableCellPr()
				oTableCellPr.SetVerticalAlign('center')
				oTable.SetWrappingStyle(params.layout == 1 ? true : false)
				oTable.SetStyle(oTableStyle)
				oTable.SetCellSpacing(150)
				oTable.SetTableBorderTop('single', 1, 0.1, 255, 255, 255)
				oTable.SetTableBorderBottom('single', 1, 0.1, 255, 255, 255)
				oTable.SetTableBorderLeft('single', 1, 0.1, 255, 255, 255)
				oTable.SetTableBorderRight('single', 1, 0.1, 255, 255, 255)
				var scoreindex = -1
				// 设置单元格文本
				for (var irow = 0; irow < rowcount; ++irow) {
					var cbegin = 0
					var cend = columncount
					if (mergecount > 0 && irow == rowcount - 1) {
						// 最后一行
						cbegin = 1
						cend = columncount - mergecount + 1
					}
					for (var icolumn = cbegin; icolumn < cend; ++icolumn) {
						var cr = irow
						var cc = icolumn
						scoreindex++
						if (scoreindex >= scores.length) {
							break
						}
						var cell = oTable.GetCell(cr, cc)
						if (cell) {
							var cellcontent = cell.GetContent()
							if (cellcontent) {
								var oCellPara = cellcontent.GetElement(0)
								if (oCellPara) {
									oCellPara.AddText(scores[scoreindex].v)
									cell.SetWidth('twips', cellWidth)
									oCellPara.SetJc('center')
									oCellPara.SetColor(0, 0, 0, false)
									oCellPara.SetFontSize(16)
									scores[scoreindex].row = cr
									scores[scoreindex].column = cc
								} else {
									console.log('oCellPra is null')
								}
							} else {
								console.log('cellcontent is null')
							}
							// console.log('cell', cell)
						} else {
							console.log('cannot get cell', cc, cr)
						}
					}
				}
				oTable.SetWidth('twips', columncount * cellWidth)
				var shapew = cell_width_mm * columncount + 3
				var shapeh = cell_height_mm * rowcount + 4
				var Props = {
					CellSelect: true,
					Locked: false,
					PositionV: {
						Align: 1,
						RelativeFrom: 2,
						UseAlign: true,
						Value: 0,
					},
					PositionH: {
						Align: 4,
						RelativeFrom: 0,
						UseAlign: true,
						Value: 0,
					},
					TableDefaultMargins: {
						Bottom: 0,
						Left: 0,
						Right: 0,
						Top: 0,
					},
				}
				oTable.Table.Set_Props(Props)
				// if (oTable.SetLockValue) {
				//   oTable.SetLockValue(true)
				// }
				var oFill = Api.CreateNoFill()
				var oStroke = Api.CreateStroke(3600, Api.CreateNoFill())
				oDrawing = Api.CreateShape(
					'rect',
					shapew * 36e3,
					shapeh * 36e3,
					oFill,
					oStroke
				)
				var drawDocument = oDrawing.GetContent()
				drawDocument.AddElement(0, oTable)
				if (options.layout == 2) {
					// 嵌入式
					oDrawing.SetWrappingStyle('square')
					oDrawing.SetHorPosition('column', (width - shapew) * 36e3)
				} else {
					// 顶部
					oDrawing.SetWrappingStyle('topAndBottom')
					oDrawing.SetHorPosition('column', (width - shapew) * 36e3)
					oDrawing.SetVerPosition('paragraph', 1 * 36e3)
					// oDrawing.SetVerAlign("paragraph")
				}
				var titleobj = {
					type: 'qscore',
					ques_control_id: controlData.control_id,
				}
				oDrawing.Drawing.Set_Props({
					title: JSON.stringify(titleobj),
				})
				// 下面是全局插入的
				// control.GetRange().Select()
				// oDocument.AddDrawingToPage(oDrawing, 0, (newRect.Right - shapew) * 36E3, newRect.Top * 36E3); // 用这种方式加入的一定是相对页面的
				// oDrawing.SetVerAlign("paragraph", newRect.Top * 36E3)
				// Api.asc_RemoveSelection();
				// 在题目内插入
				var oRun = Api.CreateRun()
				oRun.AddDrawing(oDrawing)
				var paragraphs = controlContent.GetAllParagraphs()
				console.log('paragraphs', paragraphs)
				if (paragraphs && paragraphs.length > 0) {
					paragraphs[0].AddElement(oRun, 1)
					resList.push({
						code: 1,
						ques_no: options.ques_no,
						options: options,
						paragraph_id: paragraphs[0].Paragraph.Id,
						run_id: oRun.Run.Id,
						drawing_id: oDrawing.Drawing.Id,
						table_id: oTable.Table.Id,
						run_id: oRun.Run.Id,
					})
				}
			}
			return resList
		},
		false,
		true
	).then((res) => {
		console.log('callback for handleScoreField', res)
		if (!res) {
			return
		}
		res.forEach((e) => {
			if (e.options && e.options.control_index != undefined) {
				control_list[e.options.control_index].score = e.options.score
				if (e.options.score) {
					control_list[e.options.control_index].score_options = {
						paragraph_id: e.paragraph_id,
						run_id: e.run_id,
						drawing_id: e.drawing_id,
						table_id: e.table_id,
						mode: e.options.mode,
						layout: e.options.layout,
					}
				} else {
					control_list[e.options.control_index].score_options = null
				}
			}
		})
	})
}
// 打分区用添加表格单元格分割实现
function handleScoreField2(options) {
	if (!options) {
		return
	}
	var control_list = window.BiyueCustomData.control_list
	var list = []
	options.forEach((e) => {
		var index = control_list.findIndex((control) => {
			return e.ques_no == control.ques_no && control.regionType == 'question'
		})
		if (index >= 0) {
			var controlData = control_list[index]
			var needHandle = true
			if (controlData.score == e.score) {
				if (e.score) {
					var score_options = controlData.score_options
					if (
						score_options &&
						score_options.mode == e.mode &&
						score_options.layout == e.layout
					) {
						needHandle = false
					}
				}
			}
			if (needHandle) {
				var obj = Object.assign({}, e, {
					control_index: index,
				})
				if (e.score) {
					obj.scores = getScores(e.score, e.mode)
				}
				list.push(obj)
			}
		}
	})
	if (list.length == 0) {
		console.log('没有要处理的题目')
		return
	}
	Asc.scope.control_list = control_list
	Asc.scope.list = list
	biyueCallCommand(
		window,
		function () {
			var oDocument = Api.GetDocument()
			var controls = oDocument.GetAllContentControls()
			var control_list = Asc.scope.control_list
			var list = Asc.scope.list
			console.log('list', list)
			var resList = []
			var shapes = oDocument.GetAllShapes()
			var drawings = oDocument.GetAllDrawingObjects()
			var cell_width_mm = 8
			var cell_height_mm = 8
			var spacing_width_mm = 4
			var MM2TWIPS = 25.4 / 72 / 20
			var cellWidth = cell_width_mm / MM2TWIPS
			var spacingWidth = spacing_width_mm / MM2TWIPS
			var cellHeight = cell_height_mm / MM2TWIPS
			function getJsonData(str) {
				if (!str || str == '' || typeof str != 'string') {
					return {}
				}
				try {
					return JSON.parse(str)
				} catch (error) {
					console.log('json parse error', error)
					return {}
				}
			}
			for (var idx = 0, maxidx = list.length; idx < maxidx; ++idx) {
				var options = list[idx]
				var controlData = control_list[options.control_index]
				var control = controls.find((e) => {
					return e.Sdt.GetId() == controlData.control_id
				})
				if (!control) {
					resList.push({
						ques_no: options.ques_no,
						message: '题目控件不存在',
					})
					continue
				}
				var oShape = null
				var score_options = controlData.score_options
				if (score_options && score_options.run_id) {
					var scoreParagraph = new Api.private_CreateApiParagraph(
						AscCommon.g_oTableId.Get_ById(score_options.paragraph_id)
					)
					for (var i = 0; i < scoreParagraph.Paragraph.Content.length; ++i) {
						if (
							scoreParagraph.Paragraph.Content[i].Id == score_options.run_id
						) {
							console.log('remove run')
							scoreParagraph.RemoveElement(i)
							break
						}
					}
				} else if (score_options && score_options.drawing_id) {
					oShape = shapes.find((e) => {
						return e.Drawing.Id == controlData.score_options.drawing_id
					})
				} else {
					for (var i = 0, imax = shapes.length; i < imax; ++i) {
						var dtitle = shapes[i].Drawing.docPr.title
						if (dtitle && dtitle != '') {
							var titlejson = getJsonData(dtitle)
							if (
								titlejson.type == 'qscore' &&
								titlejson.ques_control_id == controlData.control_id
							) {
								oShape = shapes[i]
								break // 暂时假设打分区不会重复
							}
						}
					}
				}
				var oDrawing = null
				if (oShape) {
					oDrawing = drawings.find((e) => {
						return e.Drawing.Id == oShape.Drawing.Id
					})
				}
				if (oDrawing) {
					oDrawing.Delete()
				}
				if (!options.score || options.score == '') {
					// 删除
					resList.push({
						ques_no: options.ques_no,
						code: 1,
						options: options,
					})
					continue
				}
				var rect = Api.asc_GetContentControlBoundingRect(
					controlData.control_id,
					true
				)
				var newRect = {
					Left: rect.X0,
					Right: rect.X1,
					Top: rect.Y0,
					Bottom: rect.Y1,
				}
				var controlContent = control.GetContent()
				if (controlContent) {
					var pageIndex = 0
					if (
						controlContent.Document &&
						controlContent.Document.Pages &&
						controlContent.Document.Pages.length > 1
					) {
						for (var p = 0; p < controlContent.Document.Pages.length; ++p) {
							if (!control.Sdt.IsEmptyPage(p)) {
								pageIndex = p
								break
							}
						}
					}
					console.log('controlContent', controlContent)
					console.log('pageIndex', pageIndex)
					var pagebounds = controlContent.Document.Get_PageBounds(pageIndex)
					if (pagebounds) {
						newRect.Right = Math.max(pagebounds.Right, newRect.Right)
					}
				}
				console.log(
					controlData.ques_no,
					controlData.control_id,
					'rect',
					rect,
					'pagebounds',
					pagebounds
				)
				var width = newRect.Right - newRect.Left
				console.log('newRect', newRect, 'width', width)
				var trips_width = width / MM2TWIPS
				console.log('trips_width', trips_width, options)
				var scores = options.scores
				var scoreCount = scores.length
				var maxTableWidth = trips_width / options.layout
				var rowcount = 1
				var columncount = scoreCount * 2 - 1
				console.log(
					'maxTableWidth',
					maxTableWidth,
					scoreCount * (cellWidth + spacingWidth)
				)
				var fillCountARow = scoreCount // 每行真正填充分数的格子数
				if (maxTableWidth < scoreCount * (cellWidth + spacingWidth)) {
					// 需要换行
					rowcount = Math.ceil(
						(scoreCount * (cellWidth + spacingWidth)) / maxTableWidth
					)
					var x = Math.floor(
						(maxTableWidth + spacingWidth) / (cellWidth + spacingWidth)
					)
					columncount = 2 * x - 1
					fillCountARow = x
				}
				console.log('rowcount', rowcount, 'columncount', columncount)
				if (rowcount <= 0 || columncount <= 0) {
					console.log('行数或列数异常', list)
					return
				}
				var fillRowCount = rowcount // 有填充分数的行数
				rowcount = fillRowCount + Math.floor(fillRowCount / 2)
				console.log(
					'fillCountARow',
					fillCountARow,
					'fillRowCount',
					fillRowCount,
					'rowcount',
					rowcount,
					'columncount',
					columncount
				)
				var oTable = Api.CreateTable(columncount, rowcount)
				var mergecount = (fillCountARow * fillRowCount - scoreCount) * 2
				for (var r = 0; r < rowcount; ++r) {
					var orow = oTable.GetRow(r)
					if (r % 2 > 0) {
						var mcell = orow.MergeCells()
						if (mcell) {
							mcell.SetCellBorderLeft('single', 1, 0.1, 255, 255, 255)
							mcell.SetCellBorderRight('single', 1, 0.1, 255, 255, 255)
							mcell.SetCellBorderTop('single', 1, 0.1, 255, 255, 255)
							mcell.SetCellBorderBottom('single', 1, 0.1, 255, 255, 255)
						}
						orow.SetHeight('atLeast', 4 / MM2TWIPS)
					} else if (r == rowcount - 1 && mergecount > 0) {
						var cells = []
						for (var k = 0; k < mergecount; ++k) {
							var cellrow = oTable.GetRow(rowcount - 1)
							if (cellrow) {
								var cell = cellrow.GetCell(k)
								if (cell) {
									cells.push(cell)
								} else {
									console.log(rowcount - 1, k, 'cell is null')
								}
							} else {
								console.log(
									'cellrow is null',
									rowcount - 1,
									oTable.GetRowsCount()
								)
							}
						}
						if (cells.length > 0) {
							var mcell = oTable.MergeCells(cells)
							if (mcell) {
								mcell.SetCellBorderLeft('single', 1, 0.1, 255, 255, 255)
								mcell.SetCellBorderRight('single', 1, 0.1, 255, 255, 255)
								mcell.SetCellBorderTop('single', 1, 0.1, 255, 255, 255)
								mcell.SetCellBorderBottom('single', 1, 0.1, 255, 255, 255)
							}
						}
					}
				}
				var oTableStyle = oDocument.CreateStyle('CustomTableStyle', 'table')
				var oTableStylePr = oTableStyle.GetConditionalTableStyle('wholeTable')
				oTable.SetTableLook(true, true, true, true, true, true)
				oTableStylePr.GetTableRowPr().SetHeight('atLeast', cellHeight) // 高度至少多少trips
				var oTableCellPr = oTableStyle.GetTableCellPr()
				oTableCellPr.SetVerticalAlign('center')
				oTable.SetWrappingStyle(params.layout == 1 ? true : false)
				oTable.SetStyle(oTableStyle)
				// oTable.SetCellSpacing(150);
				oTable.SetTableBorderTop('single', 0.1, 0.1, 255, 255, 255)
				oTable.SetTableBorderBottom('single', 0.1, 0.1, 255, 255, 255)
				oTable.SetTableBorderLeft('single', 0.1, 0.1, 255, 255, 255)
				oTable.SetTableBorderRight('single', 0.1, 0.1, 255, 255, 255)
				var scoreindex = -1
				// 设置单元格文本
				for (var irow = 0; irow < rowcount; ++irow) {
					var cbegin = 0
					var cend = columncount
					if (mergecount > 0 && irow == rowcount - 1) {
						// 最后一行
						cbegin = 1
						cend = columncount - mergecount + 1
					}
					console.log('irow', irow, 'cbegin', cbegin, 'cend', cend)
					var roww = 0
					var ww = 0
					var sw = 0
					var w = 0
					if (irow % 2 > 0) {
						continue
					}
					for (var icolumn = cbegin; icolumn < cend; ++icolumn) {
						var cr = irow
						var cc = icolumn
						var cell = oTable.GetCell(cr, cc)
						if (!cell) {
							break
						}
						var fillScore = irow % 2 == 0 && cc % 2 == 0
						if (irow == rowcount - 1 && mergecount > 0) {
							fillScore = cc % 2 > 0
						}
						var cellcontent = cell.GetContent()
						if (fillScore) {
							scoreindex++
							if (scoreindex >= scores.length) {
								break
							}
						}
						console.log('cell', cr, cc, fillScore)
						var oCellPara = cellcontent.GetElement(0)
						if (oCellPara) {
							oCellPara.AddText(fillScore ? scores[scoreindex].v : '')
							if (fillScore) {
								cell.SetCellBorderLeft('single', 1, 0.1, 212, 212, 212)
								cell.SetCellBorderRight('single', 1, 0.1, 212, 212, 212)
								cell.SetCellBorderTop('single', 1, 0.1, 212, 212, 212)
								cell.SetCellBorderBottom('single', 1, 0.1, 212, 212, 212)
								cell.SetWidth('twips', cellWidth)
								oCellPara.SetJc('center')
								oCellPara.SetColor(0, 0, 0, false)
								oCellPara.SetFontSize(16)
								scores[scoreindex].row = cr
								scores[scoreindex].column = cc
								console.log('cr', cr, 'cc', cc, scores[scoreindex].v, cellWidth)
								roww += cell_width_mm
								ww++
							} else {
								cell.SetWidth('twips', spacingWidth)
								roww += spacing_width_mm
								sw++
							}
							w += cell.CellPr.TableCellW.W
						}
					}
					console.log('roww', roww, 'ww', ww, 'sw', sw, w)
				}
				var tablew =
					(cell_width_mm * (columncount + 1)) / 2 +
					(spacing_width_mm * (columncount - 1)) / 2
				oTable.SetWidth('twips', tablew / MM2TWIPS)
				console.log(
					'tablewidth',
					columncount * (cellWidth + spacingWidth) - spacingWidth,
					100 / MM2TWIPS
				)
				var shapew = tablew + 3
				console.log('shapew', shapew - 3)
				var shapeh = cell_height_mm * fillRowCount + (fillRowCount - 1) * 4 + 4
				var Props = {
					CellSelect: true,
					Locked: false,
					PositionV: {
						Align: 1,
						RelativeFrom: 2,
						UseAlign: true,
						Value: 0,
					},
					PositionH: {
						Align: 4,
						RelativeFrom: 0,
						UseAlign: true,
						Value: 0,
					},
					TableDefaultMargins: {
						Bottom: 0,
						Left: 0,
						Right: 0,
						Top: 0,
					},
				}
				oTable.Table.Set_Props(Props)
				// if (oTable.SetLockValue) {
				//   oTable.SetLockValue(true)
				// }
				var oFill = Api.CreateNoFill()
				var oStroke = Api.CreateStroke(3600, Api.CreateNoFill())
				oDrawing = Api.CreateShape(
					'rect',
					shapew * 36e3,
					shapeh * 36e3,
					oFill,
					oStroke
				)
				var drawDocument = oDrawing.GetContent()
				drawDocument.AddElement(0, oTable)
				if (options.layout == 2) {
					// 嵌入式
					oDrawing.SetWrappingStyle('square')
					oDrawing.SetHorPosition('column', (width - shapew) * 36e3)
				} else {
					// 顶部
					oDrawing.SetWrappingStyle('topAndBottom')
					oDrawing.SetHorPosition('column', (width - shapew) * 36e3)
					oDrawing.SetVerPosition('paragraph', 1 * 36e3)
					// oDrawing.SetVerAlign("paragraph")
				}
				var titleobj = {
					type: 'qscore',
					ques_control_id: controlData.control_id,
				}
				oDrawing.Drawing.Set_Props({
					title: JSON.stringify(titleobj),
				})
				// 下面是全局插入的
				// control.GetRange().Select()
				// oDocument.AddDrawingToPage(oDrawing, 0, (newRect.Right - shapew) * 36E3, newRect.Top * 36E3); // 用这种方式加入的一定是相对页面的
				// oDrawing.SetVerAlign("paragraph", newRect.Top * 36E3)
				// Api.asc_RemoveSelection();
				// 在题目内插入
				var oRun = Api.CreateRun()
				oRun.AddDrawing(oDrawing)
				var paragraphs = controlContent.GetAllParagraphs()
				console.log('paragraphs', paragraphs)
				if (paragraphs && paragraphs.length > 0) {
					paragraphs[0].AddElement(oRun, 1)
					resList.push({
						code: 1,
						ques_no: options.ques_no,
						options: options,
						paragraph_id: paragraphs[0].Paragraph.Id,
						run_id: oRun.Run.Id,
						drawing_id: oDrawing.Drawing.Id,
						table_id: oTable.Table.Id,
						run_id: oRun.Run.Id,
					})
				}
			}
			return resList
		},
		false,
		true
	).then((res) => {
		console.log('callback for handleScoreField', res)
		if (!res) {
			return
		}
		res.forEach((e) => {
			if (e.options && e.options.control_index != undefined) {
				control_list[e.options.control_index].score = e.options.score
				if (e.options.score) {
					control_list[e.options.control_index].score_options = {
						paragraph_id: e.paragraph_id,
						run_id: e.run_id,
						drawing_id: e.drawing_id,
						table_id: e.table_id,
						mode: e.options.mode,
						layout: e.options.layout,
					}
				} else {
					control_list[e.options.control_index].score_options = null
				}
			}
		})
	})
}

// 切换权重显示
function toggleWeight() {
	Asc.scope.control_list = window.BiyueCustomData.control_list
	biyueCallCommand(
		window,
		function () {
			var control_list = Asc.scope.control_list
			var oDocument = Api.GetDocument()
			var controls = oDocument.GetAllContentControls()
			var added = false
			for (var i = 0; i < control_list.length; ++i) {
				var controlData = control_list[i]
				if (
					controlData.regionType == 'question' &&
					controlData.ask_controls &&
					controlData.ask_controls.length
				) {
					var controlRect = Api.asc_GetContentControlBoundingRect(
						controlData.control_id,
						true
					)
					for (var iask = 0; iask < controlData.ask_controls.length; ++iask) {
						var askControlId = controlData.ask_controls[iask].control_id
						var askControl = controls.find((e) => {
							return e.Sdt.GetId() == askControlId
						})
						if (!controlData.ask_controls[iask].weight_options) {
							// 还没有权重的id，需要创建
							var rect = Api.asc_GetContentControlBoundingRect(
								askControlId,
								true
							)
							var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 0, 0))
							oFill.UniFill.transparent = 40 // 透明度
							var oStroke = Api.CreateStroke(3600, Api.CreateNoFill())
							var width = rect.X1 - rect.X0
							var height = rect.Y1 - rect.Y0
							var oDrawing = Api.CreateShape(
								'rect',
								width * 36e3,
								height * 36e3,
								oFill,
								oStroke
							)
							oDrawing.SetPaddings(0, 0, 0, 0)
							oDrawing.SetWrappingStyle('inFront')
							oDrawing.SetHorPosition(
								'column',
								(rect.X0 - controlRect.X0) * 36e3
							)
							oDrawing.Drawing.Set_Props({
								title: 'ask_weight',
							})

							var drawDocument = oDrawing.GetContent()
							var oParagraph = Api.CreateParagraph()
							oParagraph.AddText(iask + 1 + '')
							drawDocument.AddElement(0, oParagraph)
							oParagraph.SetJc('center')
							var oTextPr = Api.CreateTextPr()
							oTextPr.SetColor(255, 111, 61, false)
							oTextPr.SetFontSize(14)
							oParagraph.SetTextPr(oTextPr)
							oDrawing.SetVerticalTextAlign('center')

							var oRun = Api.CreateRun()
							oRun.AddDrawing(oDrawing)
							var asktype = askControl.GetClassType()
							if (asktype == 'inlineLvlSdt') {
								askControl.AddElement(oRun)
								console.log(
									'ask oParagraph',
									oParagraph,
									'run',
									oRun,
									'oDrawing',
									oDrawing
								)
								controlData.ask_controls[iask].weight_options = {
									paragraph_id: oParagraph.Paragraph.Id,
									run_id: oRun.Run.Id,
									drawing_id: oDrawing.Drawing.Id,
									show: true,
								}
								added = true
							} else if (asktype == 'blockLvlSdt') {
								// todo..
							}
						}
					}
				}
			}
			for (var i = 0; i < control_list.length; ++i) {
				var controlData = control_list[i]
				if (controlData.regionType != 'question') {
					continue
				}
				var pcontrol = controls.find((e) => {
					return e.Sdt.GetId() == controlData.control_id
				})
				if (pcontrol) {
					if (pcontrol.GetAllDrawingObjects) {
						var drawingObjs = pcontrol.GetAllDrawingObjects()
						drawingObjs.forEach((e) => {
							if (e.Drawing.docPr.title == 'ask_weight') {
								if (!added) {
									e.Fill(Api.CreateNoFill())
								} else {
									e.Drawing.Set_Props({
										title: 'ask_weight',
										hidden: null,
									})
								}
							}
						})
						console.log('drawingObjs', drawingObjs)
					} else {
						console.log('GetAllDrawingObjects is unvalid', pcontrol)
					}
					if (controlData.ask_controls) {
						controlData.ask_controls.forEach((ask) => {
							if (ask.weight_options && ask.weight_options.paragraph_id) {
								var paragraph = new Api.private_CreateApiParagraph(
									AscCommon.g_oTableId.Get_ById(ask.weight_options.paragraph_id)
								)
								if (!added) {
									paragraph.RemoveElement(0)
								}
							}
						})
					}
				}
			}

			return control_list
		},
		false,
		true
	).then((res) => {
		window.BiyueCustomData.control_list = res
	})
}

function handleContentControlChange(params) {
	var controlId = params.InternalId
	var control_list = window.BiyueCustomData.control_list
	var tag = params.Tag
	if (tag) {
		try {
			tag = JSON.parse(params.Tag)
		} catch (error) {
			console.log('json parse error', error)
			return
		}
		if (tag.regionType == 'question') {
			if (control_list) {
				var find = control_list.find((e) => {
					return e.control_id == controlId
				})
				Asc.scope.find_controldata = find
			}
		} else {
			return
		}
	}
	Asc.scope.params = params
	biyueCallCommand(
		window,
		function () {
			var controldata = Asc.scope.find_controldata
			var params = Asc.scope.params
			var oDocument = Api.GetDocument()
			var controls = oDocument.GetAllContentControls()
			var control = controls.find((e) => {
				return e.Sdt.GetId() == params.InternalId
			})
			if (controldata) {
				// if (control && control.GetAllDrawingObjects) {
				//   var drawingObjs = control.GetAllDrawingObjects()
				//   for (var i = 0, imax = drawingObjs.length; i < imax; ++i) {
				//     var oDrawing = drawingObjs[i]
				//     if (oDrawing.Drawing.docPr.title == 'ask_weight') {
				//       oDrawing.Delete()
				//     }
				//   }
				// }
			}
		},
		false,
		true
	)
}

function deletePositions(list) {
	Asc.scope.pos_delete_list = list
	var pos_list = window.BiyueCustomData.pos_list
	Asc.scope.pos_list = pos_list
	biyueCallCommand(
		window,
		function () {
			var delete_list = Asc.scope.pos_delete_list || []
			var pos_list = Asc.scope.pos_list || []
			console.log('delete_list', delete_list)
			console.log('pos_list', pos_list)
			var oDocument = Api.GetDocument()
			var objs = oDocument.GetAllDrawingObjects()
			delete_list.forEach((e) => {
				var posdata = pos_list.find((pos) => {
					return pos.zone_type == e.zone_type && pos.v == e.v
				})
				if (posdata && posdata.drawing_id) {
					var oDrawing = objs.find((obj) => {
						return obj.Drawing.Id == posdata.drawing_id
					})
					if (oDrawing) {
						oDrawing.Delete()
					} else {
						console.log('cannot find oDrawing')
					}
				}
			})
			return delete_list
		},
		false,
		true
	).then((res) => {
		console.log('deletePositions result:', res)
		if (res) {
			res.forEach((e) => {
				var index = pos_list.findIndex((pos) => {
					return pos.zone_type == e.zone_type && pos.v == e.v
				})
				if (index >= 0) {
					pos_list.splice(index, 1)
				}
			})
		}
	})
}
// 分栏
function setSectionColumn(num) {
	Asc.scope.column_num = num
	biyueCallCommand(
		window,
		function () {
			var column_num = Asc.scope.column_num
			var oDocument = Api.GetDocument()
			var Document = oDocument.Document
			var nContentPos = Document.CurPos.ContentPos // 当前光标位置
			var oElement = Document.Content[nContentPos]
			var sections = oDocument.GetSections()
			var Pages = Document.Pages
			if (oElement.GetType() == 1) {
				// 是段落
				var pagesPos = oElement.CurPos.PagesPos // 当前页数
				// 根据页数获取
				var page = Pages[pagesPos]
				var beginPos = page.Pos
				var EndPos = page.EndPos
				var pageFirstPos = Document.GetDocumentPositionByXY(pagesPos, 0, 0)
				if (page.Sections) {
					var prePage = pagesPos > 0 ? Pages[pagesPos - 1] : null
					var nextPage =
						pagesPos < Pages.length - 1 ? Pages[pagesPos + 1] : null
					var firstIndex = page.Sections[0].Index
					var lastIndex = page.Sections[page.Sections.length - 1].Index
					var containPre =
						prePage &&
						prePage.Sections &&
						prePage.Sections[prePage.Sections.length - 1].Index == firstIndex
					var containNext =
						nextPage &&
						nextPage.Sections &&
						nextPage.Sections[0].Index == lastIndex
					var section1 = null
					if (containPre) {
						// 页面起始的section包含了上一页，需要将起始位置的段落拆分，插入新的section
						var beginEl = Document.Content[beginPos]
						Document.MoveCursorToNearestPos(
							Document.Get_NearestPos(pagesPos, 0, 0)
						)
						Document.Add_SectionBreak(0)
						var curSectPr = Document.GetCurrentSectionPr()
						curSectPr.Set_Columns_EqualWidth(true)
						curSectPr.Set_Columns_Num(column_num)
						curSectPr.Set_Columns_Space((25.4 / 72 / 20) * 200)
						curSectPr.Set_Columns_Sep(true)
					}
					if (containNext) {
						// 页面结束的section包含了下一页
					}
				}
			}
		},
		false,
		true
	)
}
// 提取单独的大题的control
function addOnlyBigControl(recalc = true) {
	Asc.scope.node_list = window.BiyueCustomData.node_list || []
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls() || []
		function getJsonData(str) {
			if (!str || str == '' || typeof str != 'string') {
				return {}
			}
			try {
				return JSON.parse(str)
			} catch (error) {
				console.log('json parse error', error)
				return {}
			}
		}
		controls.forEach(oControl => {
			var controlTag = getJsonData(oControl.GetTag())
			if (controlTag.big) {
				var oRange = null
				var count = oControl.Sdt.GetElementsCount()
				for (var i = 0; i < count; ++i) {
					var element = oControl.Sdt.GetElement(i)
					if (element.Id) {
						var oElement = Api.LookupObject(element.Id)
						if (oElement) {
							if (oElement.GetClassType() == 'paragraph') {
								if (!oRange) {
									oRange = oElement.GetRange()
								} else {
									oRange = oRange.ExpandTo(oElement.GetRange())
								}
							} else {
								break
							}
						}
					}
				}
				if (oRange) {
					oRange.Select()
					var tag = JSON.stringify({
						onlybig: 1,
						link_id: controlTag.client_id,
					})
					var oResult = Api.asc_AddContentControl(1, { Tag: tag })
					console.log('添加onlybig', oResult)
				} else {
					console.log('oRange is null')
				}
			}
		})
	}, false, recalc)
}
function removeOnlyBigControl() {
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var oControls = oDocument.GetAllContentControls()
		if (oControls) {
			oControls.forEach(e => {
				var tag = e.GetTag()
				if (tag && tag.indexOf('onlybig') >= 0) {
					Api.asc_RemoveContentControlWrapper(e.Sdt.GetId())
				}
			})
		}
	}, false, true)
}
function getAllPositions() {
	return addOnlyBigControl()
	.then(() => {
		return getAllPositions2() 
	}).then(() => {
		return removeOnlyBigControl()
	})
}
// 获取所有相关坐标，用于铺码使用，这里只给出如何获取的代码演示
function getAllPositions2() {
	Asc.scope.question_map = window.BiyueCustomData.question_map || {}
	Asc.scope.node_list = window.BiyueCustomData.node_list || []
	return biyueCallCommand(
		window,
		function () {
			var oDocument = Api.GetDocument()
			var controls = oDocument.GetAllContentControls()
			var drawings = oDocument.GetAllDrawingObjects()
			var oShapes = oDocument.GetAllShapes()
			var ques_list = []
			var pageCount = oDocument.GetPageCount()
			var question_map = Asc.scope.question_map
			var node_list = Asc.scope.node_list
			var oTables = oDocument.GetAllTables() || []
			console.log('------------------------------')
			function mmToPx(mm) {
				// 1 英寸 = 25.4 毫米
				// 1 英寸 = 96 像素（常见的屏幕分辨率）
				// 因此，1 毫米 = (96 / 25.4) 像素
				const pixelsPerMillimeter = 96 / 25.4
				return Math.floor(mm * pixelsPerMillimeter) >>> 0
			}
			function getJsonData(str) {
				if (!str || str == '' || typeof str != 'string') {
					return {}
				}
				try {
					return JSON.parse(str)
				} catch (error) {
					console.log('json parse error', error)
					return {}
				}
			}
			function getBlockControlBounds(oControl, bounds) {
				if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
					return
				}
				var oControlContent = oControl.GetContent()
				var Pages = oControlContent.Document.Pages
				if (Pages) {
					Pages.forEach((page, index) => {
						if (page.Bounds) {
							let w = page.Bounds.Right - page.Bounds.Left
							let h = page.Bounds.Bottom - page.Bounds.Top
							if (w > 0 && h > 0) {
								bounds.push({
									Page: oControl.Sdt.GetAbsolutePage(index),
									X: mmToPx(page.Bounds.Left),
									Y: mmToPx(page.Bounds.Top),
									W: mmToPx(w),
									H: mmToPx(h),
								})
							}
						}
					})
				}
			}
			function getCell(write_data) {
				for (var i = 0; i < oTables.length; ++i) {
					var oTable = oTables[i]
					if (oTable.GetPosInParent() == -1) { continue }
					var desc = getJsonData(oTable.GetTableDescription())
					var keys = Object.keys(desc)
					if (keys.length) {
						for (var j = 0; j < keys.length; ++j) {
							var key = keys[j]
							if (desc[key] == write_data.id) {
								var rc = key.split('_')
								if (write_data.row_index == undefined) {
									return oTable.GetCell(rc[0], rc[1])
								} else if (write_data.row_index == rc[0] && write_data.cell_index == rc[1]) {
									return oTable.GetCell(rc[0], rc[1])
								}
								
							}
						}
					}
				}
				return null
			}
			function getCellBounds(oCell, ask_score, order) {
				if (!oCell || oCell.GetClassType() != 'tableCell') {
					return []
				}
				var bounds = []
				var pagesCount = oCell.Cell.PagesCount
				for (var p = 0; p < pagesCount; ++p) {
					var pagebounds = oCell.Cell.GetPageBounds(p)
					if (!pagebounds) {
						continue
					}
					if (pagebounds.Right == 0 && pagebounds.Left == 0) {
						continue
					}
					bounds.push({
						order: order + '',
						page: oCell.Cell.Get_AbsolutePage(p) + 1,
						x: mmToPx(pagebounds.Left),
						y: mmToPx(pagebounds.Top),
						w: mmToPx(pagebounds.Right - pagebounds.Left),
						h: mmToPx(pagebounds.Bottom - pagebounds.Top),
						v: ask_score + '',
						mark_order: order + '',
					})
				}
				return bounds
			}
			function getSimplePos(oControl) {
				var paragraphs = oControl.GetAllParagraphs()
				for (var i = 0; i < paragraphs.length; ++i) {
					var oParagraph = paragraphs[i]
					if (oParagraph) {
						var parent1 = oParagraph.Paragraph.Parent
						var parent2 = parent1.Parent
						if (parent2 && parent2.Id == oControl.Sdt.GetId()) {
							var oNumberingLevel = oParagraph.GetNumbering()
							if (oNumberingLevel) {
								var level = oNumberingLevel.Lvl
								var oNum = oNumberingLevel.Num
								if (!oNum) {
									return null
								}
								var oNumberingLvl = oNum.GetLvl(level)
								if (!oNumberingLvl) {
									return null
								}
								var LvlText = oNumberingLvl.LvlText || []
								if (LvlText && LvlText.length && LvlText[0].Value == '\ue749') {
									var Numbering = oParagraph.Paragraph.Numbering
									var numberPage = Numbering.Page
									var numberingRect =
										Api.asc_GetParagraphNumberingBoundingRect(
											oParagraph.Paragraph.Id,
											1
										) || {}
									return {
										page: oParagraph.Paragraph.GetAbsolutePage(numberPage) + 1,
										x: mmToPx(numberingRect.X0),
										y: mmToPx(numberingRect.Y0),
										w: mmToPx(numberingRect.X1 - numberingRect.X0),
										h: mmToPx(numberingRect.Y1 - numberingRect.Y0),
									}
								}
							} else {
								var pControls = oParagraph.GetAllContentControls() || []
								var numControl = pControls.find(e => {
									var tag = getJsonData(e.GetTag())
									return e.GetClassType() == 'inlineLvlSdt' && tag.regionType == 'num'
								})
								if (numControl && numControl.Sdt && numControl.Sdt.Bounds) {
									var bounds = Object.values(numControl.Sdt.Bounds) || []
									if (bounds.length) {
										return {
											page: bounds[0].Page + 1,
											x: mmToPx(bounds[0].X),
											y: mmToPx(bounds[0].Y),
											w: mmToPx(bounds[0].W),
											h: mmToPx(bounds[0].H),
										}
									}
									
								}
							}
						}
					}
				}
				return null
			}
			function GetCorrectRegion(oControl) {
				var correct_ask_region = []
				var correct_region = {}
				var identify_region = []
				var write_region = []
				if (oControl.GetClassType() == 'blockLvlSdt') {
					var oControlContent = oControl.GetContent()
					var drawings = oControlContent.GetAllDrawingObjects()
					if (drawings) {
						for (var j = 0, jmax = drawings.length; j < jmax; ++j) {
							var oDrawing = drawings[j]
							var parentControl = oDrawing.GetParentContentControl()
							if (
								!parentControl ||
								parentControl.Sdt.GetId() != oControl.Sdt.GetId()
							) {
								continue
							}
							if (oDrawing.Drawing.docPr) {
								var title = oDrawing.Drawing.docPr.title
								if (title && title.indexOf('feature') >= 0) {
									var titleObj = getJsonData(title)
									if (
										titleObj.feature &&
										titleObj.feature.zone_type == 'question'
									) {
										var obj = {
											page: oDrawing.Drawing.PageNum + 1,
											x: mmToPx(oDrawing.Drawing.X),
											y: mmToPx(oDrawing.Drawing.Y),
											w: mmToPx(oDrawing.Drawing.Width),
											h: mmToPx(oDrawing.Drawing.Height),
										}
										if (titleObj.feature.sub_type == 'simple') {
											correct_region = obj
										} else if (titleObj.feature.sub_type == 'ask_accurate') {
											obj.v = correct_ask_region.length + 1 + ''
											correct_ask_region.push(obj)
										} else if (titleObj.feature.sub_type == 'write') {
											obj.write_id = titleObj.feature.client_id
											write_region.push(obj)
										} else if (titleObj.feature.sub_type == 'identify') {
											obj.write_id = titleObj.feature.client_id
											identify_region.push(obj)
										}
									}
								}
							}
						}
					}
					correct_region = getSimplePos(oControl)
				}
				return {
					correct_ask_region,
					correct_region,
					write_region,
					identify_region,
				}
			}
			function getControlsByClientId(cid) {
				var findControls = controls.filter(e => {
					var tag = getJsonData(e.GetTag())
					if (e.GetClassType() == 'blockLvlSdt') {
						return tag.client_id == cid && e.GetPosInParent() >= 0
					} else if (e.GetClassType() == 'inlineLvlSdt') {
						return e.Sdt && e.Sdt.GetPosInParent() >= 0 && tag.client_id == cid
					}
				})
				if (findControls && findControls.length) {
					return findControls[0]
				}
			}
			// 集中作答区坐标
			function getGatherCellRegion(nodeId) {
				var nodeData = node_list.find((e) => {
					return e.id == nodeId
				})
				if (!nodeData || !nodeData.use_gather || !nodeData.gather_cell_id) {
					return null
				}
				if (question_map[nodeId].ques_mode != 1 && question_map[nodeId].ques_mode != 5) {
					return null
				}
				var oCell = Api.LookupObject(nodeData.gather_cell_id)
				if (!oCell || oCell.GetClassType() != 'tableCell') {
					return null
				}
				// 单元格区域
				var cell_region = getCellBounds(oCell, question_map[nodeId].score, '1')
				// 互动区域
				var drawings = oCell.GetContent().GetAllDrawingObjects() || []
				var oDrawing = drawings.find((e) => {
					var title = e.Drawing.docPr.title || '{}'
					var titleObj = getJsonData(title)
					if (titleObj.feature && titleObj.feature.sub_type == 'ask_accurate') {
						return true
					}
				})
				var interaction_region = null
				if (oDrawing) {
					interaction_region = {
						page: oDrawing.Drawing.PageNum + 1,
						x: mmToPx(oDrawing.Drawing.X),
						y: mmToPx(oDrawing.Drawing.Y),
						w: mmToPx(oDrawing.Drawing.Width),
						h: mmToPx(oDrawing.Drawing.Height),
						v: '1',
					}
				}
				return { cell_region, interaction_region }
			}

			for (var i = 0, imax = controls.length; i < imax; ++i) {
				var oControl = controls[i]
				var tag = getJsonData(oControl.GetTag() || '{}')
				var question_obj = question_map[tag.client_id]
					? question_map[tag.client_id]
					: {}
				// 当前题目必须是blockLvlSdt
				if (tag.regionType != 'question' || oControl.GetClassType() != 'blockLvlSdt' || question_obj.level_type != 'question') {
					continue
				}
				var bounds = []
				var useControl = oControl
				if (oControl.GetClassType() == 'blockLvlSdt') {
					var nodeData = node_list.find((e) => {
						return e.id == tag.client_id
					})
					if (nodeData && nodeData.is_big) {
						var childcontrols = oControl.GetAllContentControls() || []
						var bigControl = childcontrols.find(e => {
							var btag = getJsonData(e.GetTag())
							return e.GetClassType() == 'blockLvlSdt' && btag.onlybig == 1 && btag.link_id == nodeData.id
						})
						if (bigControl) {
							useControl = bigControl
							console.log('===== 使用bigControl')
							getBlockControlBounds(bigControl, bounds)
						} else {
							console.log('未找到bigControl')
							getBlockControlBounds(oControl, bounds)
						}
					} else {
						getBlockControlBounds(oControl, bounds)
					}
				} else if (oControl.GetClassType() == 'inlineLvlSdt') {
					if (oControl.Sdt.Bounds) {
						bounds = Object.values(oControl.Sdt.Bounds)
					}
				}
				// todo..分数框的尚未添加
				var oRange = useControl.GetRange()
				// oRange.Select()
				// 由于数据太大或异常可能导致后端数据无法写入，因此上传卷面时不再上传Html，所有这里也不再读取
				// let text_data = {
				// 	data: '',
				// 	// 返回的数据中class属性里面有binary格式的dom信息，需要删除掉
				// 	pushData: function (format, value) {
				// 		this.data = value
				// 			? value.replace(/class="[a-zA-Z0-9-:;+"\/=]*/g, '')
				// 			: ''
				// 	},
				// }
				// Api.asc_CheckCopy(text_data, 2)
				var correctPos = GetCorrectRegion(oControl)
				var gatherRegion = getGatherCellRegion(tag.client_id)
				var is_gather_region = false
				if (
					Asc.scope.choice_params &&
					Asc.scope.choice_params.style === 'show_choice_region' &&
					gatherRegion
				) {
					// 开启集中作答区并且有集中作答区的坐标信息
					is_gather_region = true
					if (question_map[tag.client_id]) {
						question_map[tag.client_id].ask_list = []
					}
				}
				var item = {
					id: tag.client_id,
					control_id: oControl.Sdt.GetId(),
					text: oRange.GetText(), // 如果需要html, 请参考ExamTree.js的reqUploadTree
					// content: `${text_data.data || ''}`,
					title_region: [],
					correct_region: correctPos.correct_region || {},
					correct_ask_region: correctPos.correct_ask_region,
					score: parseFloat(question_obj.score) || 0,
					ask_num: 0,
					additional: false, // 是否为附加题
					answer: '',
					ref_id: question_obj.uuid || '',
					ques_type: question_obj.question_type || '',
					ques_mode: question_obj.ques_mode || 3,
					ques_name: question_obj.ques_name || question_obj.ques_default_name,
					mark_method: '1',
					mark_ask_region: {},
					write_ask_region: [],
				}
				bounds.forEach((e) => {
					item.title_region.push({
						page: e.Page + 1,
						x: e.X,
						y: e.Y,
						w: e.W,
						h: e.H,
					})
				})
				if (is_gather_region) {
					// 集中作答区的题目
					let cell_region = gatherRegion.cell_region || []
					cell_region.forEach((e) => {
						item.write_ask_region.push({
							page: e.page,
							order: e.order + '',
							v: item.score + '',
							x: e.x,
							y: e.y,
							w: e.w,
							h: e.h,
						})
					})
					let mark_ask_region = {}
					mark_ask_region['1'] = item.write_ask_region
					item.mark_ask_region = mark_ask_region
				} else if (question_obj.ask_list && question_obj.ask_list.length) {
					var nodeData = node_list.find((e) => {
						return e.id == tag.client_id
					})
					if (nodeData && nodeData.write_list) {
						let mark_order = 1
						for (var iask = 0; iask < question_obj.ask_list.length; ++iask) {
							var ids = [question_obj.ask_list[iask].id]
							if (question_obj.ask_list[iask].other_fileds) {
								ids = ids.concat(question_obj.ask_list[iask].other_fileds)
							}
							var ask_score = question_obj.ask_list[iask].score || ''
							var find = false
							for (var idx2 = 0; idx2 < ids.length; ++idx2) {
								var askData = nodeData.write_list.find((e) => {
									return e.id == ids[idx2]
								})
								if (!askData) {
									continue
								}
								if (askData.sub_type == 'control') {
									// 这里可能出现的BUG，control拖动后，业务ID不变，但control_id已改变
									var oAskControl = getControlsByClientId(askData.id)
									if (oAskControl && oAskControl.Sdt) {
										if (oAskControl.GetClassType() == 'inlineLvlSdt') {
											var askBounds = Object.values(oAskControl.Sdt.Bounds)
											askBounds.forEach((e) => {
												if (e.W) {
													item.write_ask_region.push({
														order: mark_order + '',
														page: e.Page + 1,
														x: mmToPx(e.X),
														y: mmToPx(e.Y),
														w: mmToPx(e.W),
														h: mmToPx(e.H),
														v: ask_score + '',
														mark_order: mark_order,
													})
												}
											})
										} else if (oAskControl.GetClassType() == 'blockLvlSdt') {
											var rects2 = []
											getBlockControlBounds(oAskControl, rects2)
											rects2.forEach(e => {
												item.write_ask_region.push({
													order: mark_order + '',
													page: e.Page + 1,
													x: e.X,
													y: e.Y,
													w: e.W,
													h: e.H,
													v: ask_score + '',
													mark_order: mark_order,
												})
											})
										}
										find = true
									}
								} else if (askData.sub_type == 'cell') {
									var oCell = Api.LookupObject(askData.cell_id)
									if (!oCell || oCell.GetClassType() != 'tableCell' || oCell.GetParentTable().GetPosInParent() == -1) {
										oCell = getCell(askData)
									}
									if (oCell) {
										var d = getCellBounds(oCell, ask_score, mark_order)
										item.write_ask_region = item.write_ask_region.concat(d)
										find = true
									}
								} else if ( askData.sub_type == 'write' || askData.sub_type == 'identify') {
									var oShape = oShapes.find(e => {
										if (e.Drawing) {
											var shapetitle = getJsonData(e.Drawing.docPr.title)
											return shapetitle.feature && shapetitle.feature.client_id == askData.id
										}
									})
									if (oShape && oShape.Drawing) {
										item.write_ask_region.push({
											order: mark_order + '',
											page: oShape.Drawing.PageNum + 1,
											x: mmToPx(oShape.Drawing.X),
											y: mmToPx(oShape.Drawing.Y),
											w: mmToPx(oShape.Drawing.Width),
											h: mmToPx(oShape.Drawing.Height),
											v: ask_score + '',
											mark_order: mark_order,
										})
										find = true
									}
								}
							}
							if (find) {
								mark_order++
							}
						}
					}
					// 在先不考虑出现有小问跨页的情况下，一个书写区算做一个小问批改区
					let mark_ask_region = item.write_ask_region.map((e) => {
						return { ...e, order: e.order + '' }
					})
					item.mark_ask_region = mark_ask_region.reduce((acc, obj) => {
						const order = obj.mark_order
						if (!acc.hasOwnProperty(order)) {
							acc[order] = []
						}
						// 过滤mark_order
						const { mark_order, ...newObj } = obj
						acc[order].push(newObj)
						return acc
					}, {})

					// 过滤不需要传出去的标记
					const removeField = (arr) => {
						return arr.map(({ mark_order, ...rest }) => rest)
					}
					item.write_ask_region = removeField(item.write_ask_region)
				} else {
					// 没有小问的题目 暂时使用当前的题干区域作为批改和作答区 同时如果存在多个题干区，也只算作一个题目的批改区
					bounds.forEach((e) => {
						item.write_ask_region.push({
							page: e.Page + 1,
							order: '1',
							v: item.score + '',
							x: e.X,
							y: e.Y,
							w: e.W,
							h: e.H,
						})
					})
					let mark_ask_region = {}
					mark_ask_region['1'] = item.write_ask_region
					item.mark_ask_region = mark_ask_region
				}

				item.ask_num = Object.keys(item.mark_ask_region).length
				item.mark_method = '1'

				if (item.ques_type === 3) {
					// bounds.forEach((e) => {
					// 	item.write_ask_region.push({
					// 		order: item.write_ask_region.length + 1 + '',
					// 		page: e.Page + 1,
					// 		x: e.X,
					// 		y: e.Y,
					// 		w: e.W,
					// 		h: e.H,
					// 		v: '1',
					// 	})
					// })
					// let mark_ask_region = item.write_ask_region.map((e) => {
					// 	return { ...e, order: e.order + '' }
					// })
				}
				ques_list.push(item)
			}
			var feature_list = []
			var partical_no_dot_list = []
			if (drawings) {
				for (var j = 0, jmax = drawings.length; j < jmax; ++j) {
					var oDrawing = drawings[j]
					if (oDrawing.Drawing.docPr) {
						var title = oDrawing.Drawing.docPr.title
						if (title && title.indexOf('partical_no_dot') >= 0) {
							partical_no_dot_list.push({
								page: oDrawing.Drawing.PageNum + 1,
								x: mmToPx(oDrawing.Drawing.X),
								y: mmToPx(oDrawing.Drawing.Y),
								w: mmToPx(oDrawing.Drawing.Width),
								h: mmToPx(oDrawing.Drawing.Height),
							})
						}
						if (title && title.indexOf('feature') >= 0) {
							var titleObj = getJsonData(title)
							if (
								titleObj.feature &&
								titleObj.feature.zone_type &&
								titleObj.feature.zone_type != 'question'
							) {
								var featureObj = {
									zone_type: titleObj.feature.zone_type,
									fields: [],
								}
								if (
									titleObj.feature.zone_type == 'statistics' ||
									titleObj.feature.zone_type == 'pagination'
								) {
									let statistics_arr = []
									for (var p = 0; p < pageCount; ++p) {
										statistics_arr.push({
											v: p + 1 + '',
											page: p + 1,
											x: mmToPx(oDrawing.Drawing.X),
											y: mmToPx(oDrawing.Drawing.Y),
											w: mmToPx(oDrawing.Drawing.Width),
											h: mmToPx(oDrawing.Drawing.Height),
										})
									}
									feature_list.push({
										zone_type: titleObj.feature.zone_type,
										fields: statistics_arr,
									})
								} else {
									if (
										titleObj.feature.zone_type == 'self_evaluation' ||
										titleObj.feature.zone_type == 'teacher_evaluation'
									) {
										var oShape = oShapes.find((e) => {
											return e.Drawing && e.Drawing.Id == oDrawing.Drawing.Id
										})
										if (oShape) {
											var tables = oShape.GetContent().GetAllTables()
											if (tables && tables.length) {
												var oRow = tables[0].GetRow(0)
												if (oRow) {
													var CellsInfo = oRow.Row.CellsInfo
													for (var c = 2; c < CellsInfo.length; ++c) {
														var cell = CellsInfo[c]
														featureObj.fields.push({
															v: c - 1 + '',
															page: oDrawing.Drawing.PageNum + 1,
															x: mmToPx(oDrawing.Drawing.X + cell.X_cell_start),
															y: mmToPx(oDrawing.Drawing.Y),
															w: mmToPx(cell.X_cell_end - cell.X_cell_start),
															h: mmToPx(oDrawing.Drawing.Height),
														})
													}
												}
											} else {
												console.log('cannot find tables', oShapes, oDrawing)
											}
										} else {
											console.log('cannot find oShape')
										}
									} else {
										featureObj.fields.push({
											v: titleObj.feature.v + '',
											page: oDrawing.Drawing.PageNum + 1,
											x: mmToPx(oDrawing.Drawing.X),
											y: mmToPx(oDrawing.Drawing.Y),
											w: mmToPx(oDrawing.Drawing.Width),
											h: mmToPx(oDrawing.Drawing.Height),
										})
									}
									feature_list.push(featureObj)
								}
							}
						}
					}
				}
			}
			var paper_info = {
				pageCount: pageCount,
			}
			console.log('ques_list', ques_list)
			console.log('feature_list', feature_list)
			console.log('partical_no_dot_list', partical_no_dot_list)
			return {
				ques_list,
				feature_list,
				paper_info,
				partical_no_dot_list,
			}
		},
		false,
		false
	)
	// .then(res => {
	// 	console.log('the result of getAllPositions', res)
	//   return res
	// })
}

export {
	getPaperInfo,
	initPaperInfo,
	updatePageSizeMargins,
	updateCustomControls,
	clearStruct,
	getStruct,
	showQuestionTree,
	drawPosition,
	addScoreField,
	handleContentControlChange,
	handleScoreField,
	drawPositions,
	deletePositions,
	setSectionColumn,
	getAllPositions,
	addOnlyBigControl,
	removeOnlyBigControl,
	getAllPositions2
}
