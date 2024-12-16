// 选择题区别于其他题目，有多个特殊处理，另设置一个文件进行统一处理
import { JSONPath } from '../vendor/jsonpath-plus/dist/index-browser-esm.js';
import { biyueCallCommand } from './command.js';

function major_pos(path) {
    var ret = path.match(/\d+/g);
    if (ret) {
        return parseInt(ret[0]);
    }
    return 0;
}
function last_paragraph(path) {
	if (!path) {
		return
	}
    var ids = path.match(/\d+/g).map(function (item) {
        return parseInt(item);
    });

    var pIdx  = ids.length - 3;

    for (var i = pIdx; i < ids.length; i++) {
        ids[i] = -1;
    }

    return path.replace(/\[\d+\]/g, function (match) {
        return "[" + ids.shift() + "]";
    });
}
// 遍历element，提取选项
function getChoiceQuesData() {
	Asc.scope.node_list = window.BiyueCustomData.node_list
	Asc.scope.question_map = window.BiyueCustomData.question_map
	return biyueCallCommand(window, function() {
		var question_map = Asc.scope.question_map || {}
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls() || []
		function getControlsByClientId(cid) {
			var findControls = controls.filter(e => {
				var tag = Api.ParseJSON(e.GetTag())
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
		function getPosPath(pos) {
			var path = ''
			if (pos) {
				for (var i = 0; i < pos.length; ++i) {
					if (pos[i].Class && pos[i].Class.Content) {
						path += `['content'][${pos[i].Position}]`
					}
				}
			}
			return path
		}
		function getControlPath(oControl) {
			if (!oControl) {
				return {}
			}
			var docPos = oControl.Sdt.GetDocumentPositionFromObject() || []
			var doc_path = '$' + getPosPath(docPos)
			var begin_path = ''
			var end_path = ''
			var oRange = oControl.GetRange()
			if (oRange.StartPos) {
				var sPos = oRange.StartPos.slice(docPos.length, oRange.StartPos.length)
				begin_path = `$['sdtContent']` + getPosPath(sPos)
			}
			if (oRange.EndPos) {
				var ePos = oRange.EndPos.slice(docPos.length, oRange.EndPos.length)
				end_path = `$['sdtContent']` + getPosPath(ePos)
			}
			return {
				doc_path, begin_path, end_path
			}
		}
		function getNextElement(i3, runContent, i2, oElement) {
			if (i3 + 1 < runContent.length) {
				return {
					i2: i2,
					i3: i3 + 1,
					Data: runContent[i3 + 1]
				}
			} else {
				for (var i = i2 + 1; i < oElement.GetElementsCount(); ++i) {
					var oRun = oElement.GetElement(i)
					if (oRun.GetClassType() != 'run') {
						continue
					}
					var runContent2 = oRun.Run.Content || []
					for (var i4 = 0; i4 < runContent2.length; ++i4) {
						return {
							i2: i,
							i3: i4,
							Data: runContent2[i4]
						}
					}
				}
			}
			return null
		}
		function getOption(oControl) {
			if (!oControl) {
				return null
			}
			var oControlContent = oControl.GetContent()
			var elementCount = oControlContent.GetElementsCount()
			var opt_list = []
			var flag = 0
			var prePath = {}
			var steam_end = null
			var firstType = null
			var firstValue = null
			for (var i = 0; i < elementCount; i++) {
				var oElement = oControlContent.GetElement(i)
				if (oElement.GetClassType() == 'paragraph') {
					var count2 = oElement.Paragraph.Content ? oElement.Paragraph.Content.length : 0
					for (var i2 = 0; i2 < count2; i2++) {
						var oRun = oElement.GetElement(i2)
						if (!oRun) {
							var run = oElement.Paragraph.Content[i2]
							if (run) {
								var runContent = run.Content || []
								if (!flag) {
									prePath = {
										i: i,
										i2: i2,
										i3: runContent.length - 1
									}
								}
							}
							continue
						}
						if (oRun.GetClassType() != 'run') { // 目前只针对run进行处理
							continue
						}
						var runContent = oRun.Run.Content || []
						var count3 = runContent.length
						for (var i3 = 0; i3 < count3; ++i3) {
							var element3 = runContent[i3]
							var type = element3.GetType()
							if (type == 1) { // CRunText								
								var flag2 = 0
								var beginPath = ''
								if (element3.Value >= 65 && element3.Value <= 72 && (!firstType || firstType == 'A')) { // A-H
									var nextElement = getNextElement(i3, runContent, i2, oElement)
									if (nextElement) {
										if (nextElement.Data.Value == 46) { // .
											var nextNextElement = getNextElement(nextElement.i3, oElement.GetElement(nextElement.i2).Run.Content, nextElement.i2, oElement)
											if (nextNextElement && (nextNextElement.Data.Value == 32 || nextNextElement.Data.Value == 12288)) { // 空格 or 全角空格
												flag2 = 'A'
												var optElement = getNextElement(nextNextElement.i3, oElement.GetElement(nextNextElement.i2).Run.Content, nextNextElement.i2, oElement)
												if (optElement) {
													beginPath = `$['sdtContent']` + `['content'][${i}]['content'][${optElement.i2}]['content'][${optElement.i3}]`
												}
											}
										} else if (nextElement.Data.Value == 65294) { // ．全角的点
											flag2 = 'A'
											var optElement = getNextElement(nextElement.i3, oElement.GetElement(nextElement.i2).Run.Content, nextElement.i2, oElement)
											if (optElement) {
												beginPath = `$['sdtContent']` + `['content'][${i}]['content'][${optElement.i2}]['content'][${optElement.i3}]`
											}
										}
									}
								} else if (element3.Value >= 9312 && element3.Value <= 9320 && (!firstType || firstType == 'c1')) { // 圈1-圈9
									var nextElement = getNextElement(i3, runContent, i2, oElement)
									if (nextElement && (nextElement.Data.Value == 32 || nextElement.Data.Value == 12288)) { // 空格
										flag2 = 'c1'
										var optElement = getNextElement(nextElement.i3, oElement.GetElement(nextElement.i2).Run.Content, nextElement.i2, oElement)
										if (optElement) {
											beginPath = `$['sdtContent']` + `['content'][${i}]['content'][${optElement.i2}]['content'][${optElement.i3}]`
										}
									}
								}
								if (flag2) {
									if (!firstType) {
										firstType = flag2
										firstValue = String.fromCharCode(element3.Value)
									}
									flag = 1
									if (opt_list.length) {
										opt_list[opt_list.length - 1].endPath = `$['sdtContent']` + `['content'][${i}]['content'][${i2}]['content'][${i3}]`
									} else {
										if (prePath && prePath.i != undefined) {
											steam_end = `$['sdtContent']` + `['content'][${prePath.i}]`
											if (prePath.i2 != undefined) {
												steam_end += `['content'][${prePath.i2}]`
												if (prePath.i3 != undefined) {
													steam_end += `['content'][${prePath.i3}]`
												}
											}
										}
									}
									opt_list.push({
										value: element3.Value,
										beginPath: beginPath
									})
								}
							} else if (type == 4) { // CRunParagraphMark
								if (opt_list.length && !opt_list[opt_list.length - 1].endPath) {
									opt_list[opt_list.length - 1].endPath = `$['sdtContent']` + `['content'][${i}]['content'][${i2}]['content'][${i3}]`
								}
							}
							if (flag == 0) {
								prePath = {
									i: i,
									i2: i2,
									i3: i3
								}
							}
						}
					}
				}
			}
			return {
				steam_end: steam_end,
				options: opt_list,
				option_type: firstValue
			}
		}
		function getHtml(beg, end) {
			if (beg == end) {
				return ''
			}
			var oRange = Api.asc_MakeRangeByPath(beg, end)
			oRange.Select()
			let text_data = {
				data: "",
				// 返回的数据中class属性里面有binary格式的dom信息，需要删除掉
				pushData: function (format, value) {
					this.data = value ? value.replace(/class="[a-zA-Z0-9-:;+"\/=]*/g, "") : "";
				}
			};

			Api.asc_CheckCopy(text_data, 2);
			return text_data.data
		}
		var list = []
		var choiceMap = {}
		Object.keys(question_map).forEach(qid => {
			var ques = question_map[qid]
			if (ques.level_type == 'question' && (ques.ques_mode == 1 || ques.ques_mode == 5)) {
				var obj = {
					ques_id: qid,
					items: []
				}
				choiceMap[qid] = {
					steam: '',
					options: [],
					option_type: ''
				}
				if (ques.is_merge && ques.ids) {
					ques.ids.forEach(id => {
						var oControl = getControlsByClientId(id)
						if (oControl) {
							var paths = getControlPath(oControl)
							var data = getOption(oControl)
							if (data.options.length) {
								data.options.forEach((e, index) => {
									var html = getHtml(paths.doc_path + e.beginPath, paths.doc_path + e.endPath)
									choiceMap[qid].options.push({
										value: String.fromCharCode(e.value),
										html: html
									})
								})
								choiceMap[qid].option_type = data.option_type
								if (data.steam_end) {
									choiceMap[qid].steam += getHtml(paths.doc_path + paths.begin_path, paths.doc_path + data.steam_end)
								}
							} else {
								choiceMap[qid].steam += getHtml(paths.doc_path + paths.begin_path, paths.doc_path + data.end_path)
							}
						}
					})
				} else {
					var oControl = getControlsByClientId(qid)
					if (oControl) {
						var paths = getControlPath(oControl)
						var data = getOption(oControl)

						data.options.forEach((e, index) => {
							var html = getHtml(paths.doc_path + e.beginPath, paths.doc_path + e.endPath)
							choiceMap[qid].options.push({
								value: String.fromCharCode(e.value),
								html: html
							})
						})
						choiceMap[qid].steam = getHtml(paths.doc_path + paths.begin_path, paths.doc_path + data.steam_end)
						choiceMap[qid].option_type = data.option_type
					}
				}
			}
		})
		return choiceMap
	})
	// .then(res => {
	// 		$('#choiceOptions').empty()
	// 		console.log('================ dddddd', res)
	// 		Object.values(res).forEach(ques => {
	// 			$('#choiceOptions').append(`<div>题干：</div>`)
	// 			$('#choiceOptions').append(`<div>${ques.steam}</div>`)
	// 			$('#choiceOptions').append(`<div>选项：</div>`)
	// 			ques.options.forEach((e2, index) => {
	// 				$('#choiceOptions').append(`<div>${e2}</div>`)
	// 				// $('#choiceOptions').append(e2)
	// 			})
	// 		})
	// 	})
}
// 选项已切成control, 获取题干和选项
function getChoiceOptionAndSteam(ids) {
	Asc.scope.node_list = window.BiyueCustomData.node_list
	Asc.scope.question_map = window.BiyueCustomData.question_map
	Asc.scope.ids = ids
	return biyueCallCommand(window, function() {
		var question_map = Asc.scope.question_map || {}
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls() || []
		var ids = Asc.scope.ids || []
		function getControlsByClientId(cid) {
			var findControls = controls.filter(e => {
				var tag = Api.ParseJSON(e.GetTag())
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
		function getPosPath(pos) {
			var path = ''
			if (pos) {
				for (var i = 0; i < pos.length; ++i) {
					if (pos[i].Class && pos[i].Class.Content) {
						path += `['content'][${pos[i].Position}]`
					}
				}
			}
			return path
		}
		function getControlPath(oControl) {
			if (!oControl) {
				return {}
			}
			var docPos = oControl.Sdt.GetDocumentPositionFromObject() || []
			var doc_path = '$' + getPosPath(docPos)
			var begin_path = ''
			var end_path = ''
			var oRange = oControl.GetRange()
			if (oRange.StartPos) {
				var sPos = oRange.StartPos.slice(docPos.length, oRange.StartPos.length)
				begin_path = `$['sdtContent']` + getPosPath(sPos)
			}
			if (oRange.EndPos) {
				var ePos = oRange.EndPos.slice(docPos.length, oRange.EndPos.length)
				end_path = `$['sdtContent']` + getPosPath(ePos)
			}
			return {
				doc_path, begin_path, end_path
			}
		}
		function getOptionControlPath(oControl, quesControl) {
			if (!oControl) {
				return {}
			}
			var quesPos = quesControl.Sdt.GetDocumentPositionFromObject() || []
			var docPos = oControl.Sdt.GetDocumentPositionFromObject() || []
			var dPos = docPos.slice(quesPos.length, docPos.length)
			var doc_path = `$['sdtContent']` + getPosPath(dPos)
			var begin_path = ''
			var end_path = ''
			var oRange = oControl.GetRange()
			if (oRange.StartPos) {
				var sPos = oRange.StartPos.slice(docPos.length, oRange.StartPos.length)
				begin_path = getPosPath(sPos)
			}
			if (oRange.EndPos) {
				var ePos = oRange.EndPos.slice(docPos.length, oRange.EndPos.length)
				end_path = getPosPath(ePos)
			}
			return {
				doc_path, begin_path, end_path
			}
		}
		function getNextElement(i3, runContent, i2, oElement) {
			if (i3 + 1 < runContent.length) {
				return {
					i2: i2,
					i3: i3 + 1,
					Data: runContent[i3 + 1]
				}
			} else {
				for (var i = i2 + 1; i < oElement.GetElementsCount(); ++i) {
					var oRun = oElement.GetElement(i)
					if (oRun.GetClassType() != 'run') {
						continue
					}
					var runContent2 = oRun.Run.Content || []
					for (var i4 = 0; i4 < runContent2.length; ++i4) {
						return {
							i2: i,
							i3: i4,
							Data: runContent2[i4]
						}
					}
				}
			}
			return null
		}
		function getOption(oControl) {
			if (!oControl) {
				return null
			}
			var steam_end = null
			var opt_list = []
			var firstType = null
			var firstValue = null
			var flag = 0
			var prePath = {}
			var oControlContent = oControl.GetContent()
			var elementCount = oControlContent.GetElementsCount()
			var childControls = oControl.GetAllContentControls() || []
			for (var oChildControl of childControls) {
				var childTag = Api.ParseJSON(oChildControl.GetTag())
				if (childTag.regionType != 'choiceOption') {
					continue
				}
				var childPaths = getOptionControlPath(oChildControl, oControl)
				var beginPath = ''
				if (!steam_end) {
					steam_end = childPaths.doc_path + childPaths.begin_path
				}
				if (oChildControl.GetClassType() == 'inlineLvlSdt') {
					var count = oChildControl.GetElementsCount()
					for (var i2 = 0; i2 < count; ++i2) {
						var oRun = oChildControl.GetElement(i2)
						if (oRun.GetClassType() != 'run') {
							continue
						}
						var runContent = oRun.Run.Content || []
						var count3 = runContent.length
						for (var i3 = 0; i3 < count3; ++i3) {
							var element3 = runContent[i3]
							var type = element3.GetType()
							if (type == 1) {
								var flag2 = 0
								var isecond = i2
								var ithird = i3
								if (element3.Value >= 65 && element3.Value <= 72 && (!firstType || firstType == 'A')) { // A-H
									var nextElement = getNextElement(i3, runContent, i2, oChildControl)
									if (nextElement) {
										isecond = nextElement.i2
										ithird = nextElement.i3
										if (nextElement.Data.Value == 46) { // .
											var nextNextElement = getNextElement(nextElement.i3, oChildControl.GetElement(nextElement.i2).Run.Content, nextElement.i2, oChildControl)
											if (nextNextElement) {
												isecond = nextNextElement.i2
												ithird = nextNextElement.i3
												if ((nextNextElement.Data.Value == 32 || nextNextElement.Data.Value == 12288)) { // 空格 or 全角空格
													flag2 = 'A'
													var optElement = getNextElement(nextNextElement.i3, oChildControl.GetElement(nextNextElement.i2).Run.Content, nextNextElement.i2, oChildControl)
													if (optElement) {
														isecond = optElement.i2
														ithird = optElement.i3
													}
												}
											}
										} else if (nextElement.Data.Value == 65294) { // ．全角的点
											flag2 = 'A'
											var optElement = getNextElement(nextElement.i3, oChildControl.GetElement(nextElement.i2).Run.Content, nextElement.i2, oChildControl)
											if (optElement) {
												isecond = optElement.i2
												ithird = optElement.i3
											}
										} else if (nextElement.Data.Value == 32 || nextElement.Data.Value == 12288) { // 空格或全角空格
											var nextNextElement = getNextElement(nextElement.i3, oChildControl.GetElement(nextElement.i2).Run.Content, nextElement.i2, oChildControl)
											if (nextNextElement) {
												isecond = nextNextElement.i2
												ithird = nextNextElement.i3
												if ((nextNextElement.Data.Value == 46 || nextNextElement.Data.Value == 65294)) { // .
													flag2 = 'A'
													var optElement = getNextElement(nextNextElement.i3, oChildControl.GetElement(nextNextElement.i2).Run.Content, nextNextElement.i2, oChildControl)
													if (optElement) {
														isecond = optElement.i2
														ithird = optElement.i3
													}
												}
											}
										}
									}
								} else if (element3.Value >= 9312 && element3.Value <= 9320 && (!firstType || firstType == 'c1')) { // 圈1-圈9
									var nextElement = getNextElement(i3, runContent, i2, oChildControl)
									if (nextElement) {
										isecond = nextElement.i2
										ithird = nextElement.i3
										if (nextElement.Data.Value == 32 || nextElement.Data.Value == 12288) { // 空格
											flag2 = 'c1'
											var optElement = getNextElement(nextElement.i3, oChildControl.GetElement(nextElement.i2).Run.Content, nextElement.i2, oChildControl)
											if (optElement) {
												isecond = optElement.i2
												ithird = optElement.i3
											}
										}
									}
								}
								beginPath = `['content'][${isecond}]['content'][${ithird}]`
								if (beginPath) {
									if (opt_list.length == 0) {
										firstValue = String.fromCharCode(element3.Value)
									}
									opt_list.push({
										value: element3.Value,
										beginPath: beginPath,
										...childPaths
									})
									break
								}
							}
						}
						if (beginPath) {
							break
						}
					}
				}
			}
			return {
				steam_end: steam_end,
				options: opt_list,
				option_type: firstValue
			}
		}
		function getHtml(beg, end) {
			if (beg == end) {
				return ''
			}
			var oRange = Api.asc_MakeRangeByPath(beg, end)
			oRange.Select()
			let text_data = {
				data: "",
				// 返回的数据中class属性里面有binary格式的dom信息，需要删除掉
				pushData: function (format, value) {
					this.data = value ? value.replace(/class="[a-zA-Z0-9-:;+"\/=]*/g, "") : "";
				}
			};

			Api.asc_CheckCopy(text_data, 2);
			return text_data.data
		}
		var choiceMap = {}
		Object.keys(question_map).forEach(qid => {
			var ques = question_map[qid]
			if (ques.level_type == 'question' && (ques.ques_mode == 1 || ques.ques_mode == 5) && (ids.length == 0 || ids.indexOf(qid) != -1)) {
				choiceMap[qid] = {
					steam: '',
					options: [],
					option_type: ''
				}
				if (ques.is_merge && ques.ids) {
					ques.ids.forEach(id => {
						var oControl = getControlsByClientId(id)
						if (oControl) {
							var paths = getControlPath(oControl)
							var data = getOption(oControl)
							if (data.options.length) {
								data.options.forEach((e, index) => {
									var html = getHtml(paths.doc_path + e.doc_path + e.beginPath, paths.doc_path + e.doc_path + e.end_path)
									choiceMap[qid].options.push({
										value: String.fromCharCode(e.value),
										html: html
									})
								})
								choiceMap[qid].option_type = data.option_type
								if (data.steam_end) {
									choiceMap[qid].steam += getHtml(paths.doc_path + paths.begin_path, paths.doc_path + data.steam_end)
								}
							} else {
								choiceMap[qid].steam += getHtml(paths.doc_path + paths.begin_path, paths.doc_path + data.end_path)
							}
						}
					})
				} else {
					var oControl = getControlsByClientId(qid)
					if (oControl) {
						var paths = getControlPath(oControl)
						var data = getOption(oControl)
						data.options.forEach((e, index) => {
							var html = getHtml(paths.doc_path + e.doc_path + e.beginPath, paths.doc_path + e.doc_path + e.end_path)
							choiceMap[qid].options.push({
								value: String.fromCharCode(e.value),
								html: html
							})
						})
						choiceMap[qid].steam = getHtml(paths.doc_path + paths.begin_path, paths.doc_path + data.steam_end)
						choiceMap[qid].option_type = data.option_type
					}
				}
			}
		})
		return choiceMap
	}, false, false)
}

function removeChoiceOptions(ids) {
	Asc.scope.node_list = window.BiyueCustomData.node_list
	Asc.scope.question_map = window.BiyueCustomData.question_map
	Asc.scope.ids = ids
	return biyueCallCommand(window, function() {
		var question_map = Asc.scope.question_map || {}
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls() || []
		var ids = Asc.scope.ids || []
		var list = []
		for (var oControl of controls) {
			if (oControl.GetClassType() != 'blockLvlSdt') {
				continue
			}
			var tag = Api.ParseJSON(oControl.GetTag())
			var qid = tag.mid || tag.client_id
			if (ids.length) {
				if (ids.findIndex(e => {
					return (e + '') == (qid + '')
				}) == -1) {
					continue
				}
			} else {
				var quesData = question_map[qid]
				if (quesData) {
					if (quesData.ques_mode != 1 && quesData.ques_mode != 5) {
						continue
					}
				}
			}
			list.push(oControl)
		}
		if (!list.length) {
			return
		}
		var newlist = []
		for (var oControl of list) {
			// 先移除原有选项
			var childControls = oControl.GetAllContentControls() || []
			for (var oChildControl of childControls) {
				var childTag = Api.ParseJSON(oChildControl.GetTag())
				if (childTag.regionType == 'choiceOption') {
					Api.asc_RemoveContentControlWrapper(oChildControl.Sdt.GetId())
				}
			}
			var tag = Api.ParseJSON(oControl.GetTag())
			var qid = tag.mid || tag.client_id
			var quesData = question_map[qid]
			if (!quesData || quesData.level_type != 'question' || (quesData.ques_mode != 1 && quesData.ques_mode != 5)) {
				continue
			}
			newlist.push(oControl.Sdt.GetId())
		}
		return newlist
	}, false, false)
}
// 提取选择题选项
function extractChoiceOptions(ids, calc) {
	return removeChoiceOptions(ids)
	.then((res) => {
		Asc.scope.ids = res
		return biyueCallCommand(window, function() {
			function getPosPath(pos) {
				var path = ''
				if (pos) {
					for (var i = 0; i < pos.length; ++i) {
						if (pos[i].Class && pos[i].Class.Content) {
							path += `['content'][${pos[i].Position}]`
						}
					}
				}
				return path
			}
			function getControlPath(oControl) {
				if (!oControl) {
					return {}
				}
				var docPos = oControl.Sdt.GetDocumentPositionFromObject() || []
				var doc_path = '$' + getPosPath(docPos)
				var begin_path = ''
				var end_path = ''
				var oRange = oControl.GetRange()
				if (oRange.StartPos) {
					var sPos = oRange.StartPos.slice(docPos.length, oRange.StartPos.length)
					begin_path = `$['sdtContent']` + getPosPath(sPos)
				}
				if (oRange.EndPos) {
					var ePos = oRange.EndPos.slice(docPos.length, oRange.EndPos.length)
					end_path = `$['sdtContent']` + getPosPath(ePos)
				}
				return {
					doc_path, begin_path, end_path
				}
			}
			function getNextElement(i3, runContent, i2, oElement) {
				if (i3 + 1 < runContent.length) {
					return {
						i2: i2,
						i3: i3 + 1,
						Data: runContent[i3 + 1]
					}
				} else {
					for (var i = i2 + 1; i < oElement.GetElementsCount(); ++i) {
						var oRun = oElement.GetElement(i)
						if (oRun.GetClassType() != 'run') {
							continue
						}
						var runContent2 = oRun.Run.Content || []
						for (var i4 = 0; i4 < runContent2.length; ++i4) {
							return {
								i2: i,
								i3: i4,
								Data: runContent2[i4]
							}
						}
					}
				}
				return null
			}
			function getOption(oControl) {
				if (!oControl) {
					return null
				}
				var oControlContent = oControl.GetContent()
				var elementCount = oControlContent.GetElementsCount()
				var opt_list = []
				var flag = 0
				var prePath = {}
				var steam_end = null
				var firstType = null
				var firstValue = null
				for (var i = 0; i < elementCount; i++) {
					var oElement = oControlContent.GetElement(i)
					if (oElement.GetClassType() == 'paragraph') {
						var count2 = oElement.Paragraph.Content ? oElement.Paragraph.Content.length : 0
						for (var i2 = 0; i2 < count2; i2++) {
							var oRun = oElement.GetElement(i2)
							if (!oRun) {
								var run = oElement.Paragraph.Content[i2]
								if (run) {
									var runContent = run.Content || []
									if (!flag) {
										prePath = {
											i: i,
											i2: i2,
											i3: runContent.length - 1
										}
									}
								}
								continue
							}
							if (oRun.GetClassType() != 'run') { // 目前只针对run进行处理
								continue
							}
							var runContent = oRun.Run.Content || []
							var count3 = runContent.length
							for (var i3 = 0; i3 < count3; ++i3) {
								var element3 = runContent[i3]
								var type = element3.GetType()
								if (type == 1) { // CRunText
									var flag2 = 0
									var beginPath = ''
									if (element3.Value >= 65 && element3.Value <= 72 && (!firstType || firstType == 'A')) { // A-H
										var nextElement = getNextElement(i3, runContent, i2, oElement)
										if (nextElement) {
											if (nextElement.Data.Value == 65294 || nextElement.Data.Value == 46) { // ．全角半角都算
												flag2 = 'A'
												var optElement = getNextElement(nextElement.i3, oElement.GetElement(nextElement.i2).Run.Content, nextElement.i2, oElement)
												if (optElement) {
													beginPath = `$['sdtContent']` + `['content'][${i}]['content'][${i2}]['content'][${i3}]`
												}
											} else if (nextElement.Data.Value == 32 || nextElement.Data.Value == 12288) { // 空格 or 全角空格
												var nextNextElement = getNextElement(nextElement.i3, oElement.GetElement(nextElement.i2).Run.Content, nextElement.i2, oElement)
												if (nextNextElement && (nextNextElement.Data.Value == 46 || nextNextElement.Data.Value == 65294)) {
													flag2 = 'A'
													var optElement = getNextElement(nextNextElement.i3, oElement.GetElement(nextNextElement.i2).Run.Content, nextNextElement.i2, oElement)
													if (optElement) {
														beginPath = `$['sdtContent']` + `['content'][${i}]['content'][${i2}]['content'][${i3}]`
													}
												}
											}
										}
									} else if (element3.Value >= 9312 && element3.Value <= 9320 && (!firstType || firstType == 'c1')) { // 圈1-圈9
										var nextElement = getNextElement(i3, runContent, i2, oElement)
										if (nextElement && (nextElement.Data.Value == 32 || nextElement.Data.Value == 12288)) { // 空格
											flag2 = 'c1'
											var optElement = getNextElement(nextElement.i3, oElement.GetElement(nextElement.i2).Run.Content, nextElement.i2, oElement)
											if (optElement) {
												beginPath = `$['sdtContent']` + `['content'][${i}]['content'][${i2}]['content'][${i3}]`
											}
										}
									}
									if (flag2) {
										if (!firstType) {
											firstType = flag2
											firstValue = String.fromCharCode(element3.Value)
										}
										flag = 1
										if (opt_list.length) {
											opt_list[opt_list.length - 1].endPath = `$['sdtContent']` + `['content'][${i}]['content'][${i2}]['content'][${i3}]`
										} else {
											if (prePath && prePath.i != undefined) {
												steam_end = `$['sdtContent']` + `['content'][${prePath.i}]`
												if (prePath.i2 != undefined) {
													steam_end += `['content'][${prePath.i2}]`
													if (prePath.i3 != undefined) {
														steam_end += `['content'][${prePath.i3}]`
													}
												}
											}
										}
										opt_list.push({
											value: element3.Value,
											beginPath: beginPath
										})
									}
								} else if (type == 4) { // CRunParagraphMark
									if (opt_list.length && !opt_list[opt_list.length - 1].endPath) {
										opt_list[opt_list.length - 1].endPath = `$['sdtContent']` + `['content'][${i}]['content'][${i2}]['content'][${i3}]`
									}
								}
								if (flag == 0) {
									prePath = {
										i: i,
										i2: i2,
										i3: i3
									}
								}
							}
						}
					}
				}
				return {
					steam_end: steam_end,
					options: opt_list,
					option_type: firstValue
				}
			}
			var ids = Asc.scope.ids || []
			for (var controlId of ids) {
				var oControl = Api.LookupObject(controlId)
				var paths = getControlPath(oControl)
				var data = getOption(oControl)
				data.options.reverse().forEach((e, index) => {
					var oRange = Api.asc_MakeRangeByPath(paths.doc_path + e.beginPath, paths.doc_path + e.endPath)
					oRange.Select()
					var tag = JSON.stringify({ 'regionType': 'choiceOption', 'mode': 3, 'color': '#00ff0020' });
					Api.asc_AddContentControl(2, { "Tag": tag });
					Api.asc_RemoveSelection();
				})
			}
		}, false, calc)
	})
}
// 获取选择题信息 使用jsonpath的方式提取选项，会出错，已弃用
function getChoiceQuesInfo() {
	Asc.scope.node_list = window.BiyueCustomData.node_list
	Asc.scope.question_map = window.BiyueCustomData.question_map
	return biyueCallCommand(window, function() {
		var node_list = Asc.scope.node_list || []
		var question_map = Asc.scope.question_map || {}
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls() || []
		function getControlsByClientId(cid) {
			var findControls = controls.filter(e => {
				var tag = Api.ParseJSON(e.GetTag())
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
		function getPosPath(pos) {
			var path = ''
			if (pos) {
				for (var i = 0; i < pos.length; ++i) {
					if (pos[i].Class && pos[i].Class.Content) {
						path += `['content'][${pos[i].Position}]`
					}
				}
			}
			return path
		}
		function getControlPath(oControl) {
			if (!oControl) {
				return {}
			}
			var docPos = oControl.Sdt.GetDocumentPositionFromObject() || []
			var doc_path = '$' + getPosPath(docPos)
			var begin_path = ''
			var end_path = ''
			var oRange = oControl.GetRange()
			if (oRange.StartPos) {
				var sPos = oRange.StartPos.slice(docPos.length, oRange.StartPos.length)
				begin_path = `$['sdtContent']` + getPosPath(sPos)
			}
			if (oRange.EndPos) {
				var ePos = oRange.EndPos.slice(docPos.length, oRange.EndPos.length)
				end_path = `$['sdtContent']` + getPosPath(ePos)
			}
			return {
				doc_path, begin_path, end_path
			}
		}

		var list = []
		Object.keys(question_map).forEach(qid => {
			var ques = question_map[qid]
			if (ques.level_type == 'question' && (ques.ques_mode == 1 || ques.ques_mode == 5)) {
				var obj = {
					ques_id: qid,
					items: []
				}
				if (ques.is_merge && ques.ids) {
					var oRange = null
					ques.ids.forEach(id => {
						var oControl = getControlsByClientId(id)
						if (oControl) {
							var parentcell = oControl.GetParentTableCell()
							if (parentcell) {
								if (!oRange) {
									oRange = parentcell.GetContent().GetRange()
								} else {
									oRange = oRange.ExpandTo(parentcell.GetContent().GetRange())
								}
								var pathData = getControlPath(oControl)
								obj.items.push({
									id: id,
									json: oControl.ToJSON(),
									control_id: oControl.Sdt.GetId(),
									cell_id: parentcell.Cell.Id,
									...pathData
								})
							}
						}
					})
					obj.json = oRange.ToJSON()
				} else {
					var oControl = getControlsByClientId(qid)
					if (oControl) {
						var pathData = getControlPath(oControl)
						obj.items.push({
							id: qid,
							json: oControl.ToJSON(),
							control_id: oControl.Sdt.GetId(),
							...pathData
						})
					}
				}
				if (obj.items.length) {
					list.push(obj)
				}
			}
		})
		return {
			list
		}
	}, false, false).then(res => {
		return new Promise((resolve, reject) => {
			var ranges = []
			var list = res.list
			if (!list || list.length == 0) {
				return resolve([])
			}
			const quesPatt = `$..content[?(typeof(@) == 'string' && @.match(/^[\u2460-\u24FF] |^[A-H]+[.．] /))]`
			list.forEach((ques, i) => {
				ranges.push({
					ques_id: ques.ques_id,
					items: ques.items.map(e => {
						return {
							id: e.id,
							control_id: e.control_id,
							cell_id: e.cell_id,
							doc_path: e.doc_path,
							begin_path: e.begin_path,
							end_path: e.end_path,
							options: []
						}
					})
				})
				var index = ranges.length - 1
				ques.items.forEach((item, index2) => {
					var k = JSON.parse(item.json)
					JSONPath({
						path: quesPatt, json: k, resultType: 'path', callback: function(res) {
							ranges[index].items[index2].options.push({
								beg: res,
								major_pos: major_pos(res),
								end: last_paragraph(res)
							})
						}
					})
				})
			})
			resolve(ranges)
		})
	}).then(ranges => {
		Asc.scope.options_ranges = ranges
		return biyueCallCommand(window, function() {
			var optionRanges = Asc.scope.options_ranges || []
			var choiceMap = {}
			function getHtml(beg, end) {
				if (beg == end) {
					return ''
				}
				var oRange = Api.asc_MakeRangeByPath(beg, end)
				oRange.Select()
				let text_data = {
					data: "",
					// 返回的数据中class属性里面有binary格式的dom信息，需要删除掉
					pushData: function (format, value) {
						this.data = value ? value.replace(/class="[a-zA-Z0-9-:;+"\/=]*/g, "") : "";
					}
				};

				Api.asc_CheckCopy(text_data, 2);
				return text_data.data
			}
			for (var ques of optionRanges) {
				choiceMap[ques.ques_id] = {
					options: [],
					steam: ''
				}
				var steam = ''
				var find = false
				ques.items.forEach((item, index) => {
					if (item.options && item.options.length) {
						find = true
						item.options.forEach((option, idx) => {
							var beg = item.doc_path + option.beg
							var end = item.doc_path + (idx < item.options.length - 1 ? item.options[idx + 1].beg : item.end_path)
							choiceMap[ques.ques_id].options.push(getHtml(beg, end))
						})
						steam += getHtml(item.doc_path + item.begin_path, item.doc_path + item.options[0].beg)
						find = true
					} else if (!find) {
						steam += getHtml(item.doc_path + item.begin_path, item.doc_path + item.end_path)
					}
				})
				choiceMap[ques.ques_id].steam = steam
			}
			return choiceMap
		}, false, false)
	})
	// .then(res => {
	// 	$('#choiceOptions').empty()
	// 	console.log('================ dddddd', res)
	// 	Object.values(res).forEach(ques => {
	// 		$('#choiceOptions').append(`<div>题干：</div>`)
	// 		$('#choiceOptions').append(`<div>${ques.steam}</div>`)
	// 		$('#choiceOptions').append(`<div>选项：</div>`)
	// 		ques.options.forEach((e2, index) => {
	// 			$('#choiceOptions').append(`<div>${e2}</div>`)
	// 			// $('#choiceOptions').append(e2)
	// 		})
	// 	})
	// })
}

// 获取选择题信息，使用html+正则表达式提取
function getChoiceQuesOption(e) {
	var quesMode = window.BiyueCustomData.question_map[e.id]
  if (quesMode == 1 || quesMode == 5) {
    var auto_align_patt = new RegExp(/(?<!data-latex="[^"]*)([A-F]|[\u2460-\u2469\u24EA])[.．].*?([&nbsp;|\s| |　]{2}|<\/p>)/g)
    var str = e.content_html || ''
    var arr = str.match(auto_align_patt) || []

    for (const key in arr) {
      const option = arr[key]
      const nexOption = arr[key * 1 + 1] || ''
      const index = str.indexOf(option)
      if (option.substring(option.length - 4) === '</p>') {
        arr[key] = option.substring(0, option.length - 4)
        str = str.replace(arr[key], '</p>')
      } else if (nexOption) {
        const end_index = str.indexOf(nexOption)
        const newItem = str.substring(index, end_index)

        arr[key] = str.substring(index, end_index)
        if (newItem.indexOf('</p>') >= 0) {
          arr[key] = arr[key].substring(0, newItem.indexOf('</p>'))
        }
        str = str.replace(arr[key], '')
      } else {
        const end_str = str.substring(index)
        const end_index = end_str.indexOf('</p>')
        if (end_index >= 0) {
          arr[key] = str.substring(index, index + end_index)
        }
        str = str.replace(arr[key], '')
      }
      /* 清空不完整的span标签 Start*/
      // 创建一个虚拟的 DOM 元素
      const div = document.createElement('div')
      div.innerHTML = arr[key]
      // 获取所有的 span 标签
      const spanTags = div.getElementsByTagName('span')
      // 遍历每个 span 标签，检查是否存在不完整的标签并清空
      Array.from(spanTags).forEach((span) => {
        const content = span.innerHTML
        const isCompleteTag = content && span.outerHTML.includes(content)

        if (!isCompleteTag) {
          span.innerHTML = ''
        }
      })
      // 获取处理后的文本
      arr[key] = div.innerHTML
      /* 清空不完整的span标签 End*/

      // 清除选项后面的空格
      var empty_patt = new RegExp(/.(\s*$)/g)
      const item_tail = arr[key].match(empty_patt)
      if (item_tail && item_tail[0] && item_tail[0].length > 1) {
        const end_str = item_tail[0].substring(0, 1)
        arr[key] = arr[key].replace(`/(?<!<img[^>]*>)(\b${item_tail[0]}\b)/g`, end_str)
      }
    }
    console.log('html------------->', e.content_html)
    console.log('options------------->', arr)
  }
}

// 设置选择题选项布局
function setChoiceOptionLayout(options) {
	Asc.scope.choice_align_options = options
	return biyueCallCommand(window, function() {
		var options = Asc.scope.choice_align_options
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		var part = options.part
		var identifyLeft = options.indLeft >= 0 ? (options.indLeft / ((25.4 / 72 / 20))) : 0 
		var spaceNum = options.spaceNum
		var bracket = options.bracket
		// 选项可能是A-H, 也可能是圈数字
		function handleBracket(oControl) {
			if (bracket == 'none' && spaceNum == 0) {
				return
			}
			var askControls = (oControl.GetAllContentControls() || []).filter(e => {
				if (e.GetClassType() == 'inlineLvlSdt') {
					var askTag = Api.ParseJSON(e.GetTag())
					return askTag.regionType == 'write' && askTag.client_id
				}
			})
			if (!askControls || askControls.length == 0) {
				return
			}
			for (var i = 0; i < askControls.length; ++i) {
				var askControl = askControls[i]
				var count = askControl.GetElementsCount()
				var flag = 0
				var spaceCount = 0
				for (var j = 0; j < count; ++j) {
					var oRun = askControl.GetElement(j)
					if (!oRun.GetClassType || oRun.GetClassType() != 'run') {
						continue
					}
					for (var k = 0; k < oRun.Run.GetElementsCount(); ++k) {
						var oElement2 = oRun.Run.GetElement(k)
						var newBarcket = null
						if (oElement2.Value == 65288 || oElement2.Value == 40) { // 左括号
							flag = 1
							if (oElement2.Value == 65288 && bracket == 'eng') {
								newBarcket = '('
							} else if (oElement2.Value == 40 && bracket == 'ch') {
								newBarcket = '（'
							}
						} else if (oElement2.Value == 65289 || oElement2.Value == 41) { // 右括号
							if (flag == 1 && spaceCount < spaceNum) {
								var str = ''
								for (var k2 = 0; k2 < spaceNum - spaceCount; ++k2) {
									str += ' '
								}
								spaceCount = spaceNum
								oRun.Run.AddText(str, k)
								k += str.length
								continue
							}
							if (oElement2.Value == 65289 && bracket == 'eng') {
								newBarcket = ')'
							} else if (oElement2.Value == 41 && bracket == 'ch') {
								newBarcket = '）'
							}

						} else if (flag == 1 && spaceNum > 0) {
							if (oElement2.Value == 32) { // 空格
								if (spaceCount >= spaceNum) {
									oRun.Run.RemoveElement(oElement2)
									--k
								} else {
									spaceCount++
								}
							} else {
								oRun.Run.RemoveElement(oElement2)
								--k
							}
						}
						if (newBarcket) {
							oRun.Run.RemoveElement(oElement2)
							oRun.Run.AddText(newBarcket, k)
						}
					}
				}
			}
		}
		function removeRunTab(oRun, removeBreak) {
			if (!oRun || !oRun.GetClassType || oRun.GetClassType() != 'run') {
				return
			}			
			for (var j = 0; j < oRun.Run.GetElementsCount(); ++j) {
				var oElement2 = oRun.Run.GetElement(j)
				if (!oElement2) {
					return
				}
				if (oElement2.GetType() == 21) { // tab
					oRun.Run.RemoveElement(oElement2)
					--j
				} else if (removeBreak && oElement2.GetType() == 16 && j == oRun.Run.GetElementsCount() - 1) {
					oRun.Run.RemoveElement(oElement2)
					--j
				}
			}
		}
		function removeTabs(oControl) {
			var paragrahs = oControl.GetAllParagraphs() || []
			for (var oParagraph of paragrahs) {
				for (var i = 0; i < oParagraph.GetElementsCount(); ++i) {
					var oElement = oParagraph.GetElement(i)
					var oType = oElement.GetClassType()
					if (oType != 'run' && oType != 'inlineLvlSdt') {
						continue
					}
					if (oType == 'run') {
						removeRunTab(oElement)
					} else if (oType == 'inlineLvlSdt') {
						for (var j = 0; j < oElement.GetElementsCount(); ++j) {
							var oElement2 = oElement.GetElement(j)
							removeRunTab(oElement2, true)
						}
					}
				}
			}
		}
		function getNextElement(i3, runContent, i2, oElement) {
			if (i3 + 1 < runContent.length) {
				return {
					i2: i2,
					i3: i3 + 1,
					Data: runContent[i3 + 1]
				}
			} else {
				for (var i = i2 + 1; i < oElement.GetElementsCount(); ++i) {
					var oRun = oElement.GetElement(i)
					if (oRun.GetClassType() != 'run') {
						continue
					}
					var runContent2 = oRun.Run.Content || []
					for (var i4 = 0; i4 < runContent2.length; ++i4) {
						return {
							i2: i,
							i3: i4,
							Data: runContent2[i4]
						}
					}
				}
			}
			return null
		}
		function addTabBreak2(count, newRun) {
			if (identifyLeft) {
				if (count) {
					if (part == 1) {
						// console.log('add line break 2')
						newRun.AddLineBreak()
					} else if (count % part == 0) {
						// console.log('add line break 1')
						newRun.AddLineBreak()
					}
				}
				// console.log('add tab stop 1')
				newRun.AddTabStop()
			} else if (count) {
				if (part == 1) {
					// console.log('add line break 3')
					newRun.AddLineBreak()
				} else {
					if (count % part == 0) {
						// console.log('add line break 4')
						newRun.AddLineBreak()
					} else {
						// console.log('add tab stop 2')
						newRun.AddTabStop()
					}
				}
			}
		}
		// 第一个选项是否与题干混在一起
		function isFirstOptionMixWithQuesTitle(i1, oParagraph) {
			for (var j1 = i1 - 1; j1 >= 0; --j1) {
				var lastElement = oParagraph.GetElement(i1 - 1)
				if (!lastElement) {
					break
				}
				if (lastElement.GetClassType() == 'run') {
					if (lastElement.Run.IsEmpty()) {
						continue
					}
					var runChild = lastElement.Run.GetElement(lastElement.Run.GetElementsCount() - 1)
					if (runChild && runChild.GetType() != 16) {
						return true
					}
					break
				} else if (lastElement.GetClassType() == 'inlineLvlSdt') {
					return true
				}
			}
			return false
		}
		function addTabs(oParagraph, optionMin, optionMax) {
			var optionList = []
			for (var i1 = 0; i1 < oParagraph.GetElementsCount(); ++i1) {
				var oElement = oParagraph.GetElement(i1)
				if (oElement.GetClassType() == 'run') {
					if (oElement.Run.IsEmpty()) {
						var preCount = oParagraph.GetElementsCount()
						oParagraph.RemoveElement(i1)
						var aflterCount = oParagraph.GetElementsCount()
						// 会存在precount == aflterCount 的情况，如果不加判断，会陷入死循环
						if (preCount > aflterCount) {
							--i1
						}
						continue
					}
					for (var i2 = 0; i2 < oElement.Run.GetElementsCount(); ++i2) {
						var oElement2 = oElement.Run.GetElement(i2)
						if (!oElement2) {
							continue
						}
						if (oElement2.GetType() == 1) {
							if (optionMin) {
								if (oElement2.Value < optionMin || oElement2.Value > optionMax) {
									continue
								}
							}
							var flag = 0
							if (oElement2.Value >= 65 && oElement2.Value <= 72) {
								var nextElement = getNextElement(i2, oElement.Run.Content, i1, oParagraph)
								if (nextElement && (nextElement.Data.Value == 46 || nextElement.Data.Value == 65294)) {
									flag = 1
									if (!optionMin) {
										optionMin = 65
										optionMax = 72
									}
								}
							} else if (oElement2.Value >= 9312 && oElement2.Value <= 9320) {
								var nextElement = getNextElement(i2, oElement.Run.Content, i1, oParagraph)
								if (nextElement && (nextElement.Data.Value == 32 || nextElement.Data.Value == 12288)) {
									flag = 1
									if (!optionMin) {
										optionMin = 9312
										optionMax = 9320
									}
								}
							}
							if (flag) {
								// 找到了
								if (optionList.length == 0) {
									// 若前一个字符不是break，则在前面添加break
									if (i1 > 0 && part > 0) {
										if (isFirstOptionMixWithQuesTitle(i1, oParagraph)) {
											var newRun = Api.CreateRun()
											newRun.AddLineBreak()
											oParagraph.AddElement(newRun, i1)
											++i1
										}
									}
								}
								if ((optionList.length == 0 && identifyLeft) || (optionList.length && optionList[optionList.length - 1].type == 'text')) { // 第1个, 且有左缩进
									if (i2 == 0) { // run的第1个字符, 在其前面添加tab
										var newRun = Api.CreateRun()
										addTabBreak2(optionList.length, newRun)
										oParagraph.AddElement(newRun, i1)
										++i1
									} else {
										var newRun = oElement.Run.Split_Run(i2)
										var newRun1 = Api.CreateRun()
										addTabBreak2(optionList.length, newRun1)
										oParagraph.AddElement(newRun1, i1 + 1)
										oParagraph.Paragraph.Add_ToContent(i1 + 2, newRun)
										i1 += 2
									}
								}
								optionList.push({
									index1: i1,
									index2: i2,
									type: 'text'
								})
							}
						} else if (oElement2.GetType() == 16 && optionList.length) { // linebreak
							var precount = oElement.Run.GetElementsCount()
							oElement.Run.RemoveElement(oElement2)
							if (oElement.Run.GetElementsCount() < precount) {
								--i2
							}
						}
					}
				} else if (oElement.GetClassType() == 'inlineLvlSdt') {
					var tag = Api.ParseJSON(oElement.GetTag())
					if (tag.regionType != 'choiceOption') {
						continue
					}
					if (optionList.length == 0) {
						// 若前一个字符不是break，则在前面添加break
						if (i1 > 0 && part > 0) {
							if (isFirstOptionMixWithQuesTitle(i1, oParagraph)) {
								var newRun = Api.CreateRun()
								newRun.AddLineBreak()
								oParagraph.AddElement(newRun, i1)
								++i1
							}
						}
					}
					if (identifyLeft || optionList.length) {
						var newRun = Api.CreateRun()
						addTabBreak2(optionList.length, newRun)
						oParagraph.AddElement(newRun, i1)
						++i1
					}
					optionList.push({
						index1: i1,
						type: 'control'
					})
				}
			}
		}
		for (var qdx = 0; qdx < controls.length; ++qdx) {
			var oControl = controls[qdx]
			if (oControl.GetClassType() != 'blockLvlSdt') {
				continue
			}
			if (oControl.GetPosInParent() < 0) {
				continue
			}
			var tag = Api.ParseJSON(oControl.GetTag())
			if (!tag.client_id) {
				continue
			}
			if (!(options.list.includes(tag.client_id))) {
				continue
			}
			var controlContent = oControl.GetContent()
			// 处理括号
			handleBracket(oControl)
			// 移除tab键
			removeTabs(oControl)
			var childControls = oControl.GetAllContentControls()
			var firstOptionControl = null
			var firstOption = ''
			var optionMin = 0
			var optionMax = 0
			// 获取第一个选项
			for (var j = 0; j < childControls.length; ++j) {
				var childTag = Api.ParseJSON(childControls[j].GetTag())
				if (childTag.regionType == 'choiceOption') {
					firstOptionControl = childControls[j]
					var text = firstOptionControl.GetRange().GetText()
					firstOption = text.charCodeAt(0)
					break
				}
			}

			var oOptionParagraph = null
			if (firstOption > 65 && firstOption <= 72) {
				optionMin = 65
				optionMax = 72
			} else if (firstOption > 9312 && firstOption <= 9320) {
				optionMin = 9312
				optionMax = 9320
			}
			// 获取选项段落
			for (var i1 = 0; i1 < controlContent.GetElementsCount(); ++i1) {
				var oElement1 = controlContent.GetElement(i1)
				if (!oElement1) {
					continue
				}
				if (oElement1.GetClassType() != 'paragraph') {
					continue
				}
				var text = oElement1.GetText()
				var patt = null
				if (optionMin) {
					patt = optionMin == 65 ? (new RegExp(/[A-H][.．]/g)) : (new RegExp(/^[\u2460-\u24FF][ 　]/g))
				} else {
					patt = new RegExp(/([\u2460-\u24FF][\u0020\u3000]|[A-H][.．][\u0020\u3000])/g)
				}
				var rlist = text.match(patt) || []
				if (rlist.length) {
					if (oOptionParagraph) {
						// 将所有选项挪到一个段落中
						controlContent.RemoveElement(i1)
						var elCount = oElement1.GetElementsCount()
						var insertPos = oOptionParagraph.GetElementsCount()
						for (var i = elCount; i >= 0; --i) {
							oOptionParagraph.AddElement(oElement1.GetElement(i), insertPos)
						}
						--i1
					} else {
						oOptionParagraph = oElement1
					}
				}
			}
			if (oOptionParagraph) {
				// 先移除缩进
				oOptionParagraph.SetIndLeft(0)
				oOptionParagraph.SetIndFirstLine(0)
				// 插入tab
				addTabs(oOptionParagraph, optionMin, optionMax)
				var section = oOptionParagraph.GetSection()
				var PageMargins = section.Section.PageMargins
				var PageSize = section.Section.PageSize
				var w = PageSize.W - PageMargins.Left - PageMargins.Right
				var newTabs = []
				var aligns = []
				if (options.indLeft) { // 存在左缩进
					newTabs.push(identifyLeft)
					aligns.push('left')
				}
				var tabs = []
				for (var i = 1; i < part; ++i) {
					tabs.push(i / part)
				}
				tabs.forEach(e => {
					newTabs.push((w * e) / (25.4 / 72 / 20) + identifyLeft)
					aligns.push('left')
				})
				oOptionParagraph.SetTabs(newTabs, aligns)
			} else {
				console.log('未找到选项所在段落')
			}
		}
	}, false, true)
}

export { 
	extractChoiceOptions,
	getChoiceQuesInfo, 
	getChoiceQuesOption, 
	getChoiceQuesData, 
	removeChoiceOptions,
	getChoiceOptionAndSteam,
	setChoiceOptionLayout
}