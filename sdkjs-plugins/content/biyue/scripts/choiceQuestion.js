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
			if (ques.level_type == 'question' && ques.ques_mode == 1) {
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
// 提取选择题选项
function extractChoiceOptions() {
	Asc.scope.node_list = window.BiyueCustomData.node_list
	Asc.scope.question_map = window.BiyueCustomData.question_map
	return biyueCallCommand(window, function() {
		var question_map = Asc.scope.question_map || {}
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls() || []
		var list = []
		controls.forEach(oControl => {
			var tag = Api.ParseJSON(oControl.GetTag())
			if (oControl.GetClassType() == 'inlineLvlSdt') {
				if (tag.regionType == 'choiceOption') {
					Api.asc_RemoveContentControlWrapper(oControl.Sdt.GetId())
				}
			}
		})
		controls.forEach(oControl => {
			var tag = Api.ParseJSON(oControl.GetTag())
			if (oControl.GetClassType() == 'blockLvlSdt' && oControl.GetPosInParent() >= 0) {
				if (tag.client_id &&
					question_map[tag.client_id] &&
					question_map[tag.client_id].level_type == 'question' &&
					question_map[tag.client_id].ques_mode == 1
				) {
					list.push({
						client_id: tag.client_id,
						control_id: oControl.Sdt.GetId(),
						json: oControl.ToJSON()
					})
				}
			}
		})
		return list
	}, false, true)
	.then(list => {
		return new Promise((resolve, reject) => {
			var ranges = []
			if (!list || list.length == 0) {
				reject('null')
			}
			list.forEach((ques, index) => {
				var k = JSON.parse(ques.json)
				const quesPatt = `$..content[?(typeof(@) == 'string' && @.match('^[A-F]+[.]'))]`;
				JSONPath({
					path: quesPatt, json: k, resultType: "path", callback: function (res) {
						ranges.push({
							client_id: ques.client_id,
							control_id: ques.control_id,
							beg: res,
							major_pos: major_pos(res),
							end: last_paragraph(res)
						})
					}
				});
			})
			resolve(ranges)
		})
	})
	.then(ranges => {
		Asc.scope.options_ranges = ranges
		return biyueCallCommand(window, function() {
			var optionRanges = Asc.scope.options_ranges || []
			if (optionRanges.length > 1) {
				for (var i = optionRanges.length - 1; i >= 0; --i) {
					var range = optionRanges[i]
					var oControl = Api.LookupObject(range.control_id)
					var pos = oControl.GetPosInParent()
					// 这里是假设题目一定在document的直接子集，todo 需要考虑题目在表格中的情况
					var beg = "$['content'][" + pos + "]" + range.beg
					var end = ''
					if (i < optionRanges.length - 1) {
						var nextControl = Api.LookupObject(optionRanges[i + 1].control_id)
						var nextPos = nextControl.GetPosInParent()
						end = "$['content'][" + nextPos + "]" + optionRanges[i + 1].beg
					} else {
						end = range.end
					}
					if (beg && end) {
						var newRange = Api.asc_MakeRangeByPath(beg, end)
						newRange.Select()
						var tag = JSON.stringify({ 'regionType': 'choiceOption', 'mode': 3, 'color': '#00ff0020' });
						Api.asc_AddContentControl(2, { "Tag": tag });
						Api.asc_RemoveSelection();
					}
				}
			}
		}, false, true)
	})
}
// 获取选择题信息
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
			if (ques.level_type == 'question' && ques.ques_mode == 1) {
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
		console.log('===== ranges', ranges)
		return biyueCallCommand(window, function() {
			var optionRanges = Asc.scope.options_ranges || []
			var choiceMap = {}
			function getHtml(beg, end) {
				if (beg == end) {
					return ''
				}
				console.log('beg:', beg)
				console.log('end:', end)
				var oRange = Api.asc_MakeRangeByPath(beg, end)
				console.log('range', oRange)
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

// 获取选择题信息
function getChoiceQuesOption(e) {
  if (window.BiyueCustomData.question_map[e.id].ques_mode == 1) {
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

export { extractChoiceOptions, getChoiceQuesInfo, getChoiceQuesOption, getChoiceQuesData }