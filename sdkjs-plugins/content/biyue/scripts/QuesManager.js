
import { biyueCallCommand, dispatchCommandResult } from "./command.js";
import { getQuesType, reqComplete } from '../scripts/api/paper.js'
var levelSetWindow = null
var level_map = {}
var g_click_value = null
var upload_control_list = []
function initExamTree() {
	
}

// 处理文档点击
function handleDocClick(isSelectionUse) {
	window.Asc.plugin.executeMethod('GetCurrentContentControlPr', [], function(returnValue) {
		console.log('GetCurrentContentControlPr', returnValue)
		if (returnValue && returnValue.Tag) {
			var tag = JSON.parse(returnValue.Tag || '{}')
			g_click_value = {
				InternalId: returnValue.InternalId,
				Appearance: returnValue.Appearance,
				Tag: tag,
			}
			if (tag.client_id) {
				var event = new CustomEvent('clickSingleQues', {
					detail: Object.assign({}, tag, {
						InternalId: returnValue.InternalId,
						Appearance: returnValue.Appearance,
					}),
				})
				document.dispatchEvent(event)
			}
		} else {
			g_click_value = null
		}
	})
}
// 右建显示菜单
function handleContextMenuShow(options) {
	console.log('handleContextMenuShow', options)
	if (options.type == 'Target') { // 只是点击，未选中范围，处理g_click_value
		console.log('g_click_value', g_click_value)
		if (!g_click_value) {
			return
		}
		window.Asc.plugin.executeMethod('AddContextMenuItem', [getContextMenuItems(options.type)])
	} else if (options.type == 'Selection') { // 选中范围，针对所选范围处理
		window.Asc.plugin.executeMethod('AddContextMenuItem', [getContextMenuItems(options.type)])
	} else if (options.type == 'Shape') {
		window.Asc.plugin.executeMethod('AddContextMenuItem', [getContextMenuItems(options.type)])
	}
}

function getContextMenuItems(type) {
	console.log('getContextMenuItems', type)
	if (type == 'Shape') {
		return {
			guid: window.Asc.plugin.guid,
			items: [{
				separator: true,
				id: 'handleWrite_del',
				text: '删除作答区'
			}],
		}
	}
	var splitType = {
		separator: true,
		icons:
			'resources/%theme-type%(light|dark)/%state%(normal)icon%scale%(100|200).%extension%(png)',
		id: 'updateControlType',
		text: '划分类型',
		items: []
	}
	var list = ["question", 'struct', 'setBig', 'write', 'clearBig', 'clear', 'clearAll']
	var names = ['设置为 - 题目', '设置为 - 题组(结构)', '设置为 - 大题', '设置为 - 小问', '清除 - 大题', '清除 - 选中区域', '清除 - 选中区域(含子级)']
	var valueMap = {}
	list.forEach(e => {
		valueMap[e] = 1
		if (e == 'clearBig') {
			valueMap[e] = 0
		}
	})
	var isQuestion = false
	if (type == 'Target') {
		var client_id = g_click_value.Tag.client_id
		var node_list = window.BiyueCustomData.node_list || []
		var nodeData = node_list.find(e => {
			return e.id == client_id
		})
		if (nodeData) {
			if (nodeData.level_type == 'struct') {
				valueMap['struct'] = 0
			} else if (nodeData.level_type == 'question') {
				valueMap['question'] = 0
				valueMap['setBig'] = !(nodeData.is_big)
				valueMap['clearBig'] = nodeData.is_big
				isQuestion = true
			} else if (nodeData.level_type == 'write') {
				valueMap['write'] = 0
			}
		}
	} else {
		valueMap['setBig'] = 0
	}
	var items = []
	list.forEach((e, index) => {
		if (valueMap[e]) {
			items.push({
				icons: 'resources/%theme-type%(img)/%state%(normal)x%scale%(50).%extension%(png)',
				id: `updateControlType_${e}`,
				text: names[index]
			})
		}
	})
	console.log('items', items)
	splitType.items = items
	let settings = {
		guid: window.Asc.plugin.guid,
		items: [splitType],
	}
	if (isQuestion) {
		settings.items.push({
			id: 'write',
			text: '作答区',
			items: [{
				id: 'handleWrite_add',
				text: '添加',
			}, {
				id: 'handleWrite_del',
				text: '删除',
			}] 
		})
		settings.items.push({
			id: 'identify',
			text: '识别框',
			items: [
				{
					id: 'handleIdentifyBox_add',
					text: '添加',
				},
				{
					id: 'handleIdentifyBox_del',
					text: '删除',
				},
			],
		})
	}
	return settings
}

// 划分类型
function updateRangeControlType(typeName) {
	Asc.scope.typename = typeName
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	console.log('updateRangeControlType begin:', typeName)
	return biyueCallCommand(window, function() {
		var typeName = Asc.scope.typename
		var oDocument = Api.GetDocument()
		var oRange = oDocument.GetRangeBySelect()
		var client_node_id = Asc.scope.client_node_id
		var change_list = []
		var result = {
			client_node_id: client_node_id,
			change_list: change_list,
			typeName: typeName
		}
		function getParentBlock(oControl) {
			if (!oControl) {
				return null
			}
			if (oControl.GetClassType() == 'inlineLvlSdt') {
				var parentControl = oControl.GetParentContentControl()
				if (parentControl) {
					if (parentControl.GetClassType() == 'blockLvlSdt') {
						return parentControl
					} else {
						return getParentBlock(parentControl)
					}
				}
			}
			return null
		}
		function getChildControls(oControl) {
			if (!oControl) {
				return null
			}
			var childControls = oControl.GetAllContentControls()
			if (childControls) {
				var children = []
				for (var i = 0; i < childControls.length; ++i) {
					var childControl = childControls[i]
					var tag = JSON.parse(childControl.GetTag() || '{}')
					if (!tag.client_id) {
						continue
					}
					if (childControl.GetClassType() == 'blockLvlSdt') {
						children.push({
							id: tag.client_id,
							regionType: tag.regionType
						})
					} else {
						var parentBlock = getParentBlock(childControl)
						if (parentBlock && parentBlock.Sdt.GetId() == oControl.Sdt.GetId()) {
							children.push({
								id: tag.client_id,
								regionType: tag.regionType
							})
						}
					}
				}
				return children
			}
			return null
		}
		// 删除题目互动
		function clearQuesInteraction(oControl) {
			if (!oControl) {
				return
			}
			console.log('       2', oControl, oControl.GetTag)
			var tag = JSON.parse(oControl.GetTag() || '{}')
			if (tag.regionType == 'write') {
				if (oControl.GetClassType() == 'inlineLvlSdt') {
					var elementCount = oControl.GetElementsCount()
					for (var idx = 0; idx < elementCount; ++idx) {
						var oRun = oControl.GetElement(idx)
						if (oRun && 
							oRun.Run &&
							oRun.Run.Content &&
							oRun.Run.Content[0] &&
							oRun.Run.Content[0].docPr) {
							var title = oRun.Run.Content[0].docPr.title
							if (title) {
								var titleObj = JSON.parse(title)
								if (titleObj.feature && titleObj.feature.sub_type == 'ask_accurate') {
									oRun.Delete()
									break		
								}
							}
						}
					}
				}
			} else if (tag.regionType =='question') {
				if (oControl.GetClassType() == 'blockLvlSdt') {
					var oControlContent = oControl.GetContent()
					var drawings = oControlContent.GetAllDrawingObjects()
					if (drawings) {
						for (var j = 0, jmax = drawings.length; j < jmax; ++j) {
							var oDrawing = drawings[j]
							if (oDrawing.Drawing.docPr) {
								var title = oDrawing.Drawing.docPr.title
								if (title && title.indexOf('feature') >= 0) {
									var titleObj = JSON.parse(title)
									if (titleObj.feature && titleObj.feature.zone_type == 'question') {
										oDrawing.Delete()
									}
								}
							}
						}
					}
				}
			}
		}
		function removeControl(oRemove) {
			if (!oRemove) {
				return
			}
			var tagRemove = JSON.parse(oRemove.GetTag() || '{}')
			clearQuesInteraction(oRemove)
			Api.asc_RemoveContentControlWrapper(oRemove.Sdt.GetId())
			result.change_list.push({
				control_id: oRemove.Sdt.GetId(),
				client_id: tagRemove.client_id
			})
		}
		if (typeName == 'examtitle') {
			if (oRange) {
				oRange.SetBold(true)
			} else {
				return {
					code: 0,
					message: '请先选中一个范围'
				}
			}
			return
		}
		var allControls = oDocument.GetAllContentControls()
		if (!oRange) {
			var currentContentControl = oDocument.Document.GetContentControl()
			if (!currentContentControl) {
				return {
					code: 0,
					message: '请先选中一个范围',
				}
			}
			var oControl = Api.LookupObject(currentContentControl.Id)
			if (!oControl) {
				return
			}
			console.log('oControl.GetTag', oControl.GetTag)
			var tag = JSON.parse(oControl.GetTag() || '{}')
			if (oControl) {
				if (typeName == 'clear' || typeName == 'clearAll') {
					if (typeName == 'clearAll' && oControl.GetClassType() == 'blockLvlSdt') {
						var childControls = oControl.GetAllContentControls()
						for (var i = 0; i < childControls.length; ++i) {
							removeControl(childControls[i])
						}
					}
					// 删除之前创建的精准互动
					removeControl(oControl)
				} else {
					var obj = {}
					if (!tag.client_id) {
						// 之前没有配置client_id，需要分配
						result.client_node_id += 1
						tag.client_id = result.client_node_id
						if (!tag.regionType) { // 之前没有配置regionType
							tag.regionType = text == 'write' ? 'write' : 'question'
						}
						oControl.SetTag(JSON.stringify(tag));
						obj.chidren = getChildControls(oControl)
					}
					var parentBlock = getParentBlock(oControl)
					var parent_id = 0
					if (parentBlock) {
						var parentTag = JSON.parse(parentBlock.GetTag() || '{}')
						parent_id = parentTag.client_id || 0
					}
					obj.client_id = tag.client_id
					obj.parent_id = parent_id
					obj.control_id = oControl.Sdt.GetId()
					obj.regionType = tag.regionType
					var parentId = oControl.Sdt.GetParent().Id
					var templist = []
					if (typeName == 'setBig') { // 设置为大题
						// 需要将后面的级别比他小的控件挪到它的范围内
						var index = allControls.findIndex(e => {
							return e.Sdt.GetId() == oControl.Sdt.GetId()
						})
						for (i = index + 1; i < allControls.length; ++i) {
							var nextControl = allControls[i]
							var nexParent = nextControl.Sdt.GetParent()
							if (nexParent.Id != parentId) {
								continue
							}
							var parentcontrol = getParentBlock(nextControl)
							if (parentcontrol && parentcontrol.Sdt.GetId() == oControl.Sdt.GetId()) {
								continue
							}
							var nextTag = JSON.parse(nextControl.GetTag() || '{}')
							if (nextTag.regionType == 'question') {
								if (nextTag.lvl <= tag.lvl) {
									break
								}
							}
							var posinparent = nextControl.GetPosInParent()
							var oNextParent = editor.LookupObject(nexParent.Id)
							if (oNextParent) {
								templist.push(oNextParent.GetElement(posinparent))
								oNextParent.RemoveElement(posinparent)
							}
						}
						console.log('controllist', templist)
						if (templist.length) {
							var count = oControl.GetContent().GetElementsCount()
							templist.forEach((e, index) => {
								oControl.AddElement(e, count + index)
							})
						}
					} else if (typeName == 'clearBig') { // 清除大题
						// 需要将包含的子控件挪出它的范围内
						var childControls = oControl.GetAllContentControls()
						for (var i = childControls.length - 1; i >= 0; --i) {
							if (childControls[i].GetClassType() != 'blockLvlSdt') {
								continue
							}
							var childParent = childControls[i].GetParentContentControl()
							if (childParent.Sdt.GetId() == oControl.Sdt.GetId()) {
								var posinparent = childControls[i].GetPosInParent()
								templist.push(oControl.GetContent().GetElement(posinparent))
								oControl.GetContent().RemoveElement(posinparent)
							}
						}
						if (templist.length) {
							var oParent = editor.LookupObject(parentId)
							if (oParent) {
								var posinparent = oControl.GetPosInParent() + 1
								templist.forEach(e => {
									oParent.AddElement(posinparent, e)
								})
							}
						}
					}
					obj.text = tag.regionType == 'question' ? oControl.GetRange().GetText() : null
					console.log('=============== obj.text', obj.text, tag.regionType)
					result.change_list.push(obj)
				}
			}
		} else {
			if (!oRange.Paragraphs || oRange.Paragraphs.length === 0) {
				return {
					code: 0,
					message: '选中范围内无段落',
				}
			}
			function checkPosSame(pos1, pos2) {
				if (pos1 && pos2 && pos1.length == pos2.length) {
					for (var nPos = 0; nPos < pos1.length; nPos++) {
						if (pos1[nPos].Class !== pos2[nPos].Class || pos1[nPos].Position !== pos2[nPos].Position) {
							return false;
						}
					}
					return true
				}
				return false
			}
			// 获取两个range的关系 重叠，一个是另一个的子集，相交
			function getRangeRelation(range1, range2) {
				if (checkPosSame(range1.StartPos, range2.StartPos) && checkPosSame(range1.EndPos, range2.EndPos)) {
					return 1
				} else {
					var intersectRange = range1.IntersectWith(range2)
					if (intersectRange) {
						if (checkPosSame(range1.StartPos, intersectRange.StartPos) && checkPosSame(range1.EndPos, intersectRange.EndPos)) {
							return 2
						} else if (checkPosSame(range2.StartPos, intersectRange.StartPos) && checkPosSame(range2.EndPos, intersectRange.EndPos)) {
							return 3	
						} else {
							return 4
						}
					} else {
						return 5
					}
				}
			}
			
			var controlsInRange = []
			var completeOverlapControl = null
			var parentControls = []
			var containControls = [] // 包含的control
			var intersectControls = [] // 交叉的control
			for (var i = 0, imax = allControls.length; i < imax; ++i) {
				var e = allControls[i]
				var isUse = false
				if (e.GetClassType() == 'blockLvlSdt') {
					if (e.Sdt.Content.Selection && e.Sdt.Content.Selection.Use) {
						isUse = true
					}
				}  else if (e.GetClassType() == 'inlineLvlSdt') {
					if (e.Sdt.Selection && e.Sdt.Selection.Use) {
						isUse = true
					}
				}
				if (!isUse) {
					continue
				}
				var relation = getRangeRelation(oRange, e.GetRange())
				if (relation == 1) {
					completeOverlapControl = e
					if (typeName != 'clear' && typeName != 'clearAll') {
						if (typeName != 'write') {
							break
						} else {
							continue
						}
					}
				}
				if (relation == 2) {
					parentControls.push(e)
				} else if (relation == 3) {
					containControls.push(e)
				} else if (relation == 4) {
					intersectControls.push(e)
				}
				controlsInRange.push(e)
			}

			if (typeName == 'clear') {
				if (completeOverlapControl) {
					removeControl(completeOverlapControl)
				} else if (parentControls.length) {
					removeControl(parentControls[parentControls.length - 1])
				}
			} else if (typeName == 'clearAll') {
				var controls = containControls.concat(intersectControls)
				if (completeOverlapControl) {
					controls.push(completeOverlapControl)
				}
				controls.forEach(e => {
					removeControl(e)
				})
			} else {
				// 若存在完全重叠的区域，只修改tag
				if (completeOverlapControl) {
					var tag = JSON.parse(completeOverlapControl.GetTag() || {})
					var regionType = typeName == 'write' ? 'write' : 'question'
					if (tag.regionType != regionType) {
						tag.regionType = regionType
						completeOverlapControl.SetTag(JSON.stringify(tag));
					}
					result.change_list.push({
						client_id: tag.client_id
					})
				} else {
					// 不存在完全重叠的区域，新增一个control，对其下的原本的control并不进行拆分修改 todo..
					var isInCell = false
					var level = 0
					if (oRange.Paragraphs.length == 1) {
						var oParagraph = oRange.Paragraphs[0]
						var parent1 = oParagraph.Paragraph.GetParent()
						var oParent1 = Api.LookupObject(parent1.Id)
						if (oParent1 && oParent1.GetClassType && oParent1.GetClassType() == 'documentContent') {
							var parent2 = parent1.GetParent()
							if (parent2) {
								var oParent2 = Api.LookupObject(parent2.Id)
								if (oParent2 && oParent2.GetClassType && oParent2.GetClassType() == 'tableCell') {
									isInCell = true
								}
							}
						}
					}
					if (oRange.Paragraphs.length) {
						var NumPr = oRange.Paragraphs[0].GetNumbering()
						if (NumPr) {
							level = NumPr.Lvl
						}
					}
					var type = typeName == 'write' ? 2 : 1
					result.client_node_id += 1
					var tag = {
						client_id: result.client_node_id,
						regionType: typeName == 'write' ? 'write' : 'question',
						lvl: level
					}
					var oResult = Api.asc_AddContentControl(type, {
						Tag: JSON.stringify(tag)
					})
					console.log('asc_AddContentControl', oResult)
					if(oResult) {
						result.change_list.push({
							client_id: tag.client_id,
							control_id: oResult.InternalId,
							text: oRange.GetText()
						})
					}
					// 若是在单元格里添加control后，会多出一行需要删除
					if (isInCell) {
						var oCell = editor.LookupObject(oResult.InternalId).GetParentTableCell()
						if (oCell) {
							if (oCell.GetContent().GetElementsCount() == 2) {
								var oElement2 = oCell.GetContent().GetElement(1)
								if (oElement2.GetClassType() == 'paragraph' && oElement2.Paragraph.Bounds.Bottom == 0 && oElement2.Paragraph.Bounds.Top == 0) {
									oCell.GetContent().RemoveElement(1)
								}
							}
						}
					}
				}
			}
		}
		return result
	}, false, true).then(res => {
		console.log('handleChangeType result', res)
		handleChangeType(res)
	})
}

function handleChangeType(res) {
	if (!res) {
		return
	}
	if (!res.change_list || res.change_list.length == 0) {
		return
	}
	if (res.client_node_id) {
		window.BiyueCustomData.client_node_id = res.client_node_id
	}
	var node_list = window.BiyueCustomData.node_list || []
	var question_map = window.BiyueCustomData.question_map || {}
	var level_type = res.typeName
	var targetLevel = level_type
	if (res.typeName == 'setBig' || res.typeName == 'clearBig') {
		targetLevel = 'question'
	}
	res.change_list.forEach(item => {
		var nodeIndex = node_list.findIndex(e => {
			return e.id == item.client_id
		})
		var nodeData = nodeIndex >= 0 ? node_list[nodeIndex] : null
		// 之前是小问，需要移除
		if (nodeData && nodeData.level_type == 'write' && level_type != 'write') {
			if (nodeData.parent_id) {
				var parentNode = node_list.find(e => {
					return e.id == nodeData.parent_id
				})
				if (parentNode) {
					var writeIndex = parentNode.write_list.findIndex(e => {
						return e.id == nodeData.id
					})
					if (writeIndex >= 0) {
						parentNode.write_list.splice(writeIndex, 1)
						if (question_map[nodeData.parent_id]) {
							question_map[nodeData.parent_id].ask_list.splice(writeIndex, 1)
						}
					}
				}
			}
		}
		if (nodeData) {
			if (level_type == 'clear' || level_type == 'clearAll') {
				node_list.splice(nodeIndex, 1)
			} else {
				nodeData.is_big = level_type == 'setBig'
				nodeData.level_type = targetLevel
			}
			if (targetLevel == 'question') {
				if (!question_map[item.client_id]) {
					question_map[item.client_id] = {
						text: item.text,
						level_type: targetLevel,
						ques_default_name: GetDefaultName(targetLevel, item.text),
						ask_list: nodeData.write_list.map(e => {
							return {
								id: e.id
							}
						})
					}
				} else if (level_type == 'setBig' || level_type == 'clearBig') {
					question_map[item.client_id].text = item.text
					question_map[item.client_id].ques_default_name = GetDefaultName(targetLevel, item.text)
					question_map[item.client_id].level_type = targetLevel
				}
			} else {
				if (question_map[item.client_id]) {
					delete question_map[item.client_id]
				}
			}
		} else if (level_type != 'clear' && level_type != 'clearAll') {
			// 之前没有，需要增加
			var ask_list
			if (item.children && item.chidren.length) {
				ask_list = item.chidren.filter(child => {
					return child.regionType == 'write'
				})
			}
			node_list.push({
				id: item.client_id,
				control_id: item.control_id,
				regionType: item.regionType,
				level_type: targetLevel,
				write_list: ask_list
			})
			if (targetLevel == 'question' || targetLevel == 'struct') {
				question_map[item.client_id] = {
					text: item.text,
					level_type: targetLevel,
					ques_default_name: GetDefaultName(targetLevel, item.text)
				}
				if (targetLevel == 'question') {
					question_map[item.client_id].ask_list = ask_list
				}
			}
		}
	})
	
	window.BiyueCustomData.node_list = node_list
	window.BiyueCustomData.question_map = question_map
	console.log('handleChangeType end', node_list, g_click_value)
	document.dispatchEvent(
		new CustomEvent('updateQuesData', {
			detail: {
				client_id: g_click_value ? g_click_value.Tag.client_id : 0
			}
		})
	)
}
// 批量修改题型
function batchChangeQuesType(type) {
	Asc.scope.ques_type = type
	return biyueCallCommand(window, function () {
		var oDocument = Api.GetDocument()
		var control_list = oDocument.GetAllContentControls()
		var ques_id_list = []
		control_list.forEach((e) => {
			if (
				e.Sdt &&
				e.Sdt.Content &&
				e.Sdt.Content.Selection &&
				e.Sdt.Content.Selection.Use
			) {
				var tag = JSON.parse(e.GetTag() || '{}')
				if (tag && tag.client_id) {
					ques_id_list.push({
						id: tag.client_id,
						control_id: e.Sdt.GetId()
					})
				}
			}
		})
		return {
			code: 1,
			list: ques_id_list,
			type: Asc.scope.ques_type
		}
	}, false, false).then((res) => {
		if (!res || !res.code || !res.list) {
			return
		}
		res.list.forEach(e => {
			if (window.BiyueCustomData.question_map && window.BiyueCustomData.question_map[e.id]) {
				window.BiyueCustomData.question_map[e.id].question_type = res.type
			}
		})
		// 需要同步更新单题详情
		if (window.tab_select == 'tabQues') {
			document.dispatchEvent(
				new CustomEvent('updateQuesData', {
					detail: {
						list: res.list,
						field: 'question_type',
						value: res.type,
					},
				})
			)
		}
	})
}

// 切题完成
function splitEnd() {
	console.log('splitEnd')
	return new Promise((resolve, reject) => {
		showLevelSetDialog()
		resolve()
	})
}

function showLevelSetDialog() {
	window.biyue.showDialog(levelSetWindow, '自动序号识别设置', 'levelSet.html', 592, 400)
}
// 由于从BiyueCustomData中获取的中文取出后会是乱码，需要在初始化时，再根据controls刷新一次数据
function initControls() {
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		var nodeList = []
		function getParentBlock(oControl) {
			if (!oControl) {
				return null
			}
			if (oControl.GetClassType() == 'inlineLvlSdt') {
				var parentControl = oControl.GetParentContentControl()
				if (parentControl) {
					if (parentControl.GetClassType() == 'blockLvlSdt') {
						return parentControl
					} else {
						return getParentBlock(parentControl)
					}
				}
			}
			return null
		}
		controls.forEach((oControl) => {
			var tagInfo = JSON.parse(oControl.GetTag() || '{}')
			if (tagInfo.regionType) {
				var parentid = 0
				var parentControl = getParentBlock(oControl)
				if (parentControl) {
					var parentTagInfo = JSON.parse(parentControl.GetTag() || '{}')
					if (parentTagInfo.regionType == 'question') {
						parentid = parentTagInfo.client_id
					}
				}
				nodeList.push({
					id: tagInfo.client_id,
					text: oControl.GetRange().GetText(),
					parentId: parentid
				})
			}
		})
		return nodeList
	}, false, false).then(nodeList => {
		return new Promise((resolve, reject) => {
			// todo.. 这里暂不考虑上次的数据未保存或保存失败的情况，只假设此时的control数据和nodelist里的是一致的，只是乱码而已，其他的后续再处理
			if (nodeList && nodeList.length > 0) {
				var question_map = window.BiyueCustomData.question_map || {}
				nodeList.forEach(node => {
					if (question_map[node.id]) {
						question_map[node.id].text = node.text
						question_map[node.id].ques_default_name = GetDefaultName(question_map[node.id].level_type, node.text)
					}
				})
			}
			resolve()
		})
	})
}

function confirmLevelSet(levels) {
	Asc.scope.levels = levels
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	return biyueCallCommand(window, function() {
		var levelmap = Asc.scope.levels
		var client_node_id = Asc.scope.client_node_id
		var nodeList = []
		var questionMap = {}
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		function getParentBlock(oControl) {
			if (!oControl) {
				return null
			}
			if (oControl.GetClassType() == 'inlineLvlSdt') {
				var parentControl = oControl.GetParentContentControl()
				if (parentControl) {
					if (parentControl.GetClassType() == 'blockLvlSdt') {
						return parentControl
					} else {
						return getParentBlock(parentControl)
					}
				}
			}
			return null
		}
		controls.forEach((oControl) => {
			var tagInfo = JSON.parse(oControl.GetTag() || '{}')
			 if (tagInfo.regionType == 'question' || tagInfo.regionType == 'write') {
				client_node_id += 1
				var id = client_node_id
				tagInfo.client_id = id
				oControl.SetTag(JSON.stringify(tagInfo));
				var parentControl = getParentBlock(oControl)
				var nodeData = {
					id: id,
					control_id: oControl.Sdt.GetId(),
					regionType: tagInfo.regionType
				}
				if (tagInfo.regionType == 'write') {
					nodeData.level_type = 'write'
					if (parentControl) {
						var parent_tagInfo = JSON.parse(parentControl.GetTag() || '{}')
						var parentNode = nodeList.find((node) => {
							return node.id == parent_tagInfo.client_id
						})
						if (parentNode) {
							parentNode.write_list.push({
								id: id,
								control_id: oControl.Sdt.GetId(),
							})
							questionMap[parent_tagInfo.client_id].ask_list.push({
								id: id
							})
							nodeData.parent_id = parent_tagInfo.client_id
						}
					}
				} else {
					var text = oControl.GetRange().GetText()
					var level_type = levelmap[tagInfo.lvl]
					nodeData.level_type = level_type
					var detail = {
						text: text,
						ask_list: [],
						level_type: level_type
					}
					if (tagInfo.regionType == 'question') {
						nodeData.write_list = []
						detail.ask_list = []
					}
					questionMap[id] = detail
				}
				nodeList.push(nodeData)
			 }
		})
		return {
			client_node_id: client_node_id,
			nodeList: nodeList,
			questionMap: questionMap
		}
	}, false, false).then(res => {
		return new Promise((resolve, reject) => {
			if (res) {
				window.BiyueCustomData.client_node_id = res.client_node_id
				window.BiyueCustomData.node_list = res.nodeList
				Object.keys(res.questionMap).forEach(key => {
					res.questionMap[key].ques_default_name = GetDefaultName(res.questionMap[key].level_type, res.questionMap[key].text)
				})
				window.BiyueCustomData.question_map = res.questionMap
				console.log('confirmLevelSet', window.BiyueCustomData)
			}
			resolve()
		})
	}).then(() => {
		reqGetQuestionType()
	})
}

function GetDefaultName(level_type, text) {
	if (level_type == 'struct') {
		const pattern = /^[一二三四五六七八九十0-9]+.*?(?=[：:])/
		const result = pattern.exec(text)
		return result ? result[0] : null
	} else if (level_type == 'question')  {
		const regex = /^([^.．、]*)/
		const match = text.match(regex)
		return match ? match[1] : ''
	}
	return ''
}

// 获取题型
function reqGetQuestionType() {
	Asc.scope.node_list = window.BiyueCustomData.node_list
	return biyueCallCommand(window, function() {
		var node_list = Asc.scope.node_list || []
		var target_list = []
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls() || []
		controls.forEach((oControl) => {
			var tagInfo = JSON.parse(oControl.GetTag() || '{}')
			if (tagInfo.regionType == 'question' && tagInfo.client_id) {
				var nodeData = node_list.find(item => {
					return item.id == tagInfo.client_id
				})
				if (nodeData && nodeData.level_type == 'question') {
					var oRange = oControl.GetRange()
					oRange.Select()
					let text_data = {
						data:     "",
						// 返回的数据中class属性里面有binary格式的dom信息，需要删除掉
						pushData: function (format, value) {
							this.data = value ? value.replace(/class="[a-zA-Z0-9-:;+"\/=]*/g, "") : "";
						}
					};
			
					Api.asc_CheckCopy(text_data, 2);
					target_list.push({
						id: nodeData.id + '',
						content_type: nodeData.level_type,
						content_html: text_data.data
					})
				}
			}
		})
		return target_list
	}, false, false).then(control_list => {
		console.log('[reqGetQuestionType] control_list', control_list)
		if (!window.BiyueCustomData.paper_uuid || !control_list || control_list.length == 0) {
			return
		}
		getQuesType(window.BiyueCustomData.paper_uuid, control_list).then(res => {
			console.log('getQuesType success ', res)
			var content_list = res.data.content_list
			if (content_list && content_list.length) {
				content_list.forEach(e => {
					window.BiyueCustomData.question_map[e.id].question_type = e.question_type
					window.BiyueCustomData.question_map[e.id].question_type_name = e.question_type_name
				})
			}
		}).catch(res => {
			console.log('getQuesType fail ', res)
		})
	})
}

function handleWrite(cmdType) {
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	Asc.scope.write_cmd = cmdType
	return biyueCallCommand(window, function() {
		var client_node_id = Asc.scope.client_node_id
		var write_cmd = Asc.scope.write_cmd
		var oDocument = Api.GetDocument()
		var curControl = oDocument.Document.GetContentControl()
		if (!curControl) {
			console.log('curControl is null')
			return
		}
		var oControl = Api.LookupObject(curControl.Id)
		if (oControl) {
			var tag = JSON.parse(oControl.GetTag() || '{}')
			if (write_cmd == 'add') {
				client_node_id += 1
				var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 0, 0))
				oFill.UniFill.transparent = 255 * 0.2 // 透明度
				var oStroke = Api.CreateStroke(3600, Api.CreateNoFill())
				var oDrawing = Api.CreateShape(
					'rect',
					20 * 36e3,
					10 * 36e3,
					oFill,
					oStroke
				)
				var titleobj = {
					feature: {
						zone_type: 'question',
						sub_type: 'write',
						control_id: oControl.Sdt.GetId(),
						parent_id: tag.client_id,
						client_id: client_node_id
					}
				}
				oDrawing.Drawing.Set_Props({
					title: JSON.stringify(titleobj),
				})
				oDrawing.SetWrappingStyle('inFront')
				oDrawing.SetDistances(0, 0, 2 * 36e3, 0);
				var paragraphs = oControl.GetAllParagraphs()
				if (paragraphs && paragraphs.length > 0) {
					var pParagraph = paragraphs[0]
					var oRun = Api.CreateRun()
					oRun.AddDrawing(oDrawing)
					pParagraph.AddElement(
						oRun,
						1
					)
				}
				return {
					ques_id: tag.client_id,
					write_id: client_node_id,
					client_node_id: client_node_id,
					drawing_id: oDrawing.Drawing.Id,
					cmdType: write_cmd
				}
			} else {
				var selectDrawings = oDocument.GetSelectedDrawings()
				if (selectDrawings) {
					var drawings = oDocument.GetAllDrawingObjects()
					for (var i = 0; i < selectDrawings.length; ++i) {
						if (selectDrawings[i].GetClassType() == 'shape') {
							var title = selectDrawings[i].Drawing.docPr.title || '{}'
							var titleObj = JSON.parse(title)
							if (titleObj.feature && titleObj.feature.zone_type == 'question' && titleObj.feature.sub_type == 'write') {
								var drawingId = selectDrawings[i].Drawing.Id
								var oDrawing = drawings.find(e => {
									return e.Drawing.Id == drawingId
								})
								if (oDrawing) {
									oDrawing.Delete()
									return {
										ques_id: titleObj.feature.parent_id,
										write_id: titleObj.feature.client_id,
										client_node_id: client_node_id,
										drawing_id: drawingId,
										cmdType: write_cmd
									}
								}
							}
						}
					}
				}
			}
		} else {
			console.log('======== oControl is null ========')
		}
		return null
	}, false, true).then(res => {
		handleWriteResult(res)
	})
}

// 添加或删除标识
function handleIdentifyBox(cmdType) {
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	Asc.scope.write_cmd = cmdType
	biyueCallCommand(window, function () {
			var oDocument = Api.GetDocument()
			var curPosInfo = oDocument.Document.GetContentPosition()
			var cmdType = Asc.scope.write_cmd
			var client_node_id = Asc.scope.client_node_id
			console.log('curPosInfo', curPosInfo)
			var res = {
				cmdType: cmdType
			}
			if (curPosInfo) {
				var runIdx = -1
				var paragraphIdx = -1
				for (var i = curPosInfo.length - 1; i >= 0; --i) {
					if (curPosInfo[i].Class.GetType) {
						var t = curPosInfo[i].Class.GetType()
						if (t == 1) {
							paragraphIdx = i
							break
						} else if (t == 39) {
							runIdx = i
						}
					}
				}
				function createDrawing(parent_control_id, parent_id) {
					var oFill = Api.CreateNoFill()
					var oStroke = Api.CreateStroke(
						3600,
						Api.CreateSolidFill(Api.CreateRGBColor(125, 125, 125))
					)
					var oDrawing = Api.CreateShape(
						'rect',
						8 * 36e3,
						5 * 36e3,
						oFill,
						oStroke
					)
					var drawDocument = oDrawing.GetContent()
					var oParagraph = Api.CreateParagraph()
					oParagraph.AddText('×')
					oParagraph.SetColor(125, 125, 125, false)
					oParagraph.SetFontSize(24)
					oParagraph.SetFontFamily('黑体')
					oParagraph.SetJc('center')
					drawDocument.AddElement(0, oParagraph)
					oDrawing.SetPaddings(0, 0, 0, 0.5 * 36e3)
					var titleobj = {
						feature: {
							zone_type: 'question',
							sub_type: 'identify',
							control_id: parent_control_id,
							parent_id: parent_id,
							client_id: client_node_id

						}
					}
					oDrawing.Drawing.Set_Props({
						title: JSON.stringify(titleobj),
					})
					oDrawing.SetWrappingStyle('inline')
					return oDrawing
				}
				if (paragraphIdx >= 0) {
					var pParagraph = new Api.private_CreateApiParagraph(
						AscCommon.g_oTableId.Get_ById(curPosInfo[paragraphIdx].Class.Id)
					)
					var paraentControl = pParagraph.GetParentContentControl()
					if (paraentControl) {
						var tag = JSON.parse(paraentControl.GetTag())
						if (
							tag.regionType == 'question'
						) {
							res.control_id = paraentControl.Sdt.GetId()
							res.ques_id = tag.client_id
							if (cmdType == 'add') {
								client_node_id += 1
								var oDrawing = createDrawing(paraentControl.Sdt.GetId(), tag.client_id)
								if (runIdx >= 0) {
									res.run_id = curPosInfo[runIdx].Class.Id
									curPosInfo[runIdx].Class.Add_ToContent(
										curPosInfo[runIdx].Position,
										oDrawing.Drawing
									)
									res.add_pos = 'run'
								} else {
									var oRun = Api.CreateRun()
									oRun.AddDrawing(oDrawing)
									pParagraph.AddElement(
										oRun,
										curPosInfo[paragraphIdx].Position + 1
									)
								}
							} else {
								res.remove_ids = []
								var drawings = paraentControl.GetAllDrawingObjects()
								for (var sidx = 0; sidx < drawings.length; ++sidx) {
									if (
										drawings[sidx].Drawing &&
										drawings[sidx].Drawing.docPr &&
										drawings[sidx].Drawing.docPr.title &&
										drawings[sidx].Drawing.docPr.title != ''
									) {
										try {
											var dtitle = JSON.parse(
												drawings[sidx].Drawing.docPr.title
											)
											if (dtitle.feature && dtitle.feature.sub_type == 'identify') {
												res.remove_ids.push(dtitle.feature.client_id)
												drawings[sidx].Delete()
											}
										} catch (error) {}
									}
								}
							}
						}
					}
				}
			}
			res.client_node_id = client_node_id
			return res
		},
		false,
		true
	).then((res) => {
		console.log('handleIdentifyBox result', res)
		handleWriteResult(res)
	})
}

function handleWriteResult(res) {
	console.log('handleWriteResult', res)
	if (!res) {
		return
	}
	if (res.client_node_id) {
		window.BiyueCustomData.client_node_id = res.client_node_id
	}
	if (res.ques_id) {
		var node_list = window.BiyueCustomData.node_list || []
		var question_map = window.BiyueCustomData.question_map || {}
		var nodeData = node_list.find(e => {
			return e.id == res.ques_id
		})
		if (res.cmdType == 'del' && !res.remove_ids && res.write_id) {
			res.remove_ids = [res.write_id]
		}
		if (nodeData) {
			if (res.cmdType == 'add') {
				if (!nodeData.write_list) {
					nodeData.write_list = []
				}
				nodeData.write_list.push({
					id: res.write_id,
					drawing_id: res.drawing_id
				})
			} else if (nodeData.write_list) {
				res.remove_ids.forEach(removeid => {
					var writeIndex = nodeData.write_list.findIndex(e => {
						return e.id == removeid
					})
					if (writeIndex >= 0) {
						nodeData.write_list.splice(writeIndex, 1)
					}
				})
			}
		}
		if (question_map[res.ques_id]) {
			if (res.cmdType == 'add') {
				if (!question_map[res.ques_id].ask_list) {
					question_map[res.ques_id].ask_list = []
				}
				question_map[res.ques_id].ask_list.push({
					id: res.write_id,
					score: 0
				})
			} else if (question_map[res.ques_id].ask_list) {
				res.remove_ids.forEach(removeid => {
					var askIndex = question_map[res.ques_id].ask_list.findIndex(e => {
						return e.id == removeid
					})
					if (askIndex >= 0) {
						question_map[res.ques_id].ask_list.splice(askIndex, 1)
					}
				})
			}
		}
		document.dispatchEvent(
			new CustomEvent('updateQuesData', {
				detail: {
					client_id: res.ques_id
				}
			})
		)
	}
}
function setBtnLoading(elementId, isLoading) {
	var element = $(`#${elementId}`)
	if (!element) {
		return
	}
	if (isLoading) {
		element.append('<span class="loading-spinner"></span>')
		element.addClass('btn-unable')
	} else {
		element.removeClass('btn-unable')
		var children = element.find('.loading-spinner')
		if (children) {
			children.remove()
		}
 	}
	
}

function isLoading(elementId) {
	var element = $(`#${elementId}`)
	if (!element) {
		return
	}
	var loading = element.find('.loading-spinner')
	return loading && loading.length
}
// 全量更新
function reqUploadTree() {
	if (isLoading('uploadTree')) {
		return
	}
	setBtnLoading('uploadTree', true)
	Asc.scope.node_list = window.BiyueCustomData.node_list
	Asc.scope.question_map = window.BiyueCustomData.question_map
	upload_control_list = []
	console.log('[reqUploadTree start]', Date.now())
	biyueCallCommand(window, function() {
		var target_list = []
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		var question_map = Asc.scope.question_map
		console.log('question_map', question_map)
		for (var i = 0, imax = controls.length; i < imax; ++i) {
			var oControl = controls[i]
			var tag = JSON.parse(oControl.GetTag() || '{}')
			if (tag.regionType != 'question' || !tag.client_id) {
				continue
			}
			var quesData = question_map[tag.client_id]
			if (!quesData) {
				continue
			}
			let text_data = {
				data:     "",
				// 返回的数据中class属性里面有binary格式的dom信息，需要删除掉
				pushData: function (format, value) {
					this.data = value ? value.replace(/class="[a-zA-Z0-9-:;+"\/=]*/g, "") : "";
				}
			};
			var oParentControl = oControl.GetParentContentControl()
			var parent_id = 0
			if (oParentControl) {
				var parentTag = JSON.parse(oParentControl.GetTag() || '{}')
				parent_id = parentTag.client_id
			} else {
				// 根据level, 查找在它前面的比它lvl小的struct
				for (var j = target_list.length - 1; j >= 0; --j) {
					var preNode = target_list[j]
					if (preNode.lvl < tag.lvl && preNode.content_type == 'struct') {
						parent_id = preNode.id
						break
					}
				}
			}
			Api.asc_CheckCopy(text_data, 2);
			var content_html = text_data.data
			target_list.push({
				id: tag.client_id,
				parent_id: parent_id,
				uuid: question_map[tag.client_id].uuid || '',
				regionType: question_map[tag.client_id].level_type,
				content_type: question_map[tag.client_id].level_type,
				content_xml: '',
				content_html: content_html,
				content_text: oControl.GetRange().GetText(),
				question_type: question_map[tag.client_id].question_type,
				question_name: question_map[tag.client_id].ques_name || question_map[tag.client_id].ques_default_name,
				control_id: oControl.Sdt.GetId(),
				lvl: tag.lvl
			})
		}
		console.log('target_list', target_list)
		return target_list
	}, false, false).then( control_list => {
		if (control_list) {
			upload_control_list = control_list
			if (control_list && control_list.length) {
				getXml(control_list[0].control_id)
			}
		}
	})
}

function getXml(controlId) {
	window.Asc.plugin.executeMethod("SelectContentControl", [controlId])
	window.Asc.plugin.executeMethod("GetSelectionToDownload", ["docx"], function (data) {
        // 假设这是你的 ZIP 文件的 URL  
        const zipFileUrl = data;        
        fetch(zipFileUrl).then(response => {  
            if (!response.ok) {  
            throw new Error('Failed to fetch zip file');  
            }  
            return response.arrayBuffer(); // 获取 ArrayBuffer 而不是 Blob，因为 JSZip 需要它  
        })  
        .then(arrayBuffer => {  
            return JSZip.loadAsync(arrayBuffer); // 使用 JSZip 加载 ArrayBuffer  
        })  
        .then(zip => {  
            // 现在你可以操作 zip 对象了  
            zip.forEach(function(relativePath, file) {  
                if (relativePath.indexOf('word/document.xml') === -1) {
                    return;
                }
                // 这里可以遍历 ZIP 文件中的所有文件  
                file.async("text").then(function(content) {  
                    // 假设文件是文本文件，打印文件内容和相对路径
					handleXml(controlId, content)
                });  
            });  
        })  
        .catch(error => {  
            console.error('Error:', error);
			handleXmlError()
        });
        
    });
}

function handleXml(controlId, content) {
	if (!upload_control_list) {
		return
	}
	var index = upload_control_list.findIndex(e => {
		return e.control_id == controlId
	})
	if (index == -1) {
		return
	}
	upload_control_list[index].content_xml = content
	if (index + 1 < upload_control_list.length) {
		getXml(upload_control_list[index+1].control_id)
		return
	}
	generateTreeForUpload(upload_control_list)
}

function handleXmlError() {
	generateTreeForUpload(upload_control_list)
}

function generateTreeForUpload(control_list) {
	if (!control_list) {
		setBtnLoading('uploadTree', false)
		return null
	}
	var tree = []
	control_list.forEach((e) => {
		if (e.parent_id == 0) {
			tree.push(e)
		} else {
			var parent = getDataById(tree, e.parent_id)
			if (parent) {
				if (!parent.children) {
					parent.children = []
				}
				parent.children.push(e)
			}
		}
	})
	var uploadTree = {
		id: "",
        uuid: window.BiyueCustomData.paper_uuid,
        question_type:0,
        question_name:"",
        content_type:"paper",
        content_text:"",
        content_xml:"",
        content_html:"",
		children: tree
	}
	console.log('               uploadTree', uploadTree)
	var version = getTimeString()
	// reqComplete(uploadTree, version).then(res => {
	// 	console.log('reqComplete', res)
	// 	console.log('[reqUploadTree end]', Date.now())
	// 	if (res.data.questions) {
	// 		res.data.questions.forEach(e => {
	// 			window.BiyueCustomData.question_map[e.id].uuid = e.uuid
	// 		})
	// 	}
	// 	setBtnLoading('uploadTree', false)
	// 	alert('全量更新成功')
	// }).catch(res => {
	// 	console.log('reqComplete fail', res)
	// 	console.log('[reqUploadTree end]', Date.now())
	// 	setBtnLoading('uploadTree', false)
	// 	alert('全量更新失败')
	// })
	setBtnLoading('uploadTree', false)
	console.log(tree)
}

function getDataById(list, id) {
	if (!list) {
		return null
	}
	for (var i = 0, imax = list.length; i < imax; ++i) {
		if (list[i].id == id) {
			return list[i]
		}
		if (list[i].children) {
			var v = getDataById(list[i].children, id)
			if (v) {
				return v
			}
		}
	}
	return null
}

function getTimeString() {
	var date = new Date()
	var year = date.getFullYear()
	var month = date.getMonth() + 1
	var day = date.getDate()
	var hour = date.getHours()
	var minute = date.getMinutes()
	var second = date.getSeconds()
	var micsecond = date.getMilliseconds()
	return year + '-' + month + '-' + day + ' ' + hour + ':' + minute + ':' + second + ' ' + micsecond
}

export {
	handleDocClick,
	handleContextMenuShow,
	initExamTree,
	updateRangeControlType,
	reqGetQuestionType,
	reqUploadTree,
	batchChangeQuesType,
	splitEnd,
	showLevelSetDialog,
	confirmLevelSet,
	initControls,
	handleWrite,
	handleIdentifyBox
}