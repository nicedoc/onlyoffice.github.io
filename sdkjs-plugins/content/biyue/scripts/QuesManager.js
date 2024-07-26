
import { biyueCallCommand, dispatchCommandResult } from "./command.js";
import { getQuesType, reqComplete } from '../scripts/api/paper.js'
import { setInteraction } from "./featureManager.js";
import { initExtroInfo } from "./panelFeature.js";
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
	var canBatch = false // 是否可批量操作
	if (type == 'Target') {
		var client_id = g_click_value.Tag.client_id
		var node_list = window.BiyueCustomData.node_list || []
		var nodeData = node_list.find(e => {
			return e.id == client_id
		})
		if (nodeData) {
			if (nodeData.level_type == 'struct') {
				valueMap['struct'] = 0
				canBatch = true
			} else if (nodeData.level_type == 'question') {
				valueMap['question'] = 0
				valueMap['setBig'] = !(nodeData.is_big)
				valueMap['clearBig'] = nodeData.is_big
				isQuestion = true
				canBatch = true
			} else if (nodeData.level_type == 'write') {
				valueMap['write'] = 0
			}
		}
	} else {
		valueMap['setBig'] = 0
		canBatch = true
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
	splitType.items = items
	let settings = {
		guid: window.Asc.plugin.guid,
		items: [splitType],
	}
	if (canBatch) {
		var questypes = window.BiyueCustomData.paper_options.question_type
		var itemsQuesType = questypes.map((e) => {
			return {
				id: `batchChangeQuesType_${e.value}`,
				text: e.label,
			}
		})
		var itemsProportion = []
		for (var i = 1; i <= 8; ++i) {
			itemsProportion.push({
				id: `batchChangeProportion_${i}`,
				text: i == 1 ? '默认' : `1/${i}`,
			})
		}
		settings.items.push({
			id: 'batchCmd',
			text: '批量操作',
			items: [
				{
					id: 'batchChangeQuesType',
					text: '修改题型',
					items: itemsQuesType,
				},
				{
					id: 'batchChangeScore',
					text: '修改分数',
				},
				{
					id: 'batchChangeProportion',
					text: '修改占比',
					items: itemsProportion,
				},
				{
					id: 'batchChangeInteraction',
					text: '修改互动模式',
					items: [
						{
							id: 'batchChangeInteraction_none',
							text: '无互动',
						},
						{
							id: 'batchChangeInteraction_simple',
							text: '简单互动',
						},
						{
							id: 'batchChangeInteraction_accurate',
							text: '精准互动',
						},
					],
				},
			],
		})
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

function getNodeList() {
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var node_list = []
		var controls = oDocument.GetAllContentControls() || []
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
		function getParentId(oControl) {
			var parentBlock = getParentBlock(oControl)
			var parent_id = 0
			if (parentBlock) {
				var parentTag = JSON.parse(parentBlock.GetTag() || '{}')
				parent_id = parentTag.client_id || 0
			}
			return parent_id
		}
		for (var i = 0, imax = controls.length; i < imax; ++i) {
			var oControl = controls[i]
			var tagInfo = JSON.parse(oControl.GetTag() || '{}')
			var parent_id = getParentId(oControl)
			if (tagInfo.regionType == 'question') {
				var write_list = []
				var controlContent = oControl.GetContent()
				var elementCount = controlContent.GetElementsCount()
				for (var j = 0; j < elementCount; ++j) {
					var oElement = controlContent.GetElement(j)
					if (oElement.GetClassType() == 'paragraph') {
						var childCount = oElement.GetElementsCount()
						for (var k = 0; k < childCount; ++k) {
							var oChild = oElement.GetElement(k)
							var childType = oChild.GetClassType()
							if (childType == 'run') {
								var childCount1 = oChild.Run.GetElementsCount()
								for (var k2 = 0; k2 < childCount1; ++k2) {
									var oChild2 = oChild.Run.GetElement(k2)
									if (oChild2.GetType && oChild2.GetType() == 22) {
										if (oChild2.docPr) {
											var title = oChild2.docPr.title
											if (title) {
												var titleObj = JSON.parse(title)
												if (titleObj.feature && (titleObj.feature.sub_type == 'write' || titleObj.feature.sub_type == 'identify')) {
													write_list.push({
														id: titleObj.client_id,
														sub_type: titleObj.feature.sub_type,
														drawing_id: oChild2.Id,
														shape_id: oChild2.GraphicObj.Id
													})
												}
											}
										}
									}
								}
							} else if (childType == 'inlineLvlSdt') {
								var childTag = JSON.parse(oChild.GetTag() || '{}')
								if (childTag.regionType == 'write') {
									write_list.push({
										id: childTag.client_id,
										sub_type: 'control',
										control_id: oChild.Sdt.GetId()
									})	
								}
							}
						}
					} else if (oElement.GetClassType() == 'blockLvlSdt') {
						var childTag = JSON.parse(oElement.GetTag() || '{}')
						if (childTag.regionType == 'write') {
							write_list.push({
								id: childTag.client_id,
								sub_type: 'control',
								control_id: oElement.Sdt.GetId()
							})
						}
					} else if (oElement.GetClassType() == 'table') {
						var rows = oElement.GetRowsCount()
						for (var i1 = 0; i1 < rows; ++i1) {
							var oRow = oElement.GetRow(i1)
							var cells = oRow.GetCellsCount()
							for (var i2 = 0; i2 < cells; ++i2) {
								var oCell = oRow.GetCell(i2)
								var shd = oCell.Cell.Get_Shd()
								var fill = shd.Fill
								if (fill && fill.r == 204 && fill.g == 255 && fill.b == 255) {
									write_list.push({
										id: 'c_' + oCell.Cell.Id,
										sub_type: 'cell',
										table_id: oElement.Table.Id,
										cell_id: oCell.Cell.Id
									})
								}
							}
						}
					}
				}
				node_list.push({
					id: tagInfo.client_id,
					regionType: tagInfo.regionType,
					parent_id: parent_id,
					write_list: write_list
				})
			}
		}
		return node_list
	}, false, true)
}

// 划分类型
function updateRangeControlType(typeName) {
	Asc.scope.typename = typeName
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	Asc.scope.question_map = window.BiyueCustomData.question_map
	console.log('updateRangeControlType begin:', typeName)
	return biyueCallCommand(window, function() {
		var typeName = Asc.scope.typename
		var oDocument = Api.GetDocument()
		var oRange = oDocument.GetRangeBySelect()
		var client_node_id = Asc.scope.client_node_id
		var question_map = Asc.scope.question_map || {}
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
		function getParentId(oControl) {
			var parentBlock = getParentBlock(oControl)
			var parent_id = 0
			if (parentBlock) {
				var parentTag = JSON.parse(parentBlock.GetTag() || '{}')
				parent_id = parentTag.client_id || 0
			}
			return parent_id
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
								regionType: tag.regionType,
								control_id: childControl.Sdt.GetId(),
								sub_type: 'control'
							})
						}
					}
				}
				return children
			}
			return null
		}

		function getDirectParentCell(oDrawing) {
			var drawingParentParagraph = oDrawing.GetParentParagraph()
			if (drawingParentParagraph) {
				var p2 = drawingParentParagraph.Paragraph.GetParent()
				if (p2) {
					var oP2 = Api.LookupObject(p2.Id)
					if (oP2.GetClassType() == 'documentContent') {
						var p3 = p2.GetParent()
						var oP3 = Api.LookupObject(p3.Id)
						if (oP3.GetClassType() == 'tableCell') {
							return oP3
						}
					}
				}
			}
			return null
		}
		// 删除题目互动
		function clearQuesInteraction(oControl) {
			if (!oControl) {
				return
			}
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
					var oParagraph = oControl.GetAllParagraphs()[0]
					var run1 = oParagraph.GetElement(0)
					var existSimple = run1 && run1.GetClassType() == 'run' && (run1.GetText() == '\u{e6a1}' || run1.GetText() == '▢')
					if (existSimple) {
						oParagraph.RemoveElement(0)
					}
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
										if (titleObj.feature.sub_type == 'ask_accurate') {
											var cellParent = getDirectParentCell(oDrawing)
											if (cellParent) {
												removeCellInteraction(cellParent)
											} else {
												oDrawing.Delete()
											}
										} else {
											oDrawing.Delete()
										}
									}
								}
							}
						}
					}
				}
			}
		}
		function removeCellInteraction(oCell) {
			if (!oCell) {
				return
			}
			oCell.SetBackgroundColor(204, 255, 255, true)
			var cellContent = oCell.GetContent()
			var paragraphs = cellContent.GetAllParagraphs()
			paragraphs.forEach(oParagraph => {
				var childCount = oParagraph.GetElementsCount()
				for (var i = 0; i < childCount; ++i) {
					var oRun = oParagraph.GetElement(i)
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
			})
		}
		function removeControl(oRemove) {
			if (!oRemove) {
				return
			}
			var tagRemove = JSON.parse(oRemove.GetTag() || '{}')
			clearQuesInteraction(oRemove)
			result.change_list.push({
				control_id: oRemove.Sdt.GetId(),
				client_id: tagRemove.client_id,
				parent_id: getParentId(oRemove),
				regionType: tagRemove.regionType
			})
			Api.asc_RemoveContentControlWrapper(oRemove.Sdt.GetId())
		}
		function getNewNodeList() {
			var controls = oDocument.GetAllContentControls() || []
			var nodeList = []
			controls.forEach((oControl) => {
				var tagInfo = JSON.parse(oControl.GetTag() || '{}')
				if (tagInfo.regionType) {
					nodeList.push({
						id: tagInfo.client_id,
						parentId: getParentId(oControl),
						regionType: tagInfo.regionType
					})
				}
			})
			return nodeList
		}
		function getOElements(Pos) {
			if (!Pos || Pos.length == 0) {
				return {
					cellContentId: 0,
					list: []
				}
			}
			var list = []
			var cellContentId = 0
			var controlId = 0
			var tableId = 0
			var tableIndex = -1
			var controlIndex = -1
			for (var i = 0; i < Pos.length; ++i) {
				var oElement = Api.LookupObject(Pos[i].Class.Id)
				if (oElement) {
					var classType = oElement.GetClassType()
					var container_type = ''
					var container_id = null
					if (classType == 'documentContent') {
						var parentId = oElement.Document.GetParent().Id
						var oParent = Api.LookupObject(parentId)
						if (oParent) {
							container_id = parentId
							container_type = oParent.GetClassType()
							if (oParent.GetClassType() == 'tableCell') {
								cellContentId = Pos[i].Class.Id
							} else if (oParent.GetClassType() == 'blockLvlSdt') {
								controlIndex = i
								controlId = parentId
							}
						}
					} else if (classType == 'table') {
						tableId = Pos[i].Class.Id
						tableIndex = i
					}
					list.push({
						oElement: oElement,
						Position: Pos[i].Position,
						classType: oElement.GetClassType(),
						container_type: container_type,
						container_id: container_id
					})
				}
			}
			return {
				cellContentId,
				list,
				tableId,
				tableIndex,
				controlId,
				controlIndex
			}
		}

		function cellNotControl(cellContent) {
			var elementCount = cellContent.GetElementsCount()
			if (elementCount == 0) {
				return true
			}
			for (var i = 0; i < elementCount; ++i) {
				var oElement = cellContent.GetElement(i)
				if (!oElement) {
					continue
				}
				if (oElement.GetClassType() == 'blockLvlSdt') {
					return false
				}
				if (oElement.GetClassType() == 'paragraph') {
					var count = oElement.GetAllContentControls()
					if (count > 0) {
						return false
					}
				}
			}
			return true
		}

		function setCellType(oCell, parent_id, table_id) {
			if (typeName == 'write') {
				if (!parent_id || !question_map[parent_id] || question_map[parent_id].level_type != 'question') {
					result = {
						code: 0,
						message: '未处于题目中',
					}
					return
				}
			}
			if (!oCell) {
				return
			}
			var cellContent = oCell.GetContent()
			if (typeName == 'write') { // 划分为小问，需要将选中的单元格都设置为小问
				// 需要确保所选单元格里没有control，且自己处于某个control里
				if (cellNotControl(cellContent)) {
					oCell.SetBackgroundColor(204, 255, 255, false)
					result.change_list.push({
						parent_id: parent_id,
						table_id: table_id,
						cell_id: oCell.Cell.Id,
						client_id: 'c_' + oCell.Cell.Id,
					})
				}
			} else if (typeName == 'clear' || typeName == 'clearAll') {
				removeCellInteraction(oCell)
				result.change_list.push({
					parent_id: parent_id,
					table_id: table_id,
					cell_id: oCell.Cell.Id,
					client_id: 'c_' + oCell.Cell.Id,
					regionType: 'write'
				})
			}
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
		var selectionInfo = oDocument.Document.getSelectionInfo()
		if (!oRange) {
			var elementData = getOElements(selectionInfo.curPos)
			console.log('elementList', elementData)
			var container = null
			var container_type = null
			if (elementData.list) {
				for (var i = elementData.list.length - 1; i >= 0; --i) {
					if (elementData.list[i].classType == 'inlineLvlSdt') {
						container = elementData.list[i].oElement
						container_type = elementData.list[i].classType
						break
					} else if (elementData.list[i].classType == 'documentContent') {
						if (elementData.list[i].container_type == 'tableCell' || elementData.list[i].container_type == 'blockLvlSdt') {
							container = Api.LookupObject(elementData.list[i].container_id)
							container_type = elementData.list[i].container_type
							break
						}
					}
				}
			}
			if (container) {
				if (container_type == 'blockLvlSdt' || container_type == 'inlineLvlSdt') {
					var oControl = container
					var tag = JSON.parse(oControl.GetTag() || '{}')
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
							obj.children = getChildControls(oControl)
						}
						obj.client_id = tag.client_id
						obj.parent_id = getParentId(oControl)
						obj.control_id = oControl.Sdt.GetId()
						obj.regionType = tag.regionType
						var templist = []
						var posinparent = oControl.GetPosInParent()
						var oParent = Api.LookupObject(oControl.Sdt.GetParent().Id)
						if (typeName == 'setBig') { // 设置为大题
							// 需要将后面的级别比他小的控件挪到它的范围内
							var parentElementCount = oParent.GetElementsCount()
							for (var i = posinparent + 1; i < parentElementCount; ++i) {
								var element = oParent.GetElement(i)
								if (!element) {
									break
								}
								if (!element.GetClassType) {
									break
								}
								if (element.GetClassType() == 'blockLvlSdt') {
									var nextTag = JSON.parse(element.GetTag() || '{}')
									if (nextTag.regionType == 'question') {
										if (nextTag.lvl <= tag.lvl) {
											break
										}
									}
								} else if (element.GetClassType() == 'table') {
									var oCell = element.GetCell(0, 0)
									if (oCell) {
										var oCellContent = oCell.GetContent()
										var paragraphs = oCellContent.GetAllParagraphs()
										if (paragraphs && paragraphs.length) {
											var existLvl = paragraphs.filter(p => {
												var NumPr = p.GetNumbering()
												if (NumPr && NumPr.Lvl < tag.lvl) {
													return true
												}
											})
											if (existLvl && existLvl.length) {
												break
											}
										}
									}
								} else if (element.GetClassType() == 'paragraph') {
									var NumPr = element.GetNumbering()
									if (NumPr && NumPr.Lvl <= tag.lvl) {
										break
									}
								}
								templist.push(element)
							}
							for (var i = 0; i < templist.length; ++i) {
								oParent.RemoveElement(posinparent + 1)
							}
							if (templist.length) {
								var count = oControl.GetContent().GetElementsCount()
								templist.forEach((e, index) => {
									oControl.AddElement(e, count + index)
								})
							}
						} else if (typeName == 'clearBig') { // 清除大题
							// 需要将包含的子控件挪出它的范围内
							var controlContent = oControl.GetContent()
							var childCount = controlContent.GetElementsCount()
							for (var i = childCount - 1; i >= 0; --i) {
								var element = controlContent.GetElement(i)
								if (!element) {
									break
								}
								if (!element.GetClassType) {
									break
								}
								if (element.GetClassType() == 'blockLvlSdt') {
									templist.push(element)
									controlContent.RemoveElement(i)
								} else if (element.GetClassType() == 'table') {
									if (element.GetTableTitle() != 'questionTable') {
										var find = false
										// 判断所有单元格都是control
										var rows = element.GetRowsCount()
										for (var r = 0; r < rows.length; ++r) {
											var oRow = element.GetRow(r)
											var cnt = oRow.GetElementsCount()
											for (var c = 0; c < cnt; ++c) {
												var oCell = oRow.GetCell(c)
												var oCellContent = oCell.GetContent()
												var elcount = oCellContent.GetElementsCount()
												for (var k = 0; k < elcount; ++k) {
													var el = oCellContent.GetElement(k)
													if (!el || !el.GetClassType || el.GetClassType() != 'blockLvlSdt') {
														find = true
														break
													}
												}
												if (find) {
													break
												}
											}
											if (find) {
												break
											}
										}
										if (find) {
											break
										}
									}
									templist.push(element)
									controlContent.RemoveElement(i)
								}
							}
							if (templist.length) {
								templist.forEach(e => {
									oParent.AddElement(posinparent + 1, e)
								})
							}
						}
						obj.text = tag.regionType == 'question' ? oControl.GetRange().GetText() : null
						obj.children = getChildControls(oControl)
						result.change_list.push(obj)
						if (typeName == 'struct') {
							clearQuesInteraction(oControl)
						}
					}
				} else if (container_type == 'tableCell') {
					var oParentControl = null
					var parent_client_id = 0
					for (var j = elementData.list.length - 1; j >= 0; --j) {
						if (elementData.list[j].classType == 'documentContent' && elementData.list[j].container_type == 'blockLvlSdt') {
							oParentControl = Api.LookupObject(elementData.list[j].container_id)
							if (oParentControl) {
								var parentTag = JSON.parse(oParentControl.GetTag() || '{}')
								parent_client_id = parentTag.client_id
								break
							}
						}
					}
					setCellType(container, parent_client_id, container.GetParentTable().Table.Id)
				}
			} else {
				return {
					code: 0,
					message: '请先选中一个范围',
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
			var startData = getOElements(oRange.StartPos)
			var endData = getOElements(oRange.EndPos)
			console.log('startData', startData, endData)
			// 暂时只支持选中多个单元格设置小问和清除区域
			if ((typeName == 'write' || typeName == 'clear' || typeName == 'clearAll') &&
				startData.cellContentId && endData.cellContentId
				&& startData.tableId == endData.tableId && 
				startData.cellContentId != endData.cellContentId &&
				startData.controlId && startData.controlId == endData.controlId &&
				startData.controlIndex < startData.tableIndex) {
				// 选中的是单元格
				// 需要拿到单元格坐标，修改单元格背景色
				var oTable = Api.LookupObject(startData.tableId)
				var rows = oTable.GetRowsCount()
				var oParentControl = Api.LookupObject(startData.controlId)
				var oParentTag = JSON.parse(oParentControl.GetTag() || '{}')
				if (oParentTag && oParentTag.client_id) {
					for (var i = 0; i < rows; ++i) {
						var oRow = oTable.GetRow(i)
						var cellCounts = oRow.GetCellsCount()
						for (var j = 0; j < cellCounts; ++j) {
							var oCell = oRow.GetCell(j)
							var cellContent = oCell.GetContent()
							if (cellContent.Document.Selection && cellContent.Document.Selection.Use) {
								setCellType(oCell, oParentTag.client_id, startData.tableId)
							}
						}
					}	
				}
			} else {
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
						var regionType = typeName == 'write' ? 'write' : 'question'
						var tag = {
							client_id: result.client_node_id,
							regionType: regionType,
							lvl: level
						}
						var oResult = Api.asc_AddContentControl(type, {
							Tag: JSON.stringify(tag)
						})
						console.log('asc_AddContentControl', oResult)
						if(oResult) {
							var oControl = Api.LookupObject(oResult.InternalId)
							// 需要返回新增的nodeIndex todo..
							result.change_list.push({
								client_id: tag.client_id,
								control_id: oResult.InternalId,
								text: oRange.GetText(),
								children: oControl && oControl.GetClassType() == 'blockLvlSdt' ? getChildControls(oControl) : [],
								parent_id: getParentId(oControl),
								regionType: regionType
							})
						}
						// 若是在单元格里添加control后，会多出一行需要删除
						if (isInCell) {
							var oCell = Api.LookupObject(oResult.InternalId).GetParentTableCell()
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
		}
		result.new_node_list = getNewNodeList()
		return result
	}, false, true).then(res1 => {
		console.log('handleChangeType result', res1)
		if (res1) {
			if (res1.message && res1.message != '') {
				alert(res1.message)
			} else {
				getNodeList().then(res2 => {
					handleChangeType(res1, res2)	
				})
			}
		}
	})
}

function handleChangeType(res, res2) {
	console.log('handleChangeType', res, res2)
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
	var addIds = []
	var update_node_id = g_click_value ? g_click_value.Tag.client_id : 0
	function updateAskList(qid, ask_list) {
		var old_list = question_map[qid].ask_list || []
		var new_list = []
		if (ask_list) {
			for (var a = 0; a < ask_list.length; ++a) {
				var oidx = old_list.findIndex(e1 => {
					return e1.id == ask_list[a].id
				})
				new_list.push({
					id: ask_list[a].id,
					score: oidx >= 0 ? (old_list[oidx].score || 0) : 0
				})
			}
		}
		question_map[qid].ask_list = new_list
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
		var ask_list
		if (item.children && item.children.length) {
			ask_list = item.children.filter(child => {
				return child.regionType == 'write'
			})
		}
		if (nodeData) {
			if (level_type == 'clear' || level_type == 'clearAll') {
				node_list.splice(nodeIndex, 1)
			} else {
				nodeData.is_big = level_type == 'setBig'
				if (nodeData.level_type != 'question' && targetLevel == 'question') {
					addIds.push(nodeData.id)
				}
				nodeData.level_type = targetLevel
			}
			if (targetLevel == 'question') {
				nodeData.write_list = ask_list
				if (!question_map[item.client_id]) {
					var write_list = nodeData.write_list || []
					question_map[item.client_id] = {
						text: item.text,
						level_type: targetLevel,
						ques_default_name: GetDefaultName(targetLevel, item.text),
						interaction: window.BiyueCustomData.interaction,
						ask_list: write_list.map(e => {
							return {
								id: e.id
							}
						})
					}
				} else {
					updateAskList(item.client_id, ask_list)
					if (level_type == 'setBig' || level_type == 'clearBig') {
						question_map[item.client_id].text = item.text
						question_map[item.client_id].ques_default_name = GetDefaultName(targetLevel, item.text)
						question_map[item.client_id].level_type = targetLevel
					} else {
						question_map[item.client_id].level_type = targetLevel
					}
				}
			} else {
				if (question_map[item.client_id]) {
					delete question_map[item.client_id]
				}
			}
		} else if (level_type == 'write') {
			var parent_id = item.parent_id
			if (parent_id > 0) {
				var parentNode = node_list.find(e => {
					return e.id == parent_id
				})
				if (parentNode) {
					addIds.push(parent_id)
					var ndata = res2.find(e => {
						return e.id == parent_id
					})
					if (ndata) {
						parentNode.write_list = ndata.write_list
						updateAskList(parent_id, ndata.write_list)
					}
				}
			}
		} else if (level_type != 'clear' && level_type != 'clearAll') {
			// 之前没有，需要增加
			var index = node_list.length
			if (res.new_node_list) { // 需要算出之前的node的位置
				var idx = res.new_node_list.findIndex(e => {
					return e.id == item.client_id
				})
				if (idx > 0) {
					for (var idx2 = idx - 1; idx2 >=0; --idx2) {
						if (res.new_node_list[idx2].regionType == 'question') {
							var idx3 = node_list.findIndex(e => {
								return e.id == res.new_node_list[idx2].id
							})
							if (idx3 >= 0) {
								index = idx3
								break
							}
						}

					}
				} else {
					index = 0
				}
			}
			node_list.splice(index, 0, {
				id: item.client_id,
				control_id: item.control_id,
				regionType: item.regionType,
				level_type: targetLevel,
				write_list: ask_list
			})
			if (targetLevel == 'question') {
				addIds.push(item.client_id)
			}
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
		} else {
			if (item.regionType == 'write' && item.parent_id) {
				update_node_id = item.parent_id
				var nodeData = node_list.find(e => {
					return e.id == item.parent_id
				})
				if (nodeData && nodeData.write_list) {
					var write_index = nodeData.write_list.findIndex(e => {
						return e.id == item.client_id
					})
					if (write_index >= 0) {
						nodeData.write_list.splice(write_index, 1)
					}
				}
				if (question_map[item.parent_id] && question_map[item.parent_id].ask_list) {
					var ask_index = question_map[item.parent_id].ask_list.findIndex(e => {
						return e.id == item.client_id
					})
					if (ask_index >= 0) {
						question_map[item.parent_id].ask_list.splice(ask_index, 1)
					}
				}
			}
		}
	})

	window.BiyueCustomData.node_list = node_list
	window.BiyueCustomData.question_map = question_map
	if (addIds && addIds.length) {
		if (level_type == 'write') {
			setInteraction(question_map[addIds[0]].interaction, addIds).then(() => window.biyue.StoreCustomData())
		} else if (targetLevel == 'question') {
			setInteraction(window.BiyueCustomData.interaction, addIds).then(() => window.biyue.StoreCustomData())
		}
		update_node_id = addIds[0]
	} else {
		window.biyue.StoreCustomData()
	}
	console.log('handleChangeType end', node_list, 'g_click_value', g_click_value, 'update_node_id', update_node_id)
	document.dispatchEvent(
		new CustomEvent('updateQuesData', {
			detail: {
				client_id: update_node_id
			}
		})
	)
}
// 获取批量操作列表
function getBatchList() {
	Asc.scope.node_list = window.BiyueCustomData.node_list
	return biyueCallCommand(window, function () {
		var oDocument = Api.GetDocument()
		var control_list = oDocument.GetAllContentControls()
		var ques_id_list = []
		var oRange = oDocument.GetRangeBySelect()
		var type = 'Selection'
		if (!oRange) {
			type = 'Target'
			var node_list = Asc.scope.node_list || []
			var currentContentControl = oDocument.Document.GetContentControl()
			var controls = oDocument.GetAllContentControls()
			if (currentContentControl) {
				var oControl = Api.LookupObject(currentContentControl.Id)
				if (oControl) {
					if (oControl.GetClassType() == 'inlineLvlSdt') {
						oControl = oControl.GetParentContentControl()
					}
					if (oControl) {
						var tag = JSON.parse(oControl.GetTag() || '{}')
						if (tag.regionType == 'question' && tag.client_id) {
							var nodeData = node_list.find(e => {
								return e.id == tag.client_id
							})
							if (nodeData) {
								if (nodeData.level_type == 'struct') {
									var controlIndex = controls.findIndex(e => {
										return e.Sdt.GetId() == currentContentControl.Id
									})
									for (var i = controlIndex + 1; i < controls.length; i++) {
										var nextControl = controls[i]
										if (nextControl.GetClassType() == 'blockLvlSdt') {
											var nextTag = JSON.parse(nextControl.GetTag() || '{}')
											if (nextTag.lvl <= tag.lvl) {
												break
											}
											ques_id_list.push({
												id: nextTag.client_id,
												control_id: nextControl.Sdt.GetId()
											})
										}
									}
								} else if (nodeData.level_type == 'question') {
									ques_id_list.push({
										id: tag.client_id,
										control_id: oControl.Sdt.GetId()
									})
									if (nodeData.is_big) {
										var childControls = oControl.GetAllContentControls()
										if (childControls) {
											childControls.forEach(e => {
												if (e.GetClassType() == 'blockLvlSdt') {
													var childTag = JSON.parse(e.GetTag() || '{}')
													if (childTag.regionType == 'question' && childTag.client_id) {
														ques_id_list.push({
															id: childTag.client_id,
															control_id: e.Sdt.GetId()
														})
													}
												}
											})
										}
									}
								}
							}
						}
					}
				}
			}
		} else {
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
		}
		return {
			code: 1,
			list: ques_id_list,
			type: type
		}
	}, false, false)
}
// 批量修改题型
function batchChangeQuesType(type) {
	return getBatchList().then(res => {
		if (!res || !res.code || !res.list || res.list.length == 0) {
			return
		}
		var question_map = window.BiyueCustomData.question_map
		res.list.forEach(e => {
			if (question_map && question_map[e.id]) {
				question_map[e.id].question_type = type * 1
			}
		})
		// 需要同步更新单题详情
		if (window.tab_select == 'tabQues') {
			document.dispatchEvent(
				new CustomEvent('updateQuesData', {
					detail: {
						list: res.list,
						field: 'question_type',
						value: type * 1,
					},
				})
			)
		}
		window.biyue.StoreCustomData()
	})
}
// 批量设置互动
function batchChangeInteraction(type) {
	return getBatchList().then(res => {
		if (!res || !res.code || !res.list || res.list.length == 0) {
			return
		}
		var question_map = window.BiyueCustomData.question_map
		res.list.forEach(e => {
			if (question_map && question_map[e.id]) {
				question_map[e.id].interaction = type
			}
		})
		setInteraction(type, res.list.map(e => e.id)).then(() => {
			return window.biyue.StoreCustomData()
		})
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

		function GetNumberingValue(oControl) {
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
				return null
			}
			var paragraphs = oControl.GetAllParagraphs()
			for (var i = 0; i < paragraphs.length; ++i) {
				var oParagraph = paragraphs[i]
				if (oParagraph) {
					var parent1 = oParagraph.Paragraph.Parent
					var parent2 = parent1.Parent
					if (parent2) {
						if (parent2.Id == oControl.Sdt.GetId()) {
							if (oParagraph.Paragraph.HaveNumbering()) {
								return oParagraph.Paragraph.GetNumberingText()
							}
							return null
						}
					}
				}
			}
			return null
		}
		var colors = {}

		var oTextPr = Api.CreateTextPr()
    	oTextPr.SetColor(255, 0, 0, false)
    	var oColor2 = oTextPr.TextPr.GetColor().Copy()
		colors['blockLvlSdt'] = oColor2

		var oColor3 = oColor2.Copy()
		oColor3.Set(0, 255, 0, false)
		colors['inlineLvlSdt'] = oColor3
		controls.forEach((oControl) => {
			var tagInfo = JSON.parse(oControl.GetTag() || '{}')
			// if (oControl.GetClassType() == 'inlineLvlSdt') {
				// tagInfo.color = '#ff0000'
				// oControl.SetTag(JSON.stringify(tagInfo))
			//}
			// tagInfo.color = '#ff0000'
			// oControl.Sdt.SetColor(colors[oControl.GetClassType()])
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
					parentId: parentid,
					numbing_text: GetNumberingValue(oControl)
				})
			}
		})
		return nodeList
	}, false, false).then(nodeList => {
		console.log('initControls   nodeList', nodeList)
		return new Promise((resolve, reject) => {
			// todo.. 这里暂不考虑上次的数据未保存或保存失败的情况，只假设此时的control数据和nodelist里的是一致的，只是乱码而已，其他的后续再处理
			if (nodeList && nodeList.length > 0) {
				var question_map = window.BiyueCustomData.question_map || {}
				nodeList.forEach(node => {
					if (question_map[node.id]) {
						question_map[node.id].text = node.text
						question_map[node.id].ques_default_name = node.numbing_text ? node.numbing_text : GetDefaultName(question_map[node.id].level_type, node.text)
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
		function GetNumberingValue(oControl) {
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
				return null
			}
			var paragraphs = oControl.GetAllParagraphs()
			for (var i = 0; i < paragraphs.length; ++i) {
				var oParagraph = paragraphs[i]
				if (oParagraph) {
					var parent1 = oParagraph.Paragraph.Parent
					var parent2 = parent1.Parent
					if (parent2) {
						if (parent2.Id == oControl.Sdt.GetId()) {
							if (oParagraph.Paragraph.HaveNumbering()) {
								return oParagraph.Paragraph.GetNumberingText()
							}
							return null
						}
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
								sub_type: 'control'
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
						level_type: level_type,
						numbing_text: GetNumberingValue(oControl)
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
		console.log('===== confirmLevelSet res', res)
		if (res) {
			window.BiyueCustomData.client_node_id = res.client_node_id
			window.BiyueCustomData.node_list = res.nodeList
			Object.keys(res.questionMap).forEach(key => {
				var qdata = res.questionMap[key]
				res.questionMap[key].ques_default_name = qdata.numbing_text ? qdata.numbing_text : GetDefaultName(qdata.level_type, qdata.text)
			})
			window.BiyueCustomData.question_map = res.questionMap
		}
		return initExtroInfo()
	})
	.then(() => reqGetQuestionType())
	.then(() => {
		console.log("================================ StoreCustomData")
		window.biyue.StoreCustomData()
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
		return new Promise((resolve, reject) => {
			if (!window.BiyueCustomData.paper_uuid || !control_list || control_list.length == 0) {
				resolve()
				return
			}
			getQuesType(window.BiyueCustomData.paper_uuid, control_list).then(res => {
				console.log('getQuesType success ', res)
				var content_list = res.data.content_list
				if (content_list && content_list.length) {
					content_list.forEach(e => {
						window.BiyueCustomData.question_map[e.id].question_type = e.question_type * 1
						// window.BiyueCustomData.question_map[e.id].question_type_name = e.question_type_name
						// 存储时question_type_name莫名其妙变得很大，不再存储
					})
				}
				resolve()
			}).catch(res => {
				console.log('getQuesType fail ', res)
				resolve()
			})
		})
	})
}
// 显示/隐藏/删除所有书写区
function handleAllWrite(cmdType) {
	Asc.scope.cmdType = cmdType
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var drawings = oDocument.GetAllDrawingObjects()
		var cmdType = Asc.scope.cmdType
		function updateFill(drawing, oFill) {
			if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== 'fill') {
				return false
			}
			drawing.GraphicObj.spPr.setFill(oFill.UniFill)
		}
		var list = []
		for (var i = 0, imax = drawings.length; i < imax; ++i) {
			var oDrawing = drawings[i]
			var title = oDrawing.Drawing.docPr.title || '{}'
			var titleObj = JSON.parse(title)
			if (titleObj.feature && titleObj.feature.zone_type == 'question' && titleObj.feature.sub_type == 'write') {
				if (cmdType == 'show') {
					var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 0, 0))
					oFill.UniFill.transparent = 255 * 0.2 // 透明度
					updateFill(oDrawing.Drawing, oFill)
				} else if (cmdType == 'hide') {
					updateFill(oDrawing.Drawing, Api.CreateNoFill())
				} else if (cmdType == 'del') {
					list.push({
						parent_id: titleObj.feature.parent_id,
						id: titleObj.feature.client_id
					})
					oDrawing.Delete()
				}
			}
		}
		return {
			list: list,
			cmdType: cmdType
		}
	}, false, true).then(res => {
		console.log('handleAllWrite', res)
		if (res) {
			if (res.cmdType == 'del' && res.list) {
				var node_list = window.BiyueCustomData.node_list || []
				var question_map = window.BiyueCustomData.question_map || {}
				res.list.forEach(e => {
					var nodeData = node_list.find(item => {
						return item.id == e.parent_id
					})
					if (nodeData && nodeData.write_list) {
						var writeIndex = nodeData.write_list.findIndex(item => {
							return item.id == e.id
						})
						if (writeIndex >= 0) {
							nodeData.write_list.splice(writeIndex, 1)
						}
					}
					if (question_map[e.parent_id] && question_map[e.parent_id].ask_list) {
						var askIndex = question_map[e.parent_id].ask_list.findIndex(item => {
							return item.id == e.id
						})
						if (askIndex >= 0) {
							question_map[e.parent_id].ask_list.splice(askIndex, 1)
						}
					}
				})
			}
		}
	})
}
// 显示或隐藏所有单元格小问
function showAskCells(cmdType) {
	Asc.scope.question_map = window.BiyueCustomData.question_map
	Asc.scope.node_list = window.BiyueCustomData.node_list
	Asc.scope.cmdType = cmdType
	return biyueCallCommand(window, function() {
		var question_map = Asc.scope.question_map || {}
		var node_list = Asc.scope.node_list || []
		var cmdType = Asc.scope.cmdType
		Object.keys(question_map).forEach(id => {
			if (question_map[id].level_type == 'question') {
				var nodeData = node_list.find(item => {
					return item.id == id
				})
				if (nodeData && nodeData.write_list) {
					if (question_map[id].ask_list) {
						question_map[id].ask_list.forEach(ask => {
							var writeData = nodeData.write_list.find(w => {
								return w.id == ask.id
							})
							if (writeData && writeData.sub_type == 'cell' && writeData.cell_id) {
								var oCell = Api.LookupObject(writeData.cell_id)
								if (oCell) {
									oCell.SetBackgroundColor(204, 255, 255, cmdType == 'show' ? false : true)
								}
							}
						})
					}
				}
			}	
		})
	}, false, true)
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
				oDrawing.SetPaddings(0, 0, 0, 0)
				var paragraphs = oControl.GetAllParagraphs()
				if (paragraphs && paragraphs.length > 0) {
					var paragraphs = oControl.GetAllParagraphs()
					var parentParagraph = null
					for (var i = 0; i < paragraphs.length; ++i) {
						var oParagraph = paragraphs[i]
						if (oParagraph) {
							var parent1 = oParagraph.Paragraph.Parent
							var parent2 = parent1.Parent
							if (parent2 && parent2.Id == oControl.Sdt.GetId()) {
								parentParagraph = oParagraph
								break
							}
						}
					}
					var oRun = Api.CreateRun()
					oRun.AddDrawing(oDrawing)
					parentParagraph.AddElement(
						oRun,
						1
					)
				}
				return {
					client_node_id: client_node_id,
					cmdType: write_cmd,
					sub_type: 'write',
					ques_id: tag.client_id,
					write_id: client_node_id,
					drawing_id: oDrawing.Drawing.Id,
					shape_id: oDrawing.Shape.Id,
				}
			} else if (write_cmd == 'del') {
				var selectDrawings = oDocument.GetSelectedDrawings()
				var drawings = []
				if (selectDrawings && selectDrawings.length) {
					for (var i = 0; i < selectDrawings.length; ++i) {
						if (selectDrawings[i].GetClassType() == 'shape') {
							drawings.push(selectDrawings[i])
						}
					}
				} else {
					drawings = oControl.GetAllDrawingObjects()
				}
				var allDrawings = oDocument.GetAllDrawingObjects()
				if (drawings && drawings.length > 0) {
					var result = {
						client_node_id: client_node_id,
						cmdType: write_cmd,
						ques_id: tag.client_id,
						remove_ids: []
					}
					for (var i = 0; i < drawings.length; ++i) {
						var title = drawings[i].Drawing.docPr.title || '{}'
						var titleObj = JSON.parse(title)
						if (titleObj.feature && titleObj.feature.zone_type == 'question' && titleObj.feature.sub_type == 'write') {
							var drawingId = drawings[i].Drawing.Id
							var oDrawing = allDrawings.find(e => {
								return e.Drawing.Id == drawingId
							})
							if (oDrawing) {
								oDrawing.Delete()
								result.remove_ids.push(titleObj.feature.client_id)
							}
						}
					}
					return result
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
								res.write_id = client_node_id
								var oDrawing = createDrawing(paraentControl.Sdt.GetId(), tag.client_id)
								if (runIdx >= 0) {
									res.drawing_id = oDrawing.Drawing.Id
									res.shape_id = oDrawing.Shape.Id
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
			res.sub_type = 'identify'
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
					sub_type: res.sub_type,
					drawing_id: res.drawing_id,
					shape_id: res.shape_id
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
		setInteraction(question_map[res.ques_id].interaction, [res.ques_id])
		document.dispatchEvent(
			new CustomEvent('updateQuesData', {
				detail: {
					client_id: res.ques_id
				}
			})
		)
	}
	window.biyue.StoreCustomData()
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
			e.id = e.id + ''
			tree.push(e)
		} else {
			var parent = getDataById(tree, e.parent_id)
			if (parent) {
				if (!parent.children) {
					parent.children = []
				}
				e.id = e.id + ''
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
	reqComplete(uploadTree, version).then(res => {
		console.log('reqComplete', res)
		console.log('[reqUploadTree end]', Date.now())
		if (res.data.questions) {
			res.data.questions.forEach(e => {
				window.BiyueCustomData.question_map[e.id].uuid = e.uuid
			})
		}
		setBtnLoading('uploadTree', false)
		alert('全量更新成功')
	}).catch(res => {
		console.log('reqComplete fail', res)
		console.log('[reqUploadTree end]', Date.now())
		setBtnLoading('uploadTree', false)
		alert('全量更新失败')
	})
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
// 批量修改占比
function batchChangeProportion(proportion) {
	return getBatchList().then(res => {
		if (!res || !res.code || !res.list || res.list.length == 0) {
			return
		}
		var question_map = window.BiyueCustomData.question_map
		res.list.forEach(e => {
			if (question_map && question_map[e.id]) {
				question_map[e.id].proportion = proportion
			}
		})
		batchProportion(res.list, proportion).then(() => {
			return window.biyue.StoreCustomData()
		})
	})
}

function batchProportion(idList, proportion) {
	if (!idList || idList.length == 0) {
		return
	}
	Asc.scope.change_id_list = idList
	Asc.scope.proportion = proportion
	return biyueCallCommand(window, function() {
		var idList = Asc.scope.change_id_list
		var question_map = Asc.scope.question_map || {}
		var oDocument = Api.GetDocument()
		var oControls = oDocument.GetAllContentControls()
		var target_proportion = Asc.scope.proportion || 1
		var TABLE_TITLE = 'questionTable'
		var newW = 100 / target_proportion
		// 先考虑所有题目都不在表格内的情况
		var templist = []
		var iTable = 0
		var tables = {}
		var wBefore = 0
		var toIndex = -1
		var sections = oDocument.GetSections()
		var oSection = sections[0]
		var PageSize = oSection.Section.PageSize
		var PageMargins = oSection.Section.PageMargins
		var tw = PageSize.W - PageMargins.Left - PageMargins.Right
		function isTableEmpty(tableId) {
			var oTable = Api.LookupObject(tableId)
			if (!oTable) {
				return false
			}
			if (oTable.GetClassType() != 'table') {
				return false
			}
			var rows = oTable.GetRowsCount()
			for (var row = 0; row < rows; ++row) {
				var oRow = oTable.GetRow(row)
				var cellsCount = oRow.GetCellsCount()
				for (var cell = 0; cell < cellsCount; ++cell) {
					var oCell = oRow.GetCell(0)
					var elementCount = oCell.GetContent().GetElementsCount()
					if (elementCount > 0) {
						for (var j = 0; j < elementCount; ++j) {
							var oElement = oCell.GetContent().GetElement(j)
							if (oElement.GetClassType() == 'paragraph') {
								if (oElement.GetElementsCount() > 0) {
									return false
								}
							} else {
								return false
							}
						}
					}
				}
			}
			return true
		}
		function AddControlToCell2(content, oCell) {
			if (!oCell) {
				console.log('AddControlToCell2 ocell is null')
				return
			}
			if (!content) {
				console.log('AddControlToCell2 content is null')
				return
			}
			oCell.SetCellMarginLeft(0)
			oCell.SetCellMarginRight(0)
			var cellContent = oCell.GetContent()
			content.forEach((e, index) => {
				cellContent.AddElement(index, e)
			})
			var lastindex = cellContent.GetElementsCount() - 1
			if (cellContent.GetElementsCount() > 1 && cellContent.GetElement(lastindex).GetClassType() == 'paragraph') {
				cellContent.RemoveElement(lastindex)
			}
		}
		for (var idx = 0; idx < idList.length; ++idx){
			var id = idList[idx].id
			var quesData = question_map[id]
			if (!quesData) {
				continue
			}
			var oControl = oControls.find(e => {
				var tag = JSON.parse(e.GetTag() || '{}')
				return tag.client_id == id
			})
			if (!oControl) {
				continue
			}
			var posinparent = oControl.GetPosInParent()
			var parent = oControl.Sdt.GetParent()
			var oParent = Api.LookupObject(parent.Id)
			var isInCell = false
			var parentTable = null
			var parentCell = null
			if (oParent && oParent.GetClassType && oParent.GetClassType() == 'documentContent') {
				var parent2 = oParent.Document.GetParent()
				if (parent2) {
					var oParent2 = Api.LookupObject(parent2.Id)
					if (oParent2 && oParent2.GetClassType && oParent2.GetClassType() == 'tableCell') {
						isInCell = true
						parentTable = oControl.GetParentTable()
						parentCell = oControl.GetParentTableCell()
					}
				}
			}
			if (isInCell) {
				templist.push({
					content: [oControl],
					oParent: oParent,
					posinparent: posinparent,
					tablePos: parentTable.GetPosInParent(),
					tableParent: Api.LookupObject(parentTable.Table.Parent.Id),
					tableId: parentTable.Table.Id
				})
			} else {
				templist.push({
					content: [oControl],
					oParent: oParent,
					posinparent: posinparent
				})
			}
			if (wBefore + newW > 100) {
				iTable++
				tables[iTable] = {
					cells: [{
						icell: 0,
						W: newW
					}],
					W: newW
				}
				wBefore = newW
			} else {
				if (!tables[iTable]) {
					tables[iTable] = {
						cells: [{
							icell: 0,
							W: newW
						}],
						W: newW
					}
				} else {
					tables[iTable].cells.push({
						icell: tables[iTable].cells.length,
						W: newW
					})
					tables[iTable].W += newW
				}
				wBefore += newW
			}
		}
		for (var j = templist.length - 1; j >= 0; --j) {
			templist[j].oParent.RemoveElement(templist[j].posinparent)
			if (templist[j].tableParent) {
				if (isTableEmpty(templist[j].tableId)) {
					templist[j].tableParent.RemoveElement(templist[j].tablePos)
				}
			}
		}
		var count = 0
		for (var i = 0; i <= iTable; ++i) {
			if (!tables[i]) {
				break
			}
			var targetcellscount = tables[i].cells.length
			var oTable
			if (i <= toIndex) {
				oTable = oStartTableParent.GetElement(startTablePos + i)
				if (oTable.GetClassType() != 'table') {
					break
				}
				var oRow = oTable.GetRow(0)
				var cellscount = oRow.GetCellsCount()
				if (cellscount < targetcellscount) { // 需要拆分
					oTable.Split(oTable.GetCell(0, 0), 1, targetcellscount - cellscount + 1);
				} else if (cellscount > targetcellscount) { // 需要合并
					var mergeCells = []
					for (var m = 0; m <= cellscount - targetcellscount; ++m) {
						mergeCells.push(oTable.GetCell(0, m))
					}
					oTable.MergeCells(mergeCells)
				}
			} else {
				oTable  = Api.CreateTable(targetcellscount, 1)
				oTable.SetCellSpacing(0)
				oTable.SetTableTitle(TABLE_TITLE)
			}
			oTable.SetWidth('percent', tables[i].W)
			tables[i].cells.forEach((cell, cidx) => {
				oTable.GetCell(0, cell.icell).SetWidth('percent', cell.W)
				AddControlToCell2(templist[count].content, oTable.GetCell(0, cell.icell))
				count++
				// // var tw = oTable.Table.Pages[0].XLimit - oTable.Table.Pages[0].X
				// console.log('tw', tw)
				// var twips = (tw * (cell.W / 100)) / (25.4 / 72 / 20)
				// oTable.GetCell(0, cell.icell).SetWidth('twips', twips)
			})
			if (i >= toIndex) {
				if (templist[0].tableId) {
					templist[0].tableParent.AddElement(templist[0].tablePos + i, oTable)
				} else {
					templist[0].oParent.AddElement(templist[0].posinparent + i, oTable)
				}
			}
		}
	}, false, true).then(res => {
		console.log('batchProportion', res)
	})
}

function changeProportion(idList, proportion) {
	if (!idList || idList.length == 0) {
		return
	}
	Asc.scope.node_list = window.BiyueCustomData.node_list
	Asc.scope.question_map = window.BiyueCustomData.question_map
	Asc.scope.change_id_list = idList
	Asc.scope.proportion = proportion
	return biyueCallCommand(window, function() {
		var idList = Asc.scope.change_id_list
		var question_map = Asc.scope.question_map || {}
		var target_proportion = Asc.scope.proportion || 1
		var oDocument = Api.GetDocument()
		var oControls = oDocument.GetAllContentControls()
		var TABLE_TITLE = 'questionTable'
		var newW = 100 / target_proportion
		function AddControlToCell(oControl, oCell) {
			var templist = []
			templist.push(oControl)
			var posinparent = oControl.GetPosInParent()
			var parent = oControl.Sdt.GetParent()
			var oParent = Api.LookupObject(parent.Id)
			oParent.RemoveElement(posinparent)
			oCell.SetWidth('percent', 100 / target_proportion)
			oCell.AddElement(0, templist[0])
			if (oCell.GetContent().GetElementsCount() == 2) {
				oCell.GetContent().RemoveElement(1)
			}
		}

		function AddControlToCell2(content, oCell) {
			if (!oCell) {
				console.log('AddControlToCell2 ocell is null')
				return
			}
			if (!content) {
				console.log('AddControlToCell2 content is null')
				return
			}
			oCell.SetCellMarginLeft(0)
			oCell.SetCellMarginRight(0)
			var cellContent = oCell.GetContent()
			content.forEach((e, index) => {
				cellContent.AddElement(index, e)
			})
			var lastindex = cellContent.GetElementsCount() - 1
			if (cellContent.GetElementsCount() > 1 && cellContent.GetElement(lastindex).GetClassType() == 'paragraph') {
				cellContent.RemoveElement(lastindex)
			}
		}
		// 上一个表格有足够的空间可以放，需要将上一个表格的最后一个单元格拆分，然后下面的表格内容挪到上面
		function func1(startTable, oControl) {
			var templist = []
			var iTable = 0
			var tables = {}
			var wBefore = 0
			var toIndex = -1
			var startTablePos = startTable.GetPosInParent()
			var dataHandled = {}
			dataHandled.parentTable = oControl.GetParentTable()
			dataHandled.parentTableCell = oControl.GetParentTableCell()
			if (dataHandled.parentTableCell) {
				dataHandled.cellIndex = dataHandled.parentTableCell.GetIndex()
				dataHandled.rowIndex = dataHandled.parentTableCell.GetRowIndex()
				dataHandled.table_id = dataHandled.parentTable.Table.Id
				dataHandled.table_pos = dataHandled.parentTable.GetPosInParent()
			} else {
				dataHandled.pos = oControl.GetPosInParent()
				if (dataHandled.pos < startTablePos) {
					templist.push({
						content: [oControl]
					})
					tables[iTable] = {
						cells: [{
							icell: 0,
							W: newW
						}],
						W: newW
					}
					wBefore = newW
					var oControlParent = Api.LookupObject(oControl.Sdt.GetParent().Id)
					oControlParent.RemoveElement(dataHandled.pos)
				}
			}
			startTablePos = startTable.GetPosInParent()
			var startTableParent = startTable.Table.Parent
			var oStartTableParent = Api.LookupObject(startTableParent.Id)
			var elementCount1 = oStartTableParent.GetElementsCount()
			for (var i1 = startTablePos; i1 < elementCount1; ++i1) {
				var element = oStartTableParent.GetElement(i1)
				if (element.GetClassType() != 'table') {
					if (dataHandled.pos != undefined && dataHandled.pos == i1 && element.GetClassType() == 'blockLvlSdt' && element.Sdt.GetId() == oControl.Sdt.GetId()) {
						templist.push({
							content: [oControl]
						})
						if (wBefore + newW > 100) {
							iTable++
							tables[iTable] = {
								cells: [{
									icell: 0,
									W: newW
								}],
								W: newW
							}
							wBefore = newW
						} else {
							tables[iTable].cells.push({
								icell: tables[iTable].cells.length,
								W: newW
							})
							tables[iTable].W += newW
							wBefore += newW
						}
						var oControlParent = Api.LookupObject(oControl.Sdt.GetParent().Id)
						oControlParent.RemoveElement(dataHandled.pos)
					}
					break
				}
				if (element.GetTableTitle() != TABLE_TITLE) {
					break
				}
				var rows = element.GetRowsCount()
				if (rows != 1) {
					break
				}
				var oTableRow = element.GetRow(0)
				var cellsCount = oTableRow.GetCellsCount()
				for (var icell = 0; icell < cellsCount; ++icell) {
					var oCell = oTableRow.GetCell(icell)
					var W = oCell.CellPr.TableCellW.W
					var cellContent = oCell.GetContent()
					var contents = []
					var elementcount = cellContent.GetElementsCount()
					for (var k = 0; k < elementcount; ++k) {
						contents.push(cellContent.GetElement(k))
					}
					for (var k = 0; k < elementcount; ++k) {
						cellContent.RemoveElement(0)
					}
					templist.push({
						content: contents
					})
					if (dataHandled.table_id) {
						if (i1 < dataHandled.table_pos) { // 上一个表格
							if (!tables[iTable]) {
								tables[iTable] = {
									cells: [{
										icell: 0,
										W: W
									}],
									W: W
								}
							} else {
								tables[iTable].cells.push({
									icell: tables[iTable].cells.length,
									W: W
								})
								tables[iTable].W += W
							}
							wBefore += W
						} else {
							var W2 = W
							if (dataHandled.table_pos == i1 && dataHandled.cellIndex == icell) {
								W2 = newW
							}
							if (wBefore + W2 <= 100) {
								if (!tables[iTable]) {
									tables[iTable] = {
										cells: [{
											icell: 0,
											W: W2
										}],
										W: W2
									}
									wBefore = W2
								} else {
									tables[iTable].cells.push({
										icell: tables[iTable].cells.length,
										W: W2
									})
									tables[iTable].W += W2
									wBefore += W2
								}
							} else {
								iTable++
								tables[iTable] = {
									cells: [{
										icell: 0,
										W: W2
									}],
									W: W2
								}
								wBefore = W2
							}
						}
					} else { // 当前未处于表格中
						if (!tables[iTable]) {
							tables[iTable] = {
								cells: [{
									icell: 0,
									W: W
								}],
								W: W
							}
						} else {
							tables[iTable].cells.push({
								icell: tables[iTable].cells.length,
								W: W
							})
							tables[iTable].W += W
						}
						wBefore += W
					}
				}
				++toIndex
			}
			var count = 0
			for (var i = 0; i <= iTable; ++i) {
				if (!tables[i]) {
					break
				}
				var targetcellscount = tables[i].cells.length
				var oTable
				if (i <= toIndex) {
					oTable = oStartTableParent.GetElement(startTablePos + i)
					if (oTable.GetClassType() != 'table') {
						break
					}
					var oRow = oTable.GetRow(0)
					var cellscount = oRow.GetCellsCount()
					if (cellscount < targetcellscount) { // 需要拆分
						oTable.Split(oTable.GetCell(0, 0), 1, targetcellscount - cellscount + 1);
					} else if (cellscount > targetcellscount) { // 需要合并
						var mergeCells = []
						for (var m = 0; m <= cellscount - targetcellscount; ++m) {
							mergeCells.push(oTable.GetCell(0, m))
						}
						oTable.MergeCells(mergeCells)
					}
				} else {
					oTable  = Api.CreateTable(targetcellscount, 1)
					oTable.SetCellSpacing(0)
					oTable.SetTableTitle(TABLE_TITLE)
				}
				oTable.SetWidth('percent', tables[i].W)
				tables[i].cells.forEach((cell, cidx) => {
					AddControlToCell2(templist[count].content, oTable.GetCell(0, cell.icell))
					count++
					var twips = oTable.Table.CalculatedTableW * (100 / cell.W)
					oTable.GetCell(0, cell.icell).SetWidth('twips', twips)
					// oTable.GetCell(0, cell.icell).SetWidth('percent', cell.W)
				})
				if (i >= toIndex) {
					oStartTableParent.AddElement(startTablePos + i, oTable)
				}
			}
			if (iTable < toIndex) {
				for (var j = 0; j < toIndex - iTable; ++j) {
					oStartTableParent.RemoveElement(startTablePos + iTable + j + 1)
				}
			}
		}

		for (var idx = 0; idx < idList.length; ++idx) {
			var id = idList[idx]
			var quesData = question_map[id]
			if (!quesData) {
				continue
			}
			var oControl = oControls.find(e => {
				var tag = JSON.parse(e.GetTag() || '{}')
				return tag.client_id == id
			})
			if (!oControl) {
				continue
			}
			var oTable = oControl.GetParentTable()
			var posinparent = oControl.GetPosInParent()
			var parent = oControl.Sdt.GetParent()
			var oParent = Api.LookupObject(parent.Id)
			var preTable = null
			if (oParent.GetClassType() == 'document') {
				if (posinparent) {
					var previousElement = oParent.GetElement(posinparent - 1)
					if (previousElement && previousElement.GetClassType() == 'table' && previousElement.GetTableTitle() == TABLE_TITLE) {
						var TableW = previousElement.TablePr.TableW.W
						if (TableW + newW <= 100) {
							preTable = previousElement
						}
					}
				}
			} else if (oParent.GetClassType() == 'documentContent') {
				var contentParent = Api.LookupObject(oParent.Document.GetParent().Id)
				if (contentParent && contentParent.GetClassType() == 'tableCell') {
					var tablepos = oTable.GetPosInParent()
					if (tablepos) {
						var tableParent = Api.LookupObject(oTable.Table.Parent.Id)
						var pre = tableParent.GetElement(tablepos - 1)
						if (pre.GetClassType() == 'table' && pre.GetTableTitle() == TABLE_TITLE) {
							var TableW = pre.TablePr.TableW.W
							if(TableW + newW <= 100 ) {
								preTable = pre
							}
						}
					}
				}
			}
			if (preTable) {
				func1(preTable, oControl)
			} else if (oTable) {
				func1(oTable, oControl)
			} else {
				var addToNextTable = false
				var nextTable = oParent.GetElement(posinparent + 1)
				if (nextTable && nextTable.GetClassType() == 'table' && nextTable.GetTableTitle() == TABLE_TITLE) {
					var firstCell = nextTable.GetRow(0).GetCell(0)
					var firstCellW = firstCell.CellPr.TableCellW.W
					if (firstCellW + newW <= 100) {
						addToNextTable = true
						func1(nextTable, oControl)
					}
				}
				if (!addToNextTable) {
					oTable = Api.CreateTable(1, 1);
					oTable.SetCellSpacing(0)
					oTable.SetWidth('percent', newW);
					oTable.SetTableTitle(TABLE_TITLE)
					var oCell = oTable.GetRow(0).GetCell(0)
					AddControlToCell(oControl, oCell)
					oParent.AddElement(posinparent, oTable)
				}
			}
		}
		return {
			idList: idList,
			proportion: target_proportion
		}
	}, false, true).then(res => {
		console.log('changeProportion result', res)
		if (res && res.idList && window.BiyueCustomData.question_map) {
			res.idList.forEach(id => {
				if (window.BiyueCustomData.question_map[id]) {
					window.BiyueCustomData.question_map[id].proportion = res.proportion
				}
			})
		}
	})
}

function deleteAsks(askList) {
	if (!askList || askList.length == 0) {
		return
	}
	Asc.scope.question_map = window.BiyueCustomData.question_map
	Asc.scope.node_list = window.BiyueCustomData.node_list
	Asc.scope.delete_ask_list = askList
	return biyueCallCommand(window, function() {
		var node_list = Asc.scope.node_list
		var question_map = Asc.scope.question_map
		var delete_ask_list = Asc.scope.delete_ask_list
		var oDocument = Api.GetDocument()
		var drawings = oDocument.GetAllDrawingObjects()
		function clearQuesInteraction(oControl) {
			if (!oControl) {
				return
			}
			if (oControl.GetClassType() == 'inlineLvlSdt') {
				var elementCount = oControl.GetElementsCount()
				for (var idx = 0; idx < elementCount; ++idx) {
					var oRun = oControl.GetElement(idx)
					if (oRun &&
						oRun.Run &&
						oRun.Run.Content &&
						oRun.Run.Content[0] &&
						oRun.Run.Content[0].docPr) {
						var title = oRun.Run.Content[0].docPr.title || "{}"
						var titleObj = JSON.parse(title)
						if (titleObj.feature && titleObj.feature.sub_type == 'ask_accurate') {
							oRun.Delete()
							break
						}
					}
				}
			} else {
				var drawings = oControlContent.GetAllDrawingObjects()
				if (drawings) {
					for (var j = 0, jmax = drawings.length; j < jmax; ++j) {
						var oDrawing = drawings[j]
						if (oDrawing.Drawing.docPr) {
							var title = oDrawing.Drawing.docPr.title || "{}"
							var titleObj = JSON.parse(title)
							if (titleObj.feature && titleObj.feature.zone_type == 'question') {
								oDrawing.Delete()
							}
						}
					}
				}
			}
		}
		function removeCellInteraction(oCell) {
			if (!oCell) {
				return
			}
			oCell.SetBackgroundColor(204, 255, 255, true)
			var cellContent = oCell.GetContent()
			var paragraphs = cellContent.GetAllParagraphs()
			paragraphs.forEach(oParagraph => {
				var childCount = oParagraph.GetElementsCount()
				for (var i = 0; i < childCount; ++i) {
					var oRun = oParagraph.GetElement(i)
					if (oRun &&
						oRun.Run &&
						oRun.Run.Content &&
						oRun.Run.Content[0] &&
						oRun.Run.Content[0].docPr) {
						var title = oRun.Run.Content[0].docPr.title || '{}'
						var titleObj = JSON.parse(title)
						if (titleObj.feature && titleObj.feature.sub_type == 'ask_accurate') {
							oRun.Delete()
							break
						}
					}
				}
			})
		}
		for (var i = 0, imax = delete_ask_list.length; i < imax; ++i) {
			var qid = delete_ask_list[i].ques_id
			var aid = delete_ask_list[i].ask_id
			if (question_map[qid]) {
				var askIndex = question_map[qid].ask_list.findIndex(e => e.id == aid)
				if (askIndex >= 0) {
					question_map[qid].ask_list.splice(askIndex, 1)
				}
			}
			var nodeData = node_list.find(e => {
				return e.id == qid
			})
			if (!nodeData || !nodeData.write_list) {
				continue
			}
			var writeIndex = nodeData.write_list.findIndex(e => {
				return e.id == aid
			})
			if (writeIndex == -1) {
				continue
			}
			var writeData = nodeData.write_list[writeIndex]
			if (!writeData) {
				continue
			}
			nodeData.write_list.splice(writeIndex, 1)
			if (writeData.sub_type == 'control') {
				var oControl = Api.LookupObject(writeData.control_id)
				clearQuesInteraction(oControl)
				Api.asc_RemoveContentControlWrapper(writeData.control_id)
			} else if (writeData.sub_type == 'cell') {
				var oCell = Api.LookupObject(writeData.cell_id)
				removeCellInteraction(oCell)
			} else if (writeData.sub_type == 'write' || writeData.sub_type == 'identify') {
				if (writeData.shape_id) {
					var oDrawing = drawings.find(e => {
						return writeData.drawing_id == e.Drawing.Id
					})
					if (oDrawing) {
						if (writeData.sub_type == 'identify') {
							oDrawing.Delete()
						} else {
							var run = oDrawing.Drawing.GetRun()
							if (run) {
								var paragraph = run.GetParagraph()
								if (paragraph) {
									var oParagraph = Api.LookupObject(paragraph.Id)
									var ipos = run.GetPosInParent()
									if (ipos >= 0) {
										oDrawing.Delete()
										oParagraph.RemoveElement(ipos)
									}
								}
							}
						}	
					}
				}
			}
		}
		return {
			question_map: question_map,
			node_list: node_list,
			ques_id: delete_ask_list[delete_ask_list.length - 1].ques_id
		}
	}, false, true).then(res => {
		if (res) {
			window.BiyueCustomData.question_map = res.question_map
			window.BiyueCustomData.node_list = res.node_list
			document.dispatchEvent(new CustomEvent('updateQuesData', {
				detail: {
					client_id: res.ques_id
				}
			}))
		}
	})
}

function focusAsk(writeData) {
	if (!writeData) {
		return
	}
	Asc.scope.write_data = writeData
	return biyueCallCommand(window, function() {
		var write_data = Asc.scope.write_data
		var oDocument = Api.GetDocument()
		var drawings = oDocument.GetAllDrawingObjects() || []
		if (write_data.sub_type == 'control') {
			if (write_data.control_id) {
				var oControl = Api.LookupObject(write_data.control_id)
				if (oControl) {
					oDocument.Document.MoveCursorToContentControl(write_data.control_id, true)
				}
			}
		} else if (write_data.sub_type == 'cell') {
			if (write_data.cell_id) {
				var oCell = Api.LookupObject(write_data.cell_id)
				if (oCell) {
					var cellContent = oCell.GetContent()
					if (cellContent) {
						cellContent.GetRange().Select()
					}
				}
			}
		} else if (write_data.sub_type == 'write' || write_data.sub_type == 'identify') {
			if (write_data.shape_id) {
				var oDrawing = drawings.find(e => {
					return e.Drawing.Id == write_data.drawing_id
				})
				if (oDrawing) {
					oDrawing.Select()
				}
			}
		}
	}, false, false)
}

export {
	handleDocClick,
	handleContextMenuShow,
	initExamTree,
	updateRangeControlType,
	reqGetQuestionType,
	reqUploadTree,
	batchChangeQuesType,
	batchChangeInteraction,
	batchChangeProportion,
	splitEnd,
	showLevelSetDialog,
	confirmLevelSet,
	initControls,
	handleWrite,
	handleIdentifyBox,
	handleAllWrite,
	changeProportion,
	deleteAsks,
	focusAsk,
	showAskCells
}