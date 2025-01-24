import { getNodeList, handleChangeType, handleChangeType} from './QuesManager.js'
import { setInteraction } from './featureManager.js'
// 一些旧的函数，在这里进行备份，等版本稳定后，直接删除
// 划分类型
function updateRangeControlType(typeName) {
	Asc.scope.typename = typeName
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	Asc.scope.question_map = window.BiyueCustomData.question_map
	Asc.scope.node_list = window.BiyueCustomData.node_list
	console.log('updateRangeControlType begin:', typeName)
	return biyueCallCommand(window, function() {
		var typeName = Asc.scope.typename
		var oDocument = Api.GetDocument()
		var oRange = oDocument.GetRangeBySelect()
		var client_node_id = Asc.scope.client_node_id
		var question_map = Asc.scope.question_map || {}
		var node_list = Asc.scope.node_list || []
		var change_list = []
		var result = {
			client_node_id: client_node_id,
			change_list: change_list,
			typeName: typeName
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
			} else if (oControl.GetClassType() == 'blockLvlSdt') {
				return oControl.GetParentContentControl()
			}
			return null
		}
		function getParentId(oControl) {
			var parentBlock = getParentBlock(oControl)
			var parent_id = 0
			if (parentBlock) {
				var parentTag = Api.ParseJSON(parentBlock.GetTag())
				parent_id = parentTag.client_id || 0
			}
			return parent_id
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
		function getChildControls(oControl) {
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
				return null
			}
			var childControls = oControl.GetAllContentControls()
			if (!childControls) {
				return null
			}
			var children = []
			for (var i = 0; i < childControls.length; ++i) {
				var childControl = childControls[i]
				var tag = Api.ParseJSON(childControl.GetTag())
				if (!tag.client_id) {
					continue
				}
				var added = true
				if (childControl.GetClassType() == 'inlineLvlSdt') {
					var parentBlock = getParentBlock(childControl)
					if (!parentBlock || parentBlock.Sdt.GetId() != oControl.Sdt.GetId()) {
						added = false
					}
				}
				if (added) {
					children.push({
						id: tag.client_id,
						regionType: tag.regionType,
						control_id: childControl.Sdt.GetId(),
						sub_type: 'control'
					})
				}
			}
			return children
		}

		function inRun(oRun, drawingId) {
			if (!oRun || !oRun.Run) {
				return false
			}
			var cnt = oRun.Run.GetElementsCount()
			for (var i = 0; i < cnt; ++i) {
				var oChild = oRun.Run.GetElement(i)
				if (oChild && oChild.Id == drawingId) {
					return true
				}
			}
			return false
		}

		function getDirectParentCell(oDrawing) {
			var drawingParentParagraph = oDrawing.GetParentParagraph()
			if (drawingParentParagraph) {
				var pcount = drawingParentParagraph.GetElementsCount()
				for (var i = 0; i < pcount; ++i) {
					var oChild = drawingParentParagraph.GetElement(i)
					if (oChild && oChild.GetClassType) {
						var childType = oChild.GetClassType()
						if (childType == 'run') {
							if (inRun(oChild, oDrawing.Drawing.Id)) {
								break
							}
						} else if (childType == 'inlineLvlSdt') {
							var cnt3 = oChild.GetElementsCount()
							for (var i3 = 0; i3 < cnt3; ++i3) {
								var oChild3 = oChild.GetElement(i3)
								if (oChild3 && oChild3.GetClassType() == 'run') {
									if (inRun(oChild3, oDrawing.Drawing.Id)) {
										return null
									}
								}
							}
						}
					}
				}
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
		function deleteAccurateRun(oRun) {
			if (oRun &&
				oRun.Run &&
				oRun.Run.Content &&
				oRun.Run.Content[0] &&
				oRun.Run.Content[0].docPr) {
				var title = oRun.Run.Content[0].docPr.title
				if (title) {
					var titleObj = Api.ParseJSON(title)
					if (titleObj.feature && titleObj.feature.sub_type == 'ask_accurate') {
						oRun.Delete()
						return true
					}
				}
			}
			return false
		}
		function delSimpleControl(oControl) {
			var parent = oControl.Sdt.Parent
			if (parent && parent.GetType() == 1) {
				var oParent = Api.LookupObject(parent.Id)
				if (oParent) {
					var pos = oControl.Sdt.GetPosInParent()
					if (pos >= 0) {
						oParent.RemoveElement(pos)
					}
				}
			}
		}
		function hideSimpleControl(oControl) {
			var childControls = oControl.GetAllContentControls() || []
			if (childControls.length) {
				for (var c = childControls.length - 1; c >= 0; --c) {
					var tag = Api.ParseJSON(childControls[c].GetTag())
					if (tag.regionType == 'num' && childControls[c].GetClassType() == 'inlineLvlSdt') {
						delSimpleControl(childControls[c])			
					}
				}
			}
		}
		function hideSimple(oParagraph) {
			if (!oParagraph) {
				return
			}
			var oNumberingLevel = oParagraph.GetNumbering()
			if (!oNumberingLevel) {
				return
			}
			var level = oNumberingLevel.Lvl
			var oNum = oNumberingLevel.Num
			var oNumberingLvl = oNum.GetLvl(level)
			if (!oNumberingLvl) {
				return
			}
			var LvlText = oNumberingLvl.LvlText || []
			if (LvlText && LvlText.length) {
				if (LvlText[0].Value!='\ue749') {
					return
				}
			}
			var suffix = ''
			if (LvlText.length > 1 && LvlText[LvlText.length - 1].Type == 1) {
				suffix = LvlText[LvlText.length - 1].Value
			}
			var sType = Api.GetFormatTypeString(oNumberingLvl.Format)
			oNumberingLevel.SetTemplateType('bullet', '');
			var str = ''
			var find = false
			for (var i = 0; i < LvlText.length; ++i) {
				if (LvlText[i].Type == 2) {
					if (LvlText[i].Value == level) {
						str += `%${level+1}`
					}
				} else {
					if (LvlText[i].Value == ' ') {
						if (find) {
							str += LvlText[i].Value
						}
					} else if (LvlText[i].Value != '\ue749') {
						str += LvlText[i].Value
						find = true
					}
				}
			}
			oNumberingLevel.SetCustomType(sType, str, "left")
			var oTextPr = oNumberingLevel.GetTextPr();
			oTextPr.SetFontFamily("iconfont");
		}
		// 删除题目互动
		function clearQuesInteraction(oControl) {
			if (!oControl) {
				return
			}
			var tag = Api.ParseJSON(oControl.GetTag() || '{}')
			if (tag.regionType == 'write') {
				if (oControl.GetClassType() == 'inlineLvlSdt') {
					var elementCount = oControl.GetElementsCount()
					for (var idx = 0; idx < elementCount; ++idx) {
						var oRun = oControl.GetElement(idx)
						if (deleteAccurateRun(oRun)) {
							break
						}
					}
				}
			} else if (tag.regionType == 'num') {
				delSimpleControl(oControl)
			}
			else if (tag.regionType =='question') {
				if (oControl.GetClassType() == 'blockLvlSdt') {
					hideSimpleControl(oControl)
					var oParagraph = oControl.GetAllParagraphs()[0]
					hideSimple(oParagraph)
					var oControlContent = oControl.GetContent()
					var drawings = oControlContent.GetAllDrawingObjects()
					if (drawings) {
						for (var j = 0, jmax = drawings.length; j < jmax; ++j) {
							var oDrawing = drawings[j]
							if (oDrawing.Drawing.docPr) {
								var title = oDrawing.Drawing.docPr.title
								if (title && title.indexOf('feature') >= 0) {
									var titleObj = Api.ParseJSON(title)
									if (titleObj.feature && titleObj.feature.zone_type == 'question') {
										if (titleObj.feature.sub_type == 'ask_accurate') {
											var cellParent = getDirectParentCell(oDrawing)
											if (cellParent) {
												removeCellInteraction(cellParent)
											} else {
												oDrawing.Delete()
											}
										} else {
											deleShape(oDrawing)
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
			if (!oCell || !oCell.GetClassType || oCell.GetClassType() != 'tableCell') {
				return
			}
			oCell.SetBackgroundColor(255, 191, 191, true)
			var cellContent = oCell.GetContent()
			var paragraphs = cellContent.GetAllParagraphs()
			paragraphs.forEach(oParagraph => {
				var childCount = oParagraph.GetElementsCount()
				for (var i = 0; i < childCount; ++i) {
					var oRun = oParagraph.GetElement(i)
					if (deleteAccurateRun(oRun)) {
						break
					}
				}
			})
			var oTable = oCell.GetParentTable()
			if (oTable && oTable.GetPosInParent() >= 0) {
				var desc = Api.ParseJSON(oTable.GetTableDescription())
				desc.biyue = 1
				var key = `${oCell.GetRowIndex()}_${oCell.GetIndex()}`
				if (desc[key]) {
					delete desc[key]
				}
				oTable.SetTableDescription(JSON.stringify(desc))
			}
		}
		function removeControl(oRemove) {
			if (!oRemove) {
				return
			}
			var tagRemove = Api.ParseJSON(oRemove.GetTag() || '{}')
			clearQuesInteraction(oRemove)
			result.change_list.push({
				control_id: oRemove.Sdt.GetId(),
				client_id: tagRemove.client_id,
				parent_id: getParentId(oRemove),
				regionType: tagRemove.regionType,
				type: 'remove'
			})
			oRemove.Sdt.GetLogicDocument().PreventPreDelete = true
			Api.asc_RemoveContentControlWrapper(oRemove.Sdt.GetId())
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
					var container_type = classType
					var container_id = null
					var container = oElement
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
							container = oParent
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
						container_id: container_id,
						container: container
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
			if (!oCell || !oCell.GetClassType || oCell.GetClassType() != 'tableCell') {
				return
			}
			var canadd = false
			var cellContent = oCell.GetContent()
			if (typeName == 'write') { // 划分为小问，需要将选中的单元格都设置为小问
				// 需要确保所选单元格里没有control，且自己处于某个control里
				if (cellNotControl(cellContent)) {
					oCell.SetBackgroundColor(255, 191, 191, false)
					canadd = true
				}
			} else if (typeName == 'clear' || typeName == 'clearAll') {
				removeCellInteraction(oCell)
				canadd = true
			}
			if (canadd) {
				result.change_list.push({
					parent_id: parent_id,
					table_id: table_id,
					row_index: oCell.GetRowIndex(),
					cell_index: oCell.GetIndex(),
					cell_id: oCell.Cell.Id,
					client_id: 'c_' + oCell.Cell.Id,
					regionType: 'write'
				})
				var oTable = Api.LookupObject(table_id)
				var desc = Api.ParseJSON(oTable.GetTableDescription())
				desc[`${oCell.GetRowIndex()}_${oCell.GetIndex()}`] = `c_${oCell.Cell.Id}`
				desc.biyue = 1
				oTable.SetTableDescription(JSON.stringify(desc))
			}
		}

		function getFirstElement(list, index) {
			if (!list) {
				return {
					index: -1
				}
			}
			for (var i = index; i >= 0; --i) {
				if (list[i].container_type == 'blockLvlSdt' || list[i].container_type == 'tableCell' || list[i].container_type == 'inlineLvlSdt') {
					return {
						index: i,
						container: list[i].container,
						container_type: list[i].container_type
					}
				}
			}
			return {}
		}
		function getCellNode(cellClientId) {
			for (var i = 0, imax = node_list.length; i < imax; ++i) {
				var nodeData = node_list[i]
				if (nodeData.level_type == 'question' && nodeData.write_list) {
					var writeData = nodeData.write_list.find(e => {
						return e.sub_type == 'cell' && e.id == cellClientId
					})
					if (writeData) {
						if (question_map[nodeData.id] && question_map[nodeData.id].ask_list) {
							var askIndex = question_map[nodeData.id].ask_list.findIndex(e => {
								return e.id == writeData.id
							})
							if (askIndex >= 0) {
								return nodeData
							}
						}
						break
					}
				}
			}
			return null
		}
		// 判定单元格是小问
		function getCellWriteQuestion(cellId) {
			var oCell = Api.LookupObject(cellId)
			if (!oCell || oCell.GetClassType() != 'tableCell') {
				return null
			}
			var colorAsk = false
			var shd = oCell.Cell.Get_Shd()
			if (shd) {
				var fill = shd.Fill
				if (fill && fill.r == 255 && fill.g == 191 && fill.b == 191) {
					colorAsk = true
				}
			}
			var node = getCellNode(`c_${cellId}`)
			var clientId = node ? `c_${cellId}` : null
			if (!node && colorAsk) {
				var oTable = oCell.GetParentTable()
				var desc = Api.ParseJSON(oTable.GetTableDescription())
				var rowIndex = oCell.GetRowIndex()
				var cIndex = oCell.GetIndex()
				if (desc[`${rowIndex}_${cIndex}`]) {
					var rClientId = desc[`${rowIndex}_${cIndex}`]
					node = getCellNode(rClientId)
					if (node) {
						clientId = rClientId
					}
				}
			}
			return {
				colorAsk: colorAsk,
				clientId: clientId,
				node: node
			}
		}
		function getQuestion(list, index) {
			for (var i = index - 1; i >= 0; --i) {
				var oControl = null
				if (list[i].container_type) {
					if (list[i].container_type == 'blockLvlSdt') {
						oControl = list[i].container
					}
				} else if (list[i].GetClassType) {
					if (list[i].GetClassType() == 'blockLvlSdt') {
						oControl = list[i]
					}
				}
				if (!oControl) {
					continue
				}
				var tag = Api.ParseJSON(oControl.GetTag() || '{}')
				if (tag.client_id && question_map[tag.client_id] && question_map[tag.client_id].level_type == 'question') {
					return {
						id: tag.client_id,
						quesControl: oControl
					}
				}
			}
			return null
		}
		function updateControlTag(oControl, regionType, parent_id) {
			if (!oControl) {
				return
			}
			var tag = Api.ParseJSON(oControl.GetTag() || '{}')
			var obj = {}
			if (!tag.client_id) {
				// 之前没有配置client_id，需要分配
				result.client_node_id += 1
				tag.client_id = result.client_node_id
				if (!tag.regionType) { // 之前没有配置regionType
					tag.regionType = regionType
				}
			}
			obj.client_id = tag.client_id
			obj.parent_id = parent_id
			obj.control_id = oControl.Sdt.GetId()
			obj.regionType = regionType
			obj.numbing_text = GetNumberingValue(oControl)
			if (regionType == 'write') {
				obj.sub_type = 'control'
				tag.color = '#ff000040'
			} else {
				if (typeName == 'setBig') {
					tag.big = 1
				} else {
					if (tag.big) {
						delete tag.big
					}
				}
				if (typeName == 'question' || typeName == 'setBig' || typeName == 'clearBig') {
					tag.color = '#d9d9d940'
				} else if (typeName == 'struct') {
					tag.color = '#CFF4FF80'
				} else if (tag.color) {
					delete tag.color
				}
				obj.text = oControl.GetRange().GetText()
			}
			oControl.SetTag(JSON.stringify(tag));
			result.change_list.push(obj)
		}
		function setBig(oControl) {
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
				return null
			}
			var posinparent = oControl.GetPosInParent()
			var oParent = Api.LookupObject(oControl.Sdt.GetParent().Id)
			// 需要将后面的级别比他小的控件挪到它的范围内
			var templist = []
			var parentElementCount = oParent.GetElementsCount()
			var tag = Api.ParseJSON(oControl.GetTag() || '{}')
			for (var i = posinparent + 1; i < parentElementCount; ++i) {
				var element = oParent.GetElement(i)
				if (!element) {
					break
				}
				if (!element.GetClassType) {
					break
				}
				if (element.GetClassType() == 'blockLvlSdt') {
					element.Sdt.GetLogicDocument().PreventPreDelete = true
					var nextTag = Api.ParseJSON(element.GetTag() || '{}')
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
				}
				else if (element.GetClassType() == 'paragraph') {
					break
					// var NumPr = element.GetNumbering()
					// if (NumPr && NumPr.Lvl <= tag.lvl) {
					// 	break
					// }
					// if (!NumPr) {
					// 	break
					// }
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
		}
		function clearBig(oControl) {
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
				return
			}
			var posinparent = oControl.GetPosInParent()
			var oParent = Api.LookupObject(oControl.Sdt.GetParent().Id)
			// 需要将包含的子控件挪出它的范围内
			var templist = []
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
						for (var r = 0; r < rows; ++r) {
							var oRow = element.GetRow(r)
							var cnt = oRow.GetCellsCount()
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
		function removeTableAsk(quesControl) {
			// 删除所有单元格小问
			if (!quesControl || quesControl.GetClassType() != 'blockLvlSdt') {
				return
			}
			var pageCount = quesControl.Sdt.getPageCount()
			for (var p = 0; p < pageCount; ++p) {
				var page = quesControl.Sdt.GetAbsolutePage(p)
				var tables = quesControl.GetAllTablesOnPage(page)
				if (tables) {
					for (var t = 0; t < tables.length; ++t) {
						var oTable = tables[t]
						var rowcount = oTable.GetRowsCount()
						var find = false
						for (var r = 0; r < rowcount; ++r) {
							var oRow = oTable.GetRow(r)
							var cellcount = oRow.GetCellsCount()
							for (var c = 0; c < cellcount; ++c) {
								var oCell = oRow.GetCell(c)
								if (!oCell) {
									continue
								}
								var shd = oCell.Cell.Get_Shd()
								if (shd) {
									var fill = shd.Fill
									if (fill && fill.r == 255 && fill.g == 191 && fill.b == 191) {
										oCell.SetBackgroundColor(255, 191, 191, true)
										find = true
									}
								}
							}
						}
						if (find) {
							oTable.SetTableDescription('{}')
						}
					}
				}
			}
		}
		function deleShape(oShape) {
			if (!oShape) {
				return
			}
			var run = oShape.Drawing.GetRun()
			if (run) {
				var runParent = run.GetParent()
				if (runParent) {
					var oParent = Api.LookupObject(runParent.Id)
					if (oParent && oParent.GetClassType() == 'inlineLvlSdt') {
						var count = oParent.GetElementsCount()
						for (var c = 0; c < count; ++c) {
							var child = oParent.GetElement(c)
							if (child && child.GetClassType() == 'run' && child.Run.Id == run.Id) {
								deleteAccurateRun(child)
								break
							}
						}
						return true
					}
				}
				var paragraph = run.GetParagraph()
				if (paragraph) {
					var oParagraph = Api.LookupObject(paragraph.Id)
					var ipos = run.GetPosInParent()
					if (ipos >= 0) {
						oShape.Delete()
						if (oParagraph && run.GetElementsCount() == 0) {
							oParagraph.RemoveElement(ipos)
						}
						return true
					}
				}
			}
			oShape.Delete()
			return true
		}
		function getQuesByAskId(askId) {
			var qids = Object.keys(question_map)
			for (var i = 0; i < qids.length; ++i) {
				var qid = qids[i]
				if (question_map[qid].ask_list) {
					var find = question_map[qid].ask_list.find(e => {
						return (e.id == askId) || (e.other_fields && e.other_fields.includes(askId))
					})
					if (find) {
						return {
							ques_id: qid,
							ask_ids: [find.id].concat(find.other_fields)
						}
					}
				}
			}
			return {
				ask_ids: []
			}
		}
		var allControls = oDocument.GetAllContentControls()
		var selectionInfo = oDocument.Document.getSelectionInfo()
		if (!oRange) {
			var elementData = getOElements(selectionInfo.curPos)
			console.log('elementList', elementData)
			var container = null
			var container_type = null
			var containerIndex = -1
			if (elementData.list) {
				var firstContainerData = getFirstElement(elementData.list, elementData.list.length - 1)
				if (firstContainerData.index >= 0) {
					container = firstContainerData.container
					container_type = firstContainerData.container_type
					containerIndex = firstContainerData.index
				} else if (elementData.list.length && elementData.list[0].container_type == 'shape' && elementData.list[0].container.Drawing && (typeName == 'clear' || typeName == 'clearAll')) {
					var dtitle = Api.ParseJSON(elementData.list[0].container.Drawing.docPr.title)
					if (dtitle.feature && dtitle.feature.client_id && (dtitle.feature.sub_type == 'identify' || dtitle.feature.sub_type == 'write')) {
						var adata = getQuesByAskId(dtitle.feature.client_id)
						deleShape(elementData.list[0].container)
						if (adata.ques_id) {
							result.change_list.push({
								shape_id: elementData.list[0].container.Shape.Id,
								client_id: dtitle.feature.client_id,
								parent_id: adata.ques_id,
								regionType: 'write',
								type: 'remove'
							})
						}
						return result
					}
				}
			}
			if (!container) {
				return {
					code: 0,
					message: '请先选中一个范围',
				}
			}
			switch (typeName) {
				case 'clear':
				case 'clearAll':
					{
						if (container_type == 'tableCell') {
							// 判断是否是小问，若是，删除小问，todo..
							var cellQuestion = getCellWriteQuestion(container.Cell.Id)
							if (cellQuestion && (cellQuestion.node || cellQuestion.colorAsk)) {
								removeCellInteraction(container)
								if (cellQuestion.node) {
									result.change_list.push({
										parent_id: cellQuestion.node.id,
										table_id: container.GetParentTable().Table.Id,
										cell_id: container.Cell.Id,
										client_id: cellQuestion.clientId,
										regionType: 'write'
									})
								}
							} else {
								var secondContainer = getFirstElement(elementData.list, containerIndex - 1)
								if (secondContainer && secondContainer.container_type == 'blockLvlSdt') {
									removeControl(secondContainer.container)
								}
							}
						} else {
							if (container_type == 'blockLvlSdt' && typeName == 'clearAll') {
								var childControls = container.GetAllContentControls() || []
								for (var i = 0; i < childControls.length; ++i) {
									removeControl(childControls[i])
								}
								removeTableAsk(container)
							}
							var removeIds = []
							if (container) {
								var tag2 = Api.ParseJSON(container.GetTag())
								if (tag2.regionType == 'write' && tag2.client_id) {
									removeIds = getQuesByAskId(tag2.client_id).ask_ids
								}
							}
							if (removeIds.length > 1) {
								var quesControl = getParentBlock(container)
								if (quesControl) {
									quesControl.GetAllContentControls().forEach(e => {
										var tag3 = Api.ParseJSON(e.GetTag())
										if (removeIds.includes(tag3.client_id)) {
											removeControl(e)
										}
									})
								}
							} else {
								removeControl(container)
							}
						}
					}
					break
				case 'write':
					{
						// 需要判断是否处于题目中，若不是，不可划分
						var pQuestion = getQuestion(elementData.list, containerIndex)
						if (!pQuestion) {
							return {
								code: 0,
								message: '未处于题目中，不可设为小问'
							}
						}
						if (container_type == 'tableCell') {
							// 设置单元格为小问
							setCellType(container, pQuestion.id, container.GetParentTable().Table.Id)
						} else {
							updateControlTag(container, 'write', pQuestion.id)
						}
					}
					break
				case 'question':
					{
						if (container_type == 'inlineLvlSdt') {
							// 判断是否处于单元格中，若是，将单元格内容选中设为题目
							var secondContainer = getFirstElement(elementData.list, containerIndex - 1)
							 // 处于单元格中
							if (secondContainer && secondContainer.container_type == 'tableCell') {
								var cellContent = secondContainer.container.GetContent()
								var cellRange = cellContent.GetRange()
								if (cellRange) {
									cellRange.Select()
									result.client_node_id += 1
									var tag = {
										regionType: 'question',
										mode: 1,
										column: 1,
										client_id: result.client_node_id,
										color: '#d9d9d940'
									}
									var oResult = Api.asc_AddContentControl(1, {
										Tag: JSON.stringify(tag)
									})
									if(oResult) {
										var oControl = Api.LookupObject(oResult.InternalId)
										result.change_list.push({
											client_id: tag.client_id,
											control_id: oResult.InternalId,
											text: cellRange.GetText(),
											children: [],
											parent_id: getParentId(oControl),
											regionType: 'question',
											numbing_text: GetNumberingValue(oControl)
										})
									} else {
										console.warn('添加control失败')
									}
								}
							} else {
								return {
									code: 0,
									message: '请先选中一个范围'
								}
							}
						} else if (container_type == 'blockLvlSdt') {
							updateControlTag(container, 'question', getParentId(container))
						} else if (container_type == 'tableCell') {
							// 暂不支持直接将单元格设置为题目
							return {
								code: 0,
								message: '暂不支持直接将单元格设置为题目'
							}
						}
					}
					break
				case 'struct': // 设为题组
				case 'setBig': // 设为大题
				case 'clearBig': // 清除大题
					{
						if (container_type != 'blockLvlSdt') {
							return {
								code: 0,
								message: '请先选中一个范围'
							}
						}
						if (typeName == 'setBig') {
							setBig(container)
						} else if (typeName == 'clearBig') {
							clearBig(container)
						} else if (typeName == 'struct') {
							// 需要清除小问及互动
							clearQuesInteraction(container)
							var childControls = container.GetAllContentControls() || []
							for (var i = 0; i < childControls.length; ++i) {
								Api.asc_RemoveContentControlWrapper(childControls[i].Sdt.GetId())
							}
						}
						updateControlTag(container, typeName, getParentId(container))
					}
					break
				default:
					break
			}
		} else {
			var paragrahs = oRange.GetAllParagraphs()
			if (!paragrahs || paragrahs.length === 0) {
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
			// 通过双击选中的范围，其实是inlinecontrol
			function isSelectInlineControl(control) {
				if (!startData || !startData.list || startData.list.length <= 2 || !endData || !endData.list || endData.list.length <= 2) {
					return false
				}
				// 判断startPos是否一致
				var startEndData = startData.list[startData.list.length - 1]
				var pre = startData.list[startData.list.length - 2]
				if (pre.classType == 'inlineLvlSdt') {
					if (pre.oElement.Sdt.GetId() != control.Sdt.GetId()) {
						return false
					}
					if (startEndData.Position != 0) {
						return false
					}
				} else if (pre.classType == 'paragraph') {
					if (startEndData.classType != 'run') {
						return false
					}
					var a = pre.oElement.GetElement(pre.Position + 1)
					if (a && a.GetClassType() != 'inlineLvlSdt' || a.Sdt.GetId() != control.Sdt.GetId()) {
						return false
					}
					var p = Math.min(startEndData.oElement.Run.GetElementsCount() - 1, startEndData.Position)
					if (p != 0) {
						return false
					}
				}
				// 判断endPos是否一致
				var endEndData = endData.list[endData.list.length - 1]
				var endPre = endData.list[endData.list.length - 2]
				if (endPre.classType == 'paragraph' && endPre.Position > 0 && endEndData.Position == 0) {
					var pre2 = endPre.oElement.GetElement(endPre.Position - 1)
					if (pre2 && pre2.GetClassType() == 'inlineLvlSdt' && pre2.Sdt.GetId() == control.Sdt.GetId()) {
						return true
					}
				} else if (endPre.classType == 'inlineLvlSdt' && endPre.Sdt.GetId() == control.Sdt.GetId()) {
					if (endEndData.Position == endPre.GetElementsCount() - 1) {
						return true
					}
				}
				return false
			}
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
				var oParentTag = Api.ParseJSON(oParentControl.GetTag() || '{}')
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
						if (isSelectInlineControl(e)) {
							completeOverlapControl = e
						} else {
							containControls.push(e)
						}
					} else if (relation == 4) {
						intersectControls.push(e)
					}
					controlsInRange.push(e)
				}

				if (typeName == 'clear') {
					if (completeOverlapControl) {
						removeControl(completeOverlapControl)
					} else if (containControls.length) {
						containControls.forEach(e => {
							removeControl(e)
						})
					} else if (parentControls.length) {
						removeControl(parentControls[parentControls.length - 1])
					}
				} else if (typeName == 'clearAll') {
					var controls = containControls.concat(intersectControls)
					if (completeOverlapControl) {
						controls.push(completeOverlapControl)
					}
					controls.forEach(e => {
						removeTableAsk(e)
						removeControl(e)
					})
				} else {
					// 若存在完全重叠的区域，只修改tag
					var isInCell = false
					var level = 0
					var rangeParagraphs = oRange.GetAllParagraphs()
					if (!completeOverlapControl && typeName != 'write' && rangeParagraphs.length) { // 未重叠
						var oParagraph = rangeParagraphs[0]
						var NumPr = rangeParagraphs[0].GetNumbering()
						if (NumPr) {
							level = NumPr.Lvl
						}
						var parent1 = oParagraph.Paragraph.GetParent()
						if (parent1) {
							var oParent1 = Api.LookupObject(parent1.Id)
							if (oParent1 && oParent1.GetClassType && oParent1.GetClassType() == 'documentContent') {
								var parent2 = parent1.GetParent()
								if (parent2) {
									var oParent2 = Api.LookupObject(parent2.Id)
									if (oParent2 && oParent2.GetClassType) {
										if (oParent2.GetClassType() == 'tableCell') {
											isInCell = true
										} else if (oParent2.GetClassType() == 'blockLvlSdt') {
											if (oParent2.GetAllParagraphs().length == rangeParagraphs.length) {
												completeOverlapControl = oParent2
											}
										}
									}
								}
							}
						}
					}
					if (completeOverlapControl) {
						if (typeName == 'write') {
							var elementData = getOElements(completeOverlapControl.GetRange().StartPos)
							var containerData = getFirstElement(elementData.list, elementData.list.length - 1)
							var pQuestion = getQuestion(elementData.list, containerData.index)
							if (!pQuestion) {
								return {
									code: 0,
									message: '未处于题目中，不可设为小问'
								}
							}
							updateControlTag(completeOverlapControl, 'write', pQuestion.id)
						} else if (completeOverlapControl.GetClassType() == 'blockLvlSdt') {
							if (typeName == 'setBig') {
								setBig(completeOverlapControl)
							} else if (typeName == 'clearBig') {
								clearBig(completeOverlapControl)
							} else if (typeName == 'struct') {
								// 需要清除小问及互动
								clearQuesInteraction(completeOverlapControl)
								var childControls = completeOverlapControl.GetAllContentControls() || []
								for (var i = 0; i < childControls.length; ++i) {
									Api.asc_RemoveContentControlWrapper(childControls[i].Sdt.GetId())
								}
							}
							updateControlTag(completeOverlapControl, typeName, getParentId(completeOverlapControl))
						}
					} else {
						// 不存在完全重叠的区域，新增一个control，对其下的原本的control并不进行拆分修改 todo..
						if (typeName == 'write') {
							if (controlsInRange.length == 0 || parentControls.length == 0) {
								return {
									code: 0,
									message: '未处于题目中',
								}
							} else if (parentControls.length > 0) {
								var pQuestion = getQuestion(parentControls, parentControls.length)
								if (!pQuestion) {
									return {
										code: 0,
										message: '未处于题目中',
									}
								}
							}
						}
						var type = 1
						var rangeParagraphs = oRange.GetAllParagraphs() || []
						if (rangeParagraphs.length) {
							if (typeName == 'write') {
								var f = rangeParagraphs.find((e, index) => {
									if (!e.Paragraph.IsEmpty()) {
										var cnt = e.GetElementsCount()
										if (!e.Paragraph.Selection.Use) {
											return true
										}
										var min = Math.min(e.Paragraph.Selection.StartPos, e.Paragraph.Selection.EndPos)
										var max = Math.max(e.Paragraph.Selection.StartPos, e.Paragraph.Selection.EndPos)
										if (min != 0 || max != cnt) {
											return true
										}
									}
								})
								type = f ? 2 : 1
							}
						}
						result.client_node_id += 1
						var regionType = typeName == 'write' ? 'write' : 'question'
						var tag = {
							client_id: result.client_node_id,
							regionType: regionType,
							lvl: level
						}
						// 设置小问时，先移除包含或交叉的control
						if (type == 2) {
							controlsInRange.forEach(e => {
								if (e.GetClassType() == 'inlineLvlSdt') {
									removeControl(e)
								}
							})
						}
						if (typeName == 'write') {
							tag.color = '#ff000040'
						} else if (typeName == 'question') {
							tag.color = '#d9d9d940'
						} else if (typeName == 'struct') {
							tag.color = '#CFF4FF80'
						}
						var oResult = Api.asc_AddContentControl(type, {
							Tag: JSON.stringify(tag)
						})
						if(oResult) {
							var oControl = Api.LookupObject(oResult.InternalId)
							// 需要返回新增的nodeIndex todo..
							result.change_list.push({
								client_id: tag.client_id,
								control_id: oResult.InternalId,
								text: oRange.GetText(),
								children: oControl && oControl.GetClassType() == 'blockLvlSdt' ? getChildControls(oControl) : [],
								parent_id: getParentId(oControl),
								regionType: regionType,
								numbing_text: GetNumberingValue(oControl)
							})
							// 若是在单元格里添加control后，会多出一行需要删除
							if (type == 1 && isInCell) {
								var oCell = Api.LookupObject(oResult.InternalId).GetParentTableCell()
								if (oCell) {
									if (oCell.GetContent().GetElementsCount() == 2) {
										var oElement2 = oCell.GetContent().GetElement(1)
										if (oElement2 && oElement2.GetClassType() == 'paragraph' && oElement2.Paragraph.Bounds.Bottom == 0 && oElement2.Paragraph.Bounds.Top == 0) {
											oCell.GetContent().RemoveElement(1)
										}
									}
								}
							}
						} else {
							console.warn('添加control失败')
						}
					}
				}
			}
		}
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
		function deleteShape(oDrawing) {
			var run = oDrawing.Drawing.GetRun()
			if (run) {
				var paragraph = run.GetParagraph()
				if (paragraph) {
					var oParagraph = Api.LookupObject(paragraph.Id)
					var ipos = run.GetPosInParent()
					if (ipos >= 0) {
						var cnt = oParagraph.GetElementsCount()
						oDrawing.Delete()
						cnt = oParagraph.GetElementsCount()
						var element2 = oParagraph.GetElement(ipos)
						if (element2 && element2.GetClassType() == 'run' && element2.Run.Id == run.Id) {
							oParagraph.RemoveElement(ipos)
							cnt = oParagraph.GetElementsCount()
						}
						return
					}
				}
			}
			oDrawing.Delete()
		}
		var oControl = Api.LookupObject(curControl.Id)
		var tag = oControl ? Api.ParseJSON(oControl.GetTag()) : {}
		if (write_cmd == 'add') {
			if (oControl) {
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
					if (parentParagraph) {
						var oRun = Api.CreateRun()
						oRun.AddDrawing(oDrawing)
						parentParagraph.AddElement(
							oRun,
							1
						)
					}
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
			} else if (oControl) {
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
					var titleObj = Api.ParseJSON(title)
					if (titleObj.feature && titleObj.feature.zone_type == 'question' && titleObj.feature.sub_type == 'write') {
						var drawingId = drawings[i].Drawing.Id
						var oDrawing = allDrawings.find(e => {
							return e.Drawing.Id == drawingId
						})
						if (oDrawing) {
							result.remove_ids.push(titleObj.feature.client_id)
							deleteShape(oDrawing)
						}
					}
				}
				return result
			}
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
						var tag = Api.ParseJSON(paraentControl.GetTag())
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
										var dtitle = Api.ParseJSON(
											drawings[sidx].Drawing.docPr.title
										)
										if (dtitle.feature && dtitle.feature.sub_type == 'identify') {
											res.remove_ids.push(dtitle.feature.client_id)
											drawings[sidx].Delete()
										}
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
					score: 1
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
			if (question_map[res.ques_id].ask_list) {
				var sumscore = 0
				question_map[res.ques_id].ask_list.forEach(e => {
					sumscore += e.score
				})
				question_map[res.ques_id].score = sumscore
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

export {
	updateRangeControlType
}
