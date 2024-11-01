// 划分类型，处理结构，题目，小问的增删维护
import { biyueCallCommand, dispatchCommandResult } from "./command.js";
import { getNodeList, handleChangeType } from './QuesManager.js'
function handleRangeType(options) {
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	Asc.scope.node_list = window.BiyueCustomData.node_list
	Asc.scope.question_map = window.BiyueCustomData.question_map
	Asc.scope.options = options
	return biyueCallCommand(window, function() {
		var client_node_id = Asc.scope.client_node_id || 0
		var node_list = Asc.scope.node_list || []
		var question_map = Asc.scope.question_map || {}
		var options = Asc.scope.options || {}
		var oDocument = Api.GetDocument()
		var allDrawings = oDocument.GetAllDrawingObjects()
		var selectionInfo = oDocument.Document.getSelectionInfo()
		var result = {
			change_list: [],
			client_node_id: client_node_id
		}
		var typeName = options.typeName
		// function list begin
		function addDrawingRemove(titleObj) {
			if (!titleObj.feature) {
				return
			}
			if (titleObj.feature.sub_type == 'identify' || titleObj.feature.sub_type == 'write') {
				result.change_list.push({
					client_id: titleObj.feature.client_id,
					parent_id: titleObj.feature.parent_id,
					regionType: 'write',
					type: 'remove'
				})
			}
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
			var paragraphs = oControl.GetAllParagraphs() || []
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
			if (!drawingParentParagraph) {
				return null
			}
			var pcount = drawingParentParagraph.GetElementsCount()
			for (var i = 0; i < pcount; ++i) {
				var oChild = drawingParentParagraph.GetElement(i)
				if (oChild.GetClassType) {
					var childType = oChild.GetClassType()
					if (childType == 'run') {
						if (inRun(oChild, oDrawing.Drawing.Id)) {
							break
						}
					} else if (childType == 'inlineLvlSdt') {
						var cnt3 = oChild.GetElementsCount()
						for (var i3 = 0; i3 < cnt3; ++i3) {
							var oChild3 = oChild.GetElement(i3)
							if (oChild3.GetClassType() == 'run') {
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
			return null
		}
		function deleteDrawingRun(oRun, sub_type) {
			if (oRun &&
				oRun.Run &&
				oRun.Run.Content &&
				oRun.Run.Content[0] &&
				oRun.Run.Content[0].docPr) {
				var title = oRun.Run.Content[0].docPr.title
				if (title) {
					var titleObj = Api.ParseJSON(title)
					if (titleObj.feature && titleObj.feature.sub_type == sub_type) {
						oRun.Delete()
						return true
					}
				}
			}
			return false
		}
		function delSimpleControl(oControl) {
			if (!oControl || !oControl.Sdt) {
				return
			}
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
		function clearQuesInteraction(oControl, excepetNum) {
			if (!oControl) {
				return
			}
			var tag = Api.ParseJSON(oControl.GetTag())
			if (tag.regionType == 'write') {
				if (oControl.GetClassType() == 'inlineLvlSdt') {
					var elementCount = oControl.GetElementsCount()
					for (var idx = 0; idx < elementCount; ++idx) {
						var oRun = oControl.GetElement(idx)
						deleteDrawingRun(oRun, 'ask_accurate')
					}
				}
			} else if (tag.regionType == 'num') {
				delSimpleControl(oControl)
			}
			else if (tag.regionType =='question') {
				if (oControl.GetClassType() == 'blockLvlSdt') {
					if (!excepetNum) {
						hideSimpleControl(oControl)
						var oParagraph = oControl.GetAllParagraphs()[0]
						hideSimple(oParagraph)
					}
					var oControlContent = oControl.GetContent()
					var drawings = oControlContent.GetAllDrawingObjects()
					if (drawings) {
						for (var j = 0, jmax = drawings.length; j < jmax; ++j) {
							var oDrawing = drawings[j]
							var titleObj = Api.ParseJSON(oDrawing.GetTitle())
							if (titleObj.feature && titleObj.feature.zone_type == 'question') {
								if (titleObj.feature.sub_type == 'ask_accurate') {
									var cellParent = getDirectParentCell(oDrawing)
									if (cellParent) {
										removeCellInteraction(cellParent)
									} else {
										oDrawing.Delete()
									}
								} else {
									addDrawingRemove(titleObj)
									deleteShape(oDrawing)
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
					if (deleteDrawingRun(oRun, 'ask_accurate')) {
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
		function deleteShape(oShape) {
			if (!oShape) {
				return
			}
			var paraDrawing = oShape.getParaDrawing()
			var run = paraDrawing ? paraDrawing.GetRun() : null
			if (run) {
				var runParent = run.GetParent()
				if (runParent) {
					var oParent = Api.LookupObject(runParent.Id)
					if (oParent && oParent.GetClassType() == 'inlineLvlSdt') {
						var count = oParent.GetElementsCount()
						for (var c = 0; c < count; ++c) {
							var child = oParent.GetElement(c)
							if (child.GetClassType() == 'run' && child.Run.Id == run.Id) {
								deleteDrawingRun(child, 'ask_accurate')
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
		
		function createDrawing(parent_id, sub_type) {
			var oFill = Api.CreateNoFill()
			var oStroke = Api.CreateStroke(3600, Api.CreateNoFill())
			var nWidth = 8
			var nHeight = 5
			var wrappingStyle = 'inline'
			var nBottom = 0
			if (sub_type == 'write') {
				oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 0, 0))
				oFill.UniFill.transparent = 255 * 0.2 // 透明度
				nWidth = 20
				nHeight = 10
				wrappingStyle = 'inFront'
			} else if (sub_type == 'identify') {
				oStroke = Api.CreateStroke(3600, Api.CreateSolidFill(Api.CreateRGBColor(125, 125, 125)))
				nBottom = 0.5 * 36e3				
			}
			var oDrawing = Api.CreateShape('rect', nWidth * 36e3, nHeight * 36e3, oFill, oStroke)
			oDrawing.SetWrappingStyle(wrappingStyle)
			result.client_node_id += 1
			var titleobj = {
				feature: {
					zone_type: 'question',
					sub_type: sub_type,
					parent_id: parent_id,
					client_id: result.client_node_id
				}
			}
			oDrawing.SetTitle(JSON.stringify(titleobj))
			oDrawing.SetPaddings(0, 0, 0, nBottom)
			if (sub_type == 'write') {
				oDrawing.SetDistances(0, 0, 2 * 36e3, 0);
			} else if (sub_type == 'identify') {
				var drawDocument = oDrawing.GetContent()
				var oParagraph = Api.CreateParagraph()
				oParagraph.AddText('×')
				oParagraph.SetColor(125, 125, 125, false)
				oParagraph.SetFontSize(24)
				oParagraph.SetFontFamily('黑体')
				oParagraph.SetJc('center')
				drawDocument.AddElement(0, oParagraph)
			}
			return oDrawing
		}
		function getFirstParagraph(oControl) {
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
				return null
			}
			var paragraphs = oControl.GetAllParagraphs() || []
			for (var i = 0; i < paragraphs.length; ++i) {
				var oParagraph = paragraphs[i]
				if (oParagraph) {
					var parent1 = oParagraph.Paragraph.Parent
					var parent2 = parent1.Parent
					if (parent2 && parent2.Id == oControl.Sdt.GetId()) {
						return oParagraph
					}
				}
			}
			return null
		}
		function addAskDrawing(parent_id, quesControl, sub_type, curPosInfo, paragraphIdx, runIdx) {
			var oDrawing = createDrawing(parent_id, sub_type)
			if (!oDrawing) {
				return null
			}
			if (sub_type == 'write') {
				var parentParagraph = getFirstParagraph(quesControl)
				if (parentParagraph) {
					var oRun = Api.CreateRun()
					oRun.AddDrawing(oDrawing)
					parentParagraph.AddElement(
						oRun,
						1
					)
					return oDrawing
				}
			} else if (sub_type == 'identify') {
				if (runIdx >= 0) {
					curPosInfo[runIdx].Class.Add_ToContent(
						curPosInfo[runIdx].Position,
						oDrawing.getParaDrawing()
					)
				} else {
					var oRun = Api.CreateRun()
					oRun.AddDrawing(oDrawing)
					var pParagraph = Api.LookupObject(curPosInfo[paragraphIdx].Class.Id)
					pParagraph.AddElement(
						oRun,
						curPosInfo[paragraphIdx].Position + 1
					)
				}				
			}
			return oDrawing
		}
		function deleteTypeDrawing(oDrawing, sub_type) {
			if (!oDrawing) {
				return
			}
			var titleObj = Api.ParseJSON(oDrawing.GetTitle())
			if (titleObj.feature && titleObj.feature.zone_type == 'question' && titleObj.feature.sub_type == sub_type) {
				addDrawingRemove(titleObj)
				if (sub_type == 'identify') {
					oDrawing.Delete()
				} else {
					deleteShape(oDrawing)
				}
			}
		}
		function getQuestion(quesControl) {
			if (!quesControl) {
				return false
			}
			if (quesControl.GetClassType() == 'inlineLvlSdt') {
				quesControl = getParentBlock(quesControl)
				if (!quesControl) {
					return false
				}
			}
			var quesTag = Api.ParseJSON(quesControl.GetTag())
			if (!quesTag.client_id) {
				return false
			}
			var quesData = question_map[quesTag.client_id]
			if (!quesData || quesData.level_type != 'question') { // 只有题目才能处理作答区
				return false
			}
			return quesTag
		}
		function addWrite(quesControl, sub_type, curPosInfo, runIdx, paragraphIdx) {
			var quesTag = getQuestion(quesControl)
			if (!quesTag) {
				result.message = '未处于题目中无法处理'
				return result
			}
			var shapeAdd = addAskDrawing(quesTag.client_id, quesControl, sub_type, curPosInfo, paragraphIdx, runIdx)
			if (shapeAdd) {
				result.change_list.push({
					client_id: result.client_node_id,
					parent_id: quesTag.client_id,
					regionType: 'write',
					sub_type: sub_type,
					shape_id: shapeAdd.Shape.Id,
					drawing_id: shapeAdd.Drawing.Id
				})
			}
		}
		function delShapeAsk(curBlockSdt, sub_type) {
			var drawings = []
			var selectedDrawings = oDocument.GetSelectedDrawings() // 返回的类型可能是ApiImage，ApiShape等等
			if (selectedDrawings && selectedDrawings.length) {
				for (var i = 0; i < selectedDrawings.length; ++i) {
					if (selectedDrawings[i].GetClassType() == 'shape') {
						drawings.push(selectedDrawings[i])
					}
				}
			} else if (curBlockSdt) {
				var oControl = Api.LookupObject(curBlockSdt.Id)
				drawings = oControl.GetAllDrawingObjects()
			}
			if (drawings && drawings.length) {
				for (var i = 0; i < drawings.length; ++i) {
					var drawingId = drawings[i].Drawing.Id
					var oDrawing = allDrawings.find(e => {
						return e.Drawing.Id == drawingId
					})
					if (oDrawing) {
						deleteTypeDrawing(oDrawing, sub_type)
					}
				}
			}
		}		
		
		function updateControlTag(oControl, regionType, parent_id) {
			if (!oControl) {
				return
			}
			var tag = Api.ParseJSON(oControl.GetTag())
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
		function getColor() {
			if (typeName == 'write') {
				return '#ff000040'
			} else if (typeName == 'struct') {
				return '#CFF4FF80'
			}
			return '#d9d9d940'
		}
		function addControl(type, tag) {
			result.client_node_id += 1
			var tag = {
				client_id: result.client_node_id,
				regionType: typeName == 'write' ? 'write': 'question',
				color: getColor(),
				mode: type == 1 ? 5: 3
			}
			var oResult = Api.asc_AddContentControl(type, {
				Tag: JSON.stringify(tag)
			})
			if(oResult) {
				var oControl = Api.LookupObject(oResult.InternalId)
				result.change_list.push({
					client_id: tag.client_id,
					control_id: oResult.InternalId,
					text: oControl.GetRange().GetText(),
					parent_id: getParentId(oControl),
					regionType: tag.regionType,
					numbing_text: GetNumberingValue(oControl)
				})
			}
		}
		function getParentQuestion(blockSdt) {
			if (!blockSdt) {
				return null
			}
			var tagObj = Api.ParseJSON(blockSdt.Pr.Tag)
			var quesData = null
			if (tagObj.client_id) {
				quesData = question_map[tagObj.client_id]
				if (quesData) {
					return quesData.level_type == 'question' ? tagObj.client_id : null
				}
			}
			var parentData = getParentData(Api.LookupObject(blockSdt.Id))
			if (parentData) {
				return parentData.level_type == 'question' ? parentData.id : null
			}
			return null
		}
		// 获取祖先节点信息
		function getParentData(oControl) {
			if (!oControl) {
				return null
			}
			var oParentControl = oControl.GetParentContentControl()
			if (!oParentControl) {
				return null
			}
			var tagObj = Api.ParseJSON(oParentControl.GetTag())
			if (tagObj.client_id) {
				return node_list.find(e => {
					return e.id == tagObj.client_id
				})
			}
			return getParentData(oParentControl)
		}
		function doRemoveControl(childControl, tagRemove) {
			if (!childControl) {
				return
			}
			result.change_list.push({
				control_id: childControl.Sdt.GetId(),
				client_id: tagRemove.client_id,
				parent_id: getParentId(childControl),
				regionType: tagRemove.regionType,
				type: 'remove'
			})
			Api.asc_RemoveContentControlWrapper(childControl.Sdt.GetId())

		}
		function removeControlChildren(oControl, containSelf, onlyChild, excepetNum) {
			if (!oControl) {
				return
			}
			clearQuesInteraction(oControl, excepetNum)
			if (oControl.GetClassType() == 'blockLvlSdt') {
				var childControls = oControl.GetAllContentControls() || []
				for (var i = 0; i < childControls.length; ++i) {
					// 只有父节点为我的, 且是inline或无clinet_id才删除
					if (onlyChild) {
						var parentControl = childControls[i].GetParentContentControl()
						if (parentControl.Sdt.Id == oControl.Sdt.Id) {
							var childTag = Api.ParseJSON(childControls[i].GetTag())
							if (childControls[i].GetClassType() == 'inlineLvlSdt') {
								if (excepetNum && childTag.regionType == 'num') {
									continue
								}
								doRemoveControl(childControls[i], childTag)
							} else {
								if (!childTag.client_id || childTag.regionType == 'write') {
									doRemoveControl(childControls[i], childTag)
								}
							}
						}
					} else {
						doRemoveControl(childControls[i], Api.ParseJSON(childControls[i].GetTag()))
					}
				}
			}
			if (containSelf) {
				var tagRemove = Api.ParseJSON(oControl.GetTag())
				doRemoveControl(oControl, tagRemove)
				oControl.Sdt.GetLogicDocument().PreventPreDelete = true
			}
		}
		function setBlockTypeNode(oControl, removeType) {
			var tag = Api.ParseJSON(oControl.GetTag())
			if (removeType == 'all' || tag.regionType != removeType) {
				removeControlChildren(oControl, false)
			}
			updateControlTag(oControl, 'question', getParentId(oControl))
		}
		function getNodeData(tag) {
			if (!tag.client_id) {
				return null
			}
			if (question_map[tag.client_id]) {
				var nodeData = node_list.find(e => {
					return e.id == tag.client_id
				})
				if (nodeData) {
					return {
						level_type: question_map[tag.client_id].level_type,
						is_big: nodeData.is_big
					}
				}
			}
			return null
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
		function samePos(pos1, pos2) {
			if (!pos1 || !pos2 || pos1.length != pos2.length) {
				return false
			}
			for (var i = 0; i < pos1.length; ++i) {
				if (pos1[i].Class != pos2[i].Class || pos1[i].Position != pos2[i].Position) {
					return false
				}
			}
			return true
		}
		function getSelectControls() {
			var selectControls = []
			var completeOverlapSdt = null
			// var selectParagraphs = oSelectRange.Paragraphs || []
			for (var i = 0; i < allControls.length; ++i) {
				var controlType = allControls[i].GetClassType()
				if (!allControls[i].Sdt.IsSelectionUse()) {
					continue
				}
				var relation = 0
				if (controlType == 'inlineLvlSdt') {
					relation = 3
				} else if (allControls[i].Sdt.IsSelectedAll()) {
					relation = 3
					var controlRange = allControls[i].GetRange()
					if ((typeName != 'write' && typeName != 'clear' && typeName != 'clearAll') || 
						(samePos(oSelectRange.StartPos, controlRange.StartPos) && samePos(oSelectRange.EndPos, controlRange.EndPos))) {
						relation = 2 // 完全重叠
						completeOverlapSdt = allControls[i].Sdt
					}
				} else {
					relation = 4
				}
				selectControls.push({
					relation: relation,
					classType: controlType,
					control: allControls[i]
				})
			}
			console.log('selectControls', selectControls)
			return {
				selectControls,
				completeOverlapSdt
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
		// 清除单元格中的小问
		function clearCellAsk(oCell, parent_id) {
			if (!oCell) {
				return
			}
			var cellContent = oCell.GetContent()
			if (!cellContent) {
				return
			}
			var drawings = cellContent.GetAllDrawingObjects() || []
			drawings.forEach(oDrawing => {
				deleteTypeDrawing(oDrawing, 'identify')
				deleteTypeDrawing(oDrawing, 'write')
			})
			var allControls = oDocument.GetAllContentControls() || []
			allControls.forEach(oControl => {
				var parentcell = oControl.GetParentTableCell()
				if (parentcell && parentcell.Cell.Id == oCell.Cell.Id) {
					removeControlChildren(oControl, true)
				}
			})
		}
		function setCellType(oCell, parent_id, tname) {
			if (tname == 'write') {
				if (!parent_id || !question_map[parent_id] || question_map[parent_id].level_type != 'question') {
					result = {
						code: 0,
						message: '未处于题目中',
					}
					return false
				}
			}
			if (!oCell || !oCell.GetClassType || oCell.GetClassType() != 'tableCell') {
				return false
			}
			var canadd = false
			var cellContent = oCell.GetContent()
			if (tname == 'write') { // 划分为小问，需要将选中的单元格都设置为小问
				// 需要确保所选单元格里没有control，且自己处于某个control里
				if (cellNotControl(cellContent)) {
					oCell.SetBackgroundColor(255, 191, 191, false)
					canadd = true
				}
			} else if (tname == 'remove') {
				removeCellInteraction(oCell)
				canadd = true
			}
			if (canadd) {
				var oTable = oCell.GetParentTable()
				result.change_list.push({
					parent_id: parent_id,
					table_id: oTable.Table.Id,
					row_index: oCell.GetRowIndex(),
					cell_index: oCell.GetIndex(),
					cell_id: oCell.Cell.Id,
					client_id: 'c_' + oCell.Cell.Id,
					regionType: 'write',
					type: tname == 'remove' ? 'remove' : ''
				})
				var desc = Api.ParseJSON(oTable.GetTableDescription())
				desc[`${oCell.GetRowIndex()}_${oCell.GetIndex()}`] = `c_${oCell.Cell.Id}`
				desc.biyue = 1
				oTable.SetTableDescription(JSON.stringify(desc))
				return true
			}
			return false
		}
		var selectedContent = oDocument.Document.GetSelectedContent()
		var selectedElementsInfo = oDocument.Document.GetSelectedElementsInfo()
		// 选中区域是同一表格内的单元格
		function getSelectCells(curBlockSdt) {
			if (!selectedElementsInfo.m_bTable) {
				return null
			}
			var oControl = curBlockSdt ? Api.LookupObject(curBlockSdt.Id) : null
			var quesTag = getQuestion(oControl)
			if (selectionInfo.isSelection) {
				var oCells = []
				var oCellsUse = []
				var allTables = oDocument.GetAllTables()
				for (var i = 0; i < allTables.length; ++i) {
					var oTable = allTables[i]
					if (!oTable.Table.IsSelectionUse || !(oTable.Table.IsSelectionUse())) {
						continue
					}
					var cellArray = oTable.Table.GetSelectionArray(false)
					for (var j = 0; j < cellArray.length; ++j) {
						var oCell = oTable.GetCell(cellArray[j].Row, cellArray[j].Cell)
						if (!oCell) {
							continue
						}
						var cellContent = oCell.GetContent()
						if (cellContent && cellContent.Document && cellContent.Document.IsSelectedAll()) {
							oCells.push(oCell)
						} else {
							oCellsUse.push(oCell)
						}
					}
					return {
						ques_id: quesTag ? quesTag.client_id : 0,
						oCells: oCells,
						oCellsUse: oCellsUse
					}
				}
			} else if (selectedElementsInfo.m_pParagraph) {
				var oParagraph = Api.LookupObject(selectedElementsInfo.m_pParagraph.Id)
				if (oParagraph) {
					var oCell = oParagraph.GetParentTableCell()
					if (oCell) {
						return {
							ques_id: quesTag ? quesTag.client_id : 0,
							oCells: [oCell]
						}
					}
				}
			}
			return null
		}
		function getQuesCells(quesBlockSdt) {
			if (!quesBlockSdt) {
				return null
			}
			var quesControl = Api.LookupObject(quesBlockSdt.Id)
			if (!quesControl || quesControl.GetClassType() != 'blockLvlSdt') {
				return null
			}
			var pageCount = quesControl.Sdt.getPageCount()
			var cells = []
			for (var p = 0; p < pageCount; ++p) {
				var page = quesControl.Sdt.GetAbsolutePage(p)
				var tables = quesControl.GetAllTablesOnPage(page) || []
				tables.forEach(element => {
					var rows = element.GetRowsCount()
					for (var r = 0; r < rows; ++r) {
						var oRow = element.GetRow(r)
						var cnt = oRow.GetCellsCount()
						for (var c = 0; c < cnt; ++c) {
							cells.push(oRow.GetCell(c))
						}
					}
				})
			}
			return {
				ques_id: Api.ParseJSON(quesControl.GetTag()).client_id,
				oCells: cells
			}
		}
		function handleCell(curBlockSdt, tCmd, allQuesCell) {
			var cellinfo = allQuesCell ? getQuesCells(curBlockSdt) : getSelectCells(curBlockSdt)
			if (cellinfo && cellinfo.oCells) {
				cellinfo.oCells.forEach(e => {
					// 需要删除单元格中的小问
					clearCellAsk(e, cellinfo.ques_id)
					setCellType(e, cellinfo.ques_id, tCmd)
				})
				return true
			}
			return false
		}

		// function list end
		result.typeName = typeName
		result.cmd = options.cmd
		
		var selectDrawings = selectedContent.DrawingObjects
		var arrSdts = selectedElementsInfo.m_arrSdts || []
		console.log('[selectedContent]', selectedContent)
		console.log('[arrSdts]', arrSdts)
		var curSdt = null
		var curBlockSdt = null
		for (var i = arrSdts.length - 1; i >= 0; --i) {
			if (i == arrSdts.length - 1) {
				curSdt = arrSdts[i]
			}
			if (arrSdts[i].GetType() == 3) {
				curBlockSdt = arrSdts[i]
				break
			}
		}
		if (typeName == 'writeZone') { // 作答区
			result.typeName = 'write'
			if (options.cmd == 'add') { // 添加
				var curControl = oDocument.Document.GetContentControl()
				if (!curControl) {
					result.message = '未处于题目中无法处理'
					return result
				}
				addWrite(Api.LookupObject(curControl.Id), 'write')
			} else if (options.cmd == 'del') { // 删除
				delShapeAsk(curBlockSdt, 'write')
			}
		} else if (typeName == 'identify') { // 识别框
			result.typeName = 'write'
			if (options.cmd == 'add') { // 添加
				var cellinfo = getSelectCells(curBlockSdt)
				if (cellinfo) {
					var cellUse = selectionInfo.isSelection ? cellinfo.oCellsUse : cellinfo.oCells
					if (cellUse && cellUse.length) {
						cellUse.forEach(e => {
							setCellType(e, cellinfo.ques_id, 'remove')
						})
					}
				}
				var curPosInfo = oDocument.Document.GetContentPosition()
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
					if (paragraphIdx >= 0) {
						var pParagraph = Api.LookupObject(curPosInfo[paragraphIdx].Class.Id)
						var paraentControl = pParagraph.GetParentContentControl()
						addWrite(paraentControl, 'identify', curPosInfo, runIdx, paragraphIdx)
					}
				}
			} else if (options.cmd == 'del') { // 删除
				delShapeAsk(curBlockSdt, 'identify')
			}
		} else {
			var allControls = oDocument.GetAllContentControls() || []
			var selectControls = []
			var completeOverlap = false
			var oSelectRange = oDocument.GetRangeBySelect()
			if (selectionInfo.isSelection) {
				var r = getSelectControls()
				selectControls = r.selectControls
				if (r.completeOverlapSdt) {
					completeOverlap = true
					curBlockSdt = r.completeOverlapSdt
				}
			}
			var oBlockControl = curBlockSdt ? Api.LookupObject(curBlockSdt.Id) : null
			var oCurControl = curSdt ? Api.LookupObject(curSdt.Id) : null
			if (typeName == 'struct') { // 题组
				if (selectionInfo.isSelection) {
					if (curBlockSdt && completeOverlap) {
						setBlockTypeNode(oBlockControl, 'all')
					} else {
						selectControls.reverse().forEach(e => {
							removeControlChildren(e.control, true)
						})
						addControl(1)
					}
				} else {
					if (!curBlockSdt) {
						result.message = '请先选中一个范围'
						return result
					}
					var parentData = getParentData(oBlockControl)
					if (parentData) {
						result.message = '题组不可位于其他控件中'
						return result
					}
					setBlockTypeNode(oBlockControl, 'all')
				}
			} else if (typeName == 'question') { // 题目
				if (selectionInfo.isSelection) {
					if (curBlockSdt) {
						if (completeOverlap) {
							setBlockTypeNode(oBlockControl, 'question')
						} else {
							var tag = Api.ParseJSON(oBlockControl.GetTag())
							var nodeData = getNodeData(tag)
							var isBig = nodeData.level_type == 'question' && nodeData.is_big
							selectControls.reverse().forEach(e => {
								if (!isBig || e.relation == 3) {
									removeControlChildren(e.control, true)	
								} else {
									if (!arrSdts.find(e2 => {
										return e2.Id == e.control.Sdt.Id
									})) {
										removeControlChildren(e.control, true)	
									}
								}
							})
							addControl(1)
						}
					} else {
						selectControls.reverse().forEach(e => {
							removeControlChildren(e.control, true)
						})
						addControl(1)
					}
				} else {
					if (!curBlockSdt) {
						result.message = '请先选中一个范围'
						return result
					}
					setBlockTypeNode(oBlockControl, 'question')
				}
			} else if (typeName == 'setBig' || typeName == 'clearBig') {
				if (curBlockSdt && ((selectionInfo.isSelection && completeOverlap) || (!selectionInfo.isSelection))) {
					if (curBlockSdt) {
						if (typeName == 'setBig') {
							setBig(oBlockControl)
						} else {
							clearBig(oBlockControl)
						}
						updateControlTag(oBlockControl, 'question', getParentId(oBlockControl))
					}
				}
			} else if (typeName == 'write') { // 小问
				if (selectionInfo.isSelection) {
					if (curBlockSdt) {
						var quesId = getParentQuestion(curBlockSdt)
						if(!quesId) {
							result.message = '未处于题目中'
							return result
						}
						if (selectDrawings) {
							selectDrawings.forEach(e => { // 这里的e的类型是ParaDrawing
								if (e.docPr && e.docPr.title) {
									var oDrawing = allDrawings.find(e2 => {
										return e2.GetTitle() == e.docPr.title // 这里不可用ID查找，因为selectcontent里的元素，ID都是重新分配的
									})
									if (oDrawing) {
										deleteTypeDrawing(oDrawing, 'identify')
									}
								}
							})
						}
						var cellinfo = getSelectCells(curBlockSdt)
						// 删除control时，range会修改，因而这里需要先获取cellinfo，删除，然后再设置cell ask
						for (var j = selectControls.length - 1; j >= 0; --j) {
							if (selectControls[j].relation == 3) {
								var tagobj = Api.ParseJSON(selectControls[j].control.Sdt.GetTag())
								if (tagobj.client_id != quesId) {
									removeControlChildren(selectControls[j].control, true)
								}
							}
						}
						if (cellinfo && cellinfo.oCells && cellinfo.oCells.length) {
							cellinfo.oCells.forEach(e => {
								setCellType(e, cellinfo.ques_id, 'write')
							})
						} else {
							if (cellinfo && cellinfo.oCellsUse && cellinfo.oCellsUse.length) {
								cellinfo.oCellsUse.forEach(e => {
									setCellType(e, cellinfo.ques_id, 'remove')
								})
							}
							oSelectRange.Select()
							var type = 1
							if (oSelectRange.Paragraphs.length) {
								var find = false
								for (var p = 0; p < oSelectRange.Paragraphs.length; ++p) {
									if (!oSelectRange.Paragraphs[p].Paragraph.IsEmpty() &&
										!oSelectRange.Paragraphs[p].Paragraph.IsSelectedAll() &&
										oSelectRange.Paragraphs[p].GetElementsCount()) {
										find = true
										break
									}
								}
								if (find && oSelectRange.Paragraphs.length == 1) {
									type = 2
								}
							}
							addControl(type)
						}
					} else {
						result.message = '未处于题目中'
						return result
					}
				} else {
					if (!curSdt) {
						result.message = '未处于题目中'
						return result
					}
					if (curSdt) {
						// 需要先判断下当前是否处于单元格中
						if (!handleCell(curBlockSdt, 'write')) {
							var parentData = getParentData(oCurControl)
							if (!parentData || parentData.level_type != 'question') {
								result.message = '未处于题目中'
								return result
							}
							updateControlTag(oCurControl, 'write', quesId)
						}
					}
				}
			} else if (typeName == 'clearChildren') { // 清除所有小问
				result.typeName = 'write'
				handleCell(curBlockSdt, 'remove', true)
				removeControlChildren(oCurControl, false, true, true)
			} else if (typeName == 'clear' || typeName == 'clearAll') { // 清除区域
				var askCell = false
				if (selectionInfo.isSelection) {
					handleCell(curBlockSdt, 'remove', true)
					if (selectControls) {
						for (var i = 0; i < selectControls.length; ++i) {
							if (selectControls[i].relation == 3) {
								removeControlChildren(selectControls[i].control, true)
							}
						}
					}
				} else {
					if (curSdt) {
						if (oCurControl.GetClassType() == 'blockLvlSdt') {
							if (handleCell(curBlockSdt, 'remove', false)) {
								askCell = true
							}
						}
						if (!askCell) {
							removeControlChildren(oCurControl, true, typeName == 'clear')
						}
					}
				}
			}
		}
		return result
	}, false, true).then((res1) => {
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

export {
	handleRangeType
}