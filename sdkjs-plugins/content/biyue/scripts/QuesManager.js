
import { biyueCallCommand, dispatchCommandResult } from "./command.js";
import { getQuesType, reqComplete } from '../scripts/api/paper.js'
import { handleChoiceUpdateResult, setInteraction, updateChoice } from "./featureManager.js";
import { initExtroInfo } from "./panelFeature.js";
import { addOnlyBigControl, removeOnlyBigControl, getAllPositions2 } from './business.js'
var levelSetWindow = null
var layoutRepairWindow = null
var imageRelationWindow = null
var level_map = {}
var g_click_value = null
var upload_control_list = []
function initExamTree() {

}

// 处理文档点击
function handleDocClick(options) {
	window.Asc.plugin.executeMethod('GetCurrentContentControlPr', [], function(returnValue) {
		console.log('GetCurrentContentControlPr', returnValue)
		if (returnValue && returnValue.Tag) {
			try {
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
				} else if (window.tab_select == 'tabQues') {
					var event = new CustomEvent('clickSingleQues', {
						detail: {
							InternalId: returnValue.InternalId
						},
					})
					document.dispatchEvent(event)
				}
				if (options.isSelectionUse) {
					var shortcutKey = window.BiyueCustomData ? window.BiyueCustomData.ask_shortcut : null
					if (shortcutKey && shortcutKey != '') {
						var sckey = `${shortcutKey}Key`
						if (options[sckey]) {
							// 划分小问
							updateRangeControlType('write')
						}
					}
				}
			} catch (error) {
				console.log(error)
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
		// if (!g_click_value) {
		// 	return
		// }
		window.Asc.plugin.executeMethod('AddContextMenuItem', [getContextMenuItems(options.type)])
	} else if (options.type == 'Selection' || options.type == 'Shape' || options.type == 'Image') { // 选中范围，针对所选范围处理
		window.Asc.plugin.executeMethod('AddContextMenuItem', [getContextMenuItems(options.type)])
	}
}

function getContextMenuItems(type) {
	if (type =='Image' || type == 'Shape') {
		var items = [{
			separator: true,
			id: 'handleImageIgnore',
			text: '图片铺码',
			items: [{
				id: 'handleImageIgnore_del',
				text: '开启'
			}, {
				id: 'handleImageIgnore_add',
				text: '关闭'
			}]
		}, {
			id: 'imageRelation',
			text: '图片关联'
		}]
		if (type == 'Shape') {
			items.push({
				id: 'handleWrite_del',
				text: '删除作答区'
			})
		}
		return {
			guid: window.Asc.plugin.guid,
			items: items
		}
	} else if (type == 'Target' && !g_click_value) {
		return {
			guid: window.Asc.plugin.guid,
			items: [{
				separator: true,
				id: 'setSectionColumn_2',
				text: '分为2栏'
			}, {
				id: 'setSectionColumn_1',
				text: '取消分栏'
			}]
		}
	}

	var nodeData = null
	var currentType = ''
	if (type == 'Target' && g_click_value) {
		var client_id = g_click_value.Tag.client_id
		var node_list = window.BiyueCustomData.node_list || []
		nodeData = node_list.find(e => {
			return e.id == client_id
		})
		if (nodeData) {
			if (nodeData.level_type == 'struct') {
				currentType = '(现为题组)'
			} else if (nodeData.level_type == 'question') {
				currentType = nodeData.is_big ? '(现为大题)' : '(现为题目)'
			}
		} else if (g_click_value.Tag.regionType == 'write') {
			var question_map = window.BiyueCustomData.question_map || {}
			var keys = Object.keys(question_map)
			for (var i = 0, imax = keys.length; i < imax; ++i ) {
				if (question_map[keys[i]].ask_list) {
					var askIndex = question_map[keys[i]].ask_list.findIndex(e => {
						return e.id == g_click_value.Tag.client_id
					})
					if (askIndex >= 0) {
						currentType = '(现为小问)'
						break
					}
				}
			}
		}
	}

	var splitType = {
		separator: true,
		id: 'updateControlType',
		text: '划分类型' + currentType,
		items: []
	}
	var list = ["question", 'struct', 'setBig', 'write', 'clearBig', 'clear', 'clearAll']
	var names = ['设置为 - 题目', '设置为 - 题组(结构)', '设置为 - 大题', '设置为 - 小问', '清除 - 大题', '清除 - 选中区域', '清除 - 选中区域(含子级)']
	var icons = [
		'rect',
		'struct',
		'rect',
		'rect',
		'clear',
		'clear',
		'clear'
	]
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
				icons: `resources/%theme-type%(light|dark)/%state%(normal)${icons[index]}%scale%(100|200).%extension%(png)`,
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
	if ((type == 'Target' && g_click_value && nodeData && nodeData.level_type == 'question') || type == 'Selection' ) {
		settings.items.push({
			id: `layoutRepair`,
			text: '排版修复'
		})
	}
	if (canBatch) {
		var questypes = window.BiyueCustomData.paper_options ? window.BiyueCustomData.paper_options.question_type : []
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
				// {
				// 	id: 'batchChangeScore',
				// 	text: '修改分数',
				// },
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
	if (type == 'Selection') {
		settings.items.push({
			id: 'tableRelation',
			text: '表格关联'
		})
	}
	settings.items.push({
		id: 'setSectionColumn',
		text: '分栏',
		items: [{
			id: 'setSectionColumn_2',
			text: '分为2栏'	
		}, {
			id: 'setSectionColumn_1',
			text: '取消分栏'
		}]
	})
	return settings
}

function onContextMenuClick(id) {
	var strs = id.split('_')
	if (strs && strs.length > 0) {
		var funcName = strs[0]
		switch (funcName) {
			case 'updateControlType':
				updateRangeControlType(strs[1])
				break
			case 'batchChangeQuesType':
				batchChangeQuesType(strs[1])
				break
			case 'batchChangeProportion':
				batchChangeProportion(strs[1])
				break
			case 'batchChangeInteraction':
				batchChangeInteraction(strs[1])
				break
			case 'handleIdentifyBox':
				handleIdentifyBox(strs[1])
				break
			case 'setSectionColumn': // 分栏
				setSectionColumn(strs[1] * 1)
				break
			case 'handleWrite':
				handleWrite(strs[1])
				break
			case 'handleImageIgnore':
				handleImageIgnore(strs[1])
				break
			case 'layoutRepair':
				layoutDetect()
				break
			case 'batchChangeScore':
				batchChangeScore()
				break
			case 'imageRelation':
				imageRelation()
				break
			case 'tableRelation':
				tableRelation()
				break
			default:
				break
		}
	}
}

function getNodeList() {
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var node_list = []
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
				var parentTag = getJsonData(parentBlock.GetTag())
				parent_id = parentTag.client_id || 0
			}
			return parent_id
		}
		function getParagraphWriteList(oElement, write_list) {
			if (!oElement || !oElement.GetClassType || oElement.GetClassType() != 'paragraph') {
				return
			}
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
									var titleObj = getJsonData(title)
									if (titleObj.feature && (titleObj.feature.sub_type == 'write' || titleObj.feature.sub_type == 'identify')) {
										write_list.push({
											id: titleObj.feature.client_id,
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
					var childTag = getJsonData(oChild.GetTag())
					if (childTag.regionType == 'write') {
						write_list.push({
							id: childTag.client_id,
							sub_type: 'control',
							control_id: oChild.Sdt.GetId()
						})
					}
				}
			}
		}
		function getBlockWriteList(oElement, write_list) {
			if (!oElement || !oElement.GetClassType || oElement.GetClassType() != 'blockLvlSdt') {
				return
			}
			var childTag = getJsonData(oElement.GetTag())
			if (childTag.regionType == 'write') {
				write_list.push({
					id: childTag.client_id,
					sub_type: 'control',
					control_id: oElement.Sdt.GetId()
				})
			}
		}
		for (var i = 0, imax = controls.length; i < imax; ++i) {
			var oControl = controls[i]
			var tagInfo = getJsonData(oControl.GetTag())
			var parent_id = getParentId(oControl)
			if (tagInfo.regionType == 'question') {
				var write_list = []
				var controlContent = oControl.GetContent()
				var elementCount = controlContent.GetElementsCount()
				for (var j = 0; j < elementCount; ++j) {
					var oElement = controlContent.GetElement(j)
					if (oElement.GetClassType() == 'paragraph') {
						getParagraphWriteList(oElement, write_list)
					} else if (oElement.GetClassType() == 'blockLvlSdt') {
						getBlockWriteList(oElement, write_list)
					} else if (oElement.GetClassType() == 'table') {
						var rows = oElement.GetRowsCount()
						for (var i1 = 0; i1 < rows; ++i1) {
							var oRow = oElement.GetRow(i1)
							var cells = oRow.GetCellsCount()
							for (var i2 = 0; i2 < cells; ++i2) {
								var oCell = oRow.GetCell(i2)
								var shd = oCell.Cell.Get_Shd()
								var fill = shd.Fill
								if (fill && fill.r == 255 && fill.g == 191 && fill.b == 191) {
									write_list.push({
										id: 'c_' + oCell.Cell.Id,
										sub_type: 'cell',
										table_id: oElement.Table.Id,
										cell_id: oCell.Cell.Id,
										row_index: i1,
										cell_index: i2
									})
								} else {
									var oCellContent = oCell.GetContent()
									var cnt1 = oCellContent.GetElementsCount()
									for (var i3 = 0; i3 < cnt1; ++i3) {
										var oElement2 = oCellContent.GetElement(i3)
										if (!oElement2) {
											continue
										}
										if (oElement2.GetClassType() == 'paragraph') {
											getParagraphWriteList(oElement2, write_list)
										} else if (oElement2.GetClassType() == 'blockLvlSdt') {
											getBlockWriteList(oElement2, write_list)
										}
									}
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
				var parentTag = getJsonData(parentBlock.GetTag())
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
				var tag = getJsonData(childControl.GetTag())
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
					var titleObj = getJsonData(title)
					if (titleObj.feature && titleObj.feature.sub_type == 'ask_accurate') {
						oRun.Delete()
						return true
					}
				}
			}
			return false
		}
		function getFormatTypeString(format) {
			var sType = 'decimal'
			switch(format) {
				case 8: sType = 'chineseCounting'; break
				case 9: sType = 'chineseCountingThousand'; break
				case 10: sType = 'chineseLegalSimplified'; break
				case 14: sType = 'decimalEnclosedCircle'; break
				case 15: sType = 'decimalEnclosedCircleChinese'; break
				case 21: sType = 'decimalZero'; break
				case 46: sType = 'lowerLetter'; break
				case 47: sType = 'lowerRoman'; break
				case 60: sType = 'upperLetter'; break
				case 61: sType = 'upperRoman'; break
				default: break
			}
			return sType
		}
		function hideSimpleControl(oControl) {
			var childControls = oControl.GetAllContentControls() || []
			if (childControls.length) {
				for (var c = childControls.length - 1; c >= 0; --c) {
					var tag = getJsonData(childControls[c].GetTag())
					if (tag.regionType == 'num' && childControls[c].GetClassType() == 'inlineLvlSdt') {
						var parent = childControls[c].Sdt.Parent
						if (parent && parent.GetType() == 1) {
							var oParent = Api.LookupObject(parent.Id)
							if (oParent) {
								var pos = childControls[c].Sdt.GetPosInParent()
								if (pos >= 0) {
									oParent.RemoveElement(pos)
								}
							}
						}
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
			var sType = getFormatTypeString(oNumberingLvl.Format)
			oNumberingLevel.SetTemplateType('bullet', '');
			var str = ''
			var find = false
			for (var i = 0; i < LvlText.length; ++i) {
				if (LvlText[i].Type == 2) {
					str += `%${level+1}`
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
			var tag = getJsonData(oControl.GetTag() || '{}')
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
			} else if (tag.regionType =='question') {
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
									var titleObj = getJsonData(title)
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
				var desc = getJsonData(oTable.GetTableDescription())
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
			var tagRemove = getJsonData(oRemove.GetTag() || '{}')
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
		function getNewNodeList() {
			var controls = oDocument.GetAllContentControls() || []
			var nodeList = []
			controls.forEach((oControl) => {
				var tagInfo = getJsonData(oControl.GetTag() || '{}')
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
				var desc = getJsonData(oTable.GetTableDescription())
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
				var desc = getJsonData(oTable.GetTableDescription())
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
				var tag = getJsonData(oControl.GetTag() || '{}')
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
			var tag = getJsonData(oControl.GetTag() || '{}')
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
			var tag = getJsonData(oControl.GetTag() || '{}')
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
					var nextTag = getJsonData(element.GetTag() || '{}')
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
		function removeTableAsk(quesControl) {
			// 删除所有单元格小问
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
							removeControl(container)
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
					if (a.GetClassType() != 'inlineLvlSdt' || a.Sdt.GetId() != control.Sdt.GetId()) {
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
					if (pre2.GetClassType() == 'inlineLvlSdt' && pre2.Sdt.GetId() == control.Sdt.GetId()) {
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
				var oParentTag = getJsonData(oParentControl.GetTag() || '{}')
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
										if (oElement2.GetClassType() == 'paragraph' && oElement2.Paragraph.Bounds.Bottom == 0 && oElement2.Paragraph.Bounds.Top == 0) {
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
	console.log('========== interaction', window.BiyueCustomData.interaction)
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

	function updateScore(qid) {
		if (question_map[qid] && question_map[qid].ask_list && question_map[qid].mark_mode != 2) {
			var sum = 0
			question_map[qid].ask_list.forEach(e => {
				var ascore = e.score * 1
				if (isNaN(ascore)) {
					ascore = 0
				}
				sum += ascore
			})
			question_map[qid].score = sum
		}
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
						updateScore(nodeData.parent_id)
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
						ques_default_name: item.numbing_text ? getNumberingText(item.numbing_text) : GetDefaultName(targetLevel, item.text),
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
						question_map[item.client_id].ques_default_name = item.numbing_text ? getNumberingText(item.numbing_text) : GetDefaultName(targetLevel, item.text)
						question_map[item.client_id].level_type = targetLevel
					} else {
						question_map[item.client_id].level_type = targetLevel
					}
				}
			} else if (targetLevel == 'struct') {
				question_map[item.client_id] = {
					text: item.text,
					level_type: targetLevel,
					ques_default_name: item.numbing_text ? getNumberingText(item.numbing_text) : GetDefaultName(targetLevel, item.text)
				}
				update_node_id = item.client_id
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
					if (ndata.write_list) {
						var writeIndex = ndata.write_list.findIndex(e => {
							return e.id == item.client_id
						})
						if (writeIndex == -1 && item.type != 'remove') {
							ndata.write_list.push({
								id: item.client_id,
								control_id: item.control_id,
								regionType: item.regionType
							})
							writeIndex = 0
						}
						if (question_map[parent_id]) {
							if (!question_map[parent_id].ask_list) {
								question_map[parent_id].ask_list = []
							}
							var index2 = question_map[parent_id].ask_list.findIndex(e => {
								return e.id == item.client_id
							})
							if (index2 < 0) {
								var toIndex = 0
								for (var i1 = writeIndex - 1; i1 >= 0; --i1) {
									var index1 = question_map[parent_id].ask_list.findIndex(e => {
										return e.id == ndata.write_list[i1].id
									})
									if (index1 >= 0) {
										toIndex = index1 + 1
										break
									}
								}
								question_map[parent_id].ask_list.splice(toIndex, 0, {
									id: item.client_id,
									score: 1
								})
								updateScore(parent_id)
							} else if (item.type == 'remove') {
								question_map[parent_id].ask_list.splice(index2, 1)
								updateScore(parent_id)
							}
						}
					}
					if (ndata) {
						parentNode.write_list = ndata.write_list
					}
				}
			}
		} else if (level_type != 'clear' && level_type != 'clearAll') {
			// 之前没有，需要增加
			var index = node_list.length
			if (res2) {
				index = res2.findIndex(e => {
					return e.id == item.client_id
				})
				if (index == -1) {
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
			} else if (targetLevel == 'struct') {
				update_node_id = item.client_id
			}
			if (targetLevel == 'question' || targetLevel == 'struct') {
				question_map[item.client_id] = {
					text: item.text,
					level_type: targetLevel,
					ques_default_name: item.numbing_text ? getNumberingText(item.numbing_text) : GetDefaultName(targetLevel, item.text)
				}
				if (targetLevel == 'question') {
					question_map[item.client_id].ask_list = ask_list
					question_map[item.client_id].interaction = window.BiyueCustomData.interaction
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
						updateScore(item.parent_id)
					}
					addIds.push(item.parent_id)
				}
			}
		}
	})

	window.BiyueCustomData.node_list = node_list
	window.BiyueCustomData.question_map = question_map
	var interaction = window.BiyueCustomData.interaction
	var updateinteraction = false
	if (addIds && addIds.length) {
		if (level_type == 'write' || level_type == 'clear' || level_type == 'clearAll') {
			if (question_map[addIds[0]]) {
				interaction = question_map[addIds[0]].interaction
				updateinteraction = true
			}
		} else if (targetLevel == 'question') {
			updateinteraction = true
		}
		update_node_id = addIds[0]
	}
	var use_gather = window.BiyueCustomData.choice_display && window.BiyueCustomData.choice_display.style != 'brackets_choice_region'
	if (use_gather) {
		deleteChoiceOtherWrite(null, false).then(() => {
			notifyQuestionChange(update_node_id)
			updateChoice().then(res => {
				return handleChoiceUpdateResult(res)
			}).then(() => {
				window.biyue.StoreCustomData()
			})
		})
	} else {
		if (updateinteraction) {
			deleteChoiceOtherWrite(null, false).then(res => {
				notifyQuestionChange(update_node_id)
				setInteraction(interaction, addIds).then(() => window.biyue.StoreCustomData())
			})
		} else {
			deleteChoiceOtherWrite(null, true).then(res => {
				notifyQuestionChange(update_node_id)
				window.biyue.StoreCustomData()
			})
		}
	}
	console.log('handleChangeType end', node_list, 'g_click_value', g_click_value, 'update_node_id', update_node_id)
	notifyQuestionChange(update_node_id)
}

function notifyQuestionChange(update_node_id) {
	document.dispatchEvent(
		new CustomEvent('updateQuesData', {
			detail: {
				client_id: update_node_id
			}
		})
	)
}

function updateAllChoice() {
	if (window.BiyueCustomData.choice_display && window.BiyueCustomData.choice_display.style != 'brackets_choice_region') {
		return updateChoice().then(res => {
			return handleChoiceUpdateResult(res)
		})
	} else {
		return new Promise((resolve, reject) => {
			resolve()
		})
	}
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
						var tag = getJsonData(oControl.GetTag() || '{}')
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
											var nextTag = getJsonData(nextControl.GetTag() || '{}')
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
													var childTag = getJsonData(e.GetTag() || '{}')
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
					var tag = getJsonData(e.GetTag() || '{}')
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
		var question_map = window.BiyueCustomData.question_map || {}
		var judgeChoiceAll = false
		var choice_ids = []
		var useGather = window.BiyueCustomData.choice_display && window.BiyueCustomData.choice_display.style != 'brackets_choice_region'
		var interaction_ids = []
		res.list.forEach(e => {
			if (question_map[e.id]) {
				var oldMode = question_map[e.id].ques_mode
				var newMode = getQuesMode(type)
				if (question_map[e.id].interaction != 'none' && (oldMode == 6 || newMode == 6)) {
					interaction_ids.push(e.id)
				}
				question_map[e.id].question_type = type * 1
				question_map[e.id].ques_mode = newMode
				if (newMode == 1 || newMode == 5) {
					choice_ids.push(e.id)
				}
				if (!judgeChoiceAll && useGather) {
					judgeChoiceAll = (newMode == 1 || newMode == 5 || oldMode == 1 || oldMode == 5)
				}
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
		return {
			choice_ids,
			judgeChoiceAll,
			interaction_ids
		}
	}).then(res => {
		return new Promise((resolve, reject) => {
			if (res && res.choice_ids && res.choice_ids.length) {
				return deleteChoiceOtherWrite(res.choice_ids, false).then(() => {
					resolve(res)
				})
			} else {
				resolve(res)
			}
		})
	}).then(res => {
		return new Promise((resolve, reject) => {
			if (res && res.judgeChoiceAll) {
				return updateAllChoice().then(() => {
					resolve(res)
				})
			} else {
				resolve(res)
			}
		})
	}).then((res => {
		return new Promise((resolve, reject) => {
			if (res && res.interaction_ids && res.interaction_ids.length) {
				return setInteraction('useself', res.interaction_ids).then(() => {
					resolve(res)
				})
			} else {
				resolve(res)
			}
		})
	})).then(() => {
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
	window.biyue.showDialog('levelSetWindow', '自动序号识别设置', 'levelSet.html', 592, 400)
}
function updateDataBySavedData(str) {
	if (!str || str == '') {
		return
	}
	try {
		var data = JSON.parse(str)
		var recordtime = new Date(data.time).getTime()
		var nowtime = new Date(window.BiyueCustomData.time).getTime()
		console.log('recordtime', recordtime, 'nowtime', nowtime)
		if (recordtime < nowtime) {
			console.log('当前保存的时间比后端存储的新，忽略')
			return
		}
		if (data.client_node_id) {
			console.log('========== 使用后端存储的')
			var maxId = 0
			for (var i = 0; i < data.node_list.length; i++) {
				if (data.node_list[i].id > maxId) {
					maxId = data.node_list[i].id * 1
				}
				if (data.node_list[i].write_list) {
					for (var j = 0; j < data.node_list[i].write_list.length; j++) {
						if (data.node_list[i].write_list[j].sub_type != 'cell') {
							if (data.node_list[i].write_list[j].id * 1 > maxId) {
								maxId = data.node_list[i].id * 1
							}
						}
					}
				}
			}
			window.BiyueCustomData.client_node_id = maxId >= data.client_node_id ? (maxId + 1) : data.client_node_id
			window.BiyueCustomData.node_list = data.node_list
			window.BiyueCustomData.question_map = data.question_map
		}
	} catch (error) {
		console.log(error)
	}
}
// 由于从BiyueCustomData中获取的中文取出后会是乱码，需要在初始化时，再根据controls刷新一次数据
function initControls() {
	Asc.scope.question_map = window.BiyueCustomData.question_map || {}
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	return biyueCallCommand(window, function() {
		var question_map = Asc.scope.question_map || {}
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		var oTables = oDocument.GetAllTables() || []
		var nodeList = []
		var ids = {}
		var client_node_id = Asc.scope.client_node_id
		var maxid = client_node_id
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
		for (var i = 0, imax = controls.length; i < imax; ++i) {
			var oControl = controls[i]
			var tagInfo = getJsonData(oControl.GetTag())
			if (tagInfo.onlybig == 1) {
				Api.asc_RemoveContentControlWrapper(oControl.Sdt.GetId())
				continue
			}

			// tagInfo.color = '#ff0000'
			// oControl.Sdt.SetColor(colors[oControl.GetClassType()])
			if (tagInfo.client_id) {
				if (maxid < tagInfo.client_id * 1) {
					maxid = tagInfo.client_id * 1
				}
			}
			var parentid = 0
			if (tagInfo.regionType) {
				if (tagInfo.regionType == 'question') {
					ids[tagInfo.client_id] = 1
				}
				var parentControl = getParentBlock(oControl)
				if (parentControl) {
					var parentTagInfo = getJsonData(parentControl.GetTag())
					if (parentTagInfo.regionType == 'question') {
						parentid = parentTagInfo.client_id
					}
				}
				nodeList.push({
					id: tagInfo.client_id,
					text: oControl.GetRange().GetText(),
					parentId: parentid,
					numbing_text: GetNumberingValue(oControl),
					control_id: oControl.Sdt.GetId(),
					sub_type: 'control'
				})
			}
			var changecolor = false
			if (tagInfo.client_id) {
				if (tagInfo.regionType == 'write') {
					var isWrite = false
					if (question_map[parentid] && question_map[parentid].level_type == 'question' && question_map[parentid].ask_list) {
						isWrite = question_map[parentid].ask_list.find(e => {
							return e.id == tagInfo.client_id
						})
					}
					if (isWrite) {
						changecolor = true
						tagInfo.color = '#ff000040'
					}
				} else if (oControl.GetClassType() == 'blockLvlSdt' && question_map[tagInfo.client_id]) {
					if (question_map[tagInfo.client_id].level_type == 'question') {
						tagInfo.color = '#d9d9d940'
						changecolor = true
					} else if (question_map[tagInfo.client_id].level_type == 'struct') {
						tagInfo.color = '#CFF4FF80'
						changecolor = true
					}
				}
			} else if (tagInfo.regionType == 'num') {
				tagInfo.color = '#ffffff40'
				changecolor = true
			}
			if (!changecolor && tagInfo.color) {
				delete tagInfo.color
			}
			oControl.SetTag(JSON.stringify(tagInfo))
		}
		var drawings = oDocument.GetAllDrawingObjects() || []
		var drawingList = []
		drawings.forEach(oDrawing => {
			var title = oDrawing.Drawing.docPr.title || '{}'
			var titleObj = getJsonData(title)
			if (titleObj.feature && titleObj.feature.client_id) {
				if (maxid < titleObj.feature.client_id * 1) {
					maxid = titleObj.feature.client_id * 1
				}
			}
			if(titleObj && titleObj.feature && titleObj.feature.zone_type == 'question' && (titleObj.feature.sub_type == 'write' || titleObj.feature.sub_type == 'identify')) {
				drawingList.push({
					id: titleObj.feature.client_id,
					shape_id: oDrawing.Drawing.GraphicObj.Id,
					drawing_id: oDrawing.Drawing.Id,
					sub_type: titleObj.feature.sub_type
				})
			}
		})
		var cellAskMap = {}
		oTables.forEach(oTable => {
			if (oTable.GetPosInParent() >= 0) {
				var desc = getJsonData(oTable.GetTableDescription())
				Object.keys(desc).forEach(key => {
					if (key != 'biyue') {
						var rc = key.split('_')
						var oCell = oTable.GetCell(rc[0], rc[1])
						if (oCell) {
							var obj = {
								table_id: oTable.Table.Id,
								row_index: rc[0],
								cell_index: rc[1],
								cell_id: oCell.Cell.Id
							}
							if (cellAskMap[desc[key]]) {
								cellAskMap[desc[key]].push(obj)
							} else {
								cellAskMap[desc[key]] = [obj]
							}
						}
					}
				})
			}
		})
		if (maxid > client_node_id) {
			client_node_id = maxid + 1
		}
		return {
			nodeList,
			drawingList,
			ids,
			client_node_id,
			cellAskMap
		}
	}, false, false).then(res => {
		console.log('initControls   nodeList', res)
		return new Promise((resolve, reject) => {
			// todo.. 这里暂不考虑上次的数据未保存或保存失败的情况，只假设此时的control数据和nodelist里的是一致的，只是乱码而已，其他的后续再处理
			if (res.client_node_id) {
				window.BiyueCustomData.client_node_id = res.client_node_id
			}
			var nodeList = res.nodeList
			var drawingList = res.drawingList
			var cellAskMap = res.cellAskMap
			var ids = res.ids || {}
			if (nodeList && nodeList.length > 0) {
				var question_map = window.BiyueCustomData.question_map || {}
				var node_list = window.BiyueCustomData.node_list || []
				nodeList.forEach(node => {
					if (question_map[node.id]) {
						question_map[node.id].text = node.text
						question_map[node.id].ques_default_name = node.numbing_text ? getNumberingText(node.numbing_text) : GetDefaultName(question_map[node.id].level_type, node.text)
					}
				})
				Object.keys(question_map).forEach(id => {
					if (!ids[id]) {
						delete question_map[id]
					}
				})
				for (var i = node_list.length - 1; i >= 0; --i) {
					var nodeData = node_list[i]
					var ndata = nodeList.find(e => {
						return e.id == nodeData.id
					})
					if (ndata) {
						nodeData.control_id = ndata.control_id
						if (nodeData.write_list) {
							for (var j = 0; j < nodeData.write_list.length; ++j) {
								var writeData = nodeData.write_list[j]
								if (writeData.sub_type == 'control') {
									var ndata = nodeList.find(e => {
										return e.id == writeData.id
									})
									if (ndata) {
										writeData.control_id = ndata.control_id
									}
								} else if (writeData.sub_type == 'write' || writeData.sub_type == 'identify') {
									var ddata = drawingList.find(e => {
										return e.id == writeData.id
									})
									if (ddata) {
										writeData.shape_id = ddata.shape_id
										writeData.drawing_id = ddata.drawing_id
									}
								} else if (writeData.sub_type == 'cell') {
									// todo..目前还没办法处理单元格ID改变后如何对应的情况
									var celldata = cellAskMap[writeData.id]
									if (celldata && celldata.length == 1) {
										var cdata = celldata[0]
										if (writeData.cell_index == undefined || (writeData.cell_index == cdata.cell_index && writeData.row_index == cdata.row_index)) {
											writeData.table_id = cdata.table_id
											writeData.row_index = cdata.row_index
											writeData.cell_index = cdata.cell_index
											writeData.cell_id = cdata.cell_id
										}
									}
								}
							}
						}
					} else {
						node_list.splice(i, 1)
					}
				}
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
		var controls = oDocument.GetAllContentControls() || []
		var paragrahs = oDocument.GetAllParagraphs() || []
		paragrahs.forEach(oParagraph => {
			var oParaPr = oParagraph.GetParaPr();
			oParaPr.SetShd("clear", 255, 255, 255, true);
		})
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
		function getProportion(oControl) {
			if (!oControl) {
				return 1
			}
			var oCell = oControl.GetParentTableCell()
			if (!oCell) {
				return 1
			}
			var cellContent = oCell.GetContent()
			if (!cellContent) {
				return 1
			}
			var oTable = oControl.GetParentTable()
			var tableWidth = 100
			if (oTable) {
				var tableBounds = oTable.Table.Bounds
				tableWidth = tableBounds.Right - tableBounds.Left
			}
			var cnt1 = cellContent.GetElementsCount()
			for (var i = 0; i < cnt1; ++i) {
				var oElement = cellContent.GetElement(i)
				if (oElement && oElement.GetClassType && oElement.GetClassType() == 'blockLvlSdt') {
					if (oElement.Sdt.GetId() == oControl.Sdt.GetId()) {
						var TableCellW = oCell.CellPr.TableCellW
						if (!TableCellW) {
							TableCellW = oCell.Cell.CompiledPr.Pr.TableCellW
						}
						if (TableCellW && TableCellW.W) {
							if (TableCellW.Type == 3) {
								return Math.ceil(100 / TableCellW.W)
							} else if (TableCellW.Type == 1) {
								return Math.ceil(tableWidth / TableCellW.W)
							}
						}
					}
				}
			}
			// todo..
			return 1
		}
		controls.forEach((oControl) => {
			var tagInfo = getJsonData(oControl.GetTag())
			 if (tagInfo.regionType == 'question' || tagInfo.regionType == 'write') {
				client_node_id += 1
				var id = client_node_id
				tagInfo.client_id = id
				var parentControl = getParentBlock(oControl)
				var nodeData = {
					id: id,
					control_id: oControl.Sdt.GetId(),
					regionType: tagInfo.regionType
				}
				if (tagInfo.regionType == 'write') {
					nodeData.level_type = 'write'
					if (parentControl) {
						var parent_tagInfo = getJsonData(parentControl.GetTag())
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
							tagInfo.color = '#ff000040'
						}
					}
				} else {
					var proportion = getProportion(oControl)
					var text = oControl.GetRange().GetText()
					var level_type = levelmap[tagInfo.lvl]
					nodeData.level_type = level_type
					nodeData.proportion = proportion
					var detail = {
						text: text,
						ask_list: [],
						level_type: level_type,
						numbing_text: GetNumberingValue(oControl),
						proportion: proportion
					}
					if (tagInfo.regionType == 'question') {
						nodeData.write_list = []
						detail.ask_list = []
						if (level_type == 'question') {
							tagInfo.color = '#d9d9d940'
						} else if (level_type == 'struct') {
							tagInfo.color = '#CFF4FF80'
						} else {
							tagInfo.color = '#ffffff'
						}
					}
					nodeList.push(nodeData)
					questionMap[id] = detail
				}
				oControl.SetTag(JSON.stringify(tagInfo));
			 }
		})
		Api.asc_SetGlobalContentControlShowHighlight(true, 255, 191, 191)
		return {
			client_node_id: client_node_id,
			nodeList: nodeList,
			questionMap: questionMap
		}
	}, false, false).then(res => {
		console.log('===== confirmLevelSet res', res)
		Asc.scope.control_hightlight = true
		if (res) {
			window.BiyueCustomData.client_node_id = res.client_node_id
			window.BiyueCustomData.node_list = res.nodeList
			Object.keys(res.questionMap).forEach(key => {
				var qdata = res.questionMap[key]
				res.questionMap[key].ques_default_name = qdata.numbing_text ? getNumberingText(qdata.numbing_text) : GetDefaultName(qdata.level_type, qdata.text)
			})
			window.BiyueCustomData.question_map = res.questionMap
		}
		return new Promise((resolve, reject) => {
			resolve()
		})
	})
	.then(() => {
		return reqGetQuestionType()
	})
	.then(() => {
		return deleteChoiceOtherWrite()
	})
	.then(() => {
		return initExtroInfo()
	})
	.then(() => {
		console.log("================================ StoreCustomData")
		window.biyue.StoreCustomData()
	})
}

function getNumberingText(text) {
	if (!text || typeof text != 'string') {
		return null
	}
	return text.replace(/[\ue749\ue6a1\ue607]/g, '');
}

function GetDefaultName(level_type, str) {
	if (!str || typeof str != 'string') {
		return ''
	}
	str = str.replace(/[\ue749\ue6a1\ue607]/g, '');
	var text = str
	var texts = str.split('\r\n')
	if (texts && texts.length > 0) {
		text = texts[0]
	}
	if (level_type == 'struct') {
		const pattern = /^[一二三四五六七八九十0-9]+.*?(?=[：:])/
		const result = pattern.exec(text)
		return result ? result[0] : null
	} else if (level_type == 'question')  {
		const regex = /^([^.．、]*)/
		var match = text.match(regex)
		if (match && match[1]) {
			var str = match[1]
			if (str.indexOf('(') >= 0) {
				const regex2 = /\(? *\b(\d+)\b *\)?\.?/g;
				match = str.match(regex2)
			}
		}
		if (match) {
			if (match[1]) {
				return match[1]
			} else if (match[0]) {
				return match[0]
			}
		}
		return ''
	}
	return ''
}

// 获取作答模式
function getQuesMode(question_type) {
	question_type = question_type * 1
	if (question_type == 6) {
		return 6
	}
	var ques_mode = 0
	if (question_type > 0) {
		var mark_type_info = Asc.scope.subject_mark_types
		if (mark_type_info && mark_type_info.list) {
			var find = mark_type_info.list.find(e => {
				return e.question_type_id == question_type
			})
			return find ? find.ques_mode : 3
		}
	}
	return ques_mode
}

// 获取题型
function reqGetQuestionType() {
	Asc.scope.node_list = window.BiyueCustomData.node_list
	return biyueCallCommand(window, function() {
		var node_list = Asc.scope.node_list || []
		var target_list = []
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
		controls.forEach((oControl) => {
			var tagInfo = getJsonData(oControl.GetTag())
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
						window.BiyueCustomData.question_map[e.id].ques_mode = getQuesMode(e.question_type)
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
// 删除选择题多余的空
function deleteChoiceOtherWrite(ids, recalc = true) {
	Asc.scope.question_map = window.BiyueCustomData.question_map || {}
	Asc.scope.node_list = window.BiyueCustomData.node_list || []
	Asc.scope.ids = ids ? ids : Object.keys(Asc.scope.question_map)
	Asc.scope.choice_blank = window.BiyueCustomData.choice_blank
	console.log('deleteChoiceOtherWrite begin')
	return biyueCallCommand(window, function() {
		var question_map = Asc.scope.question_map
		var node_list = Asc.scope.node_list
		var choice_blank = Asc.scope.choice_blank
		var mark_type_info = Asc.scope.subject_mark_types
		var ids = Asc.scope.ids
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls() || []
		var updateInteraction = false
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
		function deleteAccurateRun(oRun) {
			if (oRun &&
				oRun.Run &&
				oRun.Run.Content &&
				oRun.Run.Content[0] &&
				oRun.Run.Content[0].docPr) {
				var title = oRun.Run.Content[0].docPr.title
				if (title) {
					var titleObj = getJsonData(title)
					if (titleObj.feature && titleObj.feature.sub_type == 'ask_accurate') {
						oRun.Delete()
						updateInteraction = true
						return true
					}
				}
			}
			return false
		}
		function deleteControlAccurate(oControl) {
			if (!oControl || !oControl.GetElementsCount) {
				return
			}
			var elementCount = oControl.GetElementsCount()
			for (var idx = 0; idx < elementCount; ++idx) {
				var oRun = oControl.GetElement(idx)
				if (deleteAccurateRun(oRun)) {
					break
				}
			}
		}

		function updateAccurateText(blankControlId) {
			var blankControl = Api.LookupObject(blankControlId)
			if (!blankControl) {
				return
			}
			if (blankControl.GetClassType() == 'inlineLvlSdt') {
				var elementcount = blankControl.GetElementsCount()
				for (var i = 0; i < elementcount; ++i) {
					var oChild = blankControl.GetElement(i)
					if (oChild.GetClassType() == 'run' && oChild.Run.Content && oChild.Run.Content.length) { // ParaRun
						var drawing = oChild.Run.Content[0]
						if (drawing.docPr && drawing.docPr.title && drawing.GraphicObj) {
							var titleObj = getJsonData(drawing.docPr.title)
							if (titleObj.feature && titleObj.feature.sub_type == 'ask_accurate') {
								var content = drawing.GraphicObj.textBoxContent // shapeContent
								if (content.Content && content.Content.length) {
									var paragraph = content.Content[0]
									if (paragraph) {
										if (paragraph.GetElementsCount()) {
											var run = paragraph.GetElement(0)
											if (run) {
												if (run.GetText() * 1 != 1) { // 序号run
													paragraph.ReplaceCurrentWord(0, '1')
												}
											}
										} else {
											var oPara = Api.LookupObject(paragraph.Id)
											oPara.AddText('1')
											oPara.SetColor(153, 153, 153, false)
											oPara.SetJc('center')
											oPara.SetSpacingAfter(0)
										}
									}
								}
							}
						}
					}
				}
			} else {
				// todo.. write为blockLvlSdt时需要另外处理
			}
		}
		for (var i = 0; i < ids.length; ++i) {
			var id = ids[i]
			if (question_map[id].level_type != 'question' || question_map[id].question_type == 0) {
				continue
			}
			var ques_mode = 0
			if (mark_type_info && mark_type_info.list) {
				var find = mark_type_info.list.find(e => {
					return e.question_type_id == question_map[id].question_type
				})
				ques_mode = find ? find.ques_mode : 3
			}
			// 不是单选也不是多选
			if (ques_mode != 1 && ques_mode != 5) {
				continue
			}
			var nodeData = node_list.find(e => {
				return e.id == id
			})
			if (!nodeData) {
				continue
			}
			var oControl = controls.find(e => {
				var tag = getJsonData(e.GetTag())
				return e.GetClassType() == 'blockLvlSdt' && e.GetPosInParent() >= 0 && tag.client_id == nodeData.id
			})
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
				continue
			}
			var childControls = oControl.GetAllContentControls()
			if (!childControls || childControls.length == 0) {
				continue
			}
			var validIds = []
			var write_list = nodeData.write_list || []
			for (var j = 0; j < childControls.length; ++j) {
				var childTag = getJsonData(childControls[j].GetTag())
				if (childTag.client_id) {
					var childIndex = question_map[id].ask_list.findIndex(e => {
						return e.id == childTag.client_id
					})
					if (childIndex >= 0) {
						validIds.push({
							control_id: childControls[j].Sdt.GetId(),
							child_index: childIndex,
							client_id: childTag.client_id
						})
					} else {
						deleteControlAccurate(childControls[j])
						Api.asc_RemoveContentControlWrapper(childControls[j].Sdt.GetId())
						var write_index = write_list.findIndex(e => {
							return e.id = childTag.client_id
						})
						if (write_index >= 0) {
							write_list.splice(write_index, 1)
						}
					}
				} else {
					deleteControlAccurate(childControls[j])
					if (childTag.regionType != 'num') {
						Api.asc_RemoveContentControlWrapper(childControls[j].Sdt.GetId())
					}
				}
			}
			if (validIds.length > 1) {
				var begin
				var end
				var blankIndex = 0
				if (choice_blank == 'first') {
					begin = 1
					end = validIds.length -1
					question_map[id].ask_list = [].concat(question_map[id].ask_list[validIds[0].child_index])
				} else if (choice_blank == 'last') {
					begin = 0
					end = validIds.length - 2
					blankIndex = validIds.length - 1
					question_map[id].ask_list = [].concat(question_map[id].ask_list[validIds[validIds.length - 1].child_index])
				}
				for (var k = begin; k <= end; ++k) {
					var write_index = write_list.findIndex(e => {
						return e.id == validIds[k].client_id
					})
					if (write_index >= 0) {
						write_list.splice(write_index, 1)
					}
					deleteControlAccurate(Api.LookupObject(validIds[k].control_id))
					Api.asc_RemoveContentControlWrapper(validIds[k].control_id)
				}
				if (question_map[id].interaction != 'none') {
					updateAccurateText(validIds[blankIndex].control_id)
				}
			}
			nodeData.write_list = write_list
		}
		return {
			node_list,
			question_map,
			updateInteraction
		}
	}, false, recalc).then(res => {
		return new Promise((resolve, reject) => {
			if (res) {
				window.BiyueCustomData.node_list = res.node_list
				window.BiyueCustomData.question_map = res.question_map
			}
			resolve(res)
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
			var titleObj = getJsonData(title)
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
		var oTables = Api.GetDocument().GetAllTables() || []
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
								if (oCell && oCell.GetClassType && oCell.GetClassType() == 'tableCell') {
									var oTable = oCell.GetParentTable()
									if (oTable && oTable.GetPosInParent() == -1) {
										oCell = getCell(writeData)
									}
									if (oCell) {
										oCell.SetBackgroundColor(255, 191, 191, cmdType == 'show' ? false : true)
									}
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
		var oControl = Api.LookupObject(curControl.Id)
		if (oControl) {
			var tag = getJsonData(oControl.GetTag() || '{}')
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
						var titleObj = getJsonData(title)
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
						var tag = getJsonData(paraentControl.GetTag())
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
										var dtitle = getJsonData(
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
  	// 先关闭智批元素，避免智批元素在全量更新的时候被带到题目里 更新之后再打开
  	setBtnLoading('uploadTree', true)
	setInteraction('none', null, false).then(() => {
		return addOnlyBigControl(false)
	}).then(() => {
		return handleUploadPrepare('hide')
	}).then(() => {
		return getControlListForUpload()
	}).then(control_list => {
		if (control_list && control_list.length) {
			generateTreeForUpload(control_list).then(() => {
				handleCompleteResult('全量更新成功')
			  }).catch((res) => {
				handleCompleteResult(res && res.message && res.message != '' ? res.message : '全量更新失败')
			  })
		  } else {
			handleCompleteResult('未找到可更新的题目，请检查题目列表')

		  }
	})
}

function getControlListForUpload() {
	Asc.scope.node_list = window.BiyueCustomData.node_list
    Asc.scope.question_map = window.BiyueCustomData.question_map
	return biyueCallCommand(window, function() {
		var target_list = []
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		var question_map = Asc.scope.question_map
		console.log('question_map', question_map)
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
		for (var i = 0, imax = controls.length; i < imax; ++i) {
			var oControl = controls[i]
			var tag = getJsonData(oControl.GetTag() || '{}')
			if (tag.regionType != 'question' || !tag.client_id) {
				continue
			}
			var quesData = question_map[tag.client_id]
			if (!quesData) {
				continue
			}
			if (!question_map[tag.client_id].level_type) {
				continue
			}
			var oParentControl = oControl.GetParentContentControl()
			var parent_id = 0
			if (oParentControl) {
				var parentTag = getJsonData(oParentControl.GetTag() || '{}')
				parent_id = parentTag.client_id
			} else {
				// 根据level, 查找在它前面的比它lvl小的struct
				for (var j = target_list.length - 1; j >= 0; --j) {
					var preNode = target_list[j]
					// 由于struct未必有lvl，因此先将Lvl的判断移除
					if (preNode.content_type == 'struct') {
						parent_id = preNode.id
						break
					}
				}
			}
			var useControl = oControl
			if (tag.big) {
				var childcontrols = oControl.GetAllContentControls() || []
				var bigControl = childcontrols.find(e => {
					var btag = getJsonData(e.GetTag())
					console.log('btag', e.GetTag())
					return e.GetClassType() == 'blockLvlSdt' && btag.onlybig == 1 && btag.link_id == tag.client_id
				})
				if (bigControl) {
					useControl = bigControl
				}
			}
			var oRange = useControl.GetRange()
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
				content_text: oRange.GetText(),
				question_type: question_map[tag.client_id].question_type,
				question_name: question_map[tag.client_id].ques_name || question_map[tag.client_id].ques_default_name,
				control_id: oControl.Sdt.GetId(),
				lvl: tag.lvl
			})
		}
		console.log('target_list', target_list)
		return target_list
	  }, false, false)
}

// 清洗输出的html
function cleanHtml(html) {
  // 创建一个临时的div用以装载需要处理的HTML内容
  var tempDiv = document.createElement('div');

  tempDiv.innerHTML = html

  //如果没有子节点或者文本内容就可以删除的元素
  const removeEmpty = { div: 1, a: 1, abbr: 1, acronym: 1, address: 1, b: 1, bdo: 1, big: 1, cite: 1, code: 1, del: 1, dfn: 1, em: 1, font: 1, i: 1, ins: 1, label: 1, kbd: 1, q: 1, s: 1, samp: 1, small: 1, span: 1, strike: 1, strong: 1, sub: 1, sup: 1, tt: 1, u: 1, 'var': 1 };

  // 替换部分标签 为 p 标签
  tempDiv.querySelectorAll('h1, h2, h3, h4, h5, li').forEach(el => {
    const p = document.createElement('p');
    while(el.firstChild) {
      p.appendChild(el.firstChild);
    }
    el.parentNode.replaceChild(p, el);
  });

  // 移除所有 div, ul, ol 标签但是保留内容
  tempDiv.querySelectorAll('div, ul, ol').forEach(el => {
    while(el.firstChild) {
      el.parentNode.insertBefore(el.firstChild, el);
    }
    el.parentNode.removeChild(el);
  });

  // 移除 span 标签但保留内容
  tempDiv.querySelectorAll('span').forEach(el => {
    while(el.firstChild) {
      el.parentNode.insertBefore(el.firstChild, el);
    }
    el.parentNode.removeChild(el);
  });

  // 移除所有带data-zone_type="question"属性的标签
  tempDiv.querySelectorAll('[data-zone_type="question"]').forEach(el => {
    el.parentNode.removeChild(el);
  });

  // 移除所有带data-属性的元素属性
  let data_ignore_list = ['data-client_id', 'data-ques_use'] // 需要保留的data属性
  tempDiv.querySelectorAll('*').forEach(el => {
    Array.from(el.attributes).forEach(attr => {
      if (attr.name.startsWith('data-') && !data_ignore_list.includes(attr.name)) {
        el.removeAttribute(attr.name);
      }
    });
  });

  // // 移除所有style属性
  // tempDiv.querySelectorAll('[style]').forEach(el => {
  //   el.removeAttribute('style');
  // });

  // 只保留特定的 style 属性
  tempDiv.querySelectorAll('[style]').forEach(el => {
    const style = el.getAttribute('style');
    const allowedStyles = extractAllowedStyles(style);
    if (allowedStyles) {
      el.setAttribute('style', allowedStyles);
    } else {
      el.removeAttribute('style');
    }
  });

  // 移除无内容的特定标签
  Object.keys(removeEmpty).forEach(tag => {
    tempDiv.querySelectorAll(tag).forEach(el => {
      if (!el.textContent.trim()) {
        el.parentNode.removeChild(el);
      }
    });
  });

  flattenNestedP(tempDiv)

  return tempDiv.innerHTML
}

function extractAllowedStyles(style) {
  // 允许保留的样式属性列表
  const allowedProperties = ['text-align'];
  const styleRules = style.split(';');
  const filteredStyles = styleRules.filter(rule => {
    const [property] = rule.split(':');
    return allowedProperties.includes(property.trim());
  });
  return filteredStyles.join(';').trim();
}

function flattenNestedP(node) {
  // 如果有重复嵌套的p标签则保留最里面那层
  node.querySelectorAll('p').forEach(p => {
    if (p.querySelector('p')) {
      let childP = p.querySelector('p');
      p.parentNode.insertBefore(childP, p);
      p.parentNode.removeChild(p);
      flattenNestedP(node);
    }
  });
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
	return new Promise((resolve, reject) => {
		if (!control_list) {
			reject(null)
		}
		var tree = []
		control_list.forEach((e) => {
			e.content_html = cleanHtml(e.content_html || '')
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
			resolve(res)
		}).catch(res => {
			console.log('reqComplete fail', res)
			console.log('[reqUploadTree end]', Date.now())
			reject(res)
		})
	})
}

function handleCompleteResult(message) {
	setBtnLoading('uploadTree', false)
	if (message) {
		window.biyue.showMessageBox({
			content: message,
			showCancel: false
		})
	}
	removeOnlyBigControl().then(() => {
		return handleUploadPrepare('show')
	}).then(() => {
		return setInteraction('useself')
	})
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
		// console.log('============== ++++++++ 17', PageSize)
		var tw = PageSize.W - PageMargins.Left - PageMargins.Right
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
				var tag = getJsonData(e.GetTag())
				return tag.client_id == id && e.GetClassType() == 'blockLvlSdt'
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
					// console.log('============== ++++++++ 18', tables[iTable], newW)
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
			// console.log('============== ++++++++ 19', tables[i])
			oTable.SetWidth('percent', tables[i].W)
			tables[i].cells.forEach((cell, cidx) => {
				// console.log('============== ++++++++ 20', cell)
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
		function AddControlToCell(oControl, oCell) {
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
				return
			}
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
			if (!startTable || !oControl || oControl.GetClassType() != 'blockLvlSdt') {
				return
			}
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
							// console.log('============== ++++++++ 11', tables[iTable], newW)
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
					// console.log('============== ++++++++ 10', oCell)
					var TableCellW = oCell.CellPr.TableCellW
					if (!TableCellW) {
						TableCellW = oCell.Cell.CompiledPr.Pr.TableCellW
					}
					if (!TableCellW) {
						continue
					}
					var W = TableCellW.W
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
								// console.log('============== ++++++++ 9', tables[iTable], W)
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
									// console.log('============== ++++++++ 8', tables[iTable], W2)
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
							// console.log('============== ++++++++ 7', tables[iTable], W)
							tables[iTable].W += W
						}
						// console.log('============== ++++++++ 6', W)
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
				// console.log('============== ++++++++ 5', tables[i])
				oTable.SetWidth('percent', tables[i].W)
				tables[i].cells.forEach((cell, cidx) => {
					AddControlToCell2(templist[count].content, oTable.GetCell(0, cell.icell))
					count++
					// console.log('============== ++++++++ 4', cell, oTable.Table.CalculatedTableW)
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
				var tag = getJsonData(e.GetTag())
				return tag.client_id == id && e.GetClassType() == 'blockLvlSdt'
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
						// console.log('============== ++++++++ 3', previousElement)
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
							// console.log('============== ++++++++ 2', pre)
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
					// console.log('============== ++++++++ 1', firstCell)
					var TableCellW = firstCell.CellPr.TableCellW
					if (!TableCellW) {
						TableCellW = firstCell.Cell.CompiledPr.Pr.TableCellW
					}
					var firstCellW = TableCellW ? TableCellW.W : 100
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

function deleteAsks(askList, recalc = true, notify = true) {
	if (!askList || askList.length == 0) {
		return new Promise((resolve, reject) => {
			return resolve({})
		})
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
		var allControls = oDocument.GetAllContentControls() || []
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
		function deleteAccurateRun(oRun) {
			if (oRun &&
				oRun.Run &&
				oRun.Run.Content &&
				oRun.Run.Content[0] &&
				oRun.Run.Content[0].docPr) {
				var title = oRun.Run.Content[0].docPr.title
				if (title) {
					var titleObj = getJsonData(title)
					if (titleObj.feature && titleObj.feature.sub_type == 'ask_accurate') {
						oRun.Delete()
						return true
					}
				}
			}
			return false
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
							if (child.GetClassType() == 'run' && child.Run.Id == run.Id) {
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
						if (oParagraph) {
							oParagraph.RemoveElement(ipos)
						}
						return true
					}
				}
			}
			oShape.Delete()
			return true
		}
		function clearQuesInteraction(oControl) {
			if (!oControl) {
				return
			}
			if (oControl.GetClassType() == 'inlineLvlSdt') {
				var elementCount = oControl.GetElementsCount()
				for (var idx = 0; idx < elementCount; ++idx) {
					var oRun = oControl.GetElement(idx)
					deleteAccurateRun(oRun)
				}
			} else if (oControl.GetClassType() == 'blockLvlSdt') {
				var oControlContent = oControl.GetContent()
				if (!oControlContent) {
					return
				}
				var drawings = oControlContent.GetAllDrawingObjects()
				if (drawings) {
					for (var j = 0, jmax = drawings.length; j < jmax; ++j) {
						var oDrawing = drawings[j]
						if (oDrawing.Drawing.docPr) {
							var titleObj = getJsonData(oDrawing.Drawing.docPr.title)
							if (titleObj.feature && titleObj.feature.zone_type == 'question') {
								oDrawing.Delete()
							}
						}
					}
				}
			}
		}
		function removeCellAskRecord(oCell) {
			var oTable = oCell.GetParentTable()
			if (oTable && oTable.GetPosInParent() >= 0) {
				var desc = getJsonData(oTable.GetTableDescription())
				var key = `${oCell.GetRowIndex()}_${oCell.GetIndex()}`
				if (desc[key]) {
					delete desc[key]
				}
				oTable.SetTableDescription(JSON.stringify(desc))
			}
		}
		function removeCellInteraction(oCell) {
			if (!oCell) {
				return
			}
			oCell.SetBackgroundColor(255, 191, 191, true)
			var cellContent = oCell.GetContent()
			var paragraphs = cellContent.GetAllParagraphs()
			paragraphs.forEach(oParagraph => {
				var childCount = oParagraph.GetElementsCount()
				for (var i = 0; i < childCount; ++i) {
					var oRun = oParagraph.GetElement(i)
					deleteAccurateRun(oRun)
				}
			})
			removeCellAskRecord(oCell)
		}
		function getControl(client_id, control_id) {
			if (client_id) {
				if (control_id) {
					var oControl = Api.LookupObject(control_id)
					if (oControl && oControl.GetTag) {
						var pos = oControl.GetClassType() == 'inlineLvlSdt' ? oControl.Sdt.GetPosInParent() : oControl.GetPosInParent()
						if (pos >= 0) {
							var tag = getJsonData(oControl.GetTag())
							if (tag.client_id == client_id) {
								return oControl
							}
						}
					}
				}
				var oControl = allControls.find(e => {
					var tag = getJsonData(e.GetTag())
					return tag.client_id == client_id
				})
				return oControl
			}
			return null
		}
		for (var i = 0, imax = delete_ask_list.length; i < imax; ++i) {
			var qid = delete_ask_list[i].ques_id
			var aid = delete_ask_list[i].ask_id
			if (question_map[qid]) {
				if (aid == 0) {
					question_map[qid].score = 0
				} else {
					var askIndex = question_map[qid].ask_list.findIndex(e => e.id == aid)
					if (askIndex >= 0) {
						question_map[qid].ask_list.splice(askIndex, 1)
						var sum = 0
						question_map[qid].ask_list.forEach(e => {
							sum += (e.score || 0)
						})
						question_map[qid].score = sum
					}
				}
			}
			var nodeData = node_list.find(e => {
				return e.id == qid
			})
			if (!nodeData) {
				continue
			}
			var quesControl = getControl(qid, nodeData.control_id)
			if (!quesControl || quesControl.GetClassType() != 'blockLvlSdt') {
				continue
			}
			if (aid == 0) {
				// 删除所有小问，遍历出所有小问，全部删除
				// 删除所有精准互动
				var childDrawings = quesControl.GetAllDrawingObjects() || []
				for (var j = 0; j < childDrawings.length; ++j) {
					var tag = getJsonData(childDrawings[j].Drawing.docPr.title)
					if (tag.feature && tag.feature.zone_type == 'question') {
						if (tag.feature.sub_type == 'ask_accurate') {
							deleShape(childDrawings[j])
						}
					}
				}
				// 删除除订正框外的所有inlineControl
				var childControls = quesControl.GetAllContentControls() || []
				childControls.forEach(e => {
					if (e.Sdt) {
						var tag = getJsonData(e.GetTag())
						if (e.GetClassType() == 'inlineLvlSdt' && tag.regionType != 'num') {
							Api.asc_RemoveContentControlWrapper(e.Sdt.GetId())
						} else if (e.GetClassType() == 'blockLvlSdt' && tag.regionType == 'write') {
							Api.asc_RemoveContentControlWrapper(e.Sdt.GetId())
						}
					}
				})
				// 删除所有write 和 identify
				childDrawings = quesControl.GetAllDrawingObjects() || []
				for (var j = 0; j < childDrawings.length; ++j) {
					var tag = getJsonData(childDrawings[j].Drawing.docPr.title)
					if (tag.feature && tag.feature.zone_type == 'question') {
						if (tag.feature.sub_type == 'write') {
							var run = childDrawings[j].Drawing.GetRun()
							if (run) {
								var paragraph = run.GetParagraph()
								if (paragraph) {
									var oParagraph = Api.LookupObject(paragraph.Id)
									var ipos = run.GetPosInParent()
									if (ipos >= 0) {
										childDrawings[j].Delete()
										if (oParagraph) {
											oParagraph.RemoveElement(ipos)
										}
									}
								}
							}
						} else if (tag.feature.sub_type == 'identify') {
							childDrawings[j].Delete()
						}
					}
				}
				// 删除所有单元格小问
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
											removeCellAskRecord(oCell)
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
				nodeData.write_list = []
				question_map[qid].ask_list = []
			} else {
				if (!nodeData.write_list) {
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
					var oDrawing = drawings.find(e => {
						var tag = getJsonData(e.Drawing.docPr.title)
						return tag.feature && tag.feature.client_id == writeData.id
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
										if (oParagraph) {
											oParagraph.RemoveElement(ipos)
										}
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
	}, false, recalc).then(res => {
		if (res) {
			window.BiyueCustomData.question_map = res.question_map
			window.BiyueCustomData.node_list = res.node_list
			if (notify) {
				document.dispatchEvent(new CustomEvent('updateQuesData', {
					detail: {
						client_id: res.ques_id
					}
				}))
			}
			if (recalc && res.question_map[res.ques_id] && res.question_map[res.ques_id].interaction == 'accurate') {
				return setInteraction('accurate', [res.ques_id])
			} else {
				return new Promise((resolve, reject) => {
					return resolve({})
				})
			}
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
		var controls = oDocument.GetAllContentControls() || []
		var oTables = oDocument.GetAllTables() || []
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
		if (write_data.sub_type == 'control') {
			var oControls = controls.filter(e => {
				var tag = getJsonData(e.GetTag())
				if (tag.client_id == write_data.id && e.Sdt) {
					if (e.GetClassType() == 'blockLvlSdt') {
						return e.GetPosInParent() >= 0
					} else if (e.GetClassType() == 'inlineLvlSdt') {
						return e.Sdt.GetPosInParent() >= 0
					}
				}
			})
			if (oControls && oControls.length) {
				if (oControls.length == 1) {
					oDocument.Document.MoveCursorToContentControl(oControls[0].Sdt.GetId(), true)
				}
			}
		} else if (write_data.sub_type == 'cell') {
			if (write_data.cell_id) {
				var oCell = Api.LookupObject(write_data.cell_id)
				if (oCell && oCell.GetClassType() == 'tableCell') {
					var table = oCell.GetParentTable()
					if (table.GetPosInParent() == -1) {
						oCell = getCell(write_data)
					}
					if (oCell) {
						var cellContent = oCell.GetContent()
						if (cellContent) {
							cellContent.GetRange().Select()
						}
					}
				}
			}
		} else if (write_data.sub_type == 'write' || write_data.sub_type == 'identify') {
			var oDrawing = drawings.find(e => {
				var tag = getJsonData(e.Drawing.docPr.title)
				return tag.feature && tag.feature.client_id == write_data.id
			})
			if (oDrawing) {
				oDrawing.Select()
			}
		}
	}, false, false)
}

function handleImageIgnore(cmdType) {
	Asc.scope.cmdType = cmdType
	return biyueCallCommand(window, function() {
		var cmdType = Asc.scope.cmdType
		console.log('handleImageIgnore', cmdType)
		var oDocument = Api.GetDocument()
		var drawings = oDocument.GetAllDrawingObjects() || []
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
		drawings.forEach(oDrawing => {
			var GraphicObj = oDrawing.Drawing.GraphicObj
			if (GraphicObj && GraphicObj.selected) {
				var tag = getJsonData(oDrawing.Drawing.docPr.title)
				if (cmdType == 'add') {
					if (tag.feature) {
						tag.feature.partical_no_dot = 1
					} else {
						tag = {
							feature: {
								partical_no_dot: 1
							}
						}
					}
				} else {
					if (tag.feature && tag.feature.partical_no_dot) {
						delete tag.feature.partical_no_dot
					}
				}
				oDrawing.Drawing.Set_Props({
					title: JSON.stringify(tag)
				})
				var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 255, 255))
        		oFill.UniFill.transparent = 0 // 透明度
				var oStroke = Api.CreateStroke(10000, cmdType == 'add' ? Api.CreateSolidFill( Api.CreateRGBColor(255, 111, 61)) : oFill);
				oDrawing.SetOutLine(oStroke);
			}
		})
	}, false, false)
}
// todo。。分栏需要考虑的因素很多，需要之后再考虑
function setSectionColumn(column) {
	Asc.scope.column = column
	return biyueCallCommand(window, function() {
		var column = Asc.scope.column
		var oDocument = Api.GetDocument()
		Api.pluginMethod_MoveCursorToStart()
		var paragraph = oDocument.Document.GetCurrentParagraph()
		if (!paragraph) {
			return
		}
		var oParagraph = Api.LookupObject(paragraph.Id)
		var pageCount = oParagraph.Paragraph.getPageCount()
		if (pageCount > 1) {

		}
		var oSection = oParagraph.GetSection()
		var column_num =  oSection.Section.GetColumnsCount()
		if (column == column_num) {
			return
		} else {
			oSection.Section.Set_Columns_EqualWidth(true);
            oSection.Section.Set_Columns_Num(column);
            oSection.Section.Set_Columns_Space((25.4 / 72 / 20) * 640)
			oSection.Section.Set_Columns_Sep(true)
		}
	}, false, true)
}
function layoutDetect() {
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var oRange = oDocument.GetRangeBySelect()
		if (!oRange) {
			var currentContentControl = oDocument.Document.GetContentControl()
			oRange = Api.LookupObject(currentContentControl.Id).GetRange()
		}
		if (!oRange) {
			return null
		}
		var paragraphs = oRange.Paragraphs || []
		var result = {
			has32: false, // 存在空格具有下划线属性
			has160: false, // 存在ASCII码160，这是一个不可打印字符，称为不间断空格（NBSP）
			has65307: false, // 中文分号
			has12288: false, // 中文空格
			hasTab: false, // 存在tab键
			hasWhiteBg: false, // 存在背景为白色的段落
			hasSmallImage: false, // 存在宽高过小的图片
		}
		function isWhite(shd) {
			return shd && shd.Fill && shd.Fill.r == 255 && shd.Fill.g == 255 && shd.Fill.b == 255 && shd.Fill.Auto == false
		}
		var bflag = []
		function handleRun(oRun) {
			if (!oRun || oRun.GetClassType() != 'run') {
				return
			}
			var textpr = oRun.GetTextPr()
			if (textpr && textpr.TextPr) {
				if (isWhite(textpr.TextPr.Shd)) {
					result.hasWhiteBg = true
				}
			}
			var runContent = oRun.Run.Content || []
			var isUnderline = oRun.GetUnderline()
			// 判断是否在括号内存在tab
			for (var k = 0, kmax = runContent.length; k < kmax; ++k) {
				var type = runContent[k].GetType()
				if (type == 22) { // drawing
					if (runContent[k].Width <= 1 && runContent[k].Height <= 1) {
						result.hasSmallImage = true
						find = true
					}
				} else if (type == 1) {
					if (runContent[k].Value == 40 || runContent[k].Value == 65288) {
						bflag.push(1)
					} else if (runContent[k].Value == 41 || runContent[k].Value == 65289) {
						if (bflag.length) {
							if (bflag[bflag.length - 1] == 1) {
								bflag.pop()
							} else if (bflag[bflag.length - 1] == 2) {
								result.hasTab = true
								find = true
								break
							}
						}
					}
				} else if (type == 21) {
					if (isUnderline) {
						result.has32 = true
						find = true
					}
					if (bflag.length) {
						if (bflag[bflag.length - 1] == 1) {
							bflag.push(2)
						}
					}
				}
				if (runContent[k].Value == 32) {
					if (isUnderline) {
						result.has32 = true
						find = true
					}
				} else if ( runContent[k].Value && !result[`has${runContent[k].Value}`]) {
					result[`has${runContent[k].Value}`] = true
					find = true
				}
			}
		}
		for (var i = 0, imax = paragraphs.length; i < imax; ++i) {
			var oParagraph = paragraphs[i]
			bflag = []
			var controls = oParagraph.GetAllContentControls() || []
			controls.forEach(oControl => {
				if (oControl.GetClassType() == 'inlineLvlSdt') {
					var count1 = oControl.GetElementsCount()
					if (count1) {
						for (var k = 0; k < oControl.GetElementsCount(); ++k) {
							handleRun(oControl.GetElement(k))
						}
					}
				}
			})
			for (var j = 0; j < oParagraph.GetElementsCount(); ++j) {
				handleRun(oParagraph.GetElement(j))
			}
			// 判断是否为白色背景
			var paraPr = oParagraph.GetParaPr()
			if (paraPr && paraPr.ParaPr && isWhite(paraPr.ParaPr.Shd)) {
				result.hasWhiteBg = true
			}
		}
		return result
	}, false, false).then(res => {
		Asc.scope.layout_detect_result = res
		window.biyue.showDialog('layoutRepairWindow', '字符检测', 'layoutRepair.html', 250, 400, true)
	})
}

// 排版修复
function layoutRepair(cmdData) {
	Asc.scope.cmdData = cmdData
	return biyueCallCommand(window, function() {
		var cmdData = Asc.scope.cmdData
		var oDocument = Api.GetDocument()
		var oRange = oDocument.GetRangeBySelect()
		if (!oRange) {
			var currentContentControl = oDocument.Document.GetContentControl()
			oRange = Api.LookupObject(currentContentControl.Id).GetRange()
		}
		if (!oRange) {
			return null
		}
		var paragrahs = oRange.Paragraphs || []
		var oneSpaceWidth = 2.11 // 一个空格的宽度
		var bflag = []
		function getTabReplaceTarget(width, target) {
			var str = ''
			var count = Math.ceil(width / oneSpaceWidth)
			for (var i = 0; i < count; ++i) {
				str += target
			}
			return str
		}
		// 将tab下划线替换为真的下划线
		function replaceUnderline(oRun, parent, pos) {
			if (!oRun || oRun.GetClassType() != 'run') {
				return
			}
			var runContent = oRun.Run.Content || []
			if (!oRun.GetUnderline()) {
				return
			}
			var find = false
			var split = false
			for (var k = 0; k < runContent.length; ++k) {
				var len = runContent.length
				var element2 = runContent[k]
				var elementType = element2.GetType()
				if (elementType == 21 || elementType == 2) {
					if (!find && k > 0) {
						var newRun = oRun.Run.Split_Run(k + 1)
						parent.Add_ToContent(pos + 1, newRun)
						oRun.Run.RemoveElement(element2)
						var str = elementType == 2 ? '_' : getTabReplaceTarget(element2.Width, '_')
						oRun.Run.AddText(str, k)
						return
					}
					find = true
					oRun.Run.RemoveElement(element2)
					var str = elementType == 2 ? '_' : getTabReplaceTarget(element2.Width, '_')
					oRun.Run.AddText(str, k)
					var addCount = str.length - 1
					if (k < len - 1) {
						var nextType = runContent[k + 1 + addCount].GetType()
						if (nextType != 2 && nextType != 21) {
							oRun.SetUnderline(false)
							var newRun = oRun.Run.Split_Run(k + 1 + addCount)
							newRun.SetUnderline(true)
							parent.Add_ToContent(pos + 1, newRun)
							split = true
							break
						} else {
							k += addCount
						}
					}
				}
			}
			if (find) {
				if (!split) {
					oRun.SetUnderline(false)
				}
			}
		}
		function handleRun(oRun, parent, pos) {
			if (!oRun || oRun.GetClassType() != 'run') {
				return
			}
			if (cmdData.type == 1 && cmdData.value == 'whitebg') {
				var textpr = oRun.GetTextPr()
				if (textpr && textpr.TextPr) {
					var shd = textpr.TextPr.Shd
					if (shd && shd.Fill && shd.Fill.r == 255 && shd.Fill.g == 255 && shd.Fill.b == 255 && shd.Fill.Auto == false) {
						oRun.Run.Set_Shd(undefined);
					}
				}
				return
			}
			if (cmdData.type == 1 && cmdData.value == 32) {
				replaceUnderline(oRun, parent, pos)
				return
			}
 			var runContent = oRun.Run.Content || []
			for (var k = 0; k < runContent.length; ++k) {
				var element2 = runContent[k]
				if (element2.GetType() == 22) { // drawing
					if (cmdData.type == 2 && cmdData.value == 'smallimg') {
						if (element2.Width <= 1 && element2.Height <= 1) {
							element2.PreDelete();
							oRun.Run.RemoveElement(element2)
						}
					}
				}
				if (cmdData.type == 1 && cmdData.value != 32) {
					if (element2.Value == cmdData.value) {
						oRun.Run.RemoveElement(element2)
						if (cmdData.newValue == 32) {
							var replaceStr = cmdData.value == 12288 ? '  ' : ' '
							oRun.Run.AddText(replaceStr, k)
							k += replaceStr.length - 1
						} else if (cmdData.newValue == 59) {
							oRun.Run.AddText(';', k)
						}
					}
				}
			}
		}
		function hasRightBracket2(begin, runContents) {
			for (var k = begin; k < runContents.length; ++k) {
				if (runContents[k].GetType() == 1 && (runContents[k].Value == 41 || runContents[k].Value == 65289)) {
					return true
				}
			}
		}
		// 之后是否存在右括号
		function hasRightBracket(oParagraph, iElement, jRun, idxParent) {
			for (var j = iElement; j < oParagraph.GetElementsCount(); ++j) {
				var oElement = oParagraph.GetElement(j)
				if (oElement.GetClassType() == 'run') {
					var runContents = oElement.Run.Content || []
					var begin = iElement == j ? jRun : 0
					if (hasRightBracket2(begin, runContents)) {
						return true
					}
				} else if (oElement.GetClassType() == 'inlineLvlSdt') {
					var count2 = oElement.GetElementsCount()
					var begin = iElement == j ? idxParent : 0
					for (var idx = begin; idx < count2; ++idx) {
						if (oElement.GetElement(idx).GetClassType() == 'run') {
							var runContents = oElement.GetElement(idx).Run.Content || []
							var begin = iElement == j ? idxParent : 0
							if (hasRightBracket2(begin, runContents)) {
								return true
							}
						}
					}
				}
			}
		}
		function handleRun2(oParagraph, j, oElement, bflag, idxParent) {
			var runContents = oElement.Run.Content || []
			for (var k = 0; k < runContents.length; ++k) {
				var type = runContents[k].GetType()
				if (type == 1) {
					var value = runContents[k].Value
					if (value == 65288 || value == 40) { // 左括号
						bflag.push(1)
					} else if (value == 65289 || value == 41) { // 右括号
						bflag.pop()
					}
				} else if (type == 21 && bflag.length) {
					// 段落之后是否存在右括号
					var findRight = hasRightBracket(oParagraph, j, k, idxParent)
					if (findRight) {
						var element2 = runContents[k]
						oElement.Run.RemoveElement(element2)
						var str = getTabReplaceTarget(element2.Width, ' ')
						oElement.Run.AddText(str, k)
						k += str.length - 1
					}
				}
			}
		}
		if (cmdData.type == 1 && cmdData.value == 'tab') { // 将括号里的tab替换为空格
			var bflag = []
			for (var i = 0, imax = paragrahs.length; i < imax; ++i) {
				var oParagraph = paragrahs[i]
				for (var j = 0; j < oParagraph.GetElementsCount(); ++j) {
					var oElement = oParagraph.GetElement(j)
					if (oElement.GetClassType() == 'run') {
						handleRun2(oParagraph, j, oElement, bflag)
					} else if (oElement.GetClassType() == 'inlineLvlSdt') {
						var count2 = oElement.GetElementsCount()
						for (var idx = 0; idx < count2; ++idx) {
							if (oElement.GetElement(idx).GetClassType() == 'run') {
								handleRun2(oParagraph, j, oElement.GetElement(idx), bflag)
							}
						}
					}
				}
			}
		} else {
			for (var i = 0, imax = paragrahs.length; i < imax; ++i) {
				var oParagraph = paragrahs[i]
				if (cmdData.type == 1 && cmdData.value == 'whitebg') {
					var oParaPr = oParagraph.GetParaPr();
					if (oParaPr) {
						oParaPr.SetShd("clear", 255, 255, 255, true);
					}
				}
				var controls = oParagraph.GetAllContentControls() || []
				controls.forEach(oControl => {
					if (oControl.GetClassType() == 'inlineLvlSdt') {
						var count1 = oControl.GetElementsCount()
						if (count1) {
							for (var k = 0; k < oControl.GetElementsCount(); ++k) {
								handleRun(oControl.GetElement(k), oControl.Sdt, k)
							}
						}
					}
				})
				for (var j = 0; j < oParagraph.GetElementsCount(); ++j) {
					handleRun(oParagraph.GetElement(j), oParagraph.Paragraph, j)
				}
			}
		}
	}, false, true)
}

function batchChangeScore() {
	if (g_click_value && g_click_value.Tag && g_click_value.Tag.client_id) {
		window.biyue.onBatchScoreSet(g_click_value.Tag.client_id)
	} else {
		window.biyue.onBatchScoreSet()
	}
}

function imageRelation() {
	window.biyue.showDialog('imageRelationWindow', '图片关联', 'imageRelation.html', 800, 600, false)
}

function tableRelation() {
	window.biyue.showDialog('imageRelationWindow', '表格关联', 'imageRelation.html', 800, 600, false)
}

function tagImageCommon(params) {
	Asc.scope.tag_params = params
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	return biyueCallCommand(window, function() {
		var tag_params = Asc.scope.tag_params
		var client_node_id = Asc.scope.client_node_id
		var oDocument = Api.GetDocument()
		var drawings = oDocument.GetAllDrawingObjects() || []
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
		if (tag_params.target_type == 'table') {
			var oTable = Api.LookupObject(tag_params.target_id)
			if (oTable && oTable.GetClassType && oTable.GetClassType() == 'table') {
				var title = getJsonData(oTable.GetTableTitle())
				if (!title.client_id) {
					client_node_id += 1
					title.client_id = client_node_id
				}
				title.ques_use = tag_params.ques_use.join('_')
				oTable.SetTableTitle(JSON.stringify(title))
			}
		} else {
			var oDrawing = drawings.find(e => {
				return e.Drawing.Id == tag_params.target_id
			})
			if (oDrawing) {
				var tag = getJsonData(oDrawing.Drawing.docPr.title)
				if (tag.feature) {
					if (!tag.feature.client_id) {
						client_node_id += 1
						tag.feature.client_id = client_node_id
					}
					tag.feature.ques_use = tag_params.ques_use.join('_')
				} else {
					client_node_id += 1
					tag = {
						feature: {
							ques_use: tag_params.ques_use.join('_'),
							client_id: client_node_id
						}
					}
				}
				oDrawing.Drawing.Set_Props({
					title: JSON.stringify(tag)
				})
			}
		}
		return {
			client_node_id: client_node_id,
			drawing_id: tag_params.target_id,
			ques_use: tag_params.ques_use
		}
	}, false, false).then(res => {
		if (res) {
			window.BiyueCustomData.client_node_id = res.client_node_id
			if (!window.BiyueCustomData.image_use) {
				window.BiyueCustomData.image_use = {}
			}
			window.BiyueCustomData.image_use[res.drawing_id] = res.ques_use
		}
	})
}

function updateQuesScore(ids) {
	Asc.scope.node_list = window.BiyueCustomData.node_list || []
	var question_map = window.BiyueCustomData.question_map || {}
	Asc.scope.question_map = question_map
	if (!ids) {
		ids = []
		Object.keys(question_map).forEach(key => {
			if (question_map[key].level_type == 'question') {
				ids.push(key)
			}
		})
	}
	Asc.scope.ids = ids
	return biyueCallCommand(window, function() {
		var node_list = Asc.scope.node_list
		var question_map = Asc.scope.question_map
		var ids = Asc.scope.ids
		var oDocument = Api.GetDocument()
		var cell_width_mm = 7 + 3
		var cell_height_mm = 5 + 1.5
		var MM2TWIPS = 25.4 / 72 / 20
		var cellWidth = cell_width_mm / MM2TWIPS
		var cellHeight = cell_height_mm / MM2TWIPS
		var resList = []
		var shapes = oDocument.GetAllShapes()
		var drawings = oDocument.GetAllDrawingObjects()
		var controls = oDocument.GetAllContentControls()
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
		function getShape(id) {
			if (!shapes) {
				return null
			}
			for (var i = 0, imax = shapes.length; i < imax; ++i) {
				var oShape = shapes[i]
				var drawing = oShape.Drawing
				if (!drawing) {
					continue
				}
				var titleObj = getJsonData(drawing.docPr.title)
				if (titleObj.feature && titleObj.feature.type == 'score' && titleObj.feature.ques_id == id) {
					return oShape
				}
			}
			return null
		}
		function deleteShape(oShape) {
			if (!oShape) {
				return
			}
			var oDrawing = drawings.find(e => {
				return e.Drawing.Id == oShape.Drawing.Id
			})
			if (!oDrawing) {
				return
			}
			var run = oDrawing.Drawing.GetRun()
			if (run) {
				var paragraph = run.GetParagraph()
				if (paragraph) {
					var oParagraph = Api.LookupObject(paragraph.Id)
					var ipos = run.GetPosInParent()
					if (ipos >= 0) {
						oDrawing.Delete()
						var element2 = oParagraph.GetElement(ipos)
						if (element2 && element2.GetClassType() == 'run' && element2.Run.Id == run.Id) {
							oParagraph.RemoveElement(ipos)
						}
						return
					}
				}
			}
			oDrawing.Delete()
		}
		function deleteScoreControl(id) {
			var control = controls.find(e => {
				var tag = getJsonData(e.GetTag())
				return tag.regionType == 'score' && tag.ques_id == id
			})
			if (control) {
				var posinparent = oControl.GetClassType() == 'blockLvlSdt' ? control.GetPosInParent() : control.Sdt.GetPosInParent()
				var parent = control.Sdt.Parent
				if (parent) {
					var oParent = Api.LookupObject(parent.Id)
					if (oParent) {
						oParent.RemoveElement(posinparent)
					}
				}
			}
		}
		for (var i = 0, imax = ids.length; i < imax; ++i) {
			var id = ids[i]
			var quesData = question_map[id]
			if (!quesData || quesData.ques_mode == 6) {
				continue
			}
			var nodeData = node_list.find(e => {
				return e.id == id
			})
			if (!nodeData || !nodeData.control_id) {
				continue
			}
			var oControl = Api.LookupObject(nodeData.control_id)
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
				continue
			}

			var oShape = getShape(id)
			if (oShape) {
				deleteShape(oShape)
			}
			deleteScoreControl(id)
			if (quesData.mark_mode != 2) {
				continue
			}
			var scores = [0]
			var score = Math.floor(quesData.score)
			var vInteger = Math.trunc(score) // 整数部分
			var score_mode_use = !quesData.score_mode ? (score >= 15 ? 2 : 1) : quesData.score_mode
			if (score_mode_use == 2) {
				var ten = (vInteger / 10) >> 0
				for (var i1 = 1; i1 <= ten; ++i1) {
					scores.push(`${i1 * 10}+`)
				}
				var maxi2 = score < 10 ? score : 10
				for (var i2 = 1; i2 < maxi2; ++i2) {
					scores.push(i2)
				}
				scores.push('+0.5')
			} else {
				quesData.score_list.forEach(e => {
					if (e.value == 0.5 && score == quesData.score) {
						scores.push(quesData.score)
					}
					if (e.use) {
						scores.push(e.value)
					}
				})
			}

			var rect = Api.asc_GetContentControlBoundingRect(
				nodeData.control_id,
				true
			)
			var newRect = {
				Left: rect.X0,
				Right: rect.X1,
				Top: rect.Y0,
				Bottom: rect.Y1,
			}
			var controlContent = oControl.GetContent()
			if (controlContent) {
				var pageIndex = 0
				if (
					controlContent.Document &&
					controlContent.Document.Pages &&
					controlContent.Document.Pages.length > 1
				) {
					for (var p = 0; p < controlContent.Document.Pages.length; ++p) {
						if (!oControl.Sdt.IsEmptyPage(p)) {
							pageIndex = p
							break
						}
					}
				}
				var pagebounds = controlContent.Document.Get_PageBounds(pageIndex)
				if (pagebounds) {
					newRect.Right = Math.max(pagebounds.Right, newRect.Right)
				}
			}
			var width = newRect.Right - newRect.Left
			var trips_width = width / MM2TWIPS
			var cellcount = scores.length
			var rowcount = 1
			var columncount = cellcount
			var score_layout = quesData.score_layout || 2
			var maxTableWidth = trips_width / score_layout
			if (maxTableWidth < cellcount * cellWidth) {
				rowcount = Math.ceil((cellcount * cellWidth) / maxTableWidth)
				columncount = Math.floor(maxTableWidth / cellWidth)
			}
			var mergecount = rowcount * columncount - cellcount
			if (rowcount <= 0 || columncount <= 0) {
				console.log('行数或列数异常', list)
				continue
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
						}
					}
				}
				if (cells.length > 0) {
					oTable.MergeCells(cells)
				}
			}
			var oTableStyle = oDocument.CreateStyle('CustomTableStyle', 'table')
			var oTableStylePr = oTableStyle.GetConditionalTableStyle('wholeTable')
			oTable.SetTableLook(true, true, true, true, true, true)
			// oTableStylePr.GetTableRowPr().SetHeight('atLeast', cellHeight) // 高度至少多少trips
			var oTableCellPr = oTableStyle.GetTableCellPr()
			oTableCellPr.SetVerticalAlign('center')
			oTable.SetStyle(oTableStyle)
			oTable.SetCellSpacing(100)
			oTable.SetTableCellMarginLeft(0)
			oTable.SetTableCellMarginRight(0)
			oTable.SetTableBorderTop('single', 1, 0.1, 255, 255, 255)
			oTable.SetTableBorderBottom('single', 1, 0.1, 255, 255, 255)
			oTable.SetTableBorderLeft('single', 1, 0.1, 255, 255, 255)
			oTable.SetTableBorderRight('single', 1, 0.1, 255, 255, 255)
			var scoreindex = -1
			// 设置单元格文本
			for (var irow = 0; irow < rowcount; ++irow) {
				var oRow = oTable.GetRow(irow)
				if (oRow) {
					// oRow.Row.SetHeight(cell_height_mm, 'atLeast')
					oRow.SetHeight('atLeast', cellHeight)
				}
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
								oCellPara.AddText(scores[scoreindex] + '')
								cell.SetWidth('twips', cellWidth)
								oCellPara.SetJc('center')
								oCellPara.SetColor(53, 53, 53, false)
								oCellPara.SetFontSize(16)
							} else {
								console.log('oCellPra is null')
							}
						} else {
							console.log('cellcontent is null')
						}
						cell.SetCellBorderTop("single", 2, 0, 153, 153, 153)
						cell.SetCellBorderBottom("single", 2, 0, 153, 153, 153)
						cell.SetCellBorderLeft("single", 2, 0, 153, 153, 153)
						cell.SetCellBorderRight("single", 2, 0, 153, 153, 153)
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
			oTable.SetTableTitle(JSON.stringify({
				type: 'score',
				ques_id: id
			}))
			// if (oTable.SetLockValue) {
			//   oTable.SetLockValue(true)
			// }
			if (score_layout == 1) { // 顶部插入
				oTable.SetWrappingStyle(true)
				oTable.SetHAlign('right')
				oTable.SetJc('right')
				var parent = oControl.Sdt.GetParent()
				var posinparent = oControl.GetClassType() == 'blockLvlSdt' ? oControl.GetPosInParent() : oControl.Sdt.GetPosInParent()
				var oParent = Api.LookupObject(parent.Id)
				var blockcontrol = oTable.InsertInContentControl(1)
				blockcontrol.SetTag(JSON.stringify({
					regionType: "score",
					ques_id: id
				}))
				if (oParent && oParent.AddElement) {
					oParent.AddElement(posinparent, blockcontrol)
				}
				return
			}
			oTable.SetWrappingStyle(params.layout == 1 ? true : false)
			var oFill = Api.CreateNoFill()
			var oStroke = Api.CreateStroke(3600, Api.CreateNoFill())
			var oDrawing = Api.CreateShape(
				'rect',
				shapew * 36e3,
				shapeh * 36e3,
				oFill,
				oStroke
			)
			var drawDocument = oDrawing.GetContent()
			drawDocument.AddElement(0, oTable)
			var titleobj = {
				feature: {
					zone_type: 'question',
					type: 'score',
					ques_control_id: nodeData.control_id,
					ques_id: id,
				}
			}
			oDrawing.Drawing.Set_Props({
				title: JSON.stringify(titleobj),
			})
			if (score_layout == 2) {
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
			// 在题目内插入
			var oRun = Api.CreateRun()
			oRun.AddDrawing(oDrawing)
			var paragraphs = controlContent.GetAllParagraphs()
			console.log('paragraphs', paragraphs)
			if (paragraphs && paragraphs.length > 0) {
				var addParagraph = paragraphs.find(e => {
					var parent1 = e.Paragraph.Parent
					var parent2 = parent1.Parent
					if (parent2 && parent2.Id == nodeData.control_id) {
						return true
					}
				})
				if (addParagraph) {
					addParagraph.AddElement(oRun, 0)
					resList.push({
						code: 1,
						ques_id: id,
						paragraph_id: addParagraph.Paragraph.Id,
						run_id: oRun.Run.Id,
						drawing_id: oDrawing.Drawing.Id,
						table_id: oTable.Table.Id
					})
				}
			}
		}
	}, false, true)
}
// 针对单道题，进行重新切题
function splitControl(qid) {
	if (!qid) {
		return
	}
	var question_map = window.BiyueCustomData.question_map || {}
	var node_list = window.BiyueCustomData.node_list || []
	if (!question_map[qid]) {
		return
	}
	Asc.scope.node_list = node_list
	Asc.scope.question_map = question_map
	Asc.scope.qid = qid
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	return biyueCallCommand(window, function() {
		var node_list = Asc.scope.node_list
		var qid = Asc.scope.qid
		var client_node_id = Asc.scope.client_node_id
		var oDocument = Api.GetDocument()
		var result = {
			change_list: [],
			typeName: 'write'
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
				var parentTag = getJsonData(parentBlock.GetTag())
				parent_id = parentTag.client_id || 0
			}
			return parent_id
		}
		var nodeData = node_list.find(e => {
			return e.id == qid
		})
		var control_id = nodeData ? nodeData.control_id : 0
		var control = null
		if (!control_id) {
			control = Api.LookupObject(control_id)
		}
		if (!control || control.GetClassType() != 'blockLvlSdt' || control.GetPosInParent() == -1) {
			var controls = oDocument.GetAllContentControls()
			if (controls) {
				control = controls.find(e => {
					var tag = getJsonData(e.GetTag())
					return tag.client_id == qid && e.GetClassType() == 'blockLvlSdt' && e.GetPosInParent() >= 0
				})
			}
		}
		if (!control) {
			return
		}
		let marker_log = function (str, ranges) {
			let styledString = ''
			let currentIndex = 0
			const styles = []

			ranges.forEach(([start, end], index) => {
				// 添加高亮前的部分
				if (start > currentIndex) {
					styledString += '%c' + str.substring(currentIndex, start)
					styles.push('')
				}
				// 添加高亮部分
				styledString += '%c' + str.substring(start, end)
				styles.push('border: 1px solid red; padding: 2px')
				currentIndex = end
			})

			// 添加剩余的部分
			if (currentIndex < str.length) {
				styledString += '%c' + str.substring(currentIndex)
				styles.push('')
			}

			console.log(styledString, ...styles)
		}
		var obj = getJsonData(control.GetTag())
		if (obj.regionType == 'question') {
			var inlineSdts = control.GetAllContentControls().filter((e) => {
				return getJsonData(e.GetTag()).regionType == 'write'
			})
			if (inlineSdts.length > 0) {
				console.log('已有inline sdt， 删除以后再执行', inlineSdts)
				inlineSdts.forEach((e) => {
					var tag = getJsonData(e.GetTag())
					result.change_list.push({
						client_id: tag.client_id,
						control_id: e.Sdt.GetId(),
						parent_id: getParentId(e),
						regionType: 'write',
						type: 'remove'
					})
					Api.asc_RemoveContentControlWrapper(e.Sdt.GetId())
				})
			}

			// 标记inline的答题区域
			var text = control.GetRange().GetText()
			// console.log('text', text)
			var rangePatt = /(([\(]|[\（])(\s|\&nbsp\;)*([\）]|[\)]))|(___*)/gs
			var match
			var ranges = []
			var regionTexts = []
			while ((match = rangePatt.exec(text)) !== null) {
				ranges.push([match.index, match.index + match[0].length])
				regionTexts.push(match[0])
			}
			// console.log('regionTexts', regionTexts)
			if (ranges.length > 0) {
				marker_log(text, ranges)
			}
			var textSet = new Set();
			regionTexts.forEach(e => textSet.add(e));

			let includeRange = function(a, b)
			{
				return (a.Element === b.Element &&
					a.Start >= b.Start &&
					a.End <= b.End);
			};
			let mergeRange = function(arrA, arrB)
			{
				let all = arrA.concat(arrB);
				let ret = []
				for(var i = 0; i < all.length; i++) {
					var newE = true;
					for (var j = 0; j < all.length; j++) {
						if (i == j)
							continue;
						if (includeRange(all[i], all[j])) {
							newE = false;
						}
					}
					if (newE)
						ret.push(all[i]);
				}
				return ret;
			};


			//debugger;
			var apiRanges = [];
			textSet.forEach(e => {
				var ranges = control.Search(e, false);
				//debugger;;
				apiRanges = mergeRange(apiRanges, ranges);
			});

				// search 有bug少返回一个字符
			apiRanges.reverse().forEach(apiRange => {
					apiRange.Select();
					client_node_id += 1
					var tag = JSON.stringify({ regionType: 'write', mode: 3, client_id: client_node_id, color: '#ff000040' })
					var oResult = Api.asc_AddContentControl(2, { Tag: tag })
					if (oResult) {
						result.change_list.push({
							client_id: client_node_id,
							control_id: oResult.InternalId,
							parent_id: getParentId(Api.LookupObject(oResult.InternalId)),
							regionType: 'write'
						})
					}
					Api.asc_RemoveSelection();
			});
			var content = control.GetContent();
			var elements = content.GetElementsCount();
			for (var j = elements - 1; j >= 0; j--) {
				var para = content.GetElement(j);
				if (para.GetClassType() !== "paragraph") {
					break;
				}
				var text = para.GetText();
				if (text.trim() !== '') {
					break;
				}
			}

			if (j < elements - 1) {
				var range = content.GetElement(j + 1).GetRange();
				var endRange = content.GetElement(elements - 1).GetRange();
				range = range.ExpandTo(endRange);
				range.Select();
				client_node_id += 1
				var tag = JSON.stringify({ regionType: 'write', mode: 5, client_id: client_node_id, color: '#ff000040' })

				var oResult = Api.asc_AddContentControl(1, { Tag: tag })
				if (oResult) {
					result.change_list.push({
						client_id: client_node_id,
						control_id: oResult.InternalId,
						parent_id: getParentId(Api.LookupObject(oResult.InternalId)),
						regionType: 'write'
					})
				}
				Api.asc_RemoveSelection();
			}
		}
		result.client_node_id = client_node_id
		return result
	}, false, true).then(res1 => {
		if (res1) {
			if (res1.message && res1.message != '') {
				alert(res1.message)
			} else {
				return getNodeList().then(res2 => {
					return handleChangeType(res1, res2)
				})
			}
		}
	})
}
// 删除ID重复控件
function clearRepeatControl(reclac = false) {
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	Asc.scope.node_list = window.BiyueCustomData.node_list || []
	Asc.scope.question_map = window.BiyueCustomData.question_map || {}
	return biyueCallCommand(window, function() {
		var client_node_id = Asc.scope.client_node_id || 0
		var node_list = Asc.scope.node_list
		var question_map = Asc.scope.question_map
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls() || []
		var idmap = {}
		var drawings = oDocument.GetAllDrawingObjects() || []
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
		function addItem(type, id, tag) {
			if (idmap[tag.client_id]) {
				idmap[tag.client_id].push({
					type: type,
					id: id,
					tag: tag
				})
				repeatIds[tag.client_id] = 1
			} else {
				idmap[tag.client_id] = [{
					type: type,
					id: id,
					tag: tag
				}]
			}
		}
		var repeatIds = {}
		controls.forEach(e => {
			if (e.Sdt) {
				var tag = getJsonData(e.GetTag())
				if (tag.client_id) {
					addItem(e.GetClassType(), e.Sdt.GetId(), tag)
				}
			}
		})
		drawings.forEach(e => {
			if (e.Drawing && e.Drawing.docPr && e.Drawing.docPr.title) {
				var tag = getJsonData(e.Drawing.docPr.title)
				if (tag.client_id) {
					if (idmap[tag.client_id]) {
						client_node_id += client_node_id
						tag.client_id = client_node_id
						addItem('drawing', e.Drawing.Id, tag)
					}
					oDrawing.Drawing.Set_Props({
						title: JSON.stringify(tag),
					})
				}
			}
		})
		// 存在ID分配重复
		if (Object.keys(repeatIds).length) {
			Object.keys(repeatIds).forEach(id => {
				var list = idmap[id]
				for (var i = 0; i < list.length; ++i) {
					if (list[i].type == 'blockLvlSdt' || list[i].type == 'inlineLvlSdt') {
						Api.asc_RemoveContentControlWrapper(list[i].id)
					} else if (list[i].type == 'drawing') {

					}
				}
				for (var i = node_list.length - 1; i >= 0; --i) {
					if (node_list[i].id == id) {
						node_list.splice(i, 1)
					} else if (node_list[i].write_list) {
						for (var j = node_list[i].write_list.length - 1; j >= 0; --j) {
							if (node_list[i].write_list[j].id == id) {
								node_list[i].write_list.splice(j, 1)
								var qid = node_list[i].id
								if (question_map[qid] && question_map[qid].ask_list) {
									for (var k = question_map[qid].ask_list.length - 1; k >= 0; --k) {
										if (question_map[qid].ask_list[k].id == id) {
											question_map[qid].ask_list.splice(k, 1)
										}
									}
								}
							}
						}
					}
				}
				if (idx >= 0) {
					node_list.splice(idx, 1)
				}
				if (question_map[id]) {
					delete question_map[id]
				}
			})
		}
		return {
			client_node_id: client_node_id,
			node_list: node_list,
			question_map: question_map
		}
	}, false, reclac).then(res => {
		return new Promise((resolve, reject) => {
			if (res) {
				window.BiyueCustomData.client_node_id = res.client_node_id
				window.BiyueCustomData.node_list = res.node_list
				window.BiyueCustomData.question_map = res.question_map
			}
			resolve()
		})
	})
}

function tidyTree() {
	Asc.scope.node_list = window.BiyueCustomData.node_list || []
	Asc.scope.question_map = window.BiyueCustomData.question_map || {}
	return biyueCallCommand(window, function() {
		var node_list = Asc.scope.node_list
		var question_map = Asc.scope.question_map
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
		function deleteAccurateRun(oRun) {
			if (oRun &&
				oRun.Run &&
				oRun.Run.Content &&
				oRun.Run.Content[0] &&
				oRun.Run.Content[0].docPr) {
				var title = oRun.Run.Content[0].docPr.title
				if (title) {
					var titleObj = getJsonData(title)
					if (titleObj.feature && titleObj.feature.sub_type == 'ask_accurate') {
						oRun.Delete()
						return true
					}
				}
			}
			return false
		}
		function deleteControlAccurate(oControl) {
			if (!oControl || !oControl.GetElementsCount) {
				return
			}
			var elementCount = oControl.GetElementsCount()
			for (var idx = 0; idx < elementCount; ++idx) {
				var oRun = oControl.GetElement(idx)
				if (deleteAccurateRun(oRun)) {
					break
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
							if (child.GetClassType() == 'run' && child.Run.Id == run.Id) {
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
						if (oParagraph) {
							oParagraph.RemoveElement(ipos)
						}
						return true
					}
				}
			}
			oShape.Delete()
			return true
		}
		function deleteAskAccurate(drawing) {
			var tag = getJsonData(drawing.Drawing.docPr.title)
			if (tag.feature && tag.feature.zone_type == 'question' && tag.feature.sub_type == 'ask_accurate') {
				deleShape(drawing)
			}
		}
		function deleteCellAsk(oTable, write_list, nodeIndex) {
			var rowcount = oTable.GetRowsCount()
			var desc = getJsonData(oTable.GetTableDescription())
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
							var flag = 0
							if (!write_list) {
								flag = 1
							} else {
								var index = write_list.findIndex(e => {
									return e.cell_id == oCell.Cell.Id && e.table_id == oTable.Table.Id
								})
								if (index == -1) {
									var cid = desc[`${r}_${c}`]
									if (cid) {
										var index = write_list.findIndex(e => {
											return e.id == cid
										})
										if (index >= 0) {
											node_list[nodeIndex].write_list[index].cell_id = oCell.Cell.Id
											node_list[nodeIndex].write_list[index].table_id = oTable.Table.Id
										}
									}
								}
								flag = index == -1
							}
							if (flag) {
								oCell.SetBackgroundColor(255, 191, 191, true)
								var cellContent = oCell.GetContent()
								var shapes = cellContent.GetAllShapes() || []
								shapes.forEach(e => {
									deleShape(e)
								})
							}
						}
					}
				}
			}
		}
		// 删除题目control的相关子控件
		function deleteControlChild(quesControl) {
			if (!quesControl || quesControl.GetClassType() != 'blockLvlSdt') {
				return
			}
			var control_id = quesControl.Sdt.GetId()
			// 删除所有小问，遍历出所有小问，全部删除
			// 删除所有精准互动
			var childDrawings = quesControl.GetAllDrawingObjects() || []
			for (var j = 0; j < childDrawings.length; ++j) {
				deleteAskAccurate(childDrawings[j])
			}
			// 删除除订正框外的所有inlineControl
			var childControls = quesControl.GetAllContentControls() || []
			childControls.forEach(e => {
				var childTag = getJsonData(e.GetTag())
				if (childTag.regionType != 'question') {
					var parentControl = e.GetParentContentControl()
					if (parentControl && parentControl.Sdt.GetId() == control_id) {
						Api.asc_RemoveContentControlWrapper(e.Sdt.GetId())
					}
				}
			})
			// 删除所有write 和 identify
			childDrawings = quesControl.GetAllDrawingObjects() || []
			for (var j = 0; j < childDrawings.length; ++j) {
				var tag = getJsonData(childDrawings[j].Drawing.docPr.title)
				if (tag.feature && tag.feature.zone_type == 'question') {
					if (tag.feature.sub_type == 'write') {
						var run = childDrawings[j].Drawing.GetRun()
						if (run) {
							var paragraph = run.GetParagraph()
							if (paragraph) {
								var oParagraph = Api.LookupObject(paragraph.Id)
								var ipos = run.GetPosInParent()
								if (ipos >= 0) {
									childDrawings[j].Delete()
									if (oParagraph) {
										oParagraph.RemoveElement(ipos)
									}
								}
							}
						}
					} else if (tag.feature.sub_type == 'identify') {
						childDrawings[j].Delete()
					}
				}
			}
			// 删除所有单元格小问
			var pageCount = quesControl.Sdt.getPageCount()
			for (var p = 0; p < pageCount; ++p) {
				var page = quesControl.Sdt.GetAbsolutePage(p)
				var tables = quesControl.GetAllTablesOnPage(page)
				if (tables) {
					for (var t = 0; t < tables.length; ++t) {
						var oTable = tables[t]
						deleteCellAsk(oTable)
					}
				}
			}
		}
		for (var i = controls.length - 1; i >= 0; --i) {
			var oControl = controls[i]
			if (!oControl || !oControl.Sdt) {
				continue
			}
			var tagInfo = getJsonData(oControl.GetTag())
			if (!tagInfo.client_id) { // 无client_id，直接删除
				var del = true
				if (tagInfo.regionType == 'num') {
					var parentControl = oControl.GetParentContentControl()
					if (parentControl && parentControl.GetClassType() == 'blockLvlSdt') {
						var ptag = getJsonData(parentControl.GetTag())
						if (ptag.regionType == 'question') {
							del = false
						}
					}
				}
				if (del) {
					deleteControlChild(oControl)
					Api.asc_RemoveContentControlWrapper(oControl.Sdt.GetId())
				}
				continue
			}
			if (oControl.GetClassType() == 'blockLvlSdt' && tagInfo.regionType == 'question') {
				var index = node_list.findIndex(e => e.id == tagInfo.client_id)
				if (!question_map[tagInfo.client_id] || index == -1) { // 分配的ID不在业务记录里，删除control
					deleteControlChild(oControl)
					Api.asc_RemoveContentControlWrapper(oControl.Sdt.GetId())
					if (index >= 0) {
						node_list.splice(index, 1)
					}
					if (question_map[tagInfo.client_id]) {
						delete question_map[tagInfo.client_id]
					}
					continue
				}
			} else {
				var parentControl = oControl.GetParentContentControl()
				if (!parentControl) {
					deleteControlAccurate(oControl)
					Api.asc_RemoveContentControlWrapper(oControl.Sdt.GetId())
					continue
				}
			}
		}
		var tables = oDocument.GetAllTables() || []
		if (tables) {
			for (var t = 0; t < tables.length; ++t) {
				var oTable = tables[t]
				var tableParentControl = oTable.GetParentContentControl()
				if (!tableParentControl) {
					deleteCellAsk(oTable)
				} else {
					var tagInfo = getJsonData(tableParentControl.GetTag())
					if (tagInfo.client_id) {
						var index = node_list.findIndex(e => e.id == tagInfo.client_id)
						if (!question_map[tagInfo.client_id] || index == -1) {
							deleteCellAsk(oTable)
						} else {
							deleteCellAsk(oTable, node_list[index].write_list, index)
						}
					}
				}
			}
		}
		return {
			node_list: node_list
		}
	}, false, false)
}

function tidyNodes() {
	return clearRepeatControl(false).then(() => {
		return tidyTree()
	}).then((res) => {
		if (res && res.node_list) {
			window.BiyueCustomData.node_list = res.node_list
		}
		return getNodeList()
	})
	.then(list => {
		console.log('tidyNodes', list)
		if (list) {
			var node_list = window.BiyueCustomData.node_list || []
			var question_map = window.BiyueCustomData.question_map || {}
			for (var i = 0, imax = list.length; i < imax; ++i) {
				var id = list[i].id
				var index1 = node_list.findIndex(e => e.id == id)
				if (index1 == -1) { // 原本就不在node列表里，不管
					if (question_map[id]) {
						delete question_map[id] // 同步删除map里的
					}
					continue
				}
				var write_list = list[i].write_list || []
				var oldNodeData = node_list[index1]
				if (oldNodeData.write_list) {
					for (var j = oldNodeData.write_list.length - 1; j >= 0; --j) {
						var writeData = oldNodeData.write_list[j]
						var writeIndex = write_list.findIndex(e => e.cell_id == writeData.cell_id)
						if (writeIndex == -1) { // 找不到，删除
							oldNodeData.write_list.splice(j, 1)
						}
					}
				}
				if (question_map[id].ask_list) {
					for (var k = question_map[id].ask_list.length - 1; k >= 0; --k) {
						var askData = question_map[id].ask_list[k]
						var wd = oldNodeData.write_list.find(e => {
							return e.id == askData.id
						})
						if (!wd) {
							var askIndex = write_list.findIndex(e => e.cell_id == wd.cell_id)
							if (askIndex == -1) { // 找不到，删除
								question_map[id].ask_list.splice(k, 1)
							}
						}
					}
				}
			}
			for (var i = node_list.length -1; i >= 0; --i) {
				var index2 = list.findIndex(e => e.id == node_list[i].id)
				if (index2 == -1) {
					node_list.splice(i, 1)
				}
			}
			var keys = Object.keys(question_map)
			keys.forEach(id => {
				var index3 = list.findIndex(e => e.id == id)
				if (index3 == -1) {
					delete question_map[id]
				}
			})
			window.BiyueCustomData.node_list = node_list
			window.BiyueCustomData.question_map = question_map
		}
	})
}
// 处理上传前准备工作，包括
// 1、图片不铺码颜色
// 2、单元格小问颜色
// 3、作答区颜色
// 4、隐藏页码
function handleUploadPrepare(cmdType) {
	Asc.scope.cmdType = cmdType
	Asc.scope.node_list = window.BiyueCustomData.node_list || []
	Asc.scope.question_map = window.BiyueCustomData.question_map || {}
	return biyueCallCommand(window, function() {
		var cmdType = Asc.scope.cmdType
		var oDocument = Api.GetDocument()
		var drawings = oDocument.GetAllDrawingObjects() || []
		var question_map = Asc.scope.question_map || {}
		var node_list = Asc.scope.node_list || []
		var oTables = oDocument.GetAllTables() || []
		var vshow = cmdType == 'show'
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
		function updateFill(drawing, oFill) {
			if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== 'fill') {
				return false
			}
			drawing.GraphicObj.spPr.setFill(oFill.UniFill)
		}
		drawings.forEach(oDrawing => {
			let title = oDrawing.Drawing.docPr.title || ''
			var titleObj = getJsonData(title)
			// 图片不铺码的标识
			if (title.includes('partical_no_dot')) {
			  	var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 255, 255))
			  	oFill.UniFill.transparent = 0 // 透明度
			  	var oStroke = Api.CreateStroke(10000, vshow ? Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)) : oFill);
			  	oDrawing.SetOutLine(oStroke);
			} else if (titleObj.feature) {
				if (titleObj.feature.zone_type == 'pagination') { // 试卷页码
					var oShape = Api.LookupObject(oDrawing.Drawing.GraphicObj.Id)
					if (oShape) {
						var oShapeContent = oShape.GetContent()
						if (oShapeContent) {
							var paragraphs = oShapeContent.GetAllParagraphs() || []
							paragraphs.forEach(p => {p.SetColor(3, 3, 3, !vshow)})
						}
					}
				} else if (titleObj.feature.zone_type == 'question' && titleObj.feature.sub_type == 'write') { // 作答区
					if (cmdType == 'show') {
						var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 0, 0))
						oFill.UniFill.transparent = 255 * 0.2 // 透明度
						updateFill(oDrawing.Drawing, oFill)
					} else if (cmdType == 'hide') {
						updateFill(oDrawing.Drawing, Api.CreateNoFill())
					}
				}
			}
		})
		// 处理单元格小问
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
								if (oCell && oCell.GetClassType && oCell.GetClassType() == 'tableCell') {
									var oTable = oCell.GetParentTable()
									if (oTable && oTable.GetPosInParent() == -1) {
										oCell = getCell(writeData)
									}
								} else {
									oCell = getCell(writeData)
								}
								if (oCell) {
									oCell.SetBackgroundColor(255, 191, 191, cmdType == 'show' ? false : true)
								}
							}
						})
					}
				}
			}
		})

	}, false, true)
}

function importExam() {
	setBtnLoading('uploadTree', true)
	setBtnLoading('importExam', true)
	setInteraction('none', null, false).then(() => {
		return addOnlyBigControl(false)
	}).then(() => {
		return handleUploadPrepare('hide')
	}).then(() => {
		return getControlListForUpload()
	}).then(control_list => {
		if (control_list && control_list.length) {
			generateTreeForUpload(control_list).then(() => {
				setBtnLoading('uploadTree', false)
				setInteraction('useself').then(() => {
					return getAllPositions2()
				}).then(res => {
					Asc.scope.questionPositions = res
					return removeOnlyBigControl()
				}).then(() => {
					window.biyue.showDialog('exportExamWindow', '上传试卷', 'examExport.html', 1000, 800, true)
					setBtnLoading('importExam', false)
				})
			}).catch((res) => {
				setBtnLoading('uploadTree', false)
				handleCompleteResult(res && res.message && res.message != '' ? res.message : '全量更新失败')
				setBtnLoading('importExam', false)
			})
		} else {
			setBtnLoading('uploadTree', false)
			setBtnLoading('importExam', false)
			window.biyue.showMessageBox({
				content: '未找到可更新的题目，请检查题目列表',
				showCancel: false
			})
			removeOnlyBigControl().then(() => {
				return handleUploadPrepare('show')
			}).then(() => {
				return setInteraction('useself')
			})
		}
	})
}

export {
	handleDocClick,
	handleContextMenuShow,
	initExamTree,
	reqGetQuestionType,
	reqUploadTree,
	splitEnd,
	showLevelSetDialog,
	confirmLevelSet,
	initControls,
	handleAllWrite,
	changeProportion,
	deleteAsks,
	focusAsk,
	showAskCells,
	onContextMenuClick,
	g_click_value,
	updateAllChoice,
	layoutRepair,
	deleteChoiceOtherWrite,
	getQuesMode,
	tagImageCommon,
	updateQuesScore,
	splitControl,
	updateDataBySavedData,
	clearRepeatControl,
	tidyNodes,
	handleUploadPrepare,
	importExam,
	setBtnLoading,
}