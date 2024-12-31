
import { biyueCallCommand, dispatchCommandResult } from "./command.js";
import { getQuesType, reqComplete } from '../scripts/api/paper.js'
import { handleChoiceUpdateResult, setInteraction, updateChoice } from "./featureManager.js";
import { initExtroInfo } from "./panelFeature.js";
import { addOnlyBigControl, removeOnlyBigControl, getAllPositions2 } from './business.js'
import { handleRangeType, addWriteZone } from "./classifiedTypes.js"
import { imageAutoLink, ShowLinkedWhenclickImage } from './linkHandler.js'
import { layoutDetect } from './layoutFixHandler.js'
import { setBtnLoading, isLoading } from './model/util.js'
import { refreshTree } from './panelTree.js'
import { extractChoiceOptions, removeChoiceOptions, getChoiceOptionAndSteam, setChoiceOptionLayout } from './choiceQuestion.js'
import { getInteractionTypes } from './model/feature.js'
import proportionHandler from './handler/proportionHandler.js'
var g_click_value = null
var upload_control_list = []

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
				} else {
					Asc.scope.controlId = returnValue.InternalId
					biyueCallCommand(window, function() {
						var oControl = Api.LookupObject(Asc.scope.controlId)
						var parentControl = oControl.GetParentContentControl()
						return parentControl ? parentControl.GetTag() : ''
					}, false, false).then(res => {
						var parentTag = getJsonData(res)
						if (parentTag.client_id || window.tab_select == 'tabQues') {
							var event = new CustomEvent('clickSingleQues', {
								detail: {
									InternalId: returnValue.InternalId,
									tag: tag,
									parentTag: parentTag
								},
							})
							document.dispatchEvent(event)
						}
					})
				}
				if (options.isSelectionUse) {
					var shortcutKey = window.BiyueCustomData ? window.BiyueCustomData.ask_shortcut : null
					if (shortcutKey && shortcutKey != '') {
						var sckey = `${shortcutKey}Key`
						if (options[sckey]) {
							// 划分小问
							// updateRangeControlType('write')
							handleRangeType({
								typeName: 'write'
							})
							return
						}
					}
				}
				ShowLinkedWhenclickImage(options, returnValue.InternalId)
			} catch (error) {
				console.log(error)
			}
		} else {
			g_click_value = null
			ShowLinkedWhenclickImage(options)
		}
	})
}
// 右建显示菜单
function handleContextMenuShow(options) {
	console.log('handleContextMenuShow', options)
	Asc.scope.menu_options = options
	return biyueCallCommand(window, function() {
		var options = Asc.scope.menu_options
		var bTable = false
		var oDocument = Api.GetDocument()
		var selectContent = oDocument.Document.GetSelectedContent()
		var column_num = 0
		var result = {
			drawings: [],
			parentSdts: [],
			bTable: false,
			column_num: 0
		}
		var paragraph = oDocument.Document.GetCurrentParagraph()
		if (paragraph) {
			var oParagraph = Api.LookupObject(paragraph.Id)
			var oSection = oParagraph.GetSection()
			if(oSection) {
				result.column_num =  oSection.Section.GetColumnsCount()
			}
		}
		if (options.type == 'Image' || options.type == 'Shape') {
			var DrawingObjects = selectContent.DrawingObjects || []
			if (DrawingObjects.length) {
				DrawingObjects.forEach(e => {
					if (e.docPr) {
						result.drawings.push(e.docPr.title)
					}
				})
			}
		} else {
			var elementsInfo = oDocument.Document.GetSelectedElementsInfo() || {}
			result.bTable = elementsInfo.m_bTable
			var tableIds = {}
			if (elementsInfo.m_arrSdts && elementsInfo.m_arrSdts.length) {
				elementsInfo.m_arrSdts.forEach((e, index) => {
					var oControl = Api.LookupObject(e.Id)
					var lvl = null
					if (oControl.GetClassType() == 'blockLvlSdt') {
						var paragraphs = oControl.GetAllParagraphs() || []
						for (var i = 0; i < paragraphs.length; ++i) {
							var oParagraph = paragraphs[i]
							if (oParagraph) {
								var parent1 = oParagraph.Paragraph.Parent
								var parent2 = parent1.Parent
								if (parent2 && parent2.Id == oControl.Sdt.GetId()) {
									var oNumberingLevel = paragraphs[0].GetNumbering()
									if (oNumberingLevel && oNumberingLevel.Num) {
										lvl = oNumberingLevel.Num.GetLvl(oNumberingLevel.Lvl)
									}
									break
								}
							}
						}
						if (result.bTable) {
							result.cells = []
							if (elementsInfo.m_pParagraph) {
								var oParagraph = Api.LookupObject(elementsInfo.m_pParagraph.Id)
								var oCell = oParagraph.GetParentTableCell()
								if (oCell) {
									var oParentTable = oCell.GetParentTable()
									var tableDesc = Api.ParseJSON(oParentTable.GetTableDescription()) || {}
									result.cells.push({
										Id: oCell.Cell.Id,
										recordId: tableDesc[`${oCell.GetRowIndex()}_${oCell.GetIndex()}`],
										isEmpty: oCell.GetContent().Document.IsEmpty()
									})
								}
							} else {
								var pageCount = oControl.Sdt.getPageCount()
								for (var p = 0; p < pageCount; ++p) {
									var page = oControl.Sdt.GetAbsolutePage(p)
									var tables = oControl.GetAllTablesOnPage(page)
									if (tables) {
										for (var t = 0; t < tables.length; ++t) {
											var oTable = tables[t]
											if (tableIds[oTable.Table.Id]) {
												continue
											}
											tableIds[oTable.Table.Id] = 1
											if (!oTable.Table.IsSelectionUse || !(oTable.Table.IsSelectionUse())) {
												continue
											}
											var cellArray = oTable.Table.GetSelectionArray(false) || []
											var tableDesc = Api.ParseJSON(oTable.GetTableDescription()) || {}
											for (var j = 0; j < cellArray.length; ++j) {
												var oCell = oTable.GetCell(cellArray[j].Row, cellArray[j].Cell)
												if (!oCell) {
													continue
												}
												var cellContent = oCell.GetContent()
												if (!cellContent) {
													continue
												}
												result.cells.push({
													Id: oCell.Cell.Id,
													recordId: tableDesc[`${cellArray[j].Row}_${cellArray[j].Cell}`],
													isEmpty: cellContent.Document.IsEmpty()
												})
											}
										}
									}
								}
							}
						}
					}
					result.parentSdts.push({
						Id: e.Id,
						Tag: e.Pr ? e.Pr.Tag : null,
						classType: oControl.GetClassType(),
						lvl: lvl
					})
				})
			}
			
		}
		return result 
	}, false, false).then(res => {
		window.Asc.plugin.executeMethod('AddContextMenuItem', [getContextMenuItems(options.type, res)])
	})
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

function getContextMenuItems(type, selectedRes) {
	console.log('getContextMenuItems ', selectedRes)
	var items = []
	if (type == 'Image' || type == 'Shape') {
		var ignoreCount = 0
		var count = 0
		var writeCount = 0
		var identifyCount = 0
		if (selectedRes) {
			selectedRes.drawings.forEach(title => {
				var titleObj = getJsonData(title)
				if (titleObj.feature) {
					if (titleObj.feature.partical_no_dot) {
						ignoreCount++
					} else if (titleObj.feature.sub_type == 'write') {
						writeCount++
					} else if (titleObj.feature.sub_type == 'identify') {
						identifyCount++
					}
				}
				count++
			})
		}
		items.push({
			id: 'handleImageIgnore',
			text: '图片铺码',
			items: [{
				id: 'handleImageIgnore:del',
				text: '开启',
				disabled: ignoreCount == 0
			}, {
				id: 'handleImageIgnore:add',
				text: '关闭',
				disabled: ignoreCount == count
			}]
		})
		items.push({
			id: 'imageRelation',
			text: '图片关联'
		})
		if (type == 'Shape') {
			if (writeCount) {
				items.push({
					id: 'updateControlType:writeZone:del',
					text: '删除作答区'
				})
			}
			if (identifyCount) {
				items.push({
					id: 'updateControlType:identify:del',
					text: '删除识别框'
				})	
			}
		}
	} else {
		var question_map = window.BiyueCustomData.question_map || {}
		var node_list = window.BiyueCustomData.node_list || []
		function getControlData(tag) {
			var cData = null
			if (tag.client_id) {
				var keys = Object.keys(question_map)
				for (var i = 0, imax = keys.length; i < imax; ++i) {
					var id = keys[i]
					if (id == tag.client_id) {
						cData = { ques_id: id, level_type: question_map[id].level_type }
						break
					} else if (tag.regionType == 'write' && question_map[id].ask_list) {
						var askIndex = question_map[id].ask_list.findIndex(e => {
							return e.id == tag.client_id
						})
						if (askIndex == -1) {
							askIndex = question_map[id].ask_list.findIndex(e => {
								return e.other_fields && e.other_fields.includes(tag.client_id)
							})
							if (askIndex >= 0) {
								cData = { ques_id: id, ask_id: tag.client_id, is_merge: true, level_type: 'ask'}
							}
						} else {
							cData = { ques_id: id, ask_id: tag.client_id, is_merge: false, ask_index: askIndex, level_type: 'ask'}
						}
						if (askIndex >= 0) {
							if (question_map[id].ask_list[askIndex].other_fields && question_map[id].ask_list[askIndex].other_fields.length) {
								cData.is_merge_ask = true
								cData.ask_id = tag.client_id
							}
							break
						}
					} else if (tag.mid && id == tag.mid) {
						if (question_map[id].is_merge) {
							cData = { ques_id: id, level_type: 'question', is_merge: true }
							break
						}
					}
				}
				if (cData) {
					if (cData.ques_id) {
						var nodeData = node_list.find(e => {
							return e.id == cData.ques_id
						})
						if (nodeData) {
							if (cData.level_type == 'question' && nodeData.is_big) {
								cData.level_type = 'big'
							}
							if ( type != 'Selection' && 
								selectedRes.cells && 
								selectedRes.cells.length == 1 && 
								nodeData.write_list && 
								question_map[cData.ques_id] && 
								question_map[cData.ques_id].ask_list) {
								var cellId = selectedRes.cells[0].Id
								var cellRecordId = selectedRes.cells[0].recordId
								var cellWrite = nodeData.write_list.find(e => {
									return e.sub_type == 'cell' && (e.cell_id == cellId || e.id == cellRecordId)
								})
								if (cellWrite) {
									var find = question_map[cData.ques_id].ask_list.find(e => {
										return e.id == cellWrite.id || (e.other_fields && e.other_fields.includes(cellWrite.id))
									})
									if (find) {
										cData.level_type = 'ask'
										if (find.other_fields && find.other_fields.length) {
											cData.is_merge_ask = true
											cData.ask_id = cellWrite.id
										}
									}
								}
							}
						}
					}
				}
			} else {
				for (var j = selectedRes.parentSdts.length - 2; j >= 0; --j) {
					if (selectedRes.parentSdts[j].classType == 'blockLvlSdt') {
						var tag2 = getJsonData(selectedRes.parentSdts[j].Tag)
						var pData = getControlData(tag2)
						if (pData && pData.level_type) {
							cData = {
								level_type: '',
								parent_level_type: pData.level_type,
								parent_id: tag2.mid || tag2.client_id
							}
							break
						}
					}
				}
			}
			return cData
		}
		if ((type == 'Target' && (selectedRes.parentSdts.length || g_click_value)) || (type == 'Selection')) {
			items.push({
				id: `layoutRepair`,
				text: '排版修复'
			})
			// 支持划分类型，作答区，识别框
			var splitType = {
				id: 'updateControlType',
				text: '划分类型',
				items: []
			}
			var list = [{
				value: 'question',
				text: '设置为 - 题目',
				icon: 'rect'
			}, {
				value: 'mergeQuestion',
				text: '合并为 - 一题',
				icon: 'rect'
			}, {
				value: 'struct',
				text: '设置为 - 题组(结构)',
				icon: 'struct',
			}, {
				value: 'setBig',
				text: '设置为 - 大题',
				icon: 'rect'
			}, {
				value: 'write',
				text: '设置为 - 小问',
				icon: 'rect'
			}, {
				value: 'mergedAsk',
				text: '合并为 - 一问',
				icon: 'rect'
			},
			{
				value: 'choiceOption',
				text: '设置为 - 选项',
				icon: 'rect'
			}, 
			{
				value: 'clearBig',
				text: '清除 - 大题',
				icon: 'clear'
			}, {
				value: 'clearChildren',
				text: '清除 - 所有小问',
				icon: 'clear'
			}, {
				value: 'clearMergeAsk',
				text: '解除 - 小问合并',
				icon: 'clear'
			}, {
				value: 'clearMerge',
				text: '清除 - 合并题',
				icon: 'clear'
			}, {
				value: 'clear',
				text: '清除 - 选中区域',
				icon: 'clear'
			}, {
				value: 'clearAll',
				text: '清除 - 选中区域(含子级)',
				icon: 'clear'
			}]
			var valueMap = {}
			list.forEach(e => {
				valueMap[e.value] = e.value == 'clear' || e.value == 'clearAll' ? 1 : 0
				if (e.value == 'mergeQuestion') {
					valueMap[e.value] = selectedRes.bTable && type == 'Selection'
				}
			})
			if (type == 'Selection') {
				if (selectedRes.parentSdts.length == 0) {
					valueMap['struct'] = 1
					valueMap['question'] = 1
				}
			}
			var cData = null
			var lvl = null
			var curControl = null
			if (selectedRes.parentSdts.length) {
				curControl = selectedRes.parentSdts[selectedRes.parentSdts.length - 1]
				lvl = curControl.lvl
			} else if (g_click_value) {
				curControl = g_click_value
			}
			if (curControl) {
				var tag = getJsonData(curControl.Tag)
				cData = getControlData(tag)
				console.log('cData', cData)
				if (curControl.classType == 'blockLvlSdt') {
					valueMap['clearChildren'] = 1
					if (!cData || !cData.level_type) {
						valueMap['question'] = 1
						valueMap['struct'] = 1
					}
				} else if (curControl.classType == 'inlineLvlSdt' && 
					tag.regionType == 'choiceOption' && 
					cData.parent_id && 
					question_map[cData.parent_id] && 
					(question_map[cData.parent_id].ques_mode == 1 || question_map[cData.parent_id].ques_mode == 5)) {
					valueMap['choiceOption'] = 1
				}
			}
			if (cData) {
				if (type == 'Selection') {
					valueMap['question'] = 1
					valueMap['struct'] = 1
					if (cData.level_type == 'question' || cData.level_type == 'big') {
						valueMap['write'] = 1
						if (question_map[cData.ques_id] && (question_map[cData.ques_id].ques_mode == 1 || question_map[cData.ques_id].ques_mode == 5)) {
							valueMap['choiceOption'] = 1
						}
						if (selectedRes.cells && selectedRes.cells.length > 1) {
							console.log('支持合并小问')
							valueMap['mergedAsk'] = 1
						} else {
							console.log('================ 未满足单元格条件', selectedRes.cells)
						}
					}
					if (cData.level_type == 'big') {
						valueMap['clearBig'] = 1
					}
				} else {
					if (cData.level_type == 'question') {
						splitType.text += cData.is_merge ? '(现为合并题)' : '(现为题目)'
						if (cData.is_merge) {
							valueMap['clearMerge'] = 1
						} else {
							valueMap['struct'] = 1
							if (lvl) {
								valueMap['setBig'] = 1
							}
							if (selectedRes.bTable) {
								valueMap['write'] = 1
							}
						}
					} else if (cData.level_type == 'struct') {
						splitType.text += '(现为题组)'
						valueMap['question'] = 1
					} else if (cData.level_type == 'big') {
						splitType.text += '(现为大题)'
						valueMap['struct'] = 0
						valueMap['clearBig'] = 1
					} else if (cData.level_type == 'ask') {
						splitType.text += '(现为小问)'
						if (cData.is_merge_ask) {
							valueMap['clearMergeAsk'] = 1
						}
						// if (cData.classType == 'blockLvlSdt') {
						// 	valueMap['question'] = 1
						// }
					} else if (cData.parent_level_type == 'question') {
						valueMap['write'] = 1
					}
				}
			}
			
			list.forEach((e, index) => {
				if (valueMap[e.value]) {
					var id = `updateControlType:${e.value}`
					if (e.value == 'clearMergeAsk') {
						id = `clearMergeAsk:${cData.ques_id}:${cData.ask_id}`
					}
					splitType.items.push({
						icons: `resources/%theme-type%(light|dark)/%state%(normal)${e.icon}%scale%(100|200).%extension%(png)`,
						id: id,
						text: e.text
					})
				}
			})
			items.push(splitType)
			if (cData) {
				if (cData.level_type == 'ask') {
					if (!cData.is_merge && cData.ask_index > 0) {
						items.push({
							id: `mergeAsk:${cData.ques_id}:${cData.ask_id}:1`,
							text: '向前合并'
						})
					}
					if (cData.is_merge) {
						items.push({
							id: `mergeAsk:${cData.ques_id}:${cData.ask_id}:0`,
							text: '取消合并'
						})
					}
				} else if (cData.level_type == 'question') {
					items.push({
						id: 'write',
						text: '作答区',
						items: [{
							id: 'updateControlType:writeZone:add',
							text: '添加',
						}, {
							id: 'updateControlType:writeZone:del',
							text: '删除',
						}]
					})
					items.push({
						id: 'identify',
						text: '识别框',
						items: [
							{
								id: 'updateControlType:identify:add',
								text: '添加',
							},
							{
								id: 'updateControlType:identify:del',
								text: '删除',
							},
						],
					})
				}
			}
			var canBatch = type == 'Selection' || (cData && ( cData.level_type == 'question' || cData.level_type == 'struct')) 
			if (canBatch) {
				var questypes = window.BiyueCustomData.paper_options ? window.BiyueCustomData.paper_options.question_type : []
				var itemsQuesType = questypes.map((e) => {
					return {
						id: `batchChangeQuesType:${e.value}`,
						text: e.label,
					}
				})
				var itemsProportion = []
				for (var i = 1; i <= 8; ++i) {
					itemsProportion.push({
						id: `batchChangeProportion:${i}`,
						text: i == 1 ? '默认' : `1/${i}`,
					})
				}
				items.push({
					id: 'batchCmd',
					text: '批量操作',
					items: [
						{
							id: 'batchChangeQuesType',
							text: '修改题型',
							items: itemsQuesType,
						},
						{
							id: 'batchChangeProportion',
							text: '修改占比',
							items: itemsProportion,
						},
						{
							id: 'batchChangeInteraction',
							text: '修改互动模式',
							items: getInteractionTypes().map((e) => {
								return {
									id: `batchChangeInteraction:${e.value}`,
									text: e.label
								}
							})
						},
					],
				})
			}
		}
		if (selectedRes.bTable) {
			items.push({
				id: 'tableRelation',
				text: '表格关联'
			})
		}
		if (selectedRes.column_num) {
			if (selectedRes.column_num != 2) {
				items.push({
					id: 'setSectionColumn:2',
					text: '分为2栏'
				})
			}
			if (selectedRes.column_num > 1) {
				items.push({
					id: 'setSectionColumn:1',
					text: '取消分栏'
				})
			}
		}
	}
	if (items.length) {
		items[0].separator = true
	}
	return {
		guid: window.Asc.plugin.guid,
		items: items
	}
}

function onContextMenuClick(id) {
	var strs = id.split(':')
	if (strs && strs.length > 0) {
		var funcName = strs[0]
		switch (funcName) {
			case 'updateControlType':
				if (strs[1] == 'writeZone' && strs[2] == 'add') {
					addWriteZone()
				} else {
					handleRangeType({
						typeName: strs[1],
						cmd: strs[2] ? strs[2] : null
					})
				}
				break
			case 'clearMergeAsk':
				clearMergeAsk(strs)
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
			case 'setSectionColumn': // 分栏
				setSectionColumn(strs[1] * 1)
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
			case 'tableRelation':
				preGetExamTree().then(res => {
					Asc.scope.tree_info = res
					window.biyue.showDialog('imageRelationWindow', funcName == 'imageRelation' ? '图片关联' : '表格关联', 'imageRelation.html', 800, 600, false)	
				})
				break
			case 'mergeAsk':
				mergeAsk(strs)
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
				var parentTag = Api.ParseJSON(parentBlock.GetTag())
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
									var titleObj = Api.ParseJSON(title)
									if (titleObj.feature && (titleObj.feature.sub_type == 'write' || titleObj.feature.sub_type == 'identify')) {
										var obj = {
											index: write_list.length,
											id: titleObj.feature.client_id,
											sub_type: titleObj.feature.sub_type,
											drawing_id: oChild2.Id,
											shape_id: oChild2.GraphicObj.Id,
										}
										obj.bpage = oChild2.PageNum + 1
										obj.bx = oChild2.X
										obj.by = oChild2.Y
										obj.ex = oChild2.X + oChild2.Width
										obj.ey = oChild2.Y + oChild2.Height
										write_list.push(obj)
									}
								}
							}
						}
					}
				} else if (childType == 'inlineLvlSdt') {
					var childTag = Api.ParseJSON(oChild.GetTag())
					if (childTag.regionType == 'write') {
						var obj = {
							index: write_list.length,
							id: childTag.client_id,
							sub_type: 'control',
							control_id: oChild.Sdt.GetId(),
						}
						var bounds = oChild.Sdt.Bounds
						if (bounds && bounds[0]) {
							obj.bpage = bounds[0].Page + 1
							obj.bx = bounds[0].X
							obj.by = bounds[0].Y
							obj.ex = bounds[0].X + bounds[0].W
							obj.ey = bounds[0].Y + bounds[0].H
						}
						write_list.push(obj)
					}
				}
			}
		}
		function getBlockWriteList(oElement, write_list) {
			if (!oElement || !oElement.GetClassType || oElement.GetClassType() != 'blockLvlSdt') {
				return
			}
			var childTag = Api.ParseJSON(oElement.GetTag())
			if (childTag.regionType == 'write') {
				var obj = {
					index: write_list.length,
					id: childTag.client_id,
					sub_type: 'control',
					control_id: oElement.Sdt.GetId(),
				}
				var bounds = oControl.Sdt.Bounds
				if (bounds && bounds[0]) {
					obj.bpage = bounds[0].Page + 1
					obj.bx = bounds[0].X
					obj.by = bounds[0].Y
					obj.ex = bounds[0].X + bounds[0].W
					obj.ey = bounds[0].Y + bounds[0].H
				}
				write_list.push(obj)
			}
		}
		function getCellBounds(oCell) {
			if (!oCell || oCell.GetClassType() != 'tableCell') {
				return {}
			}
			var pagesCount = oCell.Cell.PagesCount
			for (var p = 0; p < pagesCount; ++p) {
				var pagebounds = oCell.Cell.GetPageBounds(p)
				if (!pagebounds) {
					continue
				}
				if (pagebounds.Right == 0 && pagebounds.Left == 0) {
					continue
				}
				return {
					bpage: oCell.Cell.Get_AbsolutePage(p) + 1,
					bx: pagebounds.Left,
					by: pagebounds.Top,
					ex: pagebounds.Right,
					ey: pagebounds.Bottom
				}
			}
			return {}
		}
		for (var i = 0, imax = controls.length; i < imax; ++i) {
			var oControl = controls[i]
			var tagInfo = Api.ParseJSON(oControl.GetTag())
			var parent_id = getParentId(oControl)
			if (tagInfo.regionType == 'question' && oControl.GetClassType() == 'blockLvlSdt') {
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
						// todo..可能需要过滤下打分区
						var rows = oElement.GetRowsCount()
						var tableTitle = Api.ParseJSON(oElement.GetTableDescription()) || {}
						for (var i1 = 0; i1 < rows; ++i1) {
							var oRow = oElement.GetRow(i1)
							var cells = oRow.GetCellsCount()
							for (var i2 = 0; i2 < cells; ++i2) {
								var oCell = oRow.GetCell(i2)
								var shd = oCell.Cell.Get_Shd()
								var fill = shd.Fill
								if (fill && fill.r == 255 && fill.g == 191 && fill.b == 191) {
									var oldId = tableTitle[`${i1}_${i2}`]
									var obj = Object.assign({}, {
										index: write_list.length,
										id: 'c_' + oCell.Cell.Id,
										sub_type: 'cell',
										table_id: oElement.Table.Id,
										cell_id: oCell.Cell.Id,
										row_index: i1,
										cell_index: i2,
										old_id: oldId
									}, getCellBounds(oCell))
									write_list.push(obj)
									tableTitle[`${i1}_${i2}`] = 'c_' + oCell.Cell.Id
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
						oElement.SetTableDescription(JSON.stringify(tableTitle))
					}
				}
				var nodeObj = {
					id: tagInfo.client_id,
					regionType: tagInfo.regionType,
					parent_id: parent_id
				}
				if (tagInfo.mid && oControl.GetParentTableCell()) {
					if (tagInfo.cask == 1) {
						var oCell = oControl.GetParentTableCell()
						var obj = Object.assign({}, {
							id: 'c_' + oCell.Cell.Id,
							sub_type: 'cell',
							table_id: oCell.GetParentTable().Table.Id,
							cell_id: oCell.Cell.Id,
							row_index: oCell.GetRowIndex(),
							cell_index: oCell.GetIndex()
						}, getCellBounds(oCell))
						nodeObj.cell_ask = true
					}
				}
				// 当有作答区小问时，才需要排序
				if (write_list.length > 1 && write_list.find(e => {
					return e.sub_type == 'write'
				})) {
					write_list = write_list.sort((a, b) => {
						if (a.bpage && b.bpage) {
							if (a.bpage != b.bpage) {
								return a.bpage - b.bpage
							} else {
								if (a.by == b.by) {
									return a.bx - b.bx
								} else {
									if (a.ey <= b.by) {
										return -1
									} else if (b.ey <= a.by) {
										return 1
									} else {
										if (b.by > (a.by + (a.ey - a.by) * 0.5)) {
											return -1
										} else if (a.by > (b.by + (b.ey - b.by) * 0.5)) {
											return 1
										} else {
											return a.bx - b.bx
										}
									}
								}
							}
						} else {
							return a.index - b.index
						}
					})
					write_list.forEach(e => {
						delete e.index
						delete e.bpage
						delete e.bx
						delete e.by
						delete e.ex
						delete e.ey
					})
				}
				nodeObj.write_list = write_list
				node_list.push(nodeObj)
			}
		}
		return node_list
	}, false, true)
}

function updateScore(qid) {
	var question_map = window.BiyueCustomData.question_map
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

function handleChangeType(res, res2) {
	console.log('handleChangeType', res, res2)
	if (!res) {
		return
	}
	var change_list = res.change_list || []
	if (change_list.length == 0) {
		if (res.typeName != 'clearChildren') {
			return
		}
	}
	if (res.client_node_id) {
		window.BiyueCustomData.client_node_id = res.client_node_id
	}
	var node_list = window.BiyueCustomData.node_list || []
	var question_map = window.BiyueCustomData.question_map || {}
	var level_type = res.typeName
	if (res.typeName == 'mergedAsk') {
		level_type = 'write'
	}
	var targetLevel = level_type
	if (res.typeName == 'setBig' || res.typeName == 'clearBig') {
		targetLevel = 'question'
	}
	var addIds = []
	var update_node_id = g_click_value ? g_click_value.Tag.client_id : 0
	var other_asks_remove = []
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
	function removeUnvalidAsk(qid, write_list) {
		let old_list = question_map[qid].ask_list || [];
		old_list = old_list.filter(item => {
			// 处理 other_fields
			if (item.other_fields) {
				item.other_fields = item.other_fields
					.map(otherField => {
						// 在 write_list 中寻找匹配的 id 或 old_id
						const validItem = write_list.find(e => e.id === otherField || e.old_id === otherField);
						return validItem ? validItem.id : null;
					})
					.filter(id => id !== null); // 过滤掉无效的 id
			}

			// 检查主 id
			const validMainItem = write_list.find(e => e.id === item.id || e.old_id === item.id);
			if (validMainItem) {
				item.id = validMainItem.id;
				return true;
			}
			return false;
		});
		// 更新 question_map 中的 ask_list
		question_map[qid].ask_list = old_list;
	}
	function reSortAsks(ques_id) {
		var quesData = question_map[ques_id]
		var ids = quesData.is_merge ? quesData.ids : [ques_id]
		var targetAsks = []
		for (var i1 = 0; i1 < ids.length; ++i1) {
			var ndata = res2.find(e => {
				return e.id == ids[i1]
			})
			if (ndata && ndata.write_list) {
				targetAsks = targetAsks.concat(ndata.write_list)
			}
		}
		var targetMap = {}
		targetAsks.forEach((e, index) => {
			targetMap[e.id] = index
		})
		var ask_list = quesData.ask_list
		if (ask_list) {
			ask_list = ask_list.sort((a, b) => {
				return targetMap[a.id] - targetMap[b.id]
			})
			quesData.ask_list = ask_list
		}
	}
	function addOtherRemove(ques_id, ask_id) {
		if (question_map[ques_id] && question_map[ques_id].ask_list) {
			var flag = 0
			for (var i = 0; i < question_map[ques_id].ask_list.length; ++i) {
				if (question_map[ques_id].ask_list[i].id == ask_id) {
					flag = 1
				}
				if (question_map[ques_id].ask_list[i].other_fields) {
					if (flag == 0) {
						var j = question_map[ques_id].ask_list[i].other_fields.findIndex(e => {
							return e == ask_id
						})
						if (j >= 0) {
							flag = 2
							question_map[ques_id].ask_list[i].other_fields.splice(j, 1)
							other_asks_remove.push({
								ques_id: ques_id,
								ask_id: question_map[ques_id].ask_list[i].id
							})
						}
					}
					if (flag) {
						question_map[ques_id].ask_list[i].other_fields.forEach(e => {
							other_asks_remove.push({
								ques_id: ques_id,
								ask_id: e
							})
						})
						if (flag == 1) {
							question_map[ques_id].ask_list.splice(i, 1)
						}
						return true
					}
				} else if (flag == 1) {
					question_map[ques_id].ask_list.splice(i, 1)
					--i
					return true
				}
			}
			return true
		}
		return false
	}	
	if (res.typeName == 'mergeQuestion') {
		var writelist = []
		change_list.forEach((item, idx) => {
			if (item.ques_state == 'merge_child') {
				var nodeData = node_list.find(e => {
					return e.id == item.client_id
				})
				var write_list = []
				if (!nodeData) {
					var index = node_list.length
					if (res2) {
						index = res2.findIndex(e => {
							return e.id == item.client_id
						})
						if (index == -1) {
							index = 0
						} else {
							write_list = res2[index].write_list
						}
					}
					node_list.splice(index, 0, {
						id: item.client_id,
						control_id: item.control_id,
						regionType: item.regionType,
						level_type: 'question',
						write_list:  write_list,
						merge_id: res.merge_data.client_id,
						cell_ask: item.cell_ask
					})
				} else {
					nodeData.merge_id = res.merge_data.client_id
					nodeData.level_type = 'question'
					nodeData.regionType = 'question'
					write_list = nodeData.write_list || []
				}
				writelist = writelist.concat(write_list.map(e => {
					return {
						...e,
						parent_node_id: item.client_id
					}
				} ))
				if (question_map[item.client_id]) {
					delete question_map[item.client_id]
				}
			}
		})
		question_map[res.merge_data.client_id] = {
			text: getQuesText(res.merge_data.text),
			ids: res.merge_data.ids,
			is_merge: true,
			level_type: 'question',
			ques_default_name: res.merge_data.numbing_text ? getNumberingText(res.merge_data.numbing_text) : GetDefaultName('question', res.merge_data.text),
			interaction: window.BiyueCustomData.interaction,
			ask_list: writelist.map(e => {
				return {
					parent_node_id: e.parent_node_id,
					id: e.id,
					score: 1
				}
			})
		}
		update_node_id = res.merge_data.client_id
	} else {
		for (var iItem = 0; iItem < change_list.length; ++iItem) {
			var item = change_list[iItem]
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
				if (level_type == 'clear' || level_type == 'clearAll' || level_type == 'clearMerge') {
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
							text: getQuesText(item.text),
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
							question_map[item.client_id].text = getQuesText(item.text)
							question_map[item.client_id].ques_default_name = item.numbing_text ? getNumberingText(item.numbing_text) : GetDefaultName(targetLevel, item.text)
						}
						question_map[item.client_id].level_type = targetLevel
					}
				} else if (targetLevel == 'struct') {
					question_map[item.client_id] = {
						text: getQuesText(item.text),
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
						if (ndata && ndata.write_list) {
							if (parentNode.merge_id) {
								parentNode.cell_ask = ndata.cell_ask
							}
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
							var real_parent_id = parentNode.merge_id ? parentNode.merge_id : parent_id
							if (question_map[real_parent_id]) {
								if (!question_map[real_parent_id].ask_list) {
									question_map[real_parent_id].ask_list = []
								}
								var index2 = question_map[real_parent_id].ask_list.findIndex(e => {
									return e.id == item.client_id
								})
								if (item.type == 'remove') {
									if (index2 >= 0) {
										question_map[real_parent_id].ask_list.splice(index2, 1)
										updateScore(real_parent_id)
									}
								} else if (index2 < 0) {
									var toIndex = 0
									question_map[real_parent_id].ask_list.splice(toIndex, 0, {
										id: item.client_id,
										score: 1
									})
									reSortAsks(real_parent_id)
									updateScore(real_parent_id)
								}
								removeUnvalidAsk(real_parent_id, ndata.write_list)
							}
						}
						if (ndata) {
							parentNode.write_list = ndata.write_list
						}
					}
				}
			} else if (level_type != 'clear' && level_type != 'clearAll' && level_type != 'clearMerge') {
				// 之前没有，需要增加
				var index = node_list.length
				if (res2) {
					index = res2.findIndex(e => {
						return e.id == item.client_id
					})
					if (index == -1) {
						if (item.type == 'remove') {
							continue
						}
						index = 0
					}
				}
				if (index >= 0) {
					node_list.splice(index, 0, {
						id: item.client_id,
						control_id: item.control_id,
						regionType: item.regionType,
						level_type: targetLevel,
						write_list: ask_list
					})
				}
				
				if (targetLevel == 'question') {
					addIds.push(item.client_id)
				} else if (targetLevel == 'struct') {
					update_node_id = item.client_id
				}
				if (targetLevel == 'question' || targetLevel == 'struct') {
					question_map[item.client_id] = {
						text: getQuesText(item.text),
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
					var real_parent_id = nodeData && nodeData.merge_id ? nodeData.merge_id : item.parent_id
					if (addOtherRemove(real_parent_id, item.client_id)) {
						updateScore(real_parent_id)
						addIds.push(real_parent_id)
					}
				}
			}
		}
		if (res.typeName == 'clearMerge' && res.merge_id) {
			if (question_map[res.merge_id]) {
				delete question_map[res.merge_id]
			}
		}
		if (res.typeName == 'clearChildren') {
			if (question_map[res.merge_id]) {
				question_map[res.merge_id].ask_list = []
				question_map[res.merge_id].score = 0
				update_node_id = res.merge_id
			}
		}
	}

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
	var typequesId = 0
	if (res.typeName == 'mergeQuestion') {
		typequesId = res.merge_data.client_id
	} else if (res.typeName == 'question' && (question_map[update_node_id] && question_map[update_node_id].level_type == 'question')) {
		typequesId = update_node_id
	}
	if (typequesId) {
		return reqGetQuestionType([typequesId], 1)
		.then((res2) => {
			if (window.BiyueCustomData.question_map[typequesId].ques_mode == 1 || window.BiyueCustomData.question_map[typequesId].ques_mode == 5) {
				return extractChoiceOptions([typequesId], false)
			} else {
				return new Promise((resolve, reject) => {
					resolve()
				})
			}
		})
		.then(() => {
			return imageAutoLink(typequesId, false)
		}).then((res3) => {
			if (res3.rev) {
				return ShowLinkedWhenclickImage({
					client_id: typequesId
				})
			} else {
				return new Promise((resolve, reject) => {
					resolve()
				})
			}
		})
		.then(() => {
			return splitControl(typequesId)
		})
	}
	if (res.typeName == 'mergedAsk') {
		return new Promise((resolve, reject) => {
			mergeOneAsk({
				ques_id: update_node_id,
				cmd: 'in',
				ask_ids: res.change_list.filter(e => {
					return e.regionType == 'write' && e.type == ''
				}).map(e => {
					return e.client_id
				})
			})
			resolve()
		}).then(() => {
			if (updateinteraction) {
				return setInteraction(interaction, addIds)
			} else {
				return new Promise((resolve, reject) => {
					resolve()
				})
			}
		}).then(() => {
			return notifyQuestionChange(update_node_id)
		})
		.then(() => {
			window.biyue.StoreCustomData()
		})
	}
	if ((res.typeName == 'clear' || res.typeName == 'clearAll') && other_asks_remove.length) {
		// 需要考虑集中作答区，互动同步
		return deleteAsks(other_asks_remove, true, true).then(() => {
			window.biyue.StoreCustomData()
		})
	}
	var updateLinked = res.link_updated
	var use_gather = window.BiyueCustomData.choice_display && window.BiyueCustomData.choice_display.style != 'brackets_choice_region'
	if (use_gather) {
		return deleteChoiceOtherWrite(null, false).then(() => {
			return notifyQuestionChange(update_node_id)
		}).then(() => {
			return updateChoice()
		}).then((res3) => {
			return handleChoiceUpdateResult(res3)
		})
		.then(() => {
			window.biyue.StoreCustomData(() => {
				if (updateLinked) {
					return ShowLinkedWhenclickImage({
						client_id: update_node_id
					})
				}
			})
		})
	} else {
		if (updateinteraction) {
			return deleteChoiceOtherWrite(null, false).then(res3 => {
				return notifyQuestionChange(update_node_id)
			}).then(() => {
				return setInteraction(interaction, addIds).then(() => {
					window.biyue.StoreCustomData(() => {
						if (updateLinked) {
							return ShowLinkedWhenclickImage({
								client_id: update_node_id
							})
						}
					})
				})
			})
		} else {
			deleteChoiceOtherWrite(null, true).then(res => {
				return notifyQuestionChange(update_node_id)
			}).then(() => {
				window.biyue.StoreCustomData(() => {
					if (updateLinked) {
						return ShowLinkedWhenclickImage({
							client_id: update_node_id
						})
					}
				})
			})
		}
	}
}
function notifyQuestionChange(update_node_id) {
	if (window.tab_select == 'tabTree' && window.tree_lock) {
		return refreshTree()
	}
	return new Promise((resolve, reject) => {
		var eventname = window.tab_select != 'tabQues' ? 'clickSingleQues' : 'updateQuesData'
		document.dispatchEvent(
			new CustomEvent(eventname, {
				detail: {
					client_id: update_node_id
				}
			})
		)
		resolve()
	})
	
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
						var tag = Api.ParseJSON(oControl.GetTag() || '{}')
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
											var nextTag = Api.ParseJSON(nextControl.GetTag() || '{}')
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
													var childTag = Api.ParseJSON(e.GetTag() || '{}')
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
					var tag = Api.ParseJSON(e.GetTag() || '{}')
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
		var choice_option_add = []
		var choice_option_remove = []
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
					if (oldMode != 1 && oldMode != 5) {
						choice_option_add.push(e.id)
					}
				} else if (oldMode == 1 || oldMode == 5) {
					choice_option_remove.push(e.id)
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
			choice_option_add,
			choice_option_remove,
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
	})).then((res) => {
		if (res.choice_option_add.length) {
			return extractChoiceOptions(res.choice_option_add, true)
		} else if (res.choice_option_remove.length) {
			return removeChoiceOptions(res.choice_option_remove)
		} else {
			return new Promise((resolve, reject) => {
				resolve()
			})
		}
	})
	.then(() => {
		window.biyue.sendMessageToWindow('batchQuestionTypeWindow', 'quesMapUpdate', {
			question_map: window.BiyueCustomData.question_map
		})
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
		window.biyue.showDialog('levelSetWindow', '自动序号识别设置', 'levelSet.html', 592, 400)
		resolve()
	})
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
			var tagInfo = Api.ParseJSON(oControl.GetTag())
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
					if (tagInfo.mid) {
						if (ids[tagInfo.mid]) {
							ids[tagInfo.mid].push(tagInfo.client_id)
						} else {
							ids[tagInfo.mid] = [tagInfo.client_id]
						}
					} else {
						ids[tagInfo.client_id] = 1
					}
				}
				var parentControl = getParentBlock(oControl)
				if (parentControl) {
					var parentTagInfo = Api.ParseJSON(parentControl.GetTag())
					if (parentTagInfo.regionType == 'question') {
						parentid = parentTagInfo.client_id
					}
				}
				var ndata = {
					id: tagInfo.client_id,
					text: oControl.GetRange().GetText(),
					parentId: parentid,
					numbing_text: GetNumberingValue(oControl),
					control_id: oControl.Sdt.GetId(),
					sub_type: 'control'
				}
				if (tagInfo.mid) {
					ndata.merge_id = tagInfo.mid
				}
				nodeList.push(ndata)
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
						tagInfo.clr = tagInfo.color = '#ff000040'
					}
				} else if (oControl.GetClassType() == 'blockLvlSdt') {
					var quesData = tagInfo.mid ? question_map[tagInfo.mid] : question_map[tagInfo.client_id]
					if (quesData) {
						if (quesData.level_type == 'question') {
							tagInfo.clr = tagInfo.color = '#d9d9d940'
							changecolor = true	
						} else if (quesData.level_type == 'struct') {
							tagInfo.clr = tagInfo.color = '#CFF4FF80'
							changecolor = true	
						}
					}
				}
			} else if (tagInfo.regionType == 'num') {
				tagInfo.clr = tagInfo.color = '#ffffff40'
				changecolor = true
			} else if (tagInfo.regionType == 'choiceOption') {
				tagInfo.clr = tagInfo.color = '#00ff0020'
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
			var title = oDrawing.GetTitle()
			var titleObj = Api.ParseJSON(title)
			if (titleObj.feature && titleObj.feature.client_id) {
				if (maxid < titleObj.feature.client_id * 1) {
					maxid = titleObj.feature.client_id * 1
				}
			}
			if(titleObj && titleObj.feature && titleObj.feature.zone_type == 'question' && (titleObj.feature.sub_type == 'write' || titleObj.feature.sub_type == 'identify')) {
				drawingList.push({
					id: titleObj.feature.client_id,
					shape_id: oDrawing.Drawing.Id,
					drawing_id: oDrawing.Drawing.Id,
					sub_type: titleObj.feature.sub_type
				})
			}
			// 图片不铺码
			if (titleObj.feature && titleObj.feature.partical_no_dot) {
				oDrawing.SetShadow(null, 0, 100, null, 0, '#0fc1fd')
			}
		})
		var cellAskMap = {}
		oTables.forEach(oTable => {
			if (oTable.GetPosInParent() >= 0) {
				var desc = Api.ParseJSON(oTable.GetTableDescription())
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
				var newNodeList = []
				nodeList.forEach(node => {
					if (question_map[node.id]) {
						question_map[node.id].text = getQuesText(node.text)
						question_map[node.id].ques_default_name = node.numbing_text ? getNumberingText(node.numbing_text) : GetDefaultName(question_map[node.id].level_type, node.text)
					}
					var nodeData = node_list.find(e => {
						return e.id == node.id
					})
					if (nodeData) {
						nodeData.control_id = node.control_id
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
						newNodeList.push(nodeData)
					}
				})
				window.BiyueCustomData.node_list = newNodeList
				Object.keys(question_map).forEach(id => {
					if (!ids[id]) {
						delete question_map[id]
					} else if (typeof ids[id] == 'object') {
						question_map[id].ids = ids[id]
						var textlist = []
						var numbing_text = ''
						question_map[id].ids.forEach(e => {
							var ndata = nodeList.find(e2 => {
								return e == e2.id && e2.merge_id == id
							})
							if (ndata) {
								textlist.push(getQuesText(ndata.text))
								var ntext = getNumberingText(ndata.numbing_text)
								if (!numbing_text && !ntext) {
									numbing_text = ntext
								}
							}
						})
						question_map[id].text = textlist.join('')
						question_map[id].numbing_text = numbing_text
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
		var controls = oDocument.GetAllContentControls() || []
		var paragrahs = oDocument.GetAllParagraphs() || []
		paragrahs.forEach(oParagraph => {
			var oParaPr = oParagraph.GetParaPr();
			oParaPr.SetShd("clear", 255, 255, 255, true);
		})
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
								var oNumberingLevel = oParagraph.GetNumbering()
								return {text: oParagraph.Paragraph.GetNumberingText(), lvl: oNumberingLevel ? oNumberingLevel.Lvl : 0} 
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
			var tagInfo = Api.ParseJSON(oControl.GetTag())
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
						var parent_tagInfo = Api.ParseJSON(parentControl.GetTag())
						var parentNode = nodeList.find((node) => {
							return node.id == parent_tagInfo.client_id
						})
						if (parentNode) {
							if (parentNode.level_type == 'struct') {
								Api.asc_RemoveContentControlWrapper(oControl.Sdt.GetId())
							} else {
								parentNode.write_list.push({
									id: id,
									control_id: oControl.Sdt.GetId(),
									sub_type: 'control'
								})
								questionMap[parent_tagInfo.client_id].ask_list.push({
									id: id,
									score: 1
								})
								questionMap[parent_tagInfo.client_id].score += 1
								nodeData.parent_id = parent_tagInfo.client_id
								tagInfo.color = '#ff000040'
							}
						}
					}
				} else {
					var proportion = getProportion(oControl)
					var text = oControl.GetRange().GetText()
					var numberingInfo = GetNumberingValue(oControl)
					var lvl = numberingInfo && numberingInfo.lvl ? numberingInfo.lvl : tagInfo.lvl
					if (lvl != tagInfo.lvl) {
						tagInfo.lvl = lvl
					}
					var level_type = levelmap[lvl] || 'question'
					nodeData.level_type = level_type
					nodeData.proportion = proportion
					var detail = {
						text: text,
						ask_list: [],
						level_type: level_type,
						numbing_text: numberingInfo ? numberingInfo.text : '',
						proportion: proportion
					}
					if (tagInfo.regionType == 'question') {
						nodeData.write_list = []
						detail.ask_list = []
						detail.score = 0
						if (level_type == 'question') {
							tagInfo.clr = tagInfo.color = '#d9d9d940'
						} else if (level_type == 'struct') {
							tagInfo.clr = tagInfo.color = '#CFF4FF80'
						} else {
							tagInfo.clr = tagInfo.color = '#ffffff'
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
		return refreshTree()
	})
	.then(() => {
		return extractChoiceOptions()
	})
	.then(() => {
		return imageAutoLink()
	}).then(() => {
		console.log("================================ StoreCustomData")
		window.biyue.StoreCustomData()
	})
}

function getQuesText(text) {
	if (!text || typeof text != 'string') {
		return ''
	}
	return text.replace(/[\ue749\ue6a1\ue607]/g, '');
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
		if (result) {
			return result[0]
		} else {
			var texts2 = text.split('\r')
			if (texts2 && texts2.length) {
				return text.substring(0, Math.min(8, texts2[0].length))
			}
			return text
		}
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

function getQuestionHtml(ids, getLatestParent) {
	Asc.scope.question_map = window.BiyueCustomData.question_map
	Asc.scope.html_ids = ids
	Asc.scope.getLatestParent = getLatestParent
	return biyueCallCommand(window, function() {
		var question_map = Asc.scope.question_map || {}
		var target_list = []
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls() || []
		var handled = {}
		var html_ids = Asc.scope.html_ids
		var getLatestParent = Asc.scope.getLatestParent
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
		function addHtml(qId, quesId, quesData, oControl, lvldata) {
			if (!quesData) {
				return
			}
			var oRange = null
			if (quesData.is_merge && quesData.ids) {
				quesData.ids.forEach(id => {
					var control = getControlsByClientId(id)
					if (control) {
						var parentcell = control.GetParentTableCell()
						if (parentcell) {
							if (!oRange) {
								oRange = parentcell.GetContent().GetRange()
							} else {
								oRange = oRange.ExpandTo(parentcell.GetContent().GetRange())
							}
						}
					}
				})
			} else {
				oRange = oControl.GetRange()
			}
			if (oRange) {
				oRange.Select()
				let text_data = {
					data:     "",
					// 返回的数据中class属性里面有binary格式的dom信息，需要删除掉
					pushData: function (format, value) {
						this.data = value ? value.replace(/class="[a-zA-Z0-9-:;+"\/=]*/g, "") : "";
					}
				};

				Api.asc_CheckCopy(text_data, 2);
				var lvltext = ''
				if (lvldata && lvldata.text) {
					lvltext = `<div>${lvldata.text}</div>`
				}
				var obj = {
					id: quesId + '',
					content_type: quesData.level_type,
					content_html:  lvltext + text_data.data
				}
				var index = target_list.findIndex(e => {
					return e.id == qId
				})
				if (quesId != qId) {
					if (getLatestParent == 1) {
						target_list.push(obj)
					} else {
						if (index >= 0) {
							target_list[index].context_list.push(obj)
						} else {
							target_list.push({
								id: qId + '',
								context_list: [obj]
							})
						}
					}
				} else {
					if (index >= 0) {
						target_list[index].content_type = quesData.level_type
						target_list[index].content_html = lvltext + text_data.data
					} else {
						target_list.push(obj)
					}
				}
			}
		}
		function getFirstParagraph(oControl) {
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
				return null
			}
			var paragraphs = oControl.GetAllParagraphs()
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
		function getLvl(oControl, paraIndex) {
			var oParagraph = getFirstParagraph(oControl)
			if (!oParagraph) {
				return null
			}
			var oNumberingLevel = oParagraph.GetNumbering()
			if (oNumberingLevel) {
				if (oNumberingLevel.Num) {
					var lvl = oNumberingLevel.Num.GetLvl(oNumberingLevel.Lvl)
					if (lvl) {
						var LvlText = lvl.GetLvlText() || []
						var list = []
						LvlText.forEach(e => {
							list.push(e.Value)
						})
						return {
							lvl: oNumberingLevel.Lvl,
							text: list.join('')
						}
					}
				}
				
				return {
					lvl: oNumberingLevel.Lvl,
					text: ''
				}
			}
			return null
		}
		function getQueId(oControl) {
			var tagInfo = Api.ParseJSON(oControl.GetTag())
			if (!tagInfo.client_id) {
				return null
			}
			return tagInfo.mid && question_map[tagInfo.mid] ? tagInfo.mid : tagInfo.client_id
		}
		for (var i = 0; i < controls.length; ++i) {
			var oControl = controls[i]
			var quesId = getQueId(oControl)
			if (handled[quesId]) {
				continue
			}
			if (html_ids && html_ids.findIndex(e => { return e == quesId }) < 0) {
				continue
			}
			var quesData = question_map[quesId]
			if (!quesData || (quesData.level_type != 'question' && quesData.level_type != 'struct')) {
				continue
			}
			handled[quesId] = true
			var lvl1 = getLvl(oControl, 0)
			if (getLatestParent) {
				var flag = 0
				var oParentControl = oControl.GetParentContentControl()
				if (oParentControl) {
					var parentId = getQueId(oParentControl)
					if (parentId && question_map[parentId]) {
						addHtml(quesId, parentId, question_map[parentId], oParentControl, getLvl(oParentControl, 0))
						flag = 1
					}
				}
				if (!flag) {
					var prelvl
					for (var j = i - 1; j >= 0; --j) {
						var oControl2 = controls[j]
						if (oControl2.GetClassType() != 'blockLvlSdt') {
							continue
						}
						var tag2 = Api.ParseJSON(oControl2.GetTag())
						if (!tag2.client_id) {
							continue
						}
						var quesId2 = tag2.mid && question_map[tag2.mid] ? tag2.mid : tag2.client_id
						if (question_map[quesId2] && question_map[quesId2].level_type == 'struct') {
							addHtml(quesId, quesId2, question_map[quesId2], oControl2)
							break
						}
						var lvl2 = getLvl(oControl2, 0)
						if (lvl1 == null) {
							if (lvl2 !== null) {
								if (prelvl) {
								} else {
									if (question_map[quesId2] && question_map[quesId2].ques_mode == 6) { // 文本题
										addHtml(quesId, quesId2, question_map[quesId2], oControl2, lvl2)
										break
									} else {
										if (prelvl) {
											if (lvl2.lvl < prelvl) {
												addHtml(quesId, quesId2, question_map[quesId2], oControl2, lvl2)
												break
											}
										} else {
											prelvl = lvl2.lvl
										}
									}
								}
							}
						} else {
							if (lvl2 === null) {
								continue
							} else if (lvl2.lvl < lvl1.lvl) {
								addHtml(quesId, quesId2, question_map[quesId2], oControl2, lvl2)
								break
							}
						}
					}
				}
			}
			addHtml(quesId, quesId, quesData, oControl, lvl1)
		}
		return target_list
	}, false, false)
}

// 获取题型
function reqGetQuestionType(ids, getLatestParent) {
	return getQuestionHtml(ids, getLatestParent).then(control_list => {
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
							var titleObj = Api.ParseJSON(drawing.docPr.title)
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
		function getParentBlock(oControl) {
			if (!oControl) {
				return null
			}
			if (oControl.GetClassType() == 'inlineLvlSdt') {
				var parentControl = oControl.GetParentContentControl()
				if (parentControl) {
					if (parentControl.GetClassType() == 'blockLvlSdt') {
						var parentTag = Api.ParseJSON(parentControl.GetTag())
						if (parentTag.client_id) {
							return {
								client_id: parentTag.client_id,
								control: parentControl
							}
						}
					}
					return getParentBlock(parentControl)
				}
			}
			return null
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
				var tag = Api.ParseJSON(e.GetTag())
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
				var childTag = Api.ParseJSON(childControls[j].GetTag())
				if (childTag.client_id) {
					if (question_map[childTag.client_id]) {
						continue
					}
					var parentData = getParentBlock(childControls[j])
					if (parentData) {
						if (parentData.client_id) {
							if (parentData.client_id != nodeData.id && question_map[parentData.client_id]) {
								continue
							}
						}
					}
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
					if (childTag.regionType != 'num' && childTag.regionType != 'choiceOption') {
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
					Api.asc_RemoveContentControlWrapper(validIds[k].control_id) // 这里会把所有的choice option control都删除掉
				}
				if (question_map[id].interaction != 'none') {
					updateAccurateText(validIds[blankIndex].control_id)
				}
				var score = 0
				question_map[id].ask_list.forEach(e => {
					score += e.score * 1
				})
				question_map[id].score = score
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
		function updateFill(oDrawing, oFill) {
			if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== 'fill') {
				return false
			}
			oDrawing.Drawing.spPr.setFill(oFill.UniFill)
		}
		var list = []
		for (var i = 0, imax = drawings.length; i < imax; ++i) {
			var oDrawing = drawings[i]
			var titleObj = Api.ParseJSON(oDrawing.GetTitle())
			if (titleObj.feature && titleObj.feature.zone_type == 'question' && titleObj.feature.sub_type == 'write') {
				if (cmdType == 'show') {
					var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 0, 0))
					oFill.UniFill.transparent = 255 * 0.2 // 透明度
					updateFill(oDrawing, oFill)
				} else if (cmdType == 'hide') {
					updateFill(oDrawing, Api.CreateNoFill())
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
		  function getCell(write_data) {
			for (var i = 0; i < oTables.length; ++i) {
				var oTable = oTables[i]
				if (oTable.GetPosInParent() == -1) { continue }
				var desc = Api.ParseJSON(oTable.GetTableDescription())
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


// 全量更新
function reqUploadTree() {
	if (isLoading('uploadTree')) {
		return
	}
  	// 先关闭智批元素，避免智批元素在全量更新的时候被带到题目里 更新之后再打开
  	setBtnLoading('uploadTree', true)
	return setInteraction('none', null, false).then(() => {
		return preGetExamTree()	// 获取题目树需要在addOnlyBigControl之前执行，否则可能出现父节点出错的情况
	}).then((res) => {
		Asc.scope.tree_info = res
		return addOnlyBigControl(false)
	}).then(() => {
		return handleUploadPrepare('hide')
	}).then(() => {
		return getChoiceOptionAndSteam() // getChoiceQuesData()
	}).then((res) => {
		Asc.scope.choice_html_map = res
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
// 后端已支持结构和题目可同级出现在结构下，取代旧代码
function getControlListForUpload() {
	Asc.scope.node_list = window.BiyueCustomData.node_list
    Asc.scope.question_map = window.BiyueCustomData.question_map
	return biyueCallCommand(window, function() {
		var target_list = []
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		var question_map = Asc.scope.question_map
		var tree_info = Asc.scope.tree_info || {}
		var handledcontrol = {}
		for (var i = 0, imax = controls.length; i < imax; ++i) {
			var oControl = controls[i]
			if (handledcontrol[oControl.Sdt.GetId()]) {
				continue
			}
			var tag = Api.ParseJSON(oControl.GetTag() || '{}')
			if (tag.regionType != 'question' || !tag.client_id) {
				continue
			}
			var quesData
			var clientid = tag.mid ? tag.mid : tag.client_id
			var quesData = question_map[clientid]
			if (!quesData) {
				continue
			}
			if (!question_map[clientid].level_type) {
				continue
			}
			if (!tree_info.list) {
				continue
			}
			var itemData = tree_info.list.find(e => {
				return e.id == clientid
			})
			if (!itemData) {
				continue
			}
			var parent_id = itemData.parent_id
			var useControl = oControl
			if (tag.big) {
				var childcontrols = oControl.GetAllContentControls() || []
				var bigControl = childcontrols.find(e => {
					var btag = Api.ParseJSON(e.GetTag())
					return e.GetClassType() == 'blockLvlSdt' && btag.onlybig == 1 && btag.link_id == tag.client_id
				})
				if (bigControl) {
					useControl = bigControl
				}
			}
			var oRange = null
			if (tag.mid) {
				for (var idkey in quesData.ids) {
					var control = controls.find(e => {
						var tag2 = Api.ParseJSON(e.GetTag())
						return e.GetClassType() == 'blockLvlSdt' && e.GetPosInParent() >= 0 && tag2.client_id == quesData.ids[idkey] && tag2.mid == tag.mid
					})
					if (control) {
						handledcontrol[control.Sdt.GetId()] = 1
						var parentcell = control.GetParentTableCell()
						if (!oRange) {
							oRange = parentcell.GetContent().GetRange()
						} else {
							oRange = oRange.ExpandTo(parentcell.GetContent().GetRange())
						}
					}
				}
			} else {
				handledcontrol[oControl.Sdt.GetId()] = 1
				oRange = useControl.GetRange()
			}
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
				id: clientid + '',
				parent_id: parent_id ? parent_id + '' : '',
				uuid: question_map[clientid].uuid || '',
				regionType: question_map[clientid].level_type,
				content_type: question_map[clientid].level_type,
				content_xml: '',
				content_html: content_html,
				content_text: oRange.GetText(),
				question_type: question_map[clientid].question_type,
				question_name: question_map[clientid].ques_name || question_map[clientid].ques_default_name,
				control_id: oControl.Sdt.GetId(),
				lvl: itemData.lvl
			})
		}
		console.log('target_list', target_list)
		return target_list
	  }, false, false)
}

function getControlListForUpload2() {
	Asc.scope.node_list = window.BiyueCustomData.node_list
    Asc.scope.question_map = window.BiyueCustomData.question_map
	return biyueCallCommand(window, function() {
		var target_list = []
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		var question_map = Asc.scope.question_map
		console.log('question_map', question_map)
		var handledcontrol = {}
		function getFirstParagraph(oControl) {
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
				return null
			}
			var paragraphs = oControl.GetAllParagraphs()
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
		function getLvl(oControl, paraIndex) {
			var oParagraph = getFirstParagraph(oControl)
			if (!oParagraph) {
				return null
			}
			var oNumberingLevel = oParagraph.GetNumbering()
			if (oNumberingLevel) {
				return oNumberingLevel.Lvl
			}
			return null
		}
		function getParentId(target_list, lvl) {
			var parent_id = 0
			for (var j = target_list.length - 1; j >= 0; --j) {
				var preNode = target_list[j]
				// 由于struct未必有lvl，因此先将Lvl的判断移除
				if (preNode.content_type == 'struct') {
					if (lvl === 0) {
						parent_id = 0
					} else if (!preNode.lvl) {
						parent_id = preNode.id
					} else if (lvl && lvl > preNode.lvl) {
						parent_id = preNode.id
					} else {
						parent_id = 0
					}
					return parent_id
				}
			}
			return parent_id
		}
		for (var i = 0, imax = controls.length; i < imax; ++i) {
			var oControl = controls[i]
			if (handledcontrol[oControl.Sdt.GetId()]) {
				continue
			}
			var tag = Api.ParseJSON(oControl.GetTag() || '{}')
			if (tag.regionType != 'question' || !tag.client_id) {
				continue
			}
			var quesData
			var clientid = tag.mid ? tag.mid : tag.client_id
			var quesData = question_map[clientid]
			if (!quesData) {
				continue
			}
			if (!question_map[clientid].level_type) {
				continue
			}
			var lvl = getLvl(oControl, 0)
			var oParentControl = oControl.GetParentContentControl()
			var parent_id = 0
			if (question_map[clientid].level_type == 'question') {
				if (oParentControl) {
					var parentTag = Api.ParseJSON(oParentControl.GetTag() || '{}')
					parent_id = parentTag.client_id
				} else {
					// 根据level, 查找在它前面的比它lvl小的struct
					parent_id = getParentId(target_list, lvl)
				}
			}
			var useControl = oControl
			if (tag.big) {
				var childcontrols = oControl.GetAllContentControls() || []
				var bigControl = childcontrols.find(e => {
					var btag = Api.ParseJSON(e.GetTag())
					return e.GetClassType() == 'blockLvlSdt' && btag.onlybig == 1 && btag.link_id == tag.client_id
				})
				if (bigControl) {
					useControl = bigControl
				}
				parent_id = getParentId(target_list, lvl)
			}
			var oRange = null
			if (tag.mid) {
				for (var idkey in quesData.ids) {
					var control = controls.find(e => {
						var tag2 = Api.ParseJSON(e.GetTag())
						return e.GetClassType() == 'blockLvlSdt' && e.GetPosInParent() >= 0 && tag2.client_id == quesData.ids[idkey] && tag2.mid == tag.mid
					})
					if (control) {
						handledcontrol[control.Sdt.GetId()] = 1
						var parentcell = control.GetParentTableCell()
						if (!oRange) {
							oRange = parentcell.GetContent().GetRange()
						} else {
							oRange = oRange.ExpandTo(parentcell.GetContent().GetRange())
						}
					}
				}
			} else {
				handledcontrol[oControl.Sdt.GetId()] = 1
				oRange = useControl.GetRange()
			}
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
				id: clientid,
				parent_id: parent_id,
				uuid: question_map[clientid].uuid || '',
				regionType: question_map[clientid].level_type,
				content_type: question_map[clientid].level_type,
				content_xml: '',
				content_html: content_html,
				content_text: oRange.GetText(),
				question_type: question_map[clientid].question_type,
				question_name: question_map[clientid].ques_name || question_map[clientid].ques_default_name,
				control_id: oControl.Sdt.GetId(),
				lvl: lvl
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

// 后端已支持结构和题目可同级出现在结构下，取代旧代码
function generateTreeForUpload(control_list) {
	return new Promise((resolve, reject) => {
		if (!control_list) {
			reject(null)
		}
		var tree = []
		var choicemap = Asc.scope.choice_html_map || {}

		const map = {};
		control_list.forEach(e => {
			e.content_html = cleanHtml(e.content_html || '')
			if (choicemap[e.id]) {
				e.content_without_opt = cleanHtml(choicemap[e.id].steam || '')
				e.options = []
				if (choicemap[e.id].options) {
					choicemap[e.id].options.forEach(option => {
						e.options.push({
							value: option.value,
							html: cleanHtml(option.html)
						})
					})
				}
				e.option_type = choicemap[e.id].option_type
			}
			map[e.id] = { ...e, children: [] };
		});
		control_list.forEach(item => {
			if (item.parent_id && item.parent_id != item.id) {
				if (map[item.parent_id] && map[item.parent_id].children) {
					map[item.parent_id].children.push(map[item.id]);
				} else {
					console.log('        cannot find parent', item);
				}
			} else {
				tree.push(map[item.id]);
			}
		});
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
				Object.keys(window.BiyueCustomData.question_map).forEach(e => {
					var index = res.data.questions.findIndex(e2 => {
						return e2.uuid == window.BiyueCustomData.question_map[e].uuid
					})
					if (index == -1) {
						window.BiyueCustomData.question_map[e].uuid = ''
					}
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

function generateTreeForUpload2(control_list) {
	return new Promise((resolve, reject) => {
		if (!control_list) {
			reject(null)
		}
		var tree = []
		var choicemap = Asc.scope.choice_html_map || {}
		control_list.forEach((e) => {
			e.content_html = cleanHtml(e.content_html || '')
			if (choicemap[e.id]) {
				e.content_without_opt = cleanHtml(choicemap[e.id].steam || '')
				e.options = []
				if (choicemap[e.id].options) {
					choicemap[e.id].options.forEach(option => {
						e.options.push({
							value: option.value,
							html: cleanHtml(option.html)
						})
					})
				}
				e.option_type = choicemap[e.id].option_type
			}
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
				Object.keys(window.BiyueCustomData.question_map).forEach(e => {
					var index = res.data.questions.findIndex(e2 => {
						return e2.uuid == window.BiyueCustomData.question_map[e].uuid
					})
					if (index == -1) {
						window.BiyueCustomData.question_map[e].uuid = ''
					}
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
		return proportionHandler.batchProportion(res.list, proportion)
	}).then(() => {
		return setChoiceOptionLayout({
			list: (Asc.scope.change_id_list || []).map(e => {
				return e.id
			}),
			from: 'proportion'
		})	
	}).then(() => {
		return window.biyue.StoreCustomData()
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

		function deleShape(oShape) {
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
						var titleObj = Api.ParseJSON(oDrawing.GetTitle())
						if (titleObj.feature && titleObj.feature.zone_type == 'question') {
							oDrawing.Delete()
						}
					}
				}
			}
		}
		function removeCellAskRecord(oCell) {
			var oTable = oCell.GetParentTable()
			if (oTable && oTable.GetPosInParent() >= 0) {
				var desc = Api.ParseJSON(oTable.GetTableDescription())
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
							var tag = Api.ParseJSON(oControl.GetTag())
							if (tag.client_id == client_id) {
								return oControl
							}
						}
					}
				}
				var oControl = allControls.find(e => {
					var tag = Api.ParseJSON(e.GetTag())
					return tag.client_id == client_id
				})
				return oControl
			}
			return null
		}
		function deleteOneNode(node_id, index, sum, quesId) {
			var nodeData = node_list.find(e => {
				return e.id == node_id
			})
			if (!nodeData) {
				return
			}
			var quesControl = getControl(node_id, nodeData.control_id)
			if (!quesControl || quesControl.GetClassType() != 'blockLvlSdt') {
				return
			}
			if (aid == 0) {
				// 删除所有小问，遍历出所有小问，全部删除
				// 删除所有精准互动
				var childDrawings = quesControl.GetAllDrawingObjects() || []
				for (var j = 0; j < childDrawings.length; ++j) {
					var tag = Api.ParseJSON(childDrawings[j].GetTitle())
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
						var tag = Api.ParseJSON(e.GetTag())
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
					var paraDrawing = childDrawings[j].getParaDrawing()
					if (!paraDrawing) {
						continue
					}
					var tag = Api.ParseJSON(childDrawings[j].GetTitle())
					if (tag.feature && tag.feature.zone_type == 'question') {
						if (tag.feature.sub_type == 'write') {
							var run = paraDrawing.GetRun()
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
				// 删除单元格小问
				if (nodeData.cell_ask) {
					var tag = Api.ParseJSON(quesControl.GetTag())
					delete tag.cask
					quesControl.SetTag(JSON.stringify(tag))
					var parentCell = quesControl.GetParentTableCell()
					if (parentCell) {
						parentCell.SetBackgroundColor(255, 191, 191, true)
						removeCellInteraction(parentCell)
						removeCellAskRecord(parentCell)
					}
					nodeData.cell_ask = false
				}
				nodeData.write_list = []
				if (index == sum) {
					question_map[quesId].ask_list = []
					question_map[quesId].score = 0
				}
			} else {
				if (!nodeData.write_list) {
					return
				}
				var writeIndex = nodeData.write_list.findIndex(e => {
					return e.id == aid
				})
				if (writeIndex == -1) {
					return
				}
				var writeData = nodeData.write_list[writeIndex]
				if (!writeData) {
					return
				}
				nodeData.write_list.splice(writeIndex, 1)
				// todo..question_map 的数据更新
				if (writeData.sub_type == 'control') {
					var oControl = Api.LookupObject(writeData.control_id)
					clearQuesInteraction(oControl)
					Api.asc_RemoveContentControlWrapper(writeData.control_id)
				} else if (writeData.sub_type == 'cell') {
					if (nodeData.cell_ask) {
						var tag = Api.ParseJSON(quesControl.GetTag())
						delete tag.cask
						quesControl.SetTag(JSON.stringify(tag))
						nodeData.cell_ask = false
					}
					var oCell = Api.LookupObject(writeData.cell_id)
					removeCellInteraction(oCell)
				} else if (writeData.sub_type == 'write' || writeData.sub_type == 'identify') {
					var oDrawing = drawings.find(e => {
						var tag = Api.ParseJSON(e.GetTitle())
						return tag.feature && tag.feature.client_id == writeData.id
					})
					if (oDrawing) {
						if (writeData.sub_type == 'identify') {
							oDrawing.Delete()
						} else {
							var paraDrawing = oDrawing.getParaDrawing()
							var run = paraDrawing ? paraDrawing.GetRun() : null
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
			if (question_map[qid] && question_map[qid].is_merge) { // 合并题
				question_map[qid].ids.forEach((e, index) => {
					deleteOneNode(e, index, question_map[qid].ids.length - 1, qid)
				})
			} else {
				deleteOneNode(qid, 1, 1, qid)
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
		} else {
			return new Promise((resolve, reject) => {
				return resolve({})
			})
		}
	})
}

function focusControl(id) {
	var quesData = window.BiyueCustomData.question_map[id]
	if (!quesData) {
		return
	}
	Asc.scope.focus_ids = quesData.level_type == 'question' && quesData.is_merge ? quesData.ids : [id]
	return biyueCallCommand(window, function() {
		var focusIds = Asc.scope.focus_ids
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
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
		for (var id of focusIds) {
			var oControl = getControlsByClientId(id)
			if (oControl) {
				oDocument.Document.MoveCursorToContentControl(oControl.Sdt.GetId(), true)
				return {
					Tag: Api.ParseJSON(oControl.GetTag()),
					InternalId: oControl.Sdt.GetId()
				}
			}
		}
		return null
	}, false, false).then((res) => {
		g_click_value = res
		return new Promise((resolve, reject) => {
			resolve()
		})
	})
}

function focusAsk(writeData) {
	if (!writeData || !(writeData.length)) {
		return
	}
	Asc.scope.write_data = writeData
	return biyueCallCommand(window, function() {
		var writeList = Asc.scope.write_data || []
		var write_data = writeList[0]
		var oDocument = Api.GetDocument()
		var drawings = oDocument.GetAllDrawingObjects() || []
		var controls = oDocument.GetAllContentControls() || []
		var oTables = oDocument.GetAllTables() || []
		function getCell(wData) {
			for (var i = 0; i < oTables.length; ++i) {
				var oTable = oTables[i]
				if (oTable.GetPosInParent() == -1) { continue }
				var desc = Api.ParseJSON(oTable.GetTableDescription())
				var keys = Object.keys(desc)
				if (keys.length) {
					for (var j = 0; j < keys.length; ++j) {
						var key = keys[j]
						if (desc[key] == wData.id) {
							var rc = key.split('_')
							if (wData.row_index == undefined) {
								return oTable.GetCell(rc[0], rc[1])
							} else if (wData.row_index == rc[0] && wData.cell_index == rc[1]) {
								return oTable.GetCell(rc[0], rc[1])
							}
						}
					}
				}
			}
			return null
		}
		if (write_data.sub_type == 'control') {
			var oRange = null
			var ids = []
			for (var wData of writeList) {
				if (wData.sub_type != 'control') {
					continue
				}
				var oControls = controls.filter(e => {
					var tag = Api.ParseJSON(e.GetTag())
					if (tag.client_id == wData.id && e.Sdt) {
						if (e.GetClassType() == 'blockLvlSdt') {
							return e.GetPosInParent() >= 0
						} else if (e.GetClassType() == 'inlineLvlSdt') {
							return e.Sdt.GetPosInParent() >= 0
						}
					}
				})
				if (oControls && oControls.length) {
					if (oControls.length == 1) {
						ids.push(oControls[0].Sdt.GetId())
						if (oRange) {
							oRange = oRange.ExpandTo(oControls[0].GetRange())
						} else {
							oRange = oControls[0].GetRange()
						}
					}
				}
			}
			if (ids.length == 1) {
				oDocument.Document.MoveCursorToContentControl(ids[0], true)
			} else if (oRange) {
				oRange.Select()
			}
		} else if (write_data.sub_type == 'cell') {
			var oRange = null
			for (var wData of writeList) {
				if (wData.cell_id) {
					var oCell = Api.LookupObject(wData.cell_id)
					if (oCell && oCell.GetClassType() == 'tableCell') {
						var table = oCell.GetParentTable()
						if (table.GetPosInParent() == -1) {
							oCell = getCell(wData)
						}
						if (oCell) {
							var cellContent = oCell.GetContent()
							if (cellContent) {
								if (oRange) {
									oRange = oRange.ExpandTo(cellContent.GetRange())
								} else {
									oRange = cellContent.GetRange()
								}
							}
						}
					}
				}
			}
			if (oRange) {
				oRange.Select()
			}
		} else if (write_data.sub_type == 'write' || write_data.sub_type == 'identify') {
			var oDrawing = drawings.find(e => {
				var tag = Api.ParseJSON(e.GetTitle())
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
		// 若调用oState的方式，oDrawing.Drawing.selected得到的值会是false，因而这里暂时注释
		// var oState = oDocument.Document.SaveDocumentState()
		drawings.forEach(oDrawing => {
			if (oDrawing.Drawing.selected) {
				var tag = Api.ParseJSON(oDrawing.GetTitle())
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
				oDrawing.SetTitle(JSON.stringify(tag))
				if (cmdType == 'add') {
					oDrawing.SetShadow(null, 0, 100, null, 0, '#0fc1fd')
				} else {
					oDrawing.ClearShadow()
				}
			}
		})
		// oDocument.Document.LoadDocumentState(oState)
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

function batchChangeScore() {
	if (g_click_value && g_click_value.Tag && g_click_value.Tag.client_id) {
		window.biyue.onBatchScoreSet(g_click_value.Tag.client_id)
	} else {
		window.biyue.onBatchScoreSet()
	}
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
		function getShape(id) {
			if (!shapes) {
				return null
			}
			for (var i = 0, imax = shapes.length; i < imax; ++i) {
				var oShape = Api.LookupObject(shapes[i].Shape.Id)
				var titleObj = Api.ParseJSON(oShape.GetTitle())
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
			var paraDrawing = oDrawing.getParaDrawing()
			var run = paraDrawing ? paraDrawing.GetRun() : null
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
				var tag = Api.ParseJSON(e.GetTag())
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
			oDrawing.SetTitle(JSON.stringify(titleobj))
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
		const WORDS = [0xe753, 0xe754, 0xe755, 0xe756, 0xe757, 0xe758]
		var oDocument = Api.GetDocument()
		var result = {
			change_list: [],
			typeName: 'write'
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
				var parentTag = Api.ParseJSON(parentBlock.GetTag())
				parent_id = parentTag.client_id || 0
			}
			return parent_id
		}
		function handleParagraph(oParagraph, parentId, quesId) {
			if (!oParagraph) {
				return
			}
			for (var i1 = 0; i1 < oParagraph.GetElementsCount(); ++i1) {
				var oElement = oParagraph.GetElement(i1)
				if (oElement.GetClassType() == 'run') {
					var fontfamily = oElement.GetFontFamily()
					if (fontfamily != 'iconfont') {
						continue
					}
					var run = oElement.Run
					var elCount2 = run.GetElementsCount()
					for (var i2 = 0; i2 < elCount2; ++i2) {
						var oElement2 = run.GetElement(i2)
						if (WORDS.includes(oElement2.Value)) {
							oElement.GetRange(i2, i2 + 1).Select()
							client_node_id += 1
							var tag = JSON.stringify({ regionType: 'write', mode: 3, client_id: client_node_id, color: '#ff000040' })
							var oResult = Api.asc_AddContentControl(2, { Tag: tag })
							if (oResult) {
								result.change_list.push({
									client_id: client_node_id,
									control_id: oResult.InternalId,
									parent_id: parentId,
									ques_id: quesId,
									regionType: 'write'
								})
							}
							Api.asc_RemoveSelection();
						}
					}
				}
			}
		}
		function addWordAsk(oControl, client_id, qid) {
			var controlContent = oControl.GetContent()
			var elementCount = controlContent.GetElementsCount()
			for (var i = 0; i < elementCount; ++i) {
				var oElement1 = controlContent.GetElement(i)
				if (oElement1.GetClassType() == 'paragraph') {
					handleParagraph(oElement1, client_id, qid)
				}
			}
		}
		function splitOneNode(nodeId) {
			var nodeData = node_list.find(e => {
				return e.id == nodeId
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
						var tag = Api.ParseJSON(e.GetTag())
						return tag.client_id == nodeId && e.GetClassType() == 'blockLvlSdt' && e.GetPosInParent() >= 0
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
			var obj = Api.ParseJSON(control.GetTag())
			if (obj.regionType == 'question') {
				var inlineSdts = control.GetAllContentControls().filter((e) => {
					return Api.ParseJSON(e.GetTag()).regionType == 'write'
				})
				if (inlineSdts.length > 0) {
					console.log('已有inline sdt， 删除以后再执行', inlineSdts)
					inlineSdts.forEach((e) => {
						var tag = Api.ParseJSON(e.GetTag())
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
	
				let mergeRanges2 = function(ranges) {
					var newRanges = []
					for (var i = 0; i < ranges.length; ++i) {
						if (i == 0) {
							newRanges.push(ranges[i])
						} else {
							var lastRange = newRanges[newRanges.length - 1]
							var canMerge = false
							if (ranges[i].Element == lastRange.Element && ranges[i].Start == lastRange.End) {
								var idx = ranges[i].Text.indexOf('\r')
								if (idx == 0 && ranges[i].Text[idx + 1] == lastRange.Text[lastRange.Text.length - 1]) {
									var Start = lastRange.Start
									var End = ranges[i].End
									var nrange = lastRange.ExpandTo(ranges[i])
									nrange.Start = Start
									nrange.End = End
									newRanges[newRanges.length - 1] = nrange
									canMerge = true
								}
							}
							if (!canMerge) {
								newRanges.push(ranges[i])
							}
						}
					}
					return newRanges
				}
				//debugger;
				var apiRanges = [];
				textSet.forEach(e => {
					var ranges = control.Search(e, false);
					//debugger;;
					apiRanges = mergeRange(apiRanges, ranges);
				});
				if (apiRanges.length > 1) {
					apiRanges = mergeRanges2(apiRanges)
				}
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
								ques_id: qid,
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
	
				// if (j < elements - 1) {
				// 	var range = content.GetElement(j + 1).GetRange();
				// 	var endRange = content.GetElement(elements - 1).GetRange();
				// 	range = range.ExpandTo(endRange);
				// 	range.Select();
					// client_node_id += 1
					// var tag = JSON.stringify({ regionType: 'write', mode: 5, client_id: client_node_id, color: '#ff000040' })
					// var oResult = Api.asc_AddContentControl(1, { Tag: tag })
					// if (oResult) {
					// 	result.change_list.push({
					// 		client_id: client_node_id,
					// 		control_id: oResult.InternalId,
					// 		parent_id: getParentId(Api.LookupObject(oResult.InternalId)),
					// 		regionType: 'write'
					// 	})
					// }
				// 	Api.asc_RemoveSelection();
				// }
			}
			if (obj.mid) {
				var oCell = control.GetParentTableCell()
				if (oCell) {
					var cellContent = oCell.GetContent()
					if (cellContent.GetElementsCount() == 1) {
						var text = control.GetRange().Text
						if (text && text.replace(/[\s\r\n]/g, '').length === 0) {
							var oTable = oCell.GetParentTable()
							result.change_list.push({
								parent_id: obj.client_id,
								table_id: oTable.Table.Id,
								row_index: oCell.GetRowIndex(),
								cell_index: oCell.GetIndex(),
								cell_id: oCell.Cell.Id,
								client_id: 'c_' + oCell.Cell.Id,
								regionType: 'write',
								type: ''
							})
							var desc = Api.ParseJSON(oTable.GetTableDescription())
							desc[`${oCell.GetRowIndex()}_${oCell.GetIndex()}`] = `c_${oCell.Cell.Id}`
							desc.biyue = 1
							oTable.SetTableDescription(JSON.stringify(desc))
							oCell.SetBackgroundColor(255, 191, 191, false)
							obj.cask = 1
							control.SetTag(JSON.stringify(obj))
						}
					}
				}
			}
			addWordAsk(control, obj.client_id, obj.client_id || obj.mid)
		}
		var quesData = Asc.scope.question_map[qid]
		if (!quesData) {
			return
		}
		if (quesData.is_merge && quesData.ids) { // 是合并题
			quesData.ids.forEach(id => {
				splitOneNode(id)
			})
		} else {
			splitOneNode(qid)
		}
		result.client_node_id = client_node_id
		result.ques_id = qid
		return result
	}, false, true).then(res1 => {
		if (res1) {
			if (res1.message && res1.message != '') {
				alert(res1.message)
			} else {
				if (!res1.change_list || res1.change_list.length == 0) {
					if (res1.ques_id) {
						return notifyQuestionChange(res1.ques_id)
					} else {
						return new Promise((resolve, reject) => {
							resolve()
						})
					}
				} else {
					return getNodeList().then(res2 => {
						return handleChangeType(res1, res2)
					})
				}
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
				var tag = Api.ParseJSON(e.GetTag())
				if (tag.client_id) {
					addItem(e.GetClassType(), e.Sdt.GetId(), tag)
				}
			}
		})
		drawings.forEach(e => {
			var tag = Api.ParseJSON(e.GetTitle())
			if (tag.client_id) {
				if (idmap[tag.client_id]) {
					client_node_id += client_node_id
					tag.client_id = client_node_id
					addItem('drawing', e.Drawing.Id, tag)
				}
				e.SetTitle(JSON.stringify(tag))
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
			var tag = Api.ParseJSON(drawing.GetTitle())
			if (tag.feature && tag.feature.zone_type == 'question' && tag.feature.sub_type == 'ask_accurate') {
				deleShape(drawing)
			}
		}
		function deleteCellAsk(oTable, write_list, nodeIndex) {
			var rowcount = oTable.GetRowsCount()
			var desc = Api.ParseJSON(oTable.GetTableDescription())
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
				var childTag = Api.ParseJSON(e.GetTag())
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
				var paraDrawing = childDrawings[j].getParaDrawing()
				if (!paraDrawing) {
					continue
				}
				var tag = Api.ParseJSON(childDrawings[j].GetTitle())
				if (tag.feature && tag.feature.zone_type == 'question') {
					if (tag.feature.sub_type == 'write') {						
						var run = paraDrawing.GetRun()
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
			var tagInfo = Api.ParseJSON(oControl.GetTag())
			if (!tagInfo.client_id) { // 无client_id，直接删除
				var del = true
				if (tagInfo.regionType == 'num') {
					var parentControl = oControl.GetParentContentControl()
					if (parentControl && parentControl.GetClassType() == 'blockLvlSdt') {
						var ptag = Api.ParseJSON(parentControl.GetTag())
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
					var tagInfo = Api.ParseJSON(tableParentControl.GetTag())
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
		var oState = oDocument.Document.SaveDocumentState()
		function updateFill(oDrawing, oFill) {
			if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== 'fill') {
				return false
			}
			oDrawing.Drawing.spPr.setFill(oFill.UniFill)
		}
		drawings.forEach(oDrawing => {
			let title = oDrawing.GetTitle()
			var titleObj = Api.ParseJSON(title)
			// 图片不铺码的标识
			if (title.includes('partical_no_dot')) {
				if (cmdType == 'show') {
					oDrawing.SetShadow(null, 0, 100, null, 0, '#0fc1fd')
				} else {
					oDrawing.ClearShadow()
				}
			} else if (title.includes('ques_use')) {
				oDrawing.ClearShadow()
			} else if (titleObj.feature) {
				if (titleObj.feature.zone_type == 'pagination') { // 试卷页码
					var oShape = Api.LookupObject(oDrawing.Drawing.Id)
					if (oShape && oShape.GetClassType() == 'shape') {
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
						updateFill(oDrawing, oFill)
					} else if (cmdType == 'hide') {
						updateFill(oDrawing, Api.CreateNoFill())
					}
				}
			}
		})
		// 处理单元格小问
		function getCell(write_data) {
			for (var i = 0; i < oTables.length; ++i) {
				var oTable = oTables[i]
				if (oTable.GetPosInParent() == -1) { continue }
				var desc = Api.ParseJSON(oTable.GetTableDescription())
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
				var ids = []
				if (question_map[id].is_merge) {
					ids = question_map[id].ids || []
				} else {
					ids.push(id)
				}
				ids.forEach(nid => {
					var nodeData = node_list.find(item => {
						return item.id == nid
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
				})
			}
		})
		oDocument.Document.LoadDocumentState(oState)
	}, false, true)
}

function importExam() {
	if (isLoading('importExam')) { 
		return
	}
	setBtnLoading('importExam', true)
	if (window.BiyueCustomData.page_type == 1) {
		return handleUploadPrepare('hide')
		.then(() => {
			return getAllPositions2()
		}).then(res => {
			Asc.scope.questionPositions = res
			window.biyue.showDialog('exportExamWindow', '上传试卷', 'examExport.html', 1000, 800, true)
			setBtnLoading('importExam', false)
		})
	}
	setBtnLoading('uploadTree', true)
	return setInteraction('none', null, false).then(() => {
		return preGetExamTree() // 获取题目树需要在addOnlyBigControl之前执行，否则可能出现父节点出错的情况
	}).then((res) => {
		Asc.scope.tree_info = res
		return addOnlyBigControl(false)
	}).then(() => {
		return handleUploadPrepare('hide')
	}).then(() => {
		return getChoiceOptionAndSteam() // getChoiceQuesData()
	}).then((res) => {
		Asc.scope.choice_html_map = res
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
function clearMergeAsk(options) {
	console.log('clearMergeAsk', options)
	if (!options.length || options.length != 3) {
		return
	}
	var question_map = window.BiyueCustomData.question_map || {}
	var quesData = question_map[options[1]]
	if (!quesData || !quesData.ask_list) {
		return
	}
	var targetItem = quesData.ask_list.find(e => {
		return e.id == options[2] || (e.other_fields && e.other_fields.findIndex(e => {
			return e == options[2]
		}) >= 0)
	})
	if (!targetItem) {
		return
	}
	var node_list = window.BiyueCustomData.node_list || []
	var nodeData = node_list.find(item => {
		return item.id == options[1]
	})
	var newFields = targetItem.other_fields.map(id => ({
		id: id,
		score: 1
	}))
	delete targetItem.other_fields;
	var ask_list = quesData.ask_list.concat(newFields);
	ask_list = nodeData.write_list.map(writeItem =>
        ask_list.find(askItem => askItem.id === writeItem.id)
    )
    .filter(item => item !== undefined);  // 过滤掉不存在的元素
	quesData.ask_list = ask_list
	updateScore(options[1])
	// 集中作答区暂时先不考虑
	return new Promise((resolve, reject) => {
		if (quesData.ques_mode == 1 || quesData.ques_mode == 5) {
			return deleteChoiceOtherWrite([options[1]], false).then(() => {
				resolve()
			})
		} else {
			resolve()
		}
	}).then(() => {
		if (quesData.interaction == 'accurate') {
			return setInteraction(quesData.interaction, [options[1]])
		} else {
			return new Promise((resolve, reject) => {
				resolve()
			})
		}
	}).then(() => {
		notifyQuestionChange(options[1])
		window.biyue.StoreCustomData()
	})
}
// 合并小问纯粹只是逻辑上的改动，和OO无关
function mergeOneAsk(options) {
	if (!options.ask_ids || options.ask_ids.length == 0) {
		return
	}
	var question_map = window.BiyueCustomData.question_map || {}
	var quesData = question_map[options.ques_id]
	if (!quesData || !quesData.ask_list) {
		return
	}
	if (options.cmd == 'in') { // 合并
		var other_fields = []
		for (var i = 1; i < options.ask_ids.length; ++i) {
			for (var j = 0; j < quesData.ask_list.length; ++j) {
				if (quesData.ask_list[j].id == options.ask_ids[i]) {
					quesData.ask_list.splice(j, 1)	
					break
				} else if (quesData.ask_list[j].other_fields) {
					var k = quesData.ask_list[j].other_fields.findIndex(e => {
						return e == options.ask_ids[i]
					})
					if (k >= 0) {
						quesData.ask_list[j].other_fields.splice(k, 1)
						break
					}
				}
			}
			other_fields.push(options.ask_ids[i])
		}
		var firstIndex = quesData.ask_list.findIndex(e => {
			return e.id == options.ask_ids[0]
		})
		if (firstIndex != -1) {
			quesData.ask_list[firstIndex].other_fields = other_fields
		}
		updateScore(options.ques_id)
	}
}
// 合并小问
function mergeAsk(options) {
	// 合并小问纯粹只是逻辑上的改动，和OO无关
	var node_list = window.BiyueCustomData.node_list || []
	var nodeData = node_list.find(item => {
		return item.id == options[1]
	})
	if (!nodeData) {
		return
	}
	var write_id = options[2] * 1
	var writeIndex = nodeData.write_list.findIndex(e => {
		return e.id == write_id
	})
	if (writeIndex == -1) {
		return
	}
	var question_map = window.BiyueCustomData.question_map || {}
	var quesData = question_map[options[1]]
	if (!quesData || !quesData.ask_list) {
		return
	}
	if (options[3] == 1) { // 向前合并
		var askIndex = quesData.ask_list.findIndex(e => {
			return e.id == write_id
		})
		if (askIndex > 0) {
			quesData.ask_list.splice(askIndex, 1)
			var preAsk = quesData.ask_list[askIndex - 1]
			if (!preAsk.other_fields) {
				preAsk.other_fields = []
			}
			preAsk.other_fields.push(write_id)
		}
	} else { // 取消合并
		for (var i = 0; i < quesData.ask_list.length; ++i) {
			if (quesData.ask_list[i].other_fields) {
				var fidx = quesData.ask_list[i].other_fields.findIndex(e => {
					return e == write_id
				})
				if (fidx >= 0) {
					quesData.ask_list[i].other_fields.splice(fidx, 1)
					quesData.ask_list.splice(i + 1, 0, {
						id: write_id,
						score: 1
					})
					break
				}
			}
		}
	}
	var sumScore = 0
	quesData.ask_list.forEach(e => {
		sumScore += e.score
	})
	quesData.score = sumScore
	document.dispatchEvent(
		new CustomEvent('clickSingleQues', {
			detail: {
				client_id: write_id,
				regionType: "write"
			}
		})
	)
	setInteraction('useself', [options[1]], true)
}
// 插入符号
function insertSymbol(unicode) {
	Asc.scope.symbol = unicode
	Asc.scope.client_node_id = window.BiyueCustomData.client_node_id
	Asc.scope.question_map = window.BiyueCustomData.question_map
	var isCalc = false
	var validUnicodes = ['e753', 'e754', 'e755', 'e756', 'e757', 'e758']
	if (!validUnicodes.indexOf(unicode)) {
		isCalc = true
	}
	return biyueCallCommand(window, function() {
		var unicode = Asc.scope.symbol
		var client_node_id = Asc.scope.client_node_id
		var question_map = Asc.scope.question_map
		var validUnicodes = ['e753', 'e754', 'e755', 'e756', 'e757', 'e758']
		var result = {
			typeName: 'write',
			change_list: [],
			client_node_id: client_node_id
		}
		var oDocument = Api.GetDocument()
		function getParentControlData(oRun) {
			var oParentControl = oRun.GetParentContentControl()
			var ques_id = 0
			var parent_id = 0
			if (oParentControl && oParentControl.GetClassType() == 'blockLvlSdt') {
				var oEle = Api.LookupObject(oParentControl.Sdt.Id)
				if (oEle) {
					if (oEle.GetClassType() == 'documentContent') {
						var parentId = oEle.Document.Parent.Id
						oParentControl = Api.LookupObject(parentId)
					}
				}
			}
			if (oParentControl && oParentControl.GetClassType() == 'blockLvlSdt') {
				var strTag = oParentControl.GetTag()
				var tag = Api.ParseJSON(strTag)
				var qid = tag.mid ? tag.mid : tag.client_id
				if (question_map[qid] && question_map[qid].level_type == 'question') {
					ques_id = qid
				}
				parent_id = tag.client_id	
			}
			return { parent_id, ques_id }
		}
		function addAsk(oRun, parent_id, ques_id, moveRight, from, end) {
			if (!(validUnicodes.includes(unicode)) || !parent_id) {
				if (moveRight) {
					oDocument.Document.MoveCursorRight()
				}
				return
			}
			var oRange = from && end ? oRun.GetRange(from, end) : oRun.GetRange()
			oRange.Select()
			result.client_node_id += 1
			var asktag = JSON.stringify({ regionType: 'write', mode: 3, client_id: result.client_node_id, color: '#ff000040' })
			var oResult = Api.asc_AddContentControl(2, { Tag: asktag })
			if (oResult) {
				result.change_list.push({
					client_id: result.client_node_id,
					control_id: oResult.InternalId,
					parent_id: parent_id,
					ques_id: ques_id,
					regionType: 'write'
				})
			}
			Api.asc_RemoveSelection();
			var oWriteControl = Api.LookupObject(oResult.InternalId)
			oWriteControl.Sdt.MoveCursorToContentControl(false)
			oDocument.Document.MoveCursorRight()
		}
		var pos = oDocument.Document.Get_CursorLogicPosition()
		if (pos) {
			var lastElement = pos[pos.length - 1].Class
			var oRun = Api.LookupObject(lastElement.Id)
			if (oRun.GetClassType() == 'run') {
				var lastPos = pos[pos.length - 1].Position
				var runParent = pos[pos.length - 2].Class
				var pos2 = pos[pos.length - 2].Position
				var fontfamily = oRun.GetFontFamily()
				const unicodeDecimal = parseInt(unicode, 16);
				const unicodeChar = String.fromCharCode(unicodeDecimal);
				const { parent_id, ques_id } = getParentControlData(oRun)
				if (fontfamily == 'iconfont') {
					oRun.Run.AddText(unicodeChar, lastPos + 1)
					addAsk(oRun, parent_id, ques_id, false, lastPos, lastPos + 1)
				} else {
					if (lastPos == 0 && pos2 > 0) {
						// 判断前一个run的字体
						var oParent = Api.LookupObject(runParent.Id)
						if (oParent.GetClassType() == 'paragraph') {
							var oRun3 = oParent.GetElement(pos2 - 1)
							if (oRun3 && oRun3.GetClassType() == 'run' && oRun3.GetFontFamily() == 'iconfont') {
								oRun3.AddText(unicodeChar)
								addAsk(oRun3, parent_id, ques_id, true)
								return
							}
						}
					}
					var newRun2 = Api.CreateRun()
					newRun2.SetFontFamily('iconfont')
					newRun2.AddText(unicodeChar)
					newRun2.SetFontSize(36 * 2)
					if (lastPos == 0) {
						pos[pos.length - 2].Class.Add_ToContent(pos[pos.length - 2].Position, newRun2.Run)
					} else {
						var newRun = oRun.Run.Split_Run(lastPos)
						pos[pos.length - 2].Class.Add_ToContent(pos2 + 1, newRun2.Run)
						pos[pos.length - 2].Class.Add_ToContent(pos2 + 2, newRun)
					}
					addAsk(newRun2, parent_id, ques_id, true)
				}
			}
		}
		return result
	}, false, isCalc).then(res1 => {
		if (res1 && res1.change_list.length) {
			return getNodeList().then(res2 => {
				return handleChangeType(res1, res2)
			})
		} else {
			return new Promise((resolve, reject) => {
				resolve()
			})
		}
	})
}

function preGetExamTree() {
	Asc.scope.node_list = window.BiyueCustomData.node_list
	Asc.scope.question_map = window.BiyueCustomData.question_map
	return biyueCallCommand(window, function() {
		var node_list = Asc.scope.node_list || []
		var question_map = Asc.scope.question_map || {}
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls() || []
		function getFirstParagraph(oControl) {
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
				return null
			}
			var paragraphs = oControl.GetAllParagraphs()
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
		function getLvl(oControl, paraIndex) {
			var oParagraph = getFirstParagraph(oControl)
			if (!oParagraph) {
				return null
			}
			var oNumberingLevel = oParagraph.GetNumbering()
			if (oNumberingLevel) {
				return oNumberingLevel.Lvl
			}
			return null
		}
		function getValidParent(oControl) {
			if (!oControl) {
				return null
			}
			var oParentControl = oControl.GetParentContentControl()
			if (oParentControl) {
				var tag = Api.ParseJSON(oParentControl.GetTag())
				var qId = tag.mid ? tag.mid : tag.client_id
				if (question_map[qId]) {
					return oParentControl
				} else {
					return getValidParent(oParentControl)
				}
			}
			return null
		}
		var list = []
		var handled = {}
		for (var oControl of controls) {
			var tag = Api.ParseJSON(oControl.GetTag())
			if (!tag.client_id) {
				continue
			}
			var qId = tag.mid ? tag.mid : tag.client_id
			if (handled[qId]) {
				continue
			}
			var quesData = question_map[qId]
			if (!quesData) {
				continue
			}
			if (quesData.level_type != 'struct' && quesData.level_type != 'question') {
				continue
			}
			handled[qId] = true
			var nodeData = node_list.find(e => {
				return e.id == tag.client_id
			})
			var is_big = nodeData ? nodeData.is_big : false
			var lvl = null
			var obj = {
				id: qId,
				level_type: quesData.level_type,
				parent_id: 0,
				parent_index: -1,
				is_big: is_big,
			}
			var oParentControl = getValidParent(oControl)
			if (quesData.level_type == 'struct') {
				lvl = getLvl(oControl)
			} else if (quesData.level_type == 'question') {
				lvl = getLvl(oControl, is_big ? 0 : -1)
			}
			obj.lvl = lvl
			if (oParentControl && quesData.level_type == 'question') {
				var parentTag = Api.ParseJSON(oParentControl.GetTag() || '{}')
				var p_id = parentTag.mid ? parentTag.mid : parentTag.client_id
				obj.parent_id = p_id
				obj.parent_index = list.findIndex(e => {
					return e.id == p_id
				})
				// console.log(qId, '1   p_id', obj.parent_id, obj.parent_index)
			} else if (lvl === 0) {
				obj.parent_id = 0
				obj.parent_index = -1
				// console.log(qId, '2   p_id', obj.parent_id, obj.parent_index)
			} else {
				// 根据level, 查找在它前面的比它lvl小的struct
				if (list.length > 0) {
					for (var i = list.length - 1; i >= 0; --i) {
						if (list[i].lvl === null) {
							if (list[i].level_type == 'struct') {
								if (lvl === null) {
									obj.parent_id = list[i].parent_id
									obj.parent_index = list[i].parent_index
								} else {
									obj.parent_id = list[i].id
									obj.parent_index = i
								}
								// console.log(qId, '3   p_id', obj.parent_id, obj.parent_index)
								break
							} else if (list[i].is_child) {
								continue
							} else {
								obj.parent_id = list[i].parent_id
								obj.parent_index = list[i].parent_index
								// console.log(qId, '4   p_id', obj.parent_id, obj.parent_index)
								break
							}
						} else if (list[i].lvl === 0) {
							if (list[i].level_type == 'struct') {
								obj.parent_id = list[i].id
								obj.parent_index = i
								// console.log(qId, '5   p_id', obj.parent_id, obj.parent_index)
								break
							} else {
								obj.parent_id = 0
								obj.parent_index = -1
								// console.log(qId, '6   p_id', obj.parent_id, obj.parent_index)
								break
							}
						} else if (list[i].lvl < lvl) {
							if (list[i].level_type == 'struct') {
								obj.parent_id = list[i].id
								obj.parent_index = i
								// console.log(qId, '7   p_id', obj.parent_id, obj.parent_index)
								break
							} else {
								if (list[i].is_child) {
									continue
								} else {
									obj.parent_id = list[i].parent_id
									obj.parent_index = list[i].parent_index
									// console.log(qId, '8   p_id', obj.parent_id, obj.parent_index)
								}
								break
							}
						} else if (list[i].lvl === lvl) {
							if (list[i].level_type == 'struct') {
								if (list[i].parent_id || quesData.level_type == 'struct') {
									obj.parent_id = list[i].parent_id
									obj.parent_index = i
									// console.log(qId, '9   p_id', obj.parent_id, obj.parent_index)
								} else {
									obj.parent_id = list[i].id
									obj.parent_index = i
									// console.log(qId, '10   p_id', obj.parent_id, obj.parent_index)
								}
								break
							} else if (list[i].level_type == 'question') {
								if (quesData.level_type == 'struct') {
									continue
								} else if (list[i].is_child) {
									continue
								} else {
									obj.parent_id = list[i].parent_id
									obj.parent_index = list[i].parent_index
									// console.log(qId, '11   p_id', obj.parent_id, obj.parent_index)
									break
								}
							}
						} else if (list[i].lvl > lvl) {
							if (list[i].level_type == 'struct' && list[i].parent_id == 0 && quesData.level_type == 'question' && lvl > 0) {
								obj.parent_id = list[i].id
								obj.parent_index = i
								break
							}
							continue
						}
					}
				}
			}
			var parentTableCell1 = oControl.GetParentTableCell()
			if (parentTableCell1) {
				obj.cell_id = parentTableCell1.Cell.Id
			}
			list.push(obj)
			if (is_big) {
				var bindex = list.length - 1
				var childControls = oControl.GetAllContentControls()
				for (var oChildControl of childControls) {
					var childTag = Api.ParseJSON(oChildControl.GetTag() || '{}')
					var childId = childTag.mid || childTag.client_id
					if (handled[childId] || oChildControl.GetClassType() != 'blockLvlSdt') {
						continue
					}
					var quesData2 = question_map[childId]
					if (!quesData2) {
						continue
					}
					if (quesData2.level_type != 'struct' && quesData2.level_type != 'question') {
						continue
					}
					handled[childId] = true
					var parentControl2 = getValidParent(oChildControl)
					if (parentControl2) {
						var parentTag2 = Api.ParseJSON(parentControl2.GetTag() || '{}')
						var parentId2 = parentTag2.mid || parentTag2.client_id
						var parentIndex2 = list.findIndex(e => {
							return e.id == parentId2
						})
						var obj2 = {
							id: childId,
							level_type: quesData2.level_type,
							parent_id: parentId2,
							parent_index: parentIndex2,
							is_big: childTag.big == 1,
							lvl: getLvl(oChildControl, childTag.big == 1 ? 0 : -1),
							is_child: true
						}
						// console.log(childId, '12   p_id', parentId2, parentIndex2)
						var parentTableCell = oChildControl.GetParentTableCell()
						if (parentTableCell) {
							obj2.cell_id = parentTableCell.Cell.Id
						}
						list.push(obj2)
						list[bindex].end_id = childId
					}
				}
			}
		}
		return list
	}, false, false).then((list => {
		// 传入OO处理的js代码的列表结构不支持层级过深，嵌套达到5级，就会导致树形结构出错，command无法返回结果
		return new Promise((resolve, reject) => {
			if (!list) {
				return resolve({})
			}
			const tree = [];
			// 使用 Map 对象，以 id 作为 key，这样能更快地找到任何一个节点
			let map = new Map();
			list.forEach(item => {
				map.set(item.id, { ...item, children: [] });
			});
	
			list.forEach(item => {
				if (item.parent_id && item.parent_id != item.id) {
					if (map.has(item.parent_id)) {
						let parent = map.get(item.parent_id);
						if (parent.parent_id != item.id) {
							parent.children.push(map.get(item.id));
						} else {
							console.error('Circular reference detected', item, parent);
						}
					}
				} else {
					tree.push(map.get(item.id));
				}
			});
			Asc.scope.tree_info = {list: list, tree: tree}
			resolve({list: list, tree: tree})
		})
	}))
}

function setNumberingLevel(ids, lvl) {
	Asc.scope.ids = ids
	Asc.scope.lvl = lvl
	return biyueCallCommand(window, function() {
		var ids = Asc.scope.ids || []
		var lvl = Asc.scope.lvl
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		var list = []
		function getFirstParagraph(oControl) {
			if (!oControl || oControl.GetClassType() != 'blockLvlSdt') {
				return null
			}
			var paragraphs = oControl.GetAllParagraphs()
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
		for (var oControl of controls) {
			if (oControl.GetClassType() != 'blockLvlSdt') {
				continue
			}
			var tag = Api.ParseJSON(oControl.GetTag() || '{}')
			var id = tag.mid || tag.client_id
			if (ids.indexOf(id) == -1) {
				continue
			}
			tag.lvl = lvl
			oControl.SetTag(JSON.stringify(tag))
			var numberingtext = ''
			var oParagraph = getFirstParagraph(oControl)
			if (oParagraph) {
				var oNumberingLevel = oParagraph.GetNumbering()
				if (oNumberingLevel) { // ApiNumberingLevel
					var oNumbering = oNumberingLevel.GetNumbering()
					var oNumLvl = oNumbering.GetLevel(lvl)
					oParagraph.SetNumbering(oNumLvl)
					numberingtext = oParagraph.Paragraph.GetNumberingText()
				} else {
					var oNumbering = Api.GetDocument().CreateNumbering("numbered")  // ApiNumbering
					for (var i = 0; i < 10; ++i) {
						var oNumLvl = oNumbering.GetLevel(i)
						oNumLvl.SetCustomType("none", '', "left");
						oNumLvl.SetRestart(false);
						oNumLvl.SetSuff("none")
						var oParaPr = oNumLvl.GetParaPr()
						oParaPr.SetIndFirstLine(0);
						oParaPr.SetIndLeft(0)
						if (lvl == i) {
							oParagraph.SetNumbering(oNumLvl)		
						}
					}
				}
			}
			list.push({
				id: id,
				numbing_text: numberingtext,
				text: oControl.GetRange().GetText()
			})
		}
		return list
	}, false, false).then(list => {
		return new Promise((resolve, reject) => {
			if (list) {
				var question_map = window.BiyueCustomData.question_map || {}
				for (var item of list) {
					var question = question_map[item.id]
					if (question) {
						question.text = item.text
						question.ques_default_name = item.numbing_text ? getNumberingText(item.numbing_text) : GetDefaultName(question.level_type, question.text)
					}
				}
			}
			resolve()
		})
	})
}
// 将单个字设置为小问
function splitWordAsk() {
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls() || []
		const WORDS = [0xe753, 0xe754, 0xe755, 0xe756, 0xe757, 0xe758]
		function handleParagraph(oParagraph) {
			if (!oParagraph) {
				return
			}
			var elCount = oParagraph.GetElementsCount()
			for (var i1 = 0; i1 < elCount; ++i1) {
				var oElement = oParagraph.GetElement(i1)
				if (oElement.GetClassType() == 'run') {
					var fontfamily = oElement.GetFontFamily()
					if (fontfamily != 'iconfont') {
						continue
					}
					var run = oElement.Run
					var elCount2 = run.GetElementsCount()
					for (var i2 = 0; i2 < elCount2; ++i2) {
						var oElement2 = run.GetElement(i2)
						if (WORDS.includes(oElement2.Value)) {
							oElement.GetRange(i2, i2 + 1).Select()
							var tag = JSON.stringify({ regionType: 'write', mode: 3, color: '#ff000040' })
							Api.asc_AddContentControl(2, { "Tag": tag });
							Api.asc_RemoveSelection();
						}
					}
				}
			}
		}
		for (var oControl of controls) {
			if (oControl.GetClassType() != 'blockLvlSdt') {
				continue
			}
			var tag = Api.ParseJSON(oControl.GetTag())
			if (tag.regionType != 'question') {
				continue
			}
			var controlContent = oControl.GetContent()
			var elementCount = controlContent.GetElementsCount()
			for (var i = 0; i < elementCount; ++i) {
				var oElement1 = controlContent.GetElement(i)
				if (oElement1.GetClassType() == 'paragraph') {
					handleParagraph(oElement1)
				}
			}
		}
	})
}

export {
	handleDocClick,
	handleContextMenuShow,
	reqGetQuestionType,
	reqUploadTree,
	splitEnd,
	confirmLevelSet,
	initControls,
	handleAllWrite,
	deleteAsks,
	focusAsk,
	showAskCells,
	onContextMenuClick,
	g_click_value,
	updateAllChoice,
	deleteChoiceOtherWrite,
	getQuesMode,
	updateQuesScore,
	splitControl,
	updateDataBySavedData,
	clearRepeatControl,
	tidyNodes,
	handleUploadPrepare,
	importExam,
	getNodeList,
	handleChangeType,
	insertSymbol,
	preGetExamTree,
	getQuestionHtml,
	focusControl,
	setNumberingLevel,
	splitWordAsk
}