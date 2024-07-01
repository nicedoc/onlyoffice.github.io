// 主要用于管理试卷结构树
import Tree from '../components/Tree.js'
import { getQuesType, reqComplete } from '../scripts/api/paper.js'
import { biyueCallCommand } from './command.js'

var g_exam_tree = null
var g_horizontal_list = []
var upload_control_list = []
function initExamTree() {
	return new Promise((resolve, reject) => {
		console.log('[initExamTree]')
		getDocList().then(res => {
			var hlist = generateListByDoc(res)
			g_horizontal_list = hlist
			console.log(' after generateListByDoc', [].concat(hlist))
			var treeInfo = genetateTreeByHList(hlist)
			console.log('******************   hlist', hlist)
			console.log('******************   treeInfo', treeInfo)
			renderTree(treeInfo)
			resolve()
		})
	})
}
// 获取文档列表
function getDocList() {
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var elementCount = oDocument.GetElementsCount()
		if (elementCount == 0) {
			return []
		}
		var allControls = oDocument.GetAllContentControls()
		var list = []
		for (var idx = 0; idx < elementCount; ++idx) {
			var oElement = oDocument.GetElement(idx)
			if (!oElement) {
				continue
			}
			var classType = oElement.GetClassType()
			if (classType == 'paragraph') {
				list.push({
					id: oElement.Paragraph.Id,
					regionType: null,
					text: oElement.GetRange().GetText(),
					classType: classType,
				})
			} else if (classType == 'blockLvlSdt') {
				var tag = JSON.parse(oElement.GetTag() || {})
				list.push({
					id: oElement.Sdt.GetId(),
					regionType: tag.regionType,
					text: oElement.GetRange().GetText(),
					classType: classType,
				})
				var controls = oElement.GetAllContentControls()
				if (controls) {
					controls.forEach((e) => {
						var subtag = JSON.parse(e.GetTag() || {})
						if (
							subtag.regionType == 'question' ||
							subtag.regionType == 'sub-question'
						) {
							list.push({
								id: e.Sdt.GetId(),
								regionType: subtag.regionType,
								text: e.GetRange().GetText(),
								classType: e.GetClassType(),
							})
						}
					})
				}
			} else if (classType == 'table') {
				var controls = []
				oElement.Table.GetAllContentControls(controls)
				if (controls && controls.length) {
					controls.forEach((e) => {
						var control = allControls.find(c => {
							return c.Sdt.GetId() == e.Id
						})
						if (control) {
							var subtag = JSON.parse(control.GetTag() || {})
							if (
								subtag.regionType == 'question' ||
								subtag.regionType == 'sub-question'
							) {
								list.push({
									id: control.Sdt.GetId(),
									regionType: subtag.regionType,
									text: control.GetRange().GetText(),
									classType: control.GetClassType(),
								})
							}
						}
					})
				}
			}
		}
		return list
	}, false, false)
}
// 根据文档列表生成list
function generateListByDoc(docList) {
	var hlist = []
	var struct_index
	var question_index
	docList.forEach((e, index) => {
		var parent_id = 0
		if (e.regionType == 'struct') {
			parent_id = 0
		} else if (e.regionType == 'question') {
			parent_id = struct_index != undefined ? docList[struct_index].id : 0
		} else if (e.regionType == 'sub-question') {
			parent_id = question_index != undefined ? docList[question_index].id : (struct_index != undefined ? docList[struct_index].id : 0)
		} else if (!e.regionType) {
			if (index > 0) {
				if (hlist[index - 1].regionType) {
					parent_id = hlist[hlist.length - 1].id
				} else {
					parent_id = hlist[hlist.length - 1].parent_id
				}
			}
		}
		var child_pos = 0
		if (parent_id) {
			var parent = hlist.find(p => {
				return p.id == parent_id
			})
			if (parent) {
				child_pos = parent.children.length
				parent.children.push(e.id)
			}
		} else {
			for (var j = hlist.length - 1; j >= 0; --j) {
				if (hlist[j].parent_id == 0) {
					child_pos = hlist[j].child_pos + 1
					break
				}
			}
		}
		var obj = {
			id: e.id,
			regionType: e.regionType,
			text: e.text,
			classType: e.classType,
			parent_id: parent_id,
			children: [],
			child_pos: child_pos
		}
		hlist.push(obj)
		if (e.regionType == 'struct') {
			struct_index = index
		} else if (e.regionType == 'question') {
			question_index = index
		}
	})
	return hlist
}
// 根据list生成渲染需要的tree
function genetateTreeByHList(list) {
	if (!list) {
		return null
	}
	var tree = []
	list.forEach((e) => {
		if (e.parent_id == 0) {
			tree.push({
				id: e.id,
				label: e.text,
				expand: true,
				is_folder: e.regionType == 'struct',
				children: [],
				regionType: e.regionType,
				parent_id: e.parent_id,
				classType: e.classType
			})
		} else {
			var parent = getDataById(tree, e.parent_id)
			if (parent && parent.children) {
				parent.children.push({
					id: e.id,
					label: e.text,
					expand: true,
					is_folder: e.regionType == 'struct',
					children: [],
					regionType: e.regionType,
					parent_id: e.parent_id,
					classType: e.classType
				})
			}
		}
	})
	return tree
}

function renderTree(tree) {
	if (!g_exam_tree) {
		g_exam_tree = new Tree($('#treeRoot'))
		g_exam_tree.addCallBack(clickItem, dropItem)
	}
	g_exam_tree.refreshList(tree)
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

function clickItem(id, item, e) {
	if (g_exam_tree) {
		g_exam_tree.setSelect([id])
	}
	Asc.scope.click_id = id
	biyueCallCommand(window, function() {
		var click_id = Asc.scope.click_id
		var oDocument = Api.GetDocument()
		oDocument.RemoveSelection()
		var clickElement = Api.LookupObject(click_id)
		if (clickElement && clickElement.GetRange) {
			var oRange = clickElement.GetRange()
			if (oRange) {
				oRange.Select()
			}
		}
	}, false, false)
}

function dropItem(list, dragId, dropId, direction) {
	console.log('dropItem', dragId, dropId, direction)
	console.log('tree', list)
	window.BiyueCustomData.exam_tree = list
	Asc.scope.drag_options = {
		dragId: dragId,
		dropId: dropId,
		direction: direction,
		horList: g_horizontal_list
	}
	biyueCallCommand(window, function() {
		var drag_options = Asc.scope.drag_options
		console.log('drag_options', drag_options)
		var horList = drag_options.horList
		if(!horList){
			return
		}
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		var dragData = horList.find((e) => {
			return e.id == drag_options.dragId
		})
		var dropData = horList.find(e => {
			return e.id == drag_options.dropId
		})
		if (!dragData || !dropData) {
			return
		}
		var regionType = dragData.regionType
		function GetElementRange(id) {
			var oElement = Api.LookupObject(id)
			if (oElement) {
				if (oElement.GetRange) {
					return oElement.GetRange()
				}
			}
			return null
		}
		var oDragRange = GetElementRange(drag_options.dragId, dragData.classType)
		if (!oDragRange) {
			return
		}
		console.log('dragData', dragData)
		console.log('dropData', dropData)

		function getRemoveIds(list, id, parent_id) {
			if (!list) {
				return
			}
			var item = horList.find(e => {
				return e.id == id
			})
			if (item) {
				var needRemove = true
				if (parent_id) {
					var oElement = Api.LookupObject(id)
					if (oElement && oElement.GetParentContentControl) {
						var parentControl = oElement.GetParentContentControl()
						if (parentControl) {
							var pid = parentControl.Sdt.GetId()
							var index = list.findIndex(e => {
								return e.id == pid
							})
							if (index != -1) {
								needRemove = false
							}
						}
					}
				}
				if (needRemove) {
					list.push({
						index: list.length,
						id: item.id
					})
				}	
				if (item.children) {
					item.children.forEach(childId => {
						getRemoveIds(list, childId, id)
					})
				}
			}
		}

		var templist = []

		var removeIds = []
		getRemoveIds(removeIds, drag_options.dragId)
		removeIds = removeIds.sort((a, b) => {
			return b.index - a.index
		})
		console.log('removeIds', removeIds)
		// 题组或题目拖拽，需要将其下的所有children都挪动位置
		if (regionType == 'struct' || regionType == 'question') {
			removeIds.forEach(e => {
				var itemRange = GetElementRange(e.id)
				var startPos = itemRange.StartPos[0].Position
				templist.push(oDocument.GetElement(startPos))
				oDocument.RemoveElement(startPos)
			})
		}

		if (regionType == 'struct') {
			// 移动到相应位置
			var dropPos
			if (drag_options.direction == 'top') {
				// 拖拽到上方，插入位置为第一个
				var dropRange = GetElementRange(drag_options.dropId, dropData.classType)
				if (dropRange) {
					dropPos = dropRange.StartPos[0].Position
				}
			} else if (drag_options.direction == 'bottom') {
				// 拖拽到下方，插入位置为children的最后一个
				if (dropData.children && dropData.children.length) {
					var childId = dropData.children[dropData.children.length - 1]
					var childData = horList.find(e => {
						return e.id == childId
					})
					if (childData) {
						var childRange = GetElementRange(childId, childData.classType)
						if (childRange) {
							dropPos = childRange.EndPos[0].Position + 1
						}
					}
				}
			}
			if (dropPos != undefined) {
				templist.forEach((e) => {
					oDocument.AddElement(dropPos, e)
				})
			}
		} else {
			// 题目位于表格的，之后再处理
			if (!regionType || regionType == 'question') {
				console.log('oDragRange', oDragRange)
				var dropRange = GetElementRange(drag_options.dropId, dropData.classType)
				console.log('dropRange', dropRange)
				if (dropRange) {
					var dropPos = dropRange.StartPos[0].Position
					if (drag_options.direction != 'top') {
						dropPos = dropPos + 1
					}
					templist.forEach((e) => {
						oDocument.AddElement(dropPos, e)
					})
				} else {
					console.log('dropRange is null')
				}
			} else if (regionType == 'sub-question') {
				var dragControl = controls.find((e) => {
					return e.Sdt.GetId() == drag_options.dragId
				})
				var parentControl = dragControl.GetParentContentControl()
				console.log('parentControl', parentControl)
				if (parentControl) {
					var startPos = oDragRange.StartPos[1].Position
					var oParentDocument = parentControl.GetContent()
					templist.push(oParentDocument.GetElement(startPos))
					oParentDocument.RemoveElement(startPos)
					var oDropRange
					var oDropParentControl
					if (dropData.classType == 'blockLvlSdt') {
						var dropControl = Api.LookupObject(drag_options.dropId)
						if (dropControl) {
							oDropRange = dropControl.GetRange()
							oDropParentControl = dropControl.GetParentContentControl()
						}
					} else if (dropData.classType == 'paragraph') {
						var oDropParagraph = Api.LookupObject(drag_options.dropId)
						if (oDropParagraph) {
							oDropRange = oDropParagraph.GetRange()
							oDropParentControl = oDropParagraph.GetParentContentControl()
						}
					}
					if (!oDropRange) {
						return
					}
					var dropPos = oDropRange.StartPos[0].Position
					if (oDropParentControl) {
						dropPos = oDropRange.StartPos[1].Position
						if (drag_options.direction != 'top') {
							dropPos += 1
						}
						oDropParentControl.GetContent().AddElement(dropPos, templist[0])
					} else {
						console.log('not oDropParentControl')
						if (drag_options.direction != 'top') {
							dropPos += 1
						}
						oDocument.AddElement(dropPos, templist[0])
					}
				} else {
					var dropElement = Api.LookupObject(drag_options.dropId)
					console.log('dropElement', dropElement)
					if (dropElement) {
						if (dropElement.GetParentContentControl) {
							var dropParentControl = dropElement.GetParentContentControl()
							console.log('dropParentControl', dropParentControl)
						}
					}
					var parentTableCell = dragControl.GetParentTableCell()
					if (parentTableCell) {
						// todo..单题在表格单元里，可能需要合并拆分单元格

					}
				}
			}
		}
	}, false, false).then(res => {
		var hlist = []
		updateHorListByTree(list, hlist)
		g_horizontal_list = hlist
		console.log('list after drop', g_horizontal_list)
	})
}
// 通过tree重新构建horlist
function updateHorListByTree(tree, hlist) {
	if (!tree) {
		return
	}
	tree.forEach(e => {
		var children = e.children || []
		hlist.push({
			id: e.id,
			regionType: e.regionType,
			text: e.label,
			parent_id: e.parent_id,
			child_pos: e.pos[e.pos.length - 1],
			classType: e.classType,
			children: children.map(child => {
				return child.id
			})
		})
		if (e.children) {
			updateHorListByTree(e.children, hlist)
		}
	})
}

function refreshExamTree() {
	handleDocUpdate()
}


function updateTreeRenderWhenClick(data) {
	console.log('data', data)
	if (g_exam_tree) {
		g_exam_tree.setSelect([data.control_id])
	}
}

function handleDocUpdate() {
	console.log('===== handleDocUpdate getDocList')
	getDocList().then(res => {
		console.log(' handleDocUpdate 1', res)
		updateHListBYDoc(res)
		console.log(' handleDocUpdate 2', [].concat(g_horizontal_list))
		rebuildRelation(g_horizontal_list)
		console.log('======= after rebuildRelation ', [].concat(g_horizontal_list))
		var treeInfo = genetateTreeByHList(g_horizontal_list)
		console.log('============= treeInfo', treeInfo)
		renderTree(treeInfo)
	})
}
// 划分类型
function updateRangeControlType(typeName) {
	Asc.scope.typename = typeName
	biyueCallCommand(window, function() {
		var typeName = Asc.scope.typename
		var oDocument = Api.GetDocument()
		var oRange = oDocument.GetRangeBySelect()
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
		var result = {}
		var changeList = []
		function GetPosData(Pos) {
			var data = {}
			if (Pos) {
				for (var i = Pos.length - 1; i >= 0; --i) {
					if (Pos[i].Class.GetType) {
						var type = Pos[i].Class.GetType()
						if (type == 1) {
							data.index_paragraph = i
							return data
						} else if (type == 39) {
							data.index_run = i
						}
					}
				}
			}
			return data
		}

		function getParagraphPosition(Pos) {
			if (Pos) {
				for (var i = Pos.length - 1; i >= 0; --i) {
					if (Pos[i].Class.GetType) {
						var type = Pos[i].Class.GetType()
						if (type == 1) {
							return Pos[i - 1].Position
						}
					}
				}
			}
			return null
		}

		function  getRangeData(range) {
			var StartData = GetPosData(range.StartPos)
			var EndData = GetPosData(range.EndPos)
			var inParagraphStart = false
			var inParagraphEnd = false
			if (StartData && StartData.index_paragraph >= 0 && StartData.index_run >= 0) {
				if (
					range.StartPos[StartData.index_paragraph].Position == 0 &&
					range.StartPos[StartData.index_run].Position == 0
				) {
					inParagraphStart = true
				}
			}
			if (EndData && EndData.index_paragraph >= 0 && EndData.index_run >= 0) {
				if (
					range.EndPos[EndData.index_paragraph].Position >=
					range.EndPos[EndData.index_paragraph].Class.Content.length - 1 &&
					range.EndPos[EndData.index_run].Position >=
					range.EndPos[EndData.index_run].Class.Content.length - 1
				) {
					inParagraphEnd = true
				}
			}
			return [inParagraphStart, inParagraphEnd]
		}

		function getNewTag(inParagraphStart, inParagraphEnd) {
			var Tag = {}
			if (typeName == 'struct') {
				Tag = { regionType: 'struct', mode: 1, column: 1 }
			} else if (typeName == 'question') {
				Tag = { regionType: 'question', mode: 2, padding: [0, 0, 0.5, 0] }
			} else if (typeName == 'write') {
				if (inParagraphStart && inParagraphEnd) {
					Tag = { regionType: 'write', mode: 5 }
				} else {
					Tag = { regionType: 'write', mode: 3 }
				}
			} else if (typeName == 'sub-question') {
				Tag = { regionType: 'sub-question', mode: 3 }
			}
			return JSON.stringify(Tag)
		}
		if (!oRange) {
			var currentContentControl = oDocument.Document.GetContentControl()
			if (!currentContentControl) {
				return {
					code: 0,
					message: '请先选中一个范围',
				}
			}
			var oControl = allControls.find(e => {
				return e.Sdt.GetId() == currentContentControl.Id
			})
			if (oControl) {
				var tag = JSON.parse(oControl.GetTag() || {})
				oRange = oControl.GetRange()
				var changeResult = {
					id_old: oControl.Sdt.GetId(),
					text: oRange ? oRange.GetText() : '',
					regionType: typeName,
					type_old: tag.regionType
				}
				if (tag.regionType != typeName) {
					var rangeData = getRangeData(oRange)
					oControl.SetTag(getNewTag(rangeData[0], rangeData[1]));
					changeResult.command_type = 'change_type'
				} else {
					changeResult.command_type = 'keep'
				}
				changeList.push(changeResult)
			}
		} else {
			if (!oRange.Paragraphs || oRange.Paragraphs.length === 0) {
				return {
					code: 0,
					message: '选中范围内无段落',
				}
			}
			var controlsInRange = allControls.filter(e => {
				if (e.GetClassType() == 'blockLvlSdt') {
					return e.Sdt.Content.Selection && e.Sdt.Content.Selection.Use
				} else if (e.GetClassType() == 'inlineLvlSdt') {
					return e.Sdt.Selection && e.Sdt.Selection.Use
				}
			})

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
			console.log('controlsInRange', controlsInRange)
			if (controlsInRange) {
				controlsInRange.forEach((e, index) => {
					var orange = e.GetRange()
					console.log(' ', index, e.GetTag(), orange.GetText())
				})
			}
			var rangeData = getRangeData(oRange)
			var rangeStartParagraphPosition = getParagraphPosition(oRange.StartPos)
			var rangeEndParagraphPosition = getParagraphPosition(oRange.EndPos)
			function addControlToRange(range, tag, forceType) {
				if (!range) {
					console.log('addControlToRange range is null')
					return null
				}
				console.log('addControlToRange', range, tag)
				oDocument.RemoveSelection()
				range.Select()
				var rangeData = getRangeData(range)
				console.log('rangeData', rangeData)
				var type = forceType ? forceType : (rangeData[0] && rangeData[1] ? 1 : 2)
				var oResult = Api.asc_AddContentControl(type, {
					Tag: tag,
				})
				console.log('oResult', oResult)
				if(oResult) {
					return {
						id: oResult.InternalId,
						regionType: JSON.parse(tag || {}).regionType,
						text: range.GetText()
					}
				} else {
					return null
				}
			}
			var needAdd = true
			var removeIdList = []
			if (controlsInRange.length > 0) {
				for (var i = 0; i < controlsInRange.length; ++i) {
					var oControl = controlsInRange[i]
					if (!oControl) {
						continue
					}
					var controlRange = oControl.GetRange()
					console.log('controlRange', controlRange)
					var controlTag = JSON.parse(oControl.GetTag()) 
					if (checkPosSame(controlRange.StartPos, oRange.StartPos) && checkPosSame(controlRange.EndPos, oRange.EndPos)) {
						// 完全重叠
						console.log('完全重叠')
						var tag = JSON.parse(oControl.GetTag() || {})
						var changeObject = {
							id_old: oControl.Sdt.GetId(),
							text: oRange ? oRange.GetText() : '',
							regionType: typeName
						}
						if (tag.regionType != typeName) {
							oControl.SetTag(getNewTag(rangeData[0], rangeData[1]));
							changeObject.command_type = 'change_type'
						}
						changeList.push(changeObject)
						needAdd = false
						break
					} else {
						if ( (oControl.GetClassType() == 'blockLvlSdt') && (typeName == 'struct' || typeName == 'question' || (typeName == 'sub-question' && controlTag.regionType == 'sub-question'))) {
							var oControlContent = oControl.GetContent()
							var elementCount = oControlContent.GetElementsCount()
							changeList.push({
								id_old: oControl.Sdt.GetId(),
								command_type: 'remove'
							})
							if (rangeEndParagraphPosition < elementCount - 1) {
								var oElement = oControlContent.GetElement(rangeEndParagraphPosition + 1)
								if (oElement) {
									oElement.Select()
									var range1 = oDocument.GetRangeBySelect()
									var backRange = Api.CreateRange(oControl, range1.StartPos, controlRange.EndPos)
									var backData = addControlToRange(backRange, oControl.GetTag())
									if (backData) {
										changeList.push({
											id_new: backData.id,
											command_type: 'add',
											regionType: backData.regionType,
											text: backData.text
										})
									}
								} else {
									console.warn('rangeend cannot find element', rangeEndParagraphPosition + 1)
								}
							} 
							if (rangeStartParagraphPosition > 0) {
								var oElement = oControlContent.GetElement(rangeStartParagraphPosition - 1)
								if (oElement) {
									oElement.Select()
									var range1 = oDocument.GetRangeBySelect()
									var frontRange = Api.CreateRange(oControl, controlRange.StartPos, range1.EndPos)
									var frontData = addControlToRange(frontRange, oControl.GetTag())
									if (frontData) {
										changeList.push({
											id_new: frontData.id,
											command_type: 'add',
											regionType: frontData.regionType,
											text: frontData.text
										})
									}
								} else {
									console.warn('rangestart cannot find element', rangeStartParagraphPosition - 1, 'elementCount', elementCount)
									console.log('oControlContent', oControl)
									console.log('oRange', oRange)
									console.log('controlRange', controlRange)
								}
							}
							removeIdList.push(oControl.Sdt.GetId())
							needAdd = true
						} else if (typeName == 'write') {
							needAdd = true
							if (controlTag.regionType == 'write') {
								changeList.push({
									id_old: oControl.Sdt.GetId(),
									command_type: 'remove'
								})
								removeIdList.push(oControl.Sdt.GetId())
							}
						}
					}
				}
			}
			if (needAdd) {
				var useType = rangeData[0] && rangeData[1] ? 1 : 2
				if (useType == 2) {
					if (typeName != 'write') {
						useType = 1
					}
				}
				var addData = addControlToRange(oRange, getNewTag(rangeData[0], rangeData[1]), useType)
				if (addData) {
					changeList.push({
						id_new: addData.id,
						command_type: 'add',
						regionType: typeName,
						text: oRange ? oRange.GetText() : '',
					})
				}
			}
			if (removeIdList.length > 0) {
				removeIdList.forEach(id => {
					Api.asc_RemoveContentControlWrapper(id);
				})
			}
		}
		result.changeList = changeList
		result.code = 1
		return result
	}, false, true).then(res => {
		console.log('updateControlType result:', res)
		if (res) {
			if (res.code == 1) {
				getDocList().then(res2 => {
					if (res2) {
						updateHListBYDoc(res2)
						updatePosBySetType(res)
					}
				})	
			} else if (res.message) {
				alert(res.message)
			}
		}
	})
}

function getItemData(list, id) {
	if (!list) {
		return null
	}
	var index = list.findIndex(e => {
		return e.id == id
	})
	if (index >= 0) {
		if (list[index].parent_id) {
			var parentIndex = list.findIndex(e => {
				return e.id == list[index].parent_id
			})
		}
		return {
			index: index,
			parent_index: parentIndex
		}
	}
	return null
}
// 文档更新后，更新horlist，先不关注父子关系
function updateHListBYDoc(docList) {
	var hlist = g_horizontal_list
	console.log('updateHListBYDoc begin ', [].concat(hlist))
	var i = 0
	var imax = docList.length
	for (; i < imax; ++i) {
		var item = docList[i]
		if (item.id == hlist[i].id) {
			// id相同，更新text
			hlist[i].text = item.text
			hlist[i].regionType = item.regionType
		} else {
			var index1 = hlist.findIndex(e => {
				return e.id == item.id
			})
			var index2 = docList.findIndex(e => {
				return e.id == hlist[i].id
			})
			if (index1 >= 0) { // 之前已在hlist中
				if (index2 == -1) {
					// 删除
					if (hlist[i].children) {
						hlist[i].children.forEach(childId => {
							setNodeParentId(hlist, childId, 0)
						})
					}
					hlist.splice(i, 1)
					--i
				} else {
					// 位置改变
					var swapElement = hlist.splice(index1, 1)
					swapElement[0].text = item.text
					swapElement[0].regionType = item.regionType
					// 父节点ID，暂时不考虑
					hlist.splice(i, 0, ...swapElement)
				}
			} else { // 之前不在hlist中，需要插入
				var parent_id = 0
				if (item.regionType != 'struct') {
					var preItem = i > 0 ? getItemData(hlist, hlist[i - 1].id) : null
					var nextItem = i < hlist.length ? getItemData(hlist, hlist[i].id) : null
					if (preItem && nextItem) {
						var pindex = Math.max(preItem.parent_index, nextItem.parent_index)
						parent_id = hlist[pindex].id
					}
				}
				hlist.splice(i, 0, {
					id: item.id,
					regionType: item.regionType,
					text: item.text,
					classType: item.classType,
					parent_id: parent_id,
					children: [],
					child_pos: 0
				})
			}
		}
	}
	if (i < hlist.length) {
		hlist.splice(i, hlist.length - i)
	}
	console.log('updateHListBYDoc end ', [].concat(hlist))
}

function updatePosBySetType(changeData) {
	console.log('updatePosBySetType', changeData)
	console.log('updatePosBySetType', g_horizontal_list)
	var changeList = changeData.changeList
	if (!changeList || changeList.length == 0) {
		return
	}
	var hlist = g_horizontal_list
	changeList.forEach(cdata => {
		if (cdata.command_type == 'change_type' || cdata.command_type == 'keep') {
			var targetIndex = hlist.findIndex(e => {
				return e.id == cdata.id_old
			})
			if (targetIndex == -1) {
				return
			}
			var targetData = hlist[targetIndex]
			if (!targetData) {
				return
			}
			if (cdata.type_old == 'struct') {
				if (cdata.regionType == 'question') { // 原本是题组，改为题目
					if (targetData.children) {
						targetData.children.forEach(childId => {
							setNodeParentId(hlist, childId, 0)
						})
					}
					targetData.parent_id = 0
				} else if (cdata.regionType == 'struct') { // 依然是题组，需要把之后无主的条目归为children
					for (var i = targetIndex + 1; i < hlist.length; ++i) {
						if (hlist[i].regionType == 'struct') {
							break
						}
						if (hlist[i].parent_id == 0) {
							setNodeParentId(hlist, hlist[i].id, targetData.id)
						}
					}
				}
			} else if (cdata.type_old == 'question') {
				if (cdata.regionType == 'struct') { // 原本是题目，改为题组
					if (targetData.parent_id) {
						var oldParent = hlist.find(e => {
							return e.id == targetData.parent_id
						})
						if (oldParent && oldParent.children) {
							for (var i = targetData.child_pos + 1; i < oldParent.children.length; ++i) {
								setNodeParentId(hlist, oldParent.children[i], targetData.id)
							}
						}
					}
					targetData.parent_id = 0
					for (var i = targetIndex - 1; i >= 0; --i) {
						if (hlist[i].parent_id == 0) {
							preStructId = hlist[i].id
						}
					}
					for (var i = targetIndex + 1; i < hlist.length; ++i) {
						if (hlist[i].regionType == 'struct') {
							break
						}
						if (hlist[i].parent_id) {
							var tpindex = hlist.findIndex(e => {
								return e.id == hlist[i].parent_id
							})
							if (tpindex >= 0 && tpindex < targetIndex) {
								hlist[i].parent_id = targetData.id
							}
						} else {
							hlist[i].parent_id = targetData.id
						}
					}
				}
			}
		} else if (cdata.command_type == 'add') {
			var targetIndex = hlist.findIndex(e => {
				return e.id == cdata.id_new
			})
			if (targetIndex == -1) {
				return
			}
			var targetData = hlist[targetIndex]
			if (cdata.regionType == 'struct') { // 新增题组
				// 其之后的题目的父节点应该改为它
				var preStructId = 0
				for (var i = targetIndex - 1; i >= 0; --i) {
					if (hlist[i].parent_id == 0) {
						preStructId = hlist[i].id
					}
				}
				var preId = targetIndex > 0 ? hlist[targetIndex - 1].id : 0
				for (var i = targetIndex + 1; i < hlist.length; ++i) {
					if (hlist[i].regionType == 'struct') {
						break
					}
					if (hlist[i].parent_id) {
						var tpindex = hlist.findIndex(e => {
							return e.id == hlist[i].parent_id
						})
						if (tpindex >= 0 && tpindex < targetIndex) {
							hlist[i].parent_id = targetData.id
						}
					} else {
						hlist[i].parent_id = targetData.id
					}
				}
				targetData.parent_id = 0
			}
		}
	})
	console.log('========= after updatePosBySetType ', [].concat(hlist))
	rebuildRelation(hlist)
	console.log('======= after rebuildRelation ', [].concat(hlist))
	var treeInfo = genetateTreeByHList(hlist)
	console.log('============= treeInfo', treeInfo)
	renderTree(treeInfo)
}

function setNodeParentId(list, id, parent_id) {
	if (!list) {
		return
	}
	var target = list.find(e => {
		return e.id == id
	})
	if (target) {
		target.parent_id = parent_id
	}
}
// 在所有节点parent_id和顺序都已正确的情况下，重新生成child_pos和children
function rebuildRelation(list) {
	if (!list) {
		return
	}
	list.forEach((e, index) => {
		var child_pos = 0
		if (e.parent_id) {
			for (var i = index - 1; i >= 0; --i) {
				if (list[i].id == e.parent_id) {
					child_pos = list[i].children.length
					list[i].children.push(e.id)
				}
			}
		} else {
			for (var i = index - 1; i >= 0; --i) {
				if (list[i].parent_id == 0) {
					child_pos = list[i].child_pos + 1
					break
				}
			}
		}
		e.children = []
		e.child_pos = child_pos
	})
}


function reqGetQuestionType() {
	Asc.scope.horlist = g_horizontal_list
	biyueCallCommand(window, function() {
		var horlist = Asc.scope.horlist
		var target_list = []
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		horlist.forEach(e => {
			if (e.regionType == 'struct' || e.regionType == 'question') {
				var control = controls.find(citem => {
					return citem.Sdt.GetId() == e.id
				})
				if (control) {
					var oRange = control.GetRange()
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
						id: e.id,
						content_type: e.regionType,
						content_html: text_data.data
					})
				}
			}
		})
		console.log('[reqGetQuestionType] target_list', target_list)
		return target_list
	}, false, false).then( control_list => {
		console.log('[reqGetQuestionType] control_list', control_list)
		if (!window.BiyueCustomData.paper_uuid || !control_list || control_list.length == 0) {
			return
		}
		getQuesType(window.BiyueCustomData.paper_uuid, control_list).then(res => {
			console.log('getQuesType success ', res)
		}).catch(res => {
			console.log('getQuesType fail ', res)
		})
	})
}

function reqUploadTree() {
	Asc.scope.horlist = g_horizontal_list
	upload_control_list = []
	console.log('[reqUploadTree start]', Date.now())
	biyueCallCommand(window, function() {
		var horlist = Asc.scope.horlist
		var target_list = []
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		horlist.forEach(e => {
			if (e.regionType == 'struct' || e.regionType == 'question' || e.regionType == 'sub-question') {
				var control = controls.find(citem => {
					return citem.Sdt.GetId() == e.id
				})
				if (control) {
					var question_name = ''
					var oRange = control.GetRange()
					oRange.Select()
					var text = oRange.GetText()
					if (e.regionType == 'struct') {
						const pattern = /^[一二三四五六七八九十0-9]+.*?(?=[：:])/
						const result = pattern.exec(text)
						question_name = result ? result[0] : null
					} else if (e.regionType) {
						const regex = /^([^.．、]*)/
						const match = text.match(regex)
						question_name = match ? match[1] : ''
					}
					let text_data = {
						data:     "",
						// 返回的数据中class属性里面有binary格式的dom信息，需要删除掉
						pushData: function (format, value) {
							this.data = value ? value.replace(/class="[a-zA-Z0-9-:;+"\/=]*/g, "") : "";
						}
					};
			
					Api.asc_CheckCopy(text_data, 2);
					var content_html = text_data.data
					var content_type = e.regionType
					if (e.regionType == 'sub-question') {
						content_type = 'question'
					}
					target_list.push({
						parent_id: e.parent_id,
						id: e.id,
						uuid: "",
						regionType: e.regionType,
						content_type: content_type,
						content_xml: '',
						content_html: content_html,
						content_text: text,
						question_type: 0,
						question_name: question_name
					})
				}
			}
		})
		return target_list
	}, false, false).then( control_list => {
		if (control_list) {
			upload_control_list = control_list
			if (control_list && control_list.length) {
				getXml(control_list[0].id)
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
		return e.id == controlId
	})
	if (index == -1) {
		return
	}
	upload_control_list[index].content_xml = content
	if (index + 1 < upload_control_list.length) {
		getXml(upload_control_list[index+1].id)
		return
	}
	generateTreeForUpload(upload_control_list)
}

function handleXmlError() {
	generateTreeForUpload(upload_control_list)
}

function generateTreeForUpload(control_list) {
	if (!control_list) {
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
	reqComplete(uploadTree, version).then(res => {
		console.log('reqComplete', res)
		console.log('[reqUploadTree end]', Date.now())
	}).catch(res => {
		console.log('reqComplete fail', res)
		console.log('[reqUploadTree end]', Date.now())
	})
	console.log(tree)
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
	initExamTree,
	refreshExamTree,
	updateTreeRenderWhenClick,
	handleDocUpdate,
	updateRangeControlType,
	reqGetQuestionType,
	reqUploadTree
}
