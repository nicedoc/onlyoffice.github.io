// 主要用于管理试卷结构树
import Tree from '../components/Tree.js'
import { getQuesType } from '../scripts/api/paper.js'
import { biyueCallCommand } from './command.js'

var g_exam_tree = null
var g_horizontal_list = []

function initExamTree() {
	return new Promise((resolve, reject) => {
		console.log('[initExamTree]')
		updateTree().then((res) => {
			console.log('updateTree result', res)
			buildTree(res)
			resolve(true)
		}).catch((err) => {
			reject(err)
		})
	})
}

function updateTree() {
	return biyueCallCommand(window, function() {
		var oDocument = Api.GetDocument()
		var elementCount = oDocument.GetElementsCount()
		if (elementCount == 0) {
			return []
		}
		var list = []
		for (var idx = 0; idx < elementCount; ++idx) {
			var oElement = oDocument.GetElement(idx)
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
				// todo..
			}
		}

		// var control_list = oDocument.GetAllContentControls()
		// var index = 0
		// for (var i = 0, imax = control_list.length; i < imax; ++i) {
		// 	var control = control_list[i]
		// 	var tag = JSON.parse(control.GetTag() || {})
		// 	if (
		// 		tag.regionType == 'struct' ||
		// 		tag.regionType == 'question' ||
		// 		tag.regionType == 'sub-question'
		// 	) {
		// 		var oRange = control.GetRange()
		// 		var sPos = oRange.StartPos[0].Position
		// 		var ePos = oRange.EndPos[0].Position
		// 		if (sPos > index) {
		// 			for (var j = index; j < sPos; ++j) {
		// 				var element = oDocument.GetElement(j)
		// 				var classType = element.GetClassType()
		// 				if (classType == 'paragraph') {
		// 					list.push({
		// 						id: element.Paragraph.Id,
		// 						regionType: null,
		// 						text: element.GetRange().GetText(),
		// 						classType: classType,
		// 					})
		// 				}
		// 			}
		// 			index = sPos
		// 		}
		// 		list.push({
		// 			id: control.Sdt.GetId(),
		// 			regionType: tag.regionType,
		// 			text: oRange.GetText(),
		// 			classType: control.GetClassType(),
		// 		})
		// 		index = ePos
		// 	}
		// }
		console.log('[handle updateTree end]')
		return list
	}, false, false)
}

function addChild(list, parentId, childId) {
	var parent = list.find((e) => {
		return e.id == parentId
	})
	if (!parent) {
		return
	}
	if (!parent.children) {
		parent.children = []
	}
	parent.children.push(childId)
}

function generateTree(list) {
	if (!list) {
		return
	}
	list.forEach((e, index) => {
		var pre = index > 0 ? list[index - 1] : null
		if (!pre) {
			e.parent_id = 0
			e.pos = [0]
			e.child_pos = 0
		} else {
			if (e.regionType == 'struct') {
				e.parent_id = 0
				e.pos = [pre.pos[0] + 1]
				e.child_pos = pre.pos[0] + 1
			} else if (e.regionType == 'question') {
				if (pre.regionType == 'question' || !pre.regionType) {
					e.parent_id = pre.parent_id
					e.pos = getNextBrotherPos(pre.pos)
					e.child_pos = pre.child_pos + 1
					e.children = []
					addChild(list, pre.parent_id, e.id)
				} else if (pre.regionType == 'struct') {
					e.parent_id = pre.id
					e.pos = [].concat(pre.pos).concat([0])
					e.child_pos = 0
					addChild(list, pre.id, e.id)
				} else if (pre.regionType == 'sub-question') {
					if (pre.parent_id) {
						var subparent = list.find((item) => {
							return item.id == pre.parent_id
						})
						if (subparent.regionType == 'struct') {
							e.parent_id = subparent.id
							e.pos = getNextBrotherPos(pre.pos)
							e.child_pos = pre.child_pos + 1
							addChild(list, subparent.id, e.id)
						} else if (subparent.regionType == 'question') {
							e.parent_id = subparent.parent_id
							e.pos = getNextBrotherPos(subparent.pos)
							e.child_pos = subparent.child_pos + 1
							addChild(list, subparent.parent_id, e.id)
						}
					} else {
						e.parent_id = pre.parent_id
						e.pos = getNextBrotherPos(pre.pos)
						e.child_pos = pre.child_pos + 1
						addChild(list, pre.parent_id, e.id)
					}
				}
			} else if (e.regionType == 'sub-question') {
				if (pre.regionType == 'sub-question' || !pre.regionType) {
					e.parent_id = pre.parent_id
					e.pos = getNextBrotherPos(pre.pos)
					e.child_pos = pre.child_pos + 1
					addChild(list, pre.parent_id, e.id)
				} else if (pre.regionType == 'question' || pre.regionType == 'struct') {
					e.parent_id = pre.id
					e.pos = [].concat(pre.pos).concat([0])
					e.child_pos = 0
					addChild(list, pre.id, e.id)
				}
			} else {
				e.parent_id = pre.parent_id
				e.pos = getNextBrotherPos(pre.pos)
				e.child_pos = pre.child_pos + 1
			}
		}
	})
	g_horizontal_list = list
	window.BiyueCustomData.node_list = list
	console.log('g_horizontal_list', list)
	var newTree = generateTreeByList(list)
	return newTree
}

function generateTreeByList(list) {
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
				child_pos: e.child_pos,
				pos: e.pos,
			})
		} else {
			var parent = getItemByPos(tree, e.pos.slice(0, e.pos.length - 1), 0)
			if (parent && parent.children) {
				parent.children.push({
					id: e.id,
					label: e.text,
					expand: true,
					is_folder: e.regionType == 'struct',
					children: [],
					regionType: e.regionType,
					parent_id: e.parent_id,
					child_pos: e.child_pos,
					pos: e.pos,
				})
			} else {
				console.log('parent is null')
				debugger
			}
		}
	})
	return tree
}

function getItemByPos(tree, pos, index) {
	if (!tree) {
		return
	}
	if (index == pos.length - 1) {
		return tree[pos[index]]
	} else {
		return getItemByPos(tree[pos[index]].children, pos, index + 1)
	}
}

function buildTree(list) {
	console.log('buildTree]')
	var tree = generateTree(list)
	g_exam_tree = tree
	g_exam_tree = new Tree($('#treeRoot'))
	g_exam_tree.addCallBack(clickItem, dropItem)
	g_exam_tree.init(tree)
}

function clickItem(id, item, e) {
	if (g_exam_tree) {
		g_exam_tree.setSelect([id])
	}
	Asc.scope.click_item = item
	biyueCallCommand(window, function() {
		var clickData = Asc.scope.click_item
		var oDocument = Api.GetDocument()
		oDocument.RemoveSelection()
		var firstRange = null
		if (clickData.regionType) {
			var controls = oDocument.GetAllContentControls()
			var control = controls.find((e) => {
				return e.Sdt.GetId() == clickData.id
			})
			firstRange = control.GetRange()
			// 需要多选时，用下面代码扩展
			// var oRange = control.GetRange()
			// firstRange = firstRange.ExpandTo(oRange)
		} else {
			var oParagraph = new Api.private_CreateApiParagraph(
				AscCommon.g_oTableId.Get_ById(clickData.id)
			)
			firstRange = oParagraph.GetRange()
		}
		firstRange.Select()
	}, false, false)
}

function dropItem(list, dragId, dropId, direction) {
	console.log('dropItem', list)
	console.log('drop data: ', dragId, dropId, direction)
	window.BiyueCustomData.exam_tree = list
	// var hlist = []
	// updateHorizontalList(list, hlist)
	// g_horizontal_list = hlist
	var dragData = g_horizontal_list.find((e) => {
		return e.id == dragId
	})
	Asc.scope.drag_options = {
		dragId: dragId,
		dropId: dropId,
		direction: direction,
	}
	biyueCallCommand(window, function() {
		var drag_options = Asc.scope.drag_options
		var oDocument = Api.GetDocument()
		var controls = oDocument.GetAllContentControls()
		var dragControl = controls.find((e) => {
			return e.Sdt.GetId() == drag_options.dragId
		})
		if (!dragControl) {
			return
		}
		var tag = JSON.parse(dragControl.GetTag() || {})
		console.log('tag ', tag)
		var oDragRange = dragControl.GetRange()
		var templist = []
		var parentControl = dragControl.GetParentContentControl()
		var parentTableCell = dragControl.GetParentTableCell()
		if (parentTableCell) {
			var cellContent = parentTableCell.GetContent()
			for (var i = 0; i < cellContent.Cotent.length; ++i) {
				if (cellContent.Content.Id == drag_options.dragId) {
					templist.push(cellContent.GetElement(i))
					cellContent.RemoveElement(i)
					break
				}
			}
			return
		}
		if (tag.regionType == 'sub-question') {
			if (parentControl) {
				var startPos = oDragRange.StartPos[1].Position
				var oParentDocument = parentControl.GetContent()
				templist.push(oParentDocument.GetElement(startPos))
				oParentDocument.RemoveElement(startPos)
				var dropControl = controls.find((e) => {
					return e.Sdt.GetId() == drag_options.dropId
				})
				var oDropRange = dropControl.GetRange()
				var oDropParentControl = dropControl.GetParentContentControl()
				var dropPos = oDropRange.StartPos[0].Position
				if (oDropParentControl) {
					dropPos = oDropRange.StartPos[1].Position
					console.log('add element', dropPos)
					if (drag_options.direction == 'top') {
						console.log('AddElement', dropPos, templist[0])
						oDropParentControl.GetContent().AddElement(dropPos, templist[0])
						console.log(
							'oDropParentControl',
							oDropParentControl.GetContent().GetElementsCount()
						)
					} else {
						oDropParentControl
							.GetContent()
							.AddElement(dropPos + 1, templist[0])
					}
				} else {
					console.log('not oDropParentControl')
					if (drag_options.direction == 'top') {
						oDocument.AddElement(dropPos, templist[0])
					} else {
						oDocument.AddElement(dropPos + 1, templist[0])
					}
				}
			}
		} else if (tag.regionType == 'question') {
			var startPos = oDragRange.StartPos[0].Position
			console.log('startPos', startPos)
			templist.push(oDocument.GetElement(startPos))

			oDocument.RemoveElement(startPos)

			var dropControl = controls.find((e) => {
				return e.Sdt.GetId() == drag_options.dropId
			})
			var oDropRange = dropControl.GetRange()
			var dropPos = oDropRange.StartPos[0].Position
			console.log('dropPos', dropPos)
			if (drag_options.direction == 'top') {
				oDocument.AddElement(dropPos, templist[0])
			} else {
				oDocument.AddElement(dropPos + 1, templist[0])
			}
		}
	}, false, false).then(res => {
		console.log('drop success')
	})
}

function updateHorizontalList(list, hlist) {
	if (!list) {
		return
	}
	list.forEach((e) => {
		hlist.push({
			id: e.id,
			label: e.label,
			pos: e.pos,
		})
		if (e.children) {
			updateHorizontalList(e.children, hlist)
		}
	})
}

function refreshExamTree() {
	handleDocUpdate()
}

function updateTreeRenderWhenClick(data) {
	if (g_exam_tree) {
		g_exam_tree.setSelect([data.control_id])
	}
}
// 根据pos, 生成其之后的兄弟pos
function getNextBrotherPos(pos) {
	if (pos || pos.length == 0) {
		var lastPosition = pos[pos.length - 1]
		var parentpos = pos.slice(0, pos.length - 1)
		parentpos.push(lastPosition + 1)
		return parentpos
	} else {
		return [0]
	}
}
// 在parent下的指定位置addPos插入一个child, 要相应的改变其之后的child的pos
function moveAfterPosWhenAdd(list, parentId, addPos, addId) {
	if (!list) {
		return
	}
	if (!parentId) {
		for (var i = 0; i < list.length; ++i) {
			if (list[i].parentId == 0) {
				if (list[i].child_pos >= addPos) {
					list[i].child_pos += 1
				}
			}
		}
		return
	}
	var parent = list.find((e) => {
		return e.id == parentId
	})
	if (!parent) {
		return
	}
	if (!parent.children) {
		parent.children = []
	}
	parent.children.splice(addPos, 0, addId)
	for (var i = addPos + 1; i < parent.children.length; ++i) {
		var child = list.find((e) => {
			return e.id == parent.children[i]
		})
		if (child) {
			child.child_pos += 1
		}
	}
}

function moveAfterPosWhenRemove(list, removeId) {
	if (!list) {
		return
	}
	var remove = list.find((e) => {
		return e.id == removeId
	})
	if (!remove) {
		return
	}
	var childCount = 0
	if (remove.children && remove.children.length > 0) {
		childCount = remove.children.length
	}
	// 修改其父节点下的children的相对位置
	for (var i = 0; i < list.length; ++i) {
		if (list[i].parent_id == remove.parent_id) {
			if (list[i].child_pos > remove.child_pos) {
				list[i].child_pos += childCount - 1
			}
		}
	}
	// 修改其子节点的相对位置
	for (var i = 0; i < childCount; ++i) {
		var child = list.find((e) => {
			return e.id == remove.children[i]
		})
		if (child) {
			child.parent_id = remove.parent_id
			child.child_pos += remove.child_pos
		}
	}
}

function moveNodeAndChildrenToPos(list, index, newindex) {
	if (!list) {
		return
	}
	var node = list[index]
	if (!node) {
		return
	}
	var childCount = (node.children || []).length
	if (newindex == 0) {
		// 受影响的范围, 需要重新调整childpos
		for (var i = newindex; i < index; ++i) {
			if (list[i].parent_id == 0) {
				list[i].child_pos += 1
			}
		}
		// 调整原本之后的兄弟节点的child_pos, 并将其从children移除
		if (node.parent_id) {
			var parentIndex = list.findIndex((e) => {
				return node.parent_id == e.id
			})
			var parent = list[parentIndex]
			if (parent.children) {
				for (var k = node.child_pos + 1; k < parent.children.length; ++k) {
					var child = list.find((e) => {
						return e.id == parent.children[k]
					})
					if (child) {
						child.child_pos -= 1
					}
				}
				parent.children.splice(node.child_pos, 1)
			}
		}
		node.parent_id = 0
		node.child_pos = 0
		var temp = list.splice(index, childCount + 1)
		list.splice(newindex, 0, ...temp)
	} else {
		if (node.parent_id) {
			var parentIndex = list.findIndex((e) => {
				return node.parent_id == e.id
			})
			var parent = list[parentIndex]
			if (parentIndex < newindex) {
				// 父节点未改变, 只是child_pos改变
				var newChildPos = node.child_pos
				for (var i = newindex; i < index; ++i) {
					if (list[i].parent_id == parent.id) {
						if (i == newindex) {
							newChildPos = list[i].child_pos
						}
						list[i].child_pos += 1
					}
				}
				parent.children.splice(node.child_pos, 1)
				node.child_pos = 1
				parent.children.splice(0, 0, node.id)
				parent.children.splice(newChildPos, 0, node.id)
				var temp = list.splice(index, childCount + 1)
				list.splice(newindex, 0, ...temp)
			} else {
				// 父节点改变
				// 将其从原有的父节点中移除
				parent.children.splice(node.child_pos, 1)
				// 更新children的child_pos
				for (var i = node.child_pos; i < parent.children.length; ++i) {
					var child = list.find((e) => {
						return e.id == parent.children[i]
					})
					if (child) {
						child.child_pos = i
					}
				}
				// 将其添加到新的父节点中
				var newPre = list[newindex - 1]
				var newParent = list.find((e) => {
					return e.id == newPre.parent_id
				})
				for (var i = newPre.child_pos + 1; i < newParent.children.length; ++i) {
					var child = list.find((e) => {
						return e.id == newParent.children[i]
					})
					if (child) {
						child.child_pos = i + 1
					}
				}
				try {
					newParent.children(newPre.child_pos + 1, 0, node.id)
				} catch (error) {
					debugger
				}

				node.parent_id = newParent.id
				node.child_pos = newPre.child_pos + 1
				var temp = list.splice(index, childCount + 1)
				list.splice(newindex, 0, ...temp)
			}
		}
	}
}

function updateHorList1(ctrllist) {
	var horlist = g_horizontal_list
	var changeList = []
	var i = 0
	var imax = ctrllist.length
	for (; i < imax; ++i) {
		if (i < horlist.length) {
			if (ctrllist[i].id == horlist[i].id) {
				if (ctrllist[i].text != horlist[i].text) {
					horlist[i].text = ctrllist[i].text
					if (horlist[i].regionType) {
						changeList.push(horlist[i].id)
					}
				}
			} else {
				var iold = horlist.findIndex((e) => {
					return e.id == ctrllist[i].id
				})
				if (iold >= 0) {
					// 之前就已存在
					var index2 = ctrllist.findIndex((e) => {
						return e.id == horlist[i].id
					})
					if (index2 == -1) {
						// 未找到, 需要删除
						moveAfterPosWhenRemove(horlist, horlist[i].id)
						horlist.splice(i, 1)
						--i
					} else {
						// 有找到, 属于位置移动
						// 将节点及其子节点都挪到对应位置
						moveNodeAndChildrenToPos(horlist, iold, i)
					}
				} else {
					// 新加入, 需要插入
					if (i > 0) {
						var pre = horlist[i - 1]
						horlist.splice(
							i,
							0,
							Object.assign({}, ctrllist[i], {
								parent_id: pre.parent_id,
								pos: getNextBrotherPos(pre.pos),
								child_pos: pre.child_pos + 1,
							})
						)
						moveAfterPosWhenAdd(
							horlist,
							pre.parent_id,
							pre.child_pos + 1,
							ctrllist[i].id
						)
					} else {
						horlist.splice(
							i,
							0,
							Object.assign({}, ctrllist[i], {
								parent_id: 0,
								pos: [0],
								child_pos: 0,
							})
						)
						moveAfterPosWhenAdd(horlist, 0, 0, ctrllist[i].id)
					}
				}
			}
		} else {
			var lastdata = horlist[horlist.length - 1]
			if (lastdata.parent_id) {
				var parent = horlist.find((e) => {
					return e.id == lastdata.parent_id
				})
				parent.children.push(ctrllist[i].id)
			}
			horlist.splice(
				i,
				0,
				Object.assign({}, ctrllist[i], {
					parent_id: lastdata.parent_id,
					pos: getNextBrotherPos(lastdata.pos),
					child_pos: lastdata.child_pos + 1,
				})
			)
		}
	}
	if (i < horlist.length) {
		// 删除题目
		horlist.splice(i, imax - i)
	}
	return horlist
}

function handleDocUpdate() {
	console.log('handleDocUpdate')
	updateTree().then((ctrllist) => {
		console.log('handleDocUpdate', ctrllist)
		var horlist = updateHorList1(ctrllist)
		console.log('after handleDocUpdate', horlist)
		var tree = generateTreeByList(horlist)
		if (g_exam_tree) {
			g_exam_tree.refreshList(tree)
		}
	})
}
// 更新选中range控件的类型
function updateRangeControlType(typeName) {
	Asc.scope.typename = typeName
	biyueCallCommand(window, function() {
		var typeName = Asc.scope.typename
		var oDocument = Api.GetDocument()
		var oRange = oDocument.GetRangeBySelect()
		if (!oRange) {
			console.log('no range')
			return
		}
		if (!oRange.Paragraphs) {
			console.log('no paragraph')
			return
		}
		if (oRange.Paragraphs.length === 0) {
			console.log('no paragraph')
			return
		}
		function GetPosData(Pos) {
			var data = {}
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
			return data
		}
		var StartData = GetPosData(oRange.StartPos)
		var EndData = GetPosData(oRange.EndPos)
		var inParagraphStart = false
		var inParagraphEnd = false
		if (StartData.index_paragraph >= 0 && StartData.index_run >= 0) {
			if (
				oRange.StartPos[StartData.index_paragraph].Position == 0 &&
				oRange.StartPos[StartData.index_run].Position == 0
			) {
				inParagraphStart = true
			}
		}
		if (EndData.index_paragraph >= 0 && EndData.index_run >= 0) {
			if (
				oRange.EndPos[EndData.index_paragraph].Position >=
					oRange.EndPos[EndData.index_paragraph].Class.Content.length - 1 &&
				oRange.EndPos[EndData.index_run].Position >=
					oRange.EndPos[EndData.index_run].Class.Content.length - 1
			) {
				inParagraphEnd = true
			}
		}
		if (typeName == 'struct' || typeName == 'question') {
			if (!inParagraphStart || !inParagraphEnd) {
				return {
					code: 0,
					message: '请选中整个段落再设置',
				}
			}
		}
		var type = 1
		var Tag = null
		if (typeName == 'examtitle') {
			type = 0
			oRange.SetBold(true)
			return {
				code: 1,
				type: type,
				typeName: typeName,
				text: oRange.GetText(),
			}
		} else {
			if (typeName == 'struct') {
				Tag = { regionType: 'struct', mode: 1, column: 1 }
			} else if (typeName == 'question') {
				Tag = { regionType: 'question', mode: 2, padding: [0, 0, 0.5, 0] }
			} else if (typeName == 'write') {
				if (inParagraphStart && inParagraphEnd) {
					type = 1
					Tag = { regionType: 'write', mode: 5 }
				} else {
					type = 2
					Tag = { regionType: 'write', mode: 3 }
				}
			} else if (typeName == 'sub-question') {
				Tag = { regionType: 'sub-question', mode: 3 }
			}
			var controlsInRange = []
			oRange.Paragraphs.forEach((paragraph) => {
				var controls = paragraph.GetAllContentControls()
				if (controls) {
					controls.forEach((controlItem) => {
						if (controlItem.Sdt.Selection && controlItem.Sdt.Selection.Use) {
							var tag1 = JSON.parse(controlItem.GetTag())
							if (tag1 && tag1.regionType == typeName) {
								controlsInRange.push(controlItem)
							}
						}
					})
				}
				var pControl = paragraph.GetParentContentControl()
				if (pControl) {
					var tag2 = JSON.parse(pControl.GetTag())
					if (tag2 && tag2.regionType == typeName) {
						if (
							controlsInRange.findIndex((e) => {
								return e.Sdt.GetId() == pControl.Sdt.GetId()
							}) == -1
						) {
							controlsInRange.push(pControl)
						}
					}
				}
			})
			var oResult = Api.asc_AddContentControl(type, {
				Tag: JSON.stringify(Tag),
			})
			console.log('oResult', oResult)
			console.log('oRange', oRange)
			var contentPos = oRange.StartPos[0].Position
			var result = {
				code: 1,
				type: type,
				InternalId: oResult.InternalId,
				regionType: type,
				text: oRange.GetText(),
				tag: Tag,
			}
			if (contentPos > 0) {
				var preElement = oDocument.GetElement(contentPos - 1)
				result.preClassType = preElement.GetClassType()
				result.preElementId = oDocument.Document.Content[contentPos - 1].Id
				result.operateType = 'add'
			}

			function getElementList() {
				var elementCount = oDocument.GetElementsCount()
				if (elementCount == 0) {
					return []
				}
				var list = []
				for (var idx = 0; idx < elementCount; ++idx) {
					var oElement = oDocument.GetElement(idx)
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
						// todo..
					}
				}
				return list
			}
			result.elementList = getElementList()
			return result
			//   console.log('controlsInRange', controlsInRange)
			//   if (controlsInRange.length > 0) {
			//     if (confirm("该区域已存在相同类型区域，是否要删除覆盖?") == true) {
			//       controlsInRange.forEach(controlItem => {
			//         Api.asc_RemoveContentControlWrapper(controlItem.Sdt.GetId());
			//       })
			//       oRange.Select()
			//     } else {
			//       return {
			//         code: 2
			//       }
			//     }
			//   }
		}
	}, false, true).then(res => {
		console.log('updateControlType result:', res)
		handleSetControlType(res)
	})
}

function handleSetControlType(res) {
	if (!res) {
		return
	}
	if (res.code != 1) {
		return
	}
	var horlist = updateHorList1(res.elementList)
	updateHorList2(horlist, res.regionType, res.InternalId, res.preElementId, res.preClassType)
	var tree = generateTreeByList(horlist)
	if (g_exam_tree) {
		g_exam_tree.refreshList(tree)
	}
}
// 划分类型后，进一步调整层次结构
function updateHorList2(horlist, regionType, InternalId, preElementId, preClassType) {
	if (!horlist) {
		return
	}
	for (var i = 0; i < horlist.length; ++i) {
		if (horlist[i].id == InternalId) {
			var item = horlist[i]
			if (regionType == 'struct') { // 设置为题组
				if (item.parent_id != 0) {
					var parent = horlist.find(e => {
						return e.id == item.parent_id
					})
					if (parent && parent.children) {
						// 修改自身的父节点
						item.parent_id = 0
						item.child_pos = parent.pos[0] + 1
						item.pos = [item.child_pos]
						// 修改其之后的兄弟节点的位置
						for (var j = item.child_pos + 1, jmax = parent.children.length; j < jmax; ++j) {
							var child = horlist.find(e => {
								return e.id == parent.children[j].id
							})
							if (child) {
								child.parent_id = InternalId
								child.child_pos = jmax - j - 1
								child.pos = [item.child_pos, child.child_pos]
							}
						}
						parent.children.splice(item.child_pos, parent.children.length - item.child_pos)
						// 修改原有父节点之后的兄弟节点的位置
						for (var k = i + 1; k < horlist.length; ++k) {
							if (horlist[k].parent_id == 0) {
								horlist[k].child_pos += 1
							}
							horlist[k].pos[0] += 1
						}
					}
				}

			} else if (regionType == 'question') { // 设置为题目

			} else if (regionType == 'sub-question') { // 设置为小题

			}
			return
		}
	}
}

function reqGetQuestType() {
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
		console.log('[reqGetQuestType] target_list', target_list)
		return target_list
	}, false, false).then( control_list => {
		console.log('[reqGetQuestType] control_list', control_list)
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

export {
	initExamTree,
	refreshExamTree,
	updateTreeRenderWhenClick,
	handleDocUpdate,
	updateRangeControlType,
	reqGetQuestType
}
