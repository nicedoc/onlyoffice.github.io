// 主要用于管理试卷结构树
import Tree from '../components/Tree.js'
import { getQuesType, reqComplete } from '../scripts/api/paper.js'
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
		var allControls = oDocument.GetAllContentControls()
		var list = []
		for (var idx = 0; idx < elementCount; ++idx) {
			var oElement = oDocument.GetElement(idx)
			if (!oElement) {
				continue
			}
			if (!oElement.GetClassType) {
				console.log('========= oElement cannot GetClassType', oElement)
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
				classType: e.classType
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
					classType: e.classType
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
			if (control) {
				firstRange = control.GetRange()
			}
			// 需要多选时，用下面代码扩展
			// var oRange = control.GetRange()
			// firstRange = firstRange.ExpandTo(oRange)
		} else {
			var oParagraph = new Api.private_CreateApiParagraph(
				AscCommon.g_oTableId.Get_ById(clickData.id)
			)
			if (oParagraph) {
				firstRange = oParagraph.GetRange()
			}
		}
		if (firstRange) {
			firstRange.Select()
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
		function GetElementRange(id, classType) {
			if (classType == 'blockLvlSdt') {
				var control = controls.find((e) => {
					return e.Sdt.GetId() == id
				})
				if (control) {
					return control.GetRange()
				}
			} else if (classType == 'paragraph') {
				var oParagraph = new Api.private_CreateApiParagraph(
					AscCommon.g_oTableId.Get_ById(id)
				)
				if (oParagraph) {
					return oParagraph.GetRange()
				}
			}
			return null
		}
		var oDragRange = GetElementRange(drag_options.dragId, dragData.classType)
		if (!oDragRange) {
			return
		}
		
		var templist = []
		if (regionType == 'struct') {
			// 题组拖拽，需要将其下的所有children都挪动位置
			if (dragData.children) {
				for (var i = dragData.children.length - 1; i >= 0; --i) {
					var childId = dragData.children[i]
					var childData = horList.find(e => {
						return e.id == childId
					})
					var childRange = GetElementRange(childId, childData.classType)
					if (childRange) {
						var startPos = childRange.StartPos[0].Position
						templist.push(oDocument.GetElement(startPos))
						oDocument.RemoveElement(startPos)
					}
				}
			}
			var startPos = oDragRange.StartPos[0].Position
			templist.push(oDocument.GetElement(startPos))
			oDocument.RemoveElement(startPos)
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
				var startPos = oDragRange.StartPos[0].Position
				templist.push(oDocument.GetElement(startPos))
				oDocument.RemoveElement(startPos)

				var dropRange = GetElementRange(drag_options.dropId, dropData.classType)
				if (dropRange) {
					var dropPos = dropRange.StartPos[0].Position
					if (drag_options.direction == 'top') {
						oDocument.AddElement(dropPos, templist[0])
					} else {
						oDocument.AddElement(dropPos + 1, templist[0])
					}
				} else {
					console.log('dropRange is null')
				}
			} else if (regionType == 'sub-question') {
				var dragControl = controls.find((e) => {
					return e.Sdt.GetId() == drag_options.dragId
				})
				var parentControl = dragControl.GetParentContentControl()
				if (parentControl) {
					var startPos = oDragRange.StartPos[1].Position
					var oParentDocument = parentControl.GetContent()
					templist.push(oParentDocument.GetElement(startPos))
					oParentDocument.RemoveElement(startPos)
					var oDropRange
					var oDropParentControl
					if (dropData.classType == 'blockLvlSdt') {
						var dropControl = controls.find((e) => {
							return e.Sdt.GetId() == drag_options.dropId
						})
						if (dropControl) {
							oDropRange = dropControl.GetRange()
							oDropParentControl = dropControl.GetParentContentControl()
						}
					} else if (dropData.classType == 'paragraph') {
						var oDropParagraph = new Api.private_CreateApiParagraph(
							AscCommon.g_oTableId.Get_ById(drag_options.dropId)
						)
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
			child_pos: e.pos[e.pos.length - 1],
			classType: e.classType,
			id: e.id,
			parent_id: e.parent_id,
			pos: e.pos,
			regionType: e.regionType,
			text: e.label,
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
		horlist.splice(i, horlist.length - i)
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
		if (typeName == 'examtitle') {
			oRange.SetBold(true)
			return
		}
		var allControls = oDocument.GetAllContentControls()
		var result = {}
		var changeList = []
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

		function getParagraphPosition(Pos) {
			for (var i = Pos.length - 1; i >= 0; --i) {
				if (Pos[i].Class.GetType) {
					var type = Pos[i].Class.GetType()
					if (type == 1) {
						return Pos[i - 1].Position
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
			if (StartData.index_paragraph >= 0 && StartData.index_run >= 0) {
				if (
					range.StartPos[StartData.index_paragraph].Position == 0 &&
					range.StartPos[StartData.index_run].Position == 0
				) {
					inParagraphStart = true
				}
			}
			if (EndData.index_paragraph >= 0 && EndData.index_run >= 0) {
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
					regionType: typeName
				}
				if (tag.regionType != typeName) {
					var rangeData = getRangeData(oRange)
					oControl.SetTag(getNewTag(rangeData[0], rangeData[1]));
					changeResult.command_type = 'change_type'
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
				return e.Sdt.Content.Selection && e.Sdt.Content.Selection.Use
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
			var rangeData = getRangeData(oRange)
			var rangeStartParagraphPosition = getParagraphPosition(oRange.StartPos)
			var rangeEndParagraphPosition = getParagraphPosition(oRange.EndPos)
			function addControlToRange(range, tag, forceType) {
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
				return {
					id: oResult.InternalId,
					regionType: JSON.parse(tag || {}).regionType,
					text: range.GetText()
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
						if (typeName == 'struct' || typeName == 'question' || (typeName == 'sub-question' && controlTag.regionType == 'sub-question')) {
							var oControlContent = oControl.GetContent()
							var elementCount = oControlContent.GetElementsCount()
							var changeobj = {
								id_old: oControl.Sdt.GetId(),
								command_type: 'remove'
							}
							if (rangeEndParagraphPosition < elementCount - 1) {
								var oElement = oControlContent.GetElement(rangeEndParagraphPosition + 1)
								oElement.Select()
								var range1 = oDocument.GetRangeBySelect()
								var backRange = Api.CreateRange(oControl, range1.StartPos, controlRange.EndPos)
								var backData = addControlToRange(backRange, oControl.GetTag())
								if (!changeobj.new_controls) {
									changeobj.new_controls = [backData]
								} else {
									changeobj.new_controls.push(backData)
								}
							} 
							if (rangeStartParagraphPosition > 0) {
								var oElement = oControlContent.GetElement(rangeStartParagraphPosition - 1)
								oElement.Select()
								var range1 = oDocument.GetRangeBySelect()
								var frontRange = Api.CreateRange(oControl, controlRange.StartPos, range1.EndPos)
								var frontData = addControlToRange(frontRange, oControl.GetTag())
								if (!changeobj.new_controls) {
									changeobj.new_controls = [frontData]
								} else {
									changeobj.new_controls.push(frontData)
								}
							}
							changeList.push(changeobj)
							removeIdList.push(oControl.Sdt.GetId())
							needAdd = true
						} else if (typeName == 'write') {
							needAdd = true
							if (controlTag.regionType == 'write') {
								var changeobj = {
									id_old: oControl.Sdt.GetId(),
									command_type: 'remove'
								}
								changeList.push(changeobj)
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
				changeList.push({
					id_new: addData.id,
					command_type: 'add',
					regionType: typeName,
					text: oRange ? oRange.GetText() : '',
				})
			}
			if (removeIdList.length > 0) {
				removeIdList.forEach(id => {
					Api.asc_RemoveContentControlWrapper(id);
				})
			}
		}
		result.code = 1
		result.changeList = changeList
		function getElementList() {
			var elementCount = oDocument.GetElementsCount()
			if (elementCount == 0) {
				return []
			}
			var list = []
			for (var idx = 0; idx < elementCount; ++idx) {
				var oElement = oDocument.GetElement(idx)
				if (!oElement) {
					continue
				}
				if (!oElement.GetClassType) {
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
		}
		result.elementList = getElementList()
		return result
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
		if (res.message) {
			alert(res.message)
		}
		return
	}
	var horlist = updateHorList1(res.elementList)
	updateHorList3(horlist, res.changeList)
	var tree = generateTreeByList(horlist)
	if (g_exam_tree) {
		g_exam_tree.refreshList(tree)
	}
}

function updateHorList3(horList, changeList) {
	if (!horList) {
		return
	}
	if (!changeList || changeList.length == 0) {
		return
	}
	var change0 = changeList[0]
	if (change0.command_type == 'add') {
		if (change0.id_new) {
			if (change0.regionType == 'struct') { // 设置新题组
				// 找到这个id所在位置，更新其父节点的children
				var index = horList.findIndex((e) => {
					return e.id == change0.id_new
				})
				var data0 = horList[index]
				if (data0.parent_id) { // 存在父节点
					var parent = horList.find((e) => {
						return e.id == data0.parent_id
					})
					if (parent) {
						// 修改其之后的兄弟节点的父节点
						for (var j = data0.child_pos + 1; j < parent.children.length; ++j) {
							// todo..
						}
					}
				}
			}
		}
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
					target_list.push({
						parent_id: e.parent_id,
						id: e.id,
						uuid: "",
						content_type: e.regionType,
						content_xml: '',
						content_html: content_html,
						content_text: text,
						question_type: 0,
						question_name: question_name
					})
				}
			}
		})
		console.log('[reqUploadTree] target_list', target_list)
		return target_list
	}, false, false).then( control_list => {
		console.log('[reqUploadTree] control_list', control_list)
		generateTreeForUpload(control_list)
		getXml(control_list[0].id)
	})
}

function getXml(controlId) {
	console.log('             getXml', controlId)
	window.Asc.plugin.executeMethod("SelectContentControl", [controlId])
	window.Asc.plugin.executeMethod("GetSelectionToDownload", ["docx"], function (data) {
        console.log(data);
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
                    console.log(relativePath, content);  
                });  
            });  
        })  
        .catch(error => {  
            console.error('Error:', error);  
        });
        
    });
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
	reqComplete(uploadTree).then(res => {
		console.log('reqComplete', res)
	}).catch(res => {
		console.log('reqComplete fail', res)
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

export {
	initExamTree,
	refreshExamTree,
	updateTreeRenderWhenClick,
	handleDocUpdate,
	updateRangeControlType,
	reqGetQuestionType,
	reqUploadTree
}