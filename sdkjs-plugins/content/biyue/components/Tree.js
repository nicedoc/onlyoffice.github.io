// 试卷结构树
const CLASS_EXPAND_OPEN = 'icon-angeldown1'
const CLASS_EXPAND_CLOSE = 'icon-arrowright-copy'
const CLASS_FOLDER_OPEN = 'icon-folder'
const CLASS_FOLDER_CLOSE = 'icon-folder1'
const CLASS_EXPAND = 'iconexpand'
const CLASS_FOLDER = 'iconfolder'
const CLASS_INNER = 'inner'
const CLASS_CHILDREN = 'tree-children'
const CLASS_ITEM = 'tree-item'
const CLASS_TEXT = 'text'
class Tree {
	constructor(rootElement) {
		this.list = []
		this.rootElement = rootElement
		this.dragging = false
		this.select_id_list = []
	}
	// 重置，清除子节点事件监听
	reset() {
		this.removeEventListener(this.rootElement)
		this.rootElement.empty()
	}

	removeEventListener(element) {
		var children = element.children()
		if (children && children.length) {
			for (var i = 0, imax = children.length; i < imax; ++i) {
				var e = $(children[i])
				if (e.hasClass(CLASS_ITEM)) {
					e.off('dragstart')
				}
				if (e.hasClass(CLASS_INNER)) {
					e.off('click')
				}
				this.removeEventListener(e)
			}
		}
	}

	init(list) {
		console.log('tree init', list)
		this.updatePos(list, [])
		this.list = list
		this.buildTree(list, this.rootElement)
		this.rootElement.on('dragover', (e) => {
			this.onDragOver(e)
		})
		this.rootElement.on('drop', (e) => {
			this.onDrop(e)
		})
		this.rootElement.mouseout((e) => {
			this.dragging = false
		})
		this.rootElement.on('dragend', (e) => {
			console.log('dragend', e)
			e.originalEvent.cancelBubble = true
			e.originalEvent.stopPropagation()
			e.originalEvent.preventDefault()
			this.updateOverClass()
		})
	}

	addCallBack(clickCB, dropCB) {
		this.callback_item_click = clickCB
		this.callback_drop = dropCB
	}

	buildTree(data, parent) {
		if (!data || !parent) {
			return
		}
		this.removeEventListener(this.rootElement)
		this.rootElement.empty()
		console.log('buildTree', data)
		data.forEach((item) => {
			this.createItem(item, parent)
		})
	}

	refreshList(list) {
		this.init(list)
	}

	updateDataById(list, id, options) {
		for (var i = 0, imax = list.length; i < imax; ++i) {
			if (list[i].id == id) {
				Object.keys(options).forEach((key) => {
					list[i][key] = options[key]
				})
				return
			}
			if (list[i].children && list[i].children.length > 0) {
				this.updateDataById(list[i].children, id, options)
			}
		}
	}

	updateDataByPos(list, pos, options, index) {
		if (!list || !pos) {
			return
		}
		if (index >= pos.length || index < 0) {
			return
		}
		var i = pos[index]
		if (!list[i]) {
			return
		}
		if (index < pos.length - 1) {
			this.updateDataByPos(list[i].children, pos, options, index + 1)
		} else {
			Object.keys(options).forEach((key) => {
				list[i][key] = options[key]
			})
		}
	}

	onItemClick(e, id, item) {
		if (this.dragging) {
			return
		}
		e.cancelBubble = true
		if (item.children && item.children.length > 0) {
			var inner = $(`#${id}`).find('.inner').find(`.${CLASS_EXPAND}`)
			if (item.expand) {
				inner.removeClass(CLASS_EXPAND_OPEN).addClass(CLASS_EXPAND_CLOSE)
				$(`#${id} .${CLASS_CHILDREN}`).hide()
			} else {
				inner.removeClass(CLASS_EXPAND_CLOSE).addClass(CLASS_EXPAND_OPEN)
				$(`#${id} .${CLASS_CHILDREN}`).show()
			}
			if (item.is_folder) {
				var foldericon = $(`#${id}`)
					.find(`.${CLASS_INNER}`)
					.find(`.${CLASS_FOLDER}`)
				if (item.expand) {
					foldericon.removeClass(CLASS_FOLDER_OPEN).addClass(CLASS_FOLDER_CLOSE)
				} else {
					foldericon.removeClass(CLASS_FOLDER_CLOSE).addClass(CLASS_FOLDER_OPEN)
				}
			}
			this.updateDataById(this.list, id, { expand: !item.expand })
		}
		if (this.callback_item_click) {
			this.callback_item_click(id, item, e)
		}
	}
	// 开始拖拽
	onDragStart(e) {
		console.log('onDragStart', e)
		this.dragging = true
		e.originalEvent.dataTransfer.setData('dragdata', e.target.id)
		// e.originalEvent.stopPropagation()
		e.originalEvent.cancelBubble = true
	}

	onDragOver(e) {
		e.stopPropagation()
		e.originalEvent.cancelBubble = true
		e.originalEvent.stopPropagation()
		e.originalEvent.preventDefault()
		var $trigger = $(e.target)

		var parentNode = $trigger.hasClass(CLASS_ITEM)
			? e.target
			: $trigger.closest(`.${CLASS_ITEM}`)[0]
		if (parentNode) {
			var rect = this.rootElement[0].getBoundingClientRect()
			const node_h = parentNode.offsetHeight // 当前节点的高度
			const subsection = node_h / 3

			const client_y = e.clientY - rect.top - parentNode.offsetTop + 1 // 鼠标在当前节点的位置
			let dragAction = ''
			if (client_y < subsection) {
				// 向上移动
				dragAction = 'top'
			} else if (client_y > subsection * 2) {
				// 向下移动
				dragAction = 'bottom'
			} else {
				// 向里移动
				dragAction = 'inner'
			}
			this.updateOverClass(parentNode, dragAction)
		}
	}

	updateOverClass(target, direction) {
		if (this.drag_action) {
			if (target) {
				if (this.drag_action.id) {
					if (
						target.id == this.drag_action.id &&
						direction == this.drag_action.direction
					) {
						return
					}
					$(`#${this.drag_action.id}`).removeClass(
						'drag-action-' + this.drag_action.direction
					)
					$(`#${target.id}`).addClass('drag-action-' + direction)
					this.drag_action = {
						id: target.id,
						direction: direction,
					}
				}
				$(`#${this.drag_action.id}`).addClass(
					'drag-action-' + this.drag_action.direction
				)
			} else {
				if (this.drag_action.id) {
					$(`#${this.drag_action.id}`).removeClass(
						'drag-action-' + this.drag_action.direction
					)
				}
			}
		} else {
			if (target) {
				$(`#${target.id}`).addClass('drag-action-' + direction)
				this.drag_action = {
					id: target.id,
					direction: direction,
				}
			} else if (this.drag_action && this.direction.id) {
				$(`#${this.drag_action.id}`).removeClass(
					'drag-action-' + this.drag_action.direction
				)
			}
		}
	}

	onDrop(e) {
		this.dragging = false
		console.log('drop', e.target, this.drag_action)
		e.stopPropagation()
		e.originalEvent.cancelBubble = true
		e.originalEvent.preventDefault()
		var $trigger = $(e.target)
		var parentNode = $trigger.hasClass(CLASS_ITEM)
			? e.target
			: $trigger.closest(`.${CLASS_ITEM}`)[0]
		var dragId = e.originalEvent.dataTransfer.getData('dragdata')
		var res = {
			dragId: dragId,
		}
		if (parentNode) {
			var dropId = parentNode.id
			if (dragId == dropId) {
				console.log('dragId == dropId', dragId)
				return
			}
			var dragData = this.getDataById(this.list, dragId)
			if (!dragData) {
				console.log('dragData is null', this.list, dragId)
				return
			}
			var dropData = this.getDataById(this.list, dropId)
			if (dropData) {
				var dragElement = document.getElementById(dragId)
				this.removeItemByPos(this.list, dragData.pos, 0) // 先移除
				this.updateItemRender(
					null,
					dragData.pos.slice(0, dragData.pos.length - 1)
				) // 更新拖拽目标的原父亲节点渲染
				this.updatePos(this.list, [])
				
				dropData = this.getDataById(this.list, dropId)
				var dropPos = dropData.pos
				var direction = this.drag_action.direction
				var dropParent = this.getDataByPos(
					this.list,
					dropPos.slice(0, dropPos.length - 1),
					0
				)
				// var oldDragParentData = this.getDataByPos(this.list, dragData.pos.slice(0, dragData.pos.length - 1), 0)
				
				if (direction == 'top' || direction == 'bottom') {
					var position = dropPos[dropPos.length - 1]
					console.log('position', position)
					if (!dropParent && dropPos.length == 1) {
						if (direction == 'top') {
							this.list.splice(position, 0, dragData)
						} else if (direction == 'bottom') {
							this.list.splice(position + 1, 0, dragData)
						}
					} else if (dropParent) {
						if (!dropParent.children) {
							dropParent.children = []
						}
						dropParent.children.splice(direction == 'top' ? position : (position + 1), 0, dragData)
					}
					if (direction == 'top') {
						parentNode.before(dragElement)
					} else {
						parentNode.after(dragElement)
					}
				} else if (direction == 'inner') {
					var children = $(parentNode).find('.tree-children')
					if (!children || !children.length) {
						children = document.createElement('div')
						children.className = 'tree-children'
						parentNode.appendChild(children)
					} else if (children.length > 0) {
						children = children[0]
					}
					children.appendChild(dragElement)

					if (!dropData.children) {
						dropData.children = [dragData]
					} else {
						dropData.children.unshift(dragData)
					}
					this.updateItemRender(dropData.id, dropData.pos)
				} else {
					console.log('direction is null ', direction)
				}
				// if (oldDragParentData) {
				// 	this.updateItemRender(oldDragParentData.id) // 更新拖拽目标的原父亲节点渲染	
				// }
				
				this.updatePos(this.list, [])
				this.updateItemRender(dragData.id)
			} else {
				console.log('dropData is null')
			}
		} else {
			console.log('can not find parent')
		}
		this.updateOverClass()
		console.log('tree after drop', this.list)
		if (this.callback_drop) {
			this.callback_drop(this.list, dragId, dropId, this.drag_action.direction)
		}
	}

	removeItemByPos(list, pos, index) {
		if (!list || !pos) {
			return null
		}
		if (index >= pos.length || index < 0) {
			return null
		}
		var i = pos[index]
		if (i >= list.length || i < 0) {
			return null
		}
		if (index < pos.length - 1) {
			if (list[i].children) {
				this.removeItemByPos(list[i].children, pos, index + 1)
			}
		} else if (index == pos.length - 1) {
			list.splice(i, 1)
		}
	}

	removeDataById(list, id) {
		if (!list) {
			return null
		}
		for (var i = 0, imax = list.length; i < imax; i++) {
			if (list[i].id == id) {
				return list.splice(i, 1)
			}
			if (list[i].children) {
				var v = this.removeDataById(list[i].children, id)
				if (v) {
					if (list[i].children.length == 0) {
						list[i].children = []
					}
					return v
				}
			}
		}
		return null
	}
	// 通过ID获取数据
	getDataById(list, id) {
		if (!list) {
			return null
		}
		for (var i = 0, imax = list.length; i < imax; ++i) {
			if (list[i].id == id) {
				return list[i]
			}
			if (list[i].children) {
				var v = this.getDataById(list[i].children, id)
				if (v) {
					return v
				}
			}
		}
		return null
	}
	// 通过pos获取数据
	getDataByPos(list, pos, index) {
		if (!list) {
			return null
		}
		var i = pos[index]
		if (i >= list.length || i < 0) {
			return null
		}
		if (index < pos.length - 1) {
			return this.getDataByPos(list[i].children, pos, index + 1)
		} else if (index == pos.length - 1) {
			return list[i]
		}
		return null
	}

	updatePos(list, pos) {
		if (!list) {
			return
		}
		list.forEach((item, index) => {
			var pos2 = [].concat(pos)
			pos2.push(index)
			item.pos = pos2
			if (item.children && item.children.length > 0) {
				this.updatePos(item.children, pos2)
			}
		})
	}
	createItem(item, parent, position = 'afterend', index = 0) {
		const div = $('<div></div>')
		const innerDiv = $('<div></div>')
		var html = ''
		html += `<div class=${CLASS_INNER} id="${item.id}_inner">`
		html += '<span>'
		html += `<i class="iconfont ${
			item.expand ? CLASS_EXPAND_CLOSE : CLASS_EXPAND_OPEN
		} ${CLASS_EXPAND}"></i>`
		if (item.is_folder) {
			html += `<i class="iconfont ${
				item.expand ? CLASS_FOLDER_OPEN : CLASS_FOLDER_CLOSE
			} ${CLASS_FOLDER}"></i>`
		}
		html += '</span>'
		if (!item.regionType) {
			html += `<span class='${CLASS_TEXT} untyped'  title="[${item.id}]:${item.label}">(未分类)${item.label}</span>`
		} else {
			html += `<span class=${CLASS_TEXT} title="[${item.id}]:${item.label}">${item.label}</span>`
		}
		html += '</div>'
		innerDiv.html(html)
		div.addClass(CLASS_ITEM).attr('id', item.id).attr('draggable', true)
		div.append(innerDiv)
		div.html(html)
		var textEl = div.find('.text')
		if (item.children && item.children.length > 0) {
			const children = $('<div></div>')
			children.addClass(CLASS_CHILDREN)
			item.children.forEach((child) => {
				this.createItem(child, children)
			})
			div.append(children)
			if (!item.expand) {
				children.hide()
			}

			if (textEl) {
				textEl.css('padding-inline-start', '0px')
			}
		} else {
			div.find(`.${CLASS_EXPAND}`).hide()
			if (item.pos && item.pos.length > 1 && textEl) {
				textEl.css('padding-inline-start', '12px')
			}
		}
		div.on('dragstart', (e) => {
			this.onDragStart(e)
		})
		if (position == 'afterend') {
			parent.append(div)
		} else if (position == 'beforebegin') {
			parent.insertAdjacentElement('beforebegin', div)
		} else if (position == 'center') {
			if (parent.children && index >= 0 && index < parent.children.length) {
				parent.insertBefore(div, parent.children[index])
			} else {
				parent.append(div)
			}
		}
		div.find(`#${item.id}_inner`).on('click', (e) => {
			e.stopPropagation()
			e.cancelBubble = true
			e.originalEvent.cancelBubble = true
			this.onItemClick(e, item.id, item)
		})
	}
	// 更新item的渲染
	updateItemRender(id, pos) {
		var data
		if (id) {
			data = this.getDataById(this.list, id)
		} else if (pos) {
			data = this.getDataByPos(this.list, pos, 0)
		}
		var targetNode = data ? $(`#${data.id}`) : id ? $(`#${id}`) : null
		if (!data) {
			if (targetNode) targetNode.remove()
			return
		}

		if (!targetNode) {
			var parent = this.getDataByPos(this.list, pos.slice(0, pos.length - 1), 0)
			if (parent) {
				this.createItem(data, $('#' + parent.id), 'center', pos[pos.length - 1])
			}
			return
		}
		var innerNode = targetNode.find(`.${CLASS_INNER}`).eq(0)
		if (!innerNode) {
			return
		}
		var expandNode = innerNode.find(`.${CLASS_EXPAND}`).eq(0)
		if (data.children && data.children.length > 0) {
			expandNode.show()
			if (data.expand) {
				if (expandNode.hasClass(CLASS_EXPAND_CLOSE)) {
					expandNode.removeClass(CLASS_EXPAND_CLOSE)
				}
				if (!expandNode.hasClass(CLASS_EXPAND_OPEN)) {
					expandNode.addClass(CLASS_EXPAND_OPEN)
				}
			} else {
				if (expandNode.hasClass(CLASS_EXPAND_OPEN)) {
					expandNode.removeClass(CLASS_EXPAND_OPEN)
				}
				if (!expandNode.hasClass(CLASS_EXPAND_CLOSE)) {
					expandNode.addClass(CLASS_EXPAND_CLOSE)
				}
			}
			targetNode.find('.text').eq(0).css('padding-inline-start', '0px')
		} else {
			expandNode.hide()
			targetNode
				.find('.text')
				.eq(0)
				.css(
					'padding-inline-start',
					data.pos && data.pos.length == 1 ? '0' : '12px'
				)
		}
		var folderNode = innerNode.find(`.${CLASS_FOLDER}`).eq(0)
		if (data.is_folder) {
			folderNode.show()
		} else {
			folderNode.hide()
		}
	}

	setSelect(idList) {
		if (this.select_id_list) {
			this.select_id_list.forEach((e) => {
				$(`#${e}_inner`).removeClass('selected')
			})
		}
		if (idList) {
			idList.forEach((e) => {
				$(`#${e}_inner`).addClass('selected')
			})
		}
		this.select_id_list = idList
	}
}

export default Tree
