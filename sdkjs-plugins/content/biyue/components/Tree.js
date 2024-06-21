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
  }

  init(list) {
    this.updatePos(list, [])
    this.list = list
    this.buildTree(list, this.rootElement)
    this.rootElement.on('dragover', e => {
      this.onDragOver(e)
    })
    this.rootElement.on('drop', e => {
      this.onDrop(e)
    })
    console.log('list', list)
  }

  buildTree(data, parent) {
    if (!data || !parent) {
      return
    }
    data.forEach((item) => {
      this.createItem(item, parent)
    })
  }

  refreshList(list) {
    this.list = list
  }

  update() {

  }

  render() {

  }

  updateDataById(list, id, options) {
    for (var i = 0, imax =  list.length; i < imax; ++i) {
      if (list[i].id == id) {
        Object.keys(options).forEach(key => {
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
      Object.keys(options).forEach(key => {
        list[i][key] = options[key]  
      })
    }
  }

  onItemClick(e, id, item) {
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
        var foldericon = $(`#${id}`).find(`.${CLASS_INNER}`).find(`.${CLASS_FOLDER}`)
        if (item.expand) {
          foldericon.removeClass(CLASS_FOLDER_OPEN).addClass(CLASS_FOLDER_CLOSE)
        } else {
          foldericon.removeClass(CLASS_FOLDER_CLOSE).addClass(CLASS_FOLDER_OPEN)
        }
      }
      this.updateDataById(this.list, id, {expand: !item.expand})
    }
  }
  // 开始拖拽
  onDragStart(e) {
    console.log('onDragStart', e.target.id)
    e.originalEvent.dataTransfer.setData('dragdata', e.target.id)
  }

  onDragOver(e) {
    e.stopPropagation()
    e.originalEvent.cancelBubble = true
    e.originalEvent.preventDefault()
    var $trigger = $(e.target)
    
    var parentNode = $trigger.hasClass(CLASS_ITEM) ? e.target : $trigger.closest(`.${CLASS_ITEM}`)[0]
    if (parentNode) {
      var inner = parentNode.querySelector(`.${CLASS_INNER}`)
      var rect = inner.getBoundingClientRect()
      // console.log('rect', rect, 'node_h', parentNode.offsetHeight, parentNode.offsetTop, e.clientY)
      const node_h = inner.offsetHeight // 当前节点的高度
      const subsection = node_h / 3
      const client_y = e.clientY - rect.top - inner.offsetTop + 1// 鼠标在当前节点的位置
      // console.log('rect', rect, parentNode, e.clientY, e, 'inner', inner.offsetHeight, inner.offsetTop)
      console.log('client_y', client_y, 'subsection', subsection, 'node_h', node_h)
      let dragAction = ''
      if (client_y < subsection) {
        // 向上移动
        dragAction = 'top'
        console.log('向上移动')
      } else if (client_y > subsection * 2) {
        // 向下移动
        dragAction = 'bottom'
        console.log('向下移动')
      } else {
        // 向里移动
        dragAction = 'inner'
        console.log('向里移动')
      }
      this.updateOverClass(parentNode, dragAction)
    }
    
    // console.log('onDragOver', e.target.id, e.target.parentNode)
  }

  updateOverClass(target, direction) {
    if (this.drag_action) {
      if (target) {
        if (this.drag_action.id) {
          if (target.id == this.drag_action.id && direction == this.drag_action.direction) {
            return
          }
          $(`#${this.drag_action.id}`).removeClass('drag-action-' + this.drag_action.direction)
          $(`#${target.id}`).addClass('drag-action-' + direction)
          this.drag_action = {
            id: target.id,
            direction: direction
          }
        }
        $(`#${this.drag_action.id}`).addClass('drag-action-' + this.drag_action.direction)
      } else {
        if (this.drag_action.id) {
          $(`#${this.drag_action.id}`).removeClass('drag-action-' + this.drag_action.direction)
        }
      }
    } else {
      if (target) {
        $(`#${target.id}`).addClass('drag-action-' + direction)
        this.drag_action = {
          id: target.id,
          direction: direction
        }
      } else if (this.drag_action && this.direction.id) {
        $(`#${this.drag_action.id}`).removeClass('drag-action-' + this.drag_action.direction)
      }
    }
  }
  
  onDrop(e) {
    e.stopPropagation()
    e.originalEvent.cancelBubble = true
    e.originalEvent.preventDefault()
    var $trigger = $(e.target)
    var parentNode = $trigger.hasClass(CLASS_ITEM) ? e.target : $trigger.closest(`.${CLASS_ITEM}`)[0]
    var dragId = e.originalEvent.dataTransfer.getData('dragdata')
    if  (parentNode) {
      var dropId = parentNode.id
      if (dragId == dropId) {
        return
      }
      // 先移除
      var dragData = this.getDataById(this.list, dragId)
      this.removeItemByPos(this.list, dragData.pos, 0)

      var dropData = this.getDataById(this.list, dropId)
      if (dropData) {
        var dropPos = dropData.pos
        var direction = this.drag_action.direction
        var dropParent = this.getDataByPos(this.list, ([].concat(dropPos).splice(dropPos.length - 1, 1)), 0)
        if (direction == 'top') {
          dropParent.children.slice(dropPos[dropPos.length - 1], 0, dragData)
        } else if (direction == 'bottom') {

        } else if (direction == 'inner') {

        }
      }

      // var targetPos = this.getPosForDrop(dropId, this.drag_action.direction)
      // var targetParent = this.addDataToPos(targetPos, dragData)

      // // var dragData = this.removeDataById(this.list, dragId)
      // // 再插入
      // // this.addDataToId(this.list, dropId, dragData, 'top', dragId)
      // this.updatePos(this.list, [])
    }
    this.updateOverClass()
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
            list[i].is_leaf = true
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
  createItem(item, parent) {
    const div = $('<div></div>')
    const innerDiv = $('<div></div>')
    var html = ''
    html += `<div class=${CLASS_INNER} id="${item.id}_inner">`
    html += '<span>'
    html += `<i class="iconfont ${item.expand ? CLASS_EXPAND_CLOSE : CLASS_EXPAND_OPEN} ${CLASS_EXPAND}"></i>`
    if (item.is_folder) {
      html += `<i class="iconfont ${item.expand ? CLASS_FOLDER_OPEN : CLASS_FOLDER_CLOSE} ${CLASS_FOLDER}"></i>`
    }
    html += '</span>'
    html += `<span class=${CLASS_TEXT} title="${item.label}">${item.label}</span>`
    html += '</div>'
    innerDiv.html(html)
    div.addClass(CLASS_ITEM).attr('id', item.id).attr('draggable', true)
    div.append(innerDiv)
    div.html(html)
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
    }
    div.on('dragstart', e => {
      this.onDragStart(e)
    })
    parent.append(div)
    div.find(`#${item.id}_inner`).on('click', e => {
      e.stopPropagation()
      e.cancelBubble = true
      e.originalEvent.cancelBubble = true
      this.onItemClick(e, item.id, item)
    })
  }
  // 更新item的渲染
  updateItemRender(id) {
    var data = this.getDataById(this.list, id)
    var targetNode = $(`#${id}`)
    if (!data) {
      if (targetNode)
      targetNode.remove()
      return
    }
    var innerNode = targetNode.find(`.${CLASS_INNER}`)
    if (!innerNode) {
      return
    }
    var expandNode = innerNode.find(`.${CLASS_EXPAND}`)
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
    } else {
      expandNode.hide()
    }
    var folderNode = innerNode.find(`.${CLASS_FOLDER}`)
    if (data.is_folder) {
      folderNode.show()

    } else {
      folderNode.hide()
    }
  }

  // 获取要置入的位置
  getPosForDrop(dropId, direction) {
    var data = this.getDataById(this.list, dropId)
    if (!data) {
      return null
    }
    var pos = data.pos
    var targetPos = [].concat(pos)
    if (direction == 'top') {
      return targetPos
    } else if (direction == 'bottom') {
      targetPos[targetPos.length - 1] + 1
    } else if (direction == 'inner') {
      targetPos.push(0)
    }
    return targetPos
  }

  // 在Pos插入数据
  addDataToPos(pos, data) {
    var pos1 = [].concat(pos)
    var index = pos1.splice(pos.length - 1, 1)
    data.pos = pos
    var parentItem = this.getDataByPos(this.list, pos1, 0)
    if (parentItem) {
      parentItem.children.slice(index, 0, data)
      return parentItem
    } else {
      return null
    }
  }

  // 渲染插入数据
  addRender(pos, draggedItem, targetParent) {
    if (targetParent) {

    } else {
      if (pos[pos.length - 1] == this.list.length - 1) {
        this.rootElement.appendChild(draggedItem)
      } else {

        this.rootElement.insertBefore(draggedItem, this.rootElement.children[pos[pos.length - 1] + 1])
      }
    }
  }

  addDataToId(list, id, data, direction, addNodeId) {
    if (!list) {
      return
    }
    for (var i = 0, imax = list.length; i < imax; i++) {
      if (list[i].id == id) {
        var originalNode = document.getElementById(addNodeId)
        var cloneNode = originalNode.cloneNode(true)
        var targetNode = document.getElementById(list[i].id)
        if (direction == 'top') {
          list.slice(i, 0, data)
          targetNode.before(cloneNode)
        } else if (direction == 'bottom') {
          list.slice(i + 1, 0, data)
          targetNode.after(cloneNode)
        } else if (direction == 'inner') {
          if (!list[i].children) {
            list[i].children = [data]
            list[i].is_leaf = false
            targetNode.append(cloneNode)
          } else {
            targetNode = document.getElementById(list[i].children[0].id)
            targetNode.before(cloneNode)
            list[i].children.slice(0, 0, data)
          }
        }
        originalNode.remove()
        return
      }
      if (list[i].children) {
        this.addDataToId(list[i].children, id, data, direction, addNodeId)
      }
    }
  }
}

export default Tree