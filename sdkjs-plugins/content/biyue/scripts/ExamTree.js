// 主要用于管理试卷结构树
import Tree from "../components/Tree.js";

var g_exam_tree = null
var g_horizontal_list = []
// 构建试卷结构
function initTree(list) {
  const { horizontal_list, tree } = generateList(list)
  window.BiyueCustomData.exam_tree = tree
  g_horizontal_list = horizontal_list
  if (g_exam_tree) {
    g_exam_tree.refreshList(tree)
  } else {
    g_exam_tree = new Tree($('#treeRoot'))
    g_exam_tree.addCallBack(clickItem, dropItem)
    g_exam_tree.init(tree)
  }
}

function generateList(list) {
  var control_list = (list ? list : window.BiyueCustomData.control_list) || []
  var list = []
  var struct_index = -1
  var question_index = -1
  var horizontal_list = [] // 横向结构
  control_list.forEach((e, index) => {
    var obj = {
      id: e.control_id,
      label: e.text,
      expand: true,
      is_folder: false,
      children: []
    }
    if (e.regionType == 'struct') {
      obj.is_folder = true
      obj.pos = [list.length]
      list.push(obj)
      struct_index = list.length - 1
    } else if (e.regionType == 'question') {
      if (struct_index != -1) {
        obj.pos = [struct_index, list[struct_index].children.length]
        list[struct_index].children.push(obj)
        question_index = list[struct_index].children.length - 1
      } else {
        obj.pos = [list.length]
        list.push(obj)
        struct_index = list.length - 1
      }
    } else if (e.regionType == 'sub-question') {
      if (struct_index != -1 && question_index != -1) {
        obj.pos = [struct_index, question_index, list[struct_index].children[question_index].children.length]
        list[struct_index].children[question_index].children.push(obj)
      } else if (struct_index != -1) {
        obj.pos = [struct_index, list[struct_index].children.length]
        list[struct_index].children.push(obj)
      } else if (question_index != -1) {
        obj.pos = [list.length]
        list.push(obj)
      }
    }
    horizontal_list.push({
      id: obj.id,
      label: obj.label,
      pos: obj.pos
    })
  })
  return {
    horizontal_list: horizontal_list,
    tree: list
  }
}

function clickItem(id, item, e) {
  if (g_exam_tree) {
    g_exam_tree.setSelect([id])
  }
  console.log('clickItem', id, item)
  var control_list = window.BiyueCustomData.control_list || []
  var controlData = control_list.find(item => {
    return item.control_id == id
  })
  if (!controlData) {
    console.log('onQuesTreeClick cannot find ', id)
    return
  }
  Asc.scope.click_ids = [id]
  window.Asc.plugin.callCommand(function() {
    var ids = Asc.scope.click_ids
    var oDocument = Api.GetDocument()
    oDocument.RemoveSelection()
    var controls = oDocument.GetAllContentControls()
    var firstRange = null
    ids.forEach((id, index) => {
      var control = controls.find(e => {
        return e.Sdt.GetId() == id
      })
      if (control) {
        if (index == 0) {
          firstRange = control.GetRange()
        } else {
          var oRange = control.GetRange()
          firstRange = firstRange.ExpandTo(oRange)
        }
      }
    })
    firstRange.Select()
  }, false, false, undefined)
}

function dropItem(list) {
  console.log('dropItem', list)
  window.BiyueCustomData.exam_tree = list
  var hlist = []
  updateHorizontalList(list, hlist)
  g_horizontal_list = hlist
}

function updateHorizontalList(list, hlist) {
  if (!list) {
    return
  }
  list.forEach(e => {
    hlist.push({
      id: e.id,
      label: e.label,
      pos: e.pos
    })
    if (e.children) {
      updateHorizontalList(e.children, hlist)
    }
  })
}

function refreshExamTree() {
  window.Asc.plugin.callCommand(function() {
    var controls = Api.GetDocument().GetAllContentControls()
    var list = []
    controls.forEach(control => {
      let tagInfo = JSON.parse(control.GetTag())
      var text = control.GetRange().GetText()
      list.push({
        control_id: control.Sdt.GetId(),
        text: text,
        regionType: tagInfo.regionType
      })
    })
    return list
  }, false, false, function(list) {
    syncTreeWhenUpdate(list)
  })
}
// 同步试卷结构当右边文档更新时
function syncTreeWhenUpdate(list) {
  const { horizontal_list, tree } = generateList(list)
  // var oldList = [].concat(g_horizontal_list)
  // var updateList = []
  // for (var i = 0, imax = oldList.length; i < imax; ++i) {
  //   var index = horizontal_list.findIndex(item => {
  //     return item.id == oldList[i].id
  //   })
  //   if (index >= 0) {

  //     break
  //   }
  // }

  // for (var i = 0, imax = horizontal_list.length; i < imax; ++i) {
  //   var index = oldList.findIndex(item => {
  //     return item.id == horizontal_list[i].id
  //   })
  //   if (index == -1) {
  //     oldList.splice(i, 0, horizontal_list[i])
  //   }
  // }
  g_horizontal_list = horizontal_list
}

function updateTreeRenderWhenClick(data) {
  if (g_exam_tree) {
    g_exam_tree.setSelect([data.control_id])
  }
}

export {
  initTree,
  refreshExamTree,
  updateTreeRenderWhenClick
}