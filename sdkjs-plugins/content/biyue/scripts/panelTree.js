import { preGetExamTree, focusControl } from './QuesManager.js'
import { showCom } from './model/util.js'
import { handleRangeType } from "./classifiedTypes.js"
function generateTree() {
	return preGetExamTree().then(res => {
		Asc.scope.tree_info = res
		renderTree()
	})
}

function refreshTree() {
	if (window.tab_select == 'tabTree') {
		return generateTree()
	} else {
		return Promise.resolve()
	}
}

function renderTreeNode(parent, item, identation = 0) {
	if (!parent) {
		return
	}
	var question_map = window.BiyueCustomData.question_map || {}
	var quesData = question_map[item.id]
	if (!quesData) {
		return
	}
	var html = ''
	if (item.level_type == 'struct') {
		html += `<div class="row-align-center" style="padding-left:${identation}px" id="group-${item.id}">
					<div class="struct font-12">构</div>
					<div class="itemques text-over-ellipsis flex-1 clicked" style="margin-left: 4px;" title="${quesData.text}">${quesData.text}</div>
				</div>`
	} else if (item.level_type == 'question') {
		html += `<div class="itemques" style="margin-left: ${identation}px;">
					<div title="${quesData.text}" id="ques-${item.id}" class="text-over-ellipsis clicked flex-1">${quesData.text}</div>
					<div class="children" id="ques-${item.id}-children"></div>
				</div>`
	}
	parent.append(html)
	if (item.children && item.children.length > 0) {
		identation += 24
		for (var child of item.children) {
			if (item.level_type == 'struct') {
				renderTreeNode(parent, child, identation)
			} else {
				renderTreeNode($(`#ques-${item.id}-children`), child, 16)
			}
		}
	}
}

function renderTree() {
	var tree_info = Asc.scope.tree_info || {}
	var rootElement = $('#tree')
	rootElement.empty()
	if (tree_info.tree && tree_info.tree.length) {
		showCom('#panelTree .none', false)
		tree_info.tree.forEach(item => {
			renderTreeNode(rootElement, item, 0)
		})
		tree_info.list.forEach(item => {
			var com = $(`#${item.level_type == 'question' ? 'ques' : 'group'}-${item.id}`)
			if (com) {
				function clickHandler() {
					clickTreeItem(item.id)
				}
				com.off('click', clickHandler)
				com.on('click', clickHandler)

				function contextmenuHandler(event) {
					event.preventDefault()
					var quesData = window.BiyueCustomData.question_map[item.id]
					if (quesData.level_type == 'question') {
						var nodeData = window.BiyueCustomData.node_list.find(e => {
							return e.id == item.id
						})
						if (nodeData) {
							if (nodeData.is_big) {
								generateMenuItems(['clearBig'], item.id); // 生成动态菜单
							} else {
								generateMenuItems(['setBig'], item.id); // 生成动态菜单
							}
						}
					}
					
					$('#dynamicMenu').css({
						display: 'block',
						left: event.pageX,
						top: event.pageY
					});
				}
				com.off('contextmenu', contextmenuHandler)
				com.on('contextmenu', contextmenuHandler)
			}
		})
	} else {
		showCom('#panelTree .none', true)
	}
}

function generateMenuItems(options, id) {
	const menuContent = $('#menuContent');
	menuContent.empty(); // 清除旧的菜单项
	options.forEach(e => {
		var name = ''
		switch(e) {
			case 'setBig': 
				name = '设置 - 大题'
				break
			case 'clearBig':
				name = '清除 - 大题'
				break
			default:
				break
		}
		if (name) {
			//menuContent.append('<li onclick="alert(\'操作1-1\')">动态操作1-1</li>');
			// 使用匿名函数包裹 clickMenu 函数并传递参数
			menuContent.append(`<li>${name}</li>`).children().last().on('click', function() {
				clickMenu(id, e);
			});
		}
	})
}

function clickTreeItem(id) {
	focusControl(id)
}

function clickMenu(id, cmd) {
	console.log('clickMenu', id, cmd)
	if (cmd == 'setBig' || cmd == 'clearBig') {
		focusControl(id).then(res => {
			handleRangeType({
				typeName: cmd
			})
		})
	}
}

export {
	generateTree,
	refreshTree
}