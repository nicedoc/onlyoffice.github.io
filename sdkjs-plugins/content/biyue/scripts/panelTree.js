import { preGetExamTree, focusControl } from './QuesManager.js'
import { showCom, updateText, addClickEvent } from './model/util.js'
import { handleRangeType } from "./classifiedTypes.js"
import { getDataByParams } from '../scripts/model/ques.js'
var g_tree_info = {}
var big_info = {
	big_id: 0,
	visible_big_set: false,
	child_ids: [],
	end_id: 0
}
function generateTree() {
	return preGetExamTree().then(res => {
		Asc.scope.tree_info = res
		g_tree_info = res || {}
		console.log('=========generateTree ', g_tree_info)
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

function renderTreeNode(parent, item, parentData) {
	if (!parent) {
		return
	}
	var question_map = window.BiyueCustomData.question_map || {}
	var quesData = question_map[item.id]
	if (!quesData) {
		return
	}
	var html = ''
	var identation = 0
	var offset = 24
	if (parentData) {
		if (parentData.is_big) {
			if (item.lvl != null && parentData.lvl !== null) {
				identation = 20 + offset * (item.lvl - parentData.lvl - 1)
			} else {
				identation = 20
			}
		} else if (item.lvl === null) {
			if (parentData.lvl) {
				identation = offset * (parentData.lvl + 1)
			} else {
				identation = offset
			}
		} else {
			identation = (item.level_type =='struct' ? 10 : offset) * item.lvl
		}
	} else {
		if (item.lvl != null) {
			identation = item.lvl * offset
		} else {
			identation = 0
		}
	}
	if (item.level_type == 'struct') {
		html += `<div class="row-align-center" id="group-${item.id}">
					<div class="struct font-12">构</div>
					<div class="itemques text-over-ellipsis flex-1 clicked" id="box-${item.id}" style="margin-left: ${identation}px;" title="${quesData.text}">${quesData.text}</div>
				</div>`
	} else if (item.level_type == 'question') {
		html += `<div class="itemques" id="box-${item.id}"  style="margin-left: ${identation}px;">
					<div title="${quesData.text}" id="ques-${item.id}" class="text-over-ellipsis clicked flex-1">${quesData.text}</div>
					<div class="children" id="ques-${item.id}-children"></div>
				</div>`
	}
	parent.append(html)
	if (item.children && item.children.length > 0) {
		for (var child of item.children) {
			if (item.level_type == 'struct') {
				renderTreeNode(parent, child, item)
			} else {
				renderTreeNode($(`#ques-${item.id}-children`), child, item)
			}
		}
	}
}

function renderTree() {
	var tree_info = g_tree_info
	big_info = null
	showCom('#panelTree #bigbox', false)
	var rootElement = $('#tree')
	rootElement.empty()
	if (tree_info.tree && tree_info.tree.length) {
		showCom('#panelTree .none', false)
		tree_info.tree.forEach(item => {
			renderTreeNode(rootElement, item, null)
		})
		var structNum = 0
		var quesNum = 0
		tree_info.list.forEach(item => {
			if (item.level_type == 'struct') {
				structNum++
			} else if (item.level_type == 'question') {
				quesNum++
			}
			var com = $(`#panelTree #${item.level_type == 'question' ? 'ques' : 'group'}-${item.id}`)
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
						if (nodeData && !quesData.is_merge) { // 合并题不可设置为大题
							if (nodeData.is_big) {
								generateMenuItems(['clearBig', 'extendBig'], item.id); // 生成动态菜单
							} else {
								generateMenuItems(['setBig', 'setBig2'], item.id); // 生成动态菜单
							}
							updateMenuPos(event)
						}
					} else if (quesData.level_type == 'struct') {
						generateMenuItems(['question'], item.id)
						updateMenuPos(event)
					}
				}
				com.off('contextmenu', contextmenuHandler)
				com.on('contextmenu', contextmenuHandler)
			}
		})
		updateText('#panelTree #sum', `总计：结构${structNum}个，题目${quesNum}个`)
	} else {
		showCom('#panelTree .none', true)
		updateText('#panelTree #sum', '')
	}
}

function updateMenuPos(event) {
	const menu = document.getElementById('dynamicMenu');
	// 获取浏览器窗口的宽度、高度
    const windowWidth = window.innerWidth;
    const windowHeight = window.innerHeight;
	var x = event.pageX
	var y = event.pageY
	// 设置菜单初始位置
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;

    // 检查菜单是否会超出窗口边界
    if (x + menu.offsetWidth > windowWidth) {
        menu.style.left = `${windowWidth - menu.offsetWidth}px`;
    }
    if (y + menu.offsetHeight > windowHeight) {
        menu.style.top = `${windowHeight - menu.offsetHeight}px`;
    }

    // 显示菜单
    menu.style.display = 'block';
	window.show_dynamic_menu = true
}

function clickTreeLock() {
	window.tree_lock = !(window.tree_lock)
	var com = $('#panelTree #lock .iconfont')
	if (!com) {
		return
	}
	if (window.tree_lock) {
		com.removeClass('icon-dingzi-u')
		com.addClass('icon-dingzi1')
	} else {
		com.addClass('icon-dingzi-u')
		com.removeClass('icon-dingzi1')
	}
}

function updateBig(id) {
	if (big_info && big_info.visible_big_set) {
		resetBig()
	}
	var childrenName = `#panelTree #ques-${id}-children`
	var children = $(childrenName)
	if (!children.length) {
		return
	}
	var index = g_tree_info.list.findIndex(e => {
		return e.id == id
	})
	if (index == -1) {
		return
	}
	var lvl = g_tree_info.list[index].lvl
	var i = index + 1
	var toIndex = -1
	for (; i < g_tree_info.list.length; ++i) {
		if (g_tree_info.list[i].level_type == 'struct') {
			toIndex = i - 1
			break
		}
		if (g_tree_info.list[i].lvl == 0) {
			toIndex = i - 1
			break
		}
		if (lvl === null) {
			continue
		} else if (g_tree_info.list[i].lvl == null) {
			continue
		} else if (lvl >= g_tree_info.list[i].lvl) {
			toIndex = i - 1
			break	
		}
	}
	if (toIndex > index) {
		big_info = {
			big_id: id,
			visible_big_set: true,
			child_ids: [],
			end_id: g_tree_info.list[index].end_id || 0,
			lock: window.tree_lock
		}
		if (!window.tree_lock) {
			clickTreeLock()
		}
		var bigques = $(`#panelTree #box-${id}`)
		bigques.css({
			'background-color': '#fff',
			'border': '1px solid #bbb'
		})
		var mleft = parseInt((bigques).css('margin-left'), 10);
		var minMarginLeft = 100
		for (var j = index + 1; j <= toIndex; ++j) {
			var item = $(`#panelTree #box-${g_tree_info.list[j].id}`)
			var parentElement = item.parent();
			var prevElement = item.prev();
			var marginLeft = parseInt((item).css('margin-left'), 10);
			big_info.child_ids.push({
				id: g_tree_info.list[j].id,
				parent: parentElement ? parentElement.attr('id') : null,
				prev: prevElement ? prevElement.attr('id') : null,
				marginLeft: marginLeft
			})
			item.css({
				'background-color': '#fff',
				'border': '1px solid #bbb'
			});
			minMarginLeft = Math.min(minMarginLeft, marginLeft)
			item.appendTo(childrenName);
			if (g_tree_info.list[j].parent_id && g_tree_info.list[j].parent_id != id) {
				var parentData = g_tree_info.list.find(e => {
					return e.id == g_tree_info.list[j].parent_id
				})
				if (parentData && parentData.level_type == 'question') {
					item.hide()
				}
			}
		}
		if (minMarginLeft > mleft) {
			var offset = minMarginLeft - mleft
			big_info.child_ids.forEach(e => {
				$(`#panelTree #box-${e.id}`).css({
					'margin-left': `${e.marginLeft - offset}px`
				})
			})
		}
		if (big_info.end_id) {
			$(`#panelTree #box-${big_info.end_id}`).css({
				'border': '1px solid #95c8ff',
				'background-color': '#deedfe',
				'color': '#2489f6'
			})
		}
		showBigBox(childrenName)
		updateSelect(0)
	}
}

function showBigBox(childrenName) {
	var children = $(childrenName)
	children.append(`<div id="bigbox" class="row-between">
				<div>tip: 请点击截止题目</div>
				<div>
					<span id="bigcancel" class="clicked">取消</span>
					<span id="bigconfirm" class="clicked">确定</span>
				</div>
			</div>`)
	showCom('#panelTree #bigbox', true)
	showCom('#panelTree #bigconfirm', big_info.end_id != 0)
	addClickEvent('#panelTree #bigcancel', onBigCancel)
	addClickEvent('#panelTree #bigconfirm', onBigConfirm)
}

function onBigCancel(e) {
	e.cancelBubble = true
	e.preventDefault()
	resetBig()
}

function onBigConfirm(e) {
	e.cancelBubble = true
	e.preventDefault()
	if (big_info) {
		handleRangeType({
			typeName: 'setBig',
			big_id: big_info.big_id,
			end_id: big_info.end_id
		})
		resetBig()
	}
}

function resetBig() {
	if (!big_info) {
		return
	}
	if (!big_info.lock) {
		clickTreeLock()
	}
	renderTree()
	big_info = null
}

function generateMenuItems(options, id) {
	const menuContent = $('#menuContent');
	menuContent.empty(); // 清除旧的菜单项
	options.forEach(e => {
		var name = ''
		switch(e) {
			case 'setBig': 
				name = '构建大小题 - 编号'
				break
			case 'setBig2':
				name = '构建大小题 - 手动'
				break
			case 'clearBig':
				name = '清除 - 大题'
				break
			case 'extendBig':
				name = '扩展大小题'
				break
			case 'question':
				name = '设置为 - 题目'
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
	if (big_info && big_info.visible_big_set) {
		updateSelect(0)
		if (big_info.end_id && big_info.end_id != id) {
			$(`#panelTree #box-${big_info.end_id}`).css({
				'background-color': '#fff',
				'border': '1px solid #bbb',
				'color': '#444'
			})
		}
		if (big_info.child_ids.findIndex(e => {
			return e.id == id
		}) >= 0) {
			big_info.end_id = id
			$(`#panelTree #box-${big_info.end_id}`).css({
				'border': '1px solid #95c8ff',
				'background-color': '#deedfe',
				'color': '#2489f6'
			})
			showCom('#panelTree #bigconfirm', true)
		}
	} else {
		updateSelect(id)
	}
}

function clickMenu(id, cmd) {
	if (cmd == 'setBig' || cmd == 'clearBig' || cmd == 'question') {
		focusControl(id).then(res => {
			handleRangeType({
				typeName: cmd
			})
		})
	} else if (cmd == 'setBig2') {
		updateBig(id)
	} else if (cmd == 'extendBig') {
		updateBig(id)
	}
}

function updateTreeSelect(params) {
	var data = getDataByParams(params)
	if (!data || !data.data) {
		return
	}
	var qid = data.ques_client_id || data.client_id
	updateSelect(qid, true)
}

function updateSelect(qid, updateScroll) {
	var oldSelected = $('#panelTree #tree .selected')
	if (oldSelected) {
		oldSelected.removeClass('selected')
	}
	var $target = $(`#panelTree #box-${qid}`)
	if ($target && $target.length) {
		$target.addClass('selected')
		if (updateScroll) {
			var $container = $('#panelTree #tree')
			if ($container.length) {
				$container.animate({
					scrollTop: $target.offset().top - $container.offset().top + $container.scrollTop()
				}, 500); // 500 is the duration of the animation in milliseconds
			}
		}
	}
}

export {
	generateTree,
	refreshTree,
	updateTreeSelect,
	clickTreeLock
}