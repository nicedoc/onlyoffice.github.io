import { preGetExamTree, focusControl, setNumberingLevel } from './QuesManager.js'
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
	if (window.BiyueCustomData.page_type * 1) {
		showCom('#panelTree .none', false)
		updateText('#panelTree #sum', '')
		showCom('#panelTree #lock', false)
		showCom('#introPageWrapper', true)
		return
	}
	showCom('#introPageWrapper', false)
	showCom('#panel #lock', true)
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

function initTreeListener() {
	document.addEventListener('focusQuestion', (params) => {
		if (window.tab_select != 'tabTree') return
		if (params.detail && params.detail.ques_id) {
			updateSelect(params.detail.ques_id, true)
		}
	})
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
	var offset = 16
	if (parentData) {
		if (parentData.is_big) {
			if (item.lvl != null && parentData.lvl !== null) {
				identation = 16 * (item.lvl - parentData.lvl - 1)
			} else {
				identation = 0
			}
		} else if (item.lvl === null) {
			if (parentData.lvl) {
				identation = offset * (parentData.lvl + 1)
			} else {
				identation = 0
			}
		} else {
			if (parentData.lvl != null) {
				identation = 16 * (item.lvl - parentData.lvl - 1)
			} else {
				identation = (item.lvl - 1) * 16 // 0
			}
		}
	} else {
		if (item.lvl != null) {
			identation = item.lvl * offset
		} else {
			identation = 0
		}
	}
	if (identation < 0) {
		identation = 0
	}
	if (item.level_type == 'struct') {
		var vlineleft = 0 // (identation > 16 ? (identation - 16) : 0)
		html += `<div class="qwrapper"  style="margin-left: ${identation}px;">
					<div id="group-${item.id}">
						<div class="struct font-12" style="left:${-24 - identation}px" id="struct${item.id}">构</div>
						<div class="itemques">
							<div class="text-over-ellipsis flex-1 clicked" id="box-${item.id}" title="${quesData.text}">${quesData.ques_name || quesData.text}</div>
						</div>
					</div>
					<div class="children" id="ques-${item.id}-children"></div>
					${item.children && item.children.length > 0 ? `<div class="vline" id="vline-${item.id}" style="left: ${vlineleft}px"></div>` : ''}
					${item.parent_id ? `<div class="hline" style="left: ${-16-identation}px;width:${identation + 16}px;" id="hline-${item.id}"></div>` : ''}
				 </div>`
	} else if (item.level_type == 'question') {
		html += `<div class="qwrapper" style="margin-left: ${identation}px;">
					<div class="itemques" id="box-${item.id}" >
						<div title="${quesData.text}" id="ques-${item.id}" class="text-over-ellipsis clicked flex-1">${quesData.text}</div>
						<div class="children" id="ques-${item.id}-children"></div>
					</div>
					${item.children && item.children.length && !item.is_big ? `<div class="vline" style="left: ${identation}px" id="vline-${item.id}"></div>` : ''}
					${item.parent_id ? `<div class="hline" style="left: ${-16-identation}px;width:${identation + 16}px;}" id="hline-${item.id}"></div>` : ''}
				</div>`
	}
	parent.append(html)
	if (item.children && item.children.length > 0) {
		for (var child of item.children) {
			renderTreeNode($(`#ques-${item.id}-children`), child, item)
		}
	}
}

function countAllDescendants(node, isLastChild) {
	if (!node.children || node.children.length === 0) {
        return isLastChild ? 1 : 0; // 叶子节点; 只有最后一个算作 1
    }
    let count = 0;
    const numChildren = node.children.length;
    for (let i = 0; i < numChildren; i++) {
        const child = node.children[i];
		count += 1 + countAllDescendants(child);
    }
    return count;
}
function countDescendants(node, isLastChild = false) {
    if (!node.children || node.children.length === 0) {
        return isLastChild ? 1 : 0; // 叶子节点; 只有最后一个算作 1
    }
    let count = 0;
    const numChildren = node.children.length;
    for (let i = 0; i < numChildren; i++) {
        const child = node.children[i];
        const childIsLast = (i === numChildren - 1);

        if (childIsLast && !isLastChild) {
            // 如果是最后一个孩子，则只加1
            count += 1;
        } else {
            // 非最后一个孩子，需要计算自己+其子孙的叶子节点数量
            count += 1 + countAllDescendants(child);
        }
    }
    return count;
}
function traverse(node, result, result2) {
	if (!node) return;
	result[node.id] = countDescendants(node);
	if (node.children && node.children.length) {
		result2[node.id] = node.children[node.children.length - 1].id
		for (let child of node.children) {
			traverse(child, result, result2);
		}
	} else {
		result2[node.id] = 0
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
		var vCountMap = {}
		let vLastChildMap = {}
		for (let rootNode of tree_info.tree) {
			traverse(rootNode, vCountMap, vLastChildMap);
		}
		tree_info.list.forEach(item => {
			var itemCount = vCountMap[item.id]
			if (itemCount > 0 && vLastChildMap[item.id]) {
				var lastHLine = $(`#hline-${vLastChildMap[item.id]}`)
				var hlineParentRect = lastHLine.parent()[0].getBoundingClientRect()
				var lastHlineTop = lastHLine.css('top').replace('px', '') * 1
				var linecom = $(`#panelTree #vline-${item.id}`)
				if (linecom && linecom.length > 0) {
					var parent = linecom.parent()
					var parentRect = parent[0].getBoundingClientRect()
					var vlineTop = linecom.css('top').replace('px', '') * 1
					linecom.css('height', (hlineParentRect.top - parentRect.top + lastHlineTop - vlineTop) + 'px')
				}
			}
			if (item.parent_id) {
				if (item.level_type == 'struct') {
					var structCom = $(`#panelTree #struct${item.id}`)
					if (item.lvl !== null) {
						structCom.css('left', (-24 - item.lvl * 16) + 'px')
					} else {
						structCom.css('left', (-24 - 16) + 'px')
					}
				}
			}
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
					var menuItems = []
					if (quesData.level_type == 'question') {
						var nodeData = window.BiyueCustomData.node_list.find(e => {
							return e.id == item.id
						})
						if (nodeData && !quesData.is_merge) { // 合并题不可设置为大题, 当题目处于单元格中时，只能清除大题，不可构建大小题
							if (nodeData.is_big) {
								menuItems.push('clearBig')
								if (!item.cell_id) {
									menuItems.push('setBig2')
								}
								updateMenuPos(event)
							} else if (!item.cell_id) {
								menuItems.push('setBig2')
								if (item.lvl !== null) {
									menuItems.push('setBig')
								}
							}
						}
					} else if (quesData.level_type == 'struct') {
						menuItems.push('question')
					}
					menuItems.push('setLevel')
					if (menuItems.length) {
						generateMenuItems(menuItems, item.id, item.lvl)
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
	updateText('#panelTree #edit', '编辑')
	addClickEvent('#panelTree #edit', (e) => {
		onEdit(e)
	})
}
function updateMenuPos(event) {
	const menu = document.getElementById('dynamicMenu');
	// 获取浏览器窗口的宽度、高度
    const windowWidth = window.innerWidth;
    const windowHeight = window.innerHeight;
	var x = event.pageX
	var y = event.pageY
	// 显示菜单
	menu.style.display = 'block';
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
		if (i == g_tree_info.list.length - 1) {
			toIndex = i
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
		var minlvl = 10
		for (var j = index + 1; j <= toIndex; ++j) {
			var item = $(`#panelTree #box-${g_tree_info.list[j].id}`)
			var parentElement = item.parent();
			var prevElement = item.prev();
			var marginLeft = parseInt((item).css('margin-left'), 10);
			var itemParent = item.parent()
			marginLeft = parseInt((itemParent).css('margin-left'), 10);
			var itemLvl = g_tree_info.list[j].lvl
			big_info.child_ids.push({
				id: g_tree_info.list[j].id,
				parent: parentElement ? parentElement.attr('id') : null,
				prev: prevElement ? prevElement.attr('id') : null,
				marginLeft: marginLeft,
				lvl: itemLvl
			})
			item.css({
				'background-color': '#fff',
				'border': '1px solid #bbb'
			});
			
			if (itemLvl !== null) {
				minlvl = Math.min(minlvl, itemLvl)
			}
			minMarginLeft = Math.min(minMarginLeft, marginLeft)
			$(`#panelTree #hline-${g_tree_info.list[j].id}`).hide()
			item.parent().appendTo(childrenName);
			if (g_tree_info.list[j].parent_id && g_tree_info.list[j].parent_id != id) {
				var parentData = g_tree_info.list.find(e => {
					return e.id == g_tree_info.list[j].parent_id
				})
				if (parentData && parentData.level_type == 'question') {
					item.hide()
				}
			}
		}
		big_info.child_ids.forEach(e => {
			var mleft = 0
			if (e.lvl === null) {
				mleft = 0
			} else if (e.lvl > minlvl) {
				mleft = (e.lvl - minlvl) * 16
			}
			$(`#panelTree #box-${e.id}`).parent().css({
				'margin-left': `${mleft}px`
			})
		})
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

function generateMenuItems(options, id, currentLevel) {
	const menuContent = $('#menuContent');
	menuContent.empty(); // 清除旧的菜单项
	options.forEach(e => {
		var name = ''
		var isSubMenu = false
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
			case 'setLevel':
				name = currentLevel === null ? '设置级别' : '调整级别'
				isSubMenu = true
				break
			default:
				break
		}
		if (name) {
			const menuItem = $('<li>').text(name);
            menuItem.on('click', function() {
                clickMenu(id, e);
            });
            if (isSubMenu) {
				const submenu = $('<ul class="submenu">');
                for (let i = 0; i < 10; i++) {
                    const levelItem = $(`<li>${i + 1}级</li>`);
                    if (i === currentLevel) {
                        levelItem.css('color', '#2489f6');  // 突出显示当前级别
                    }
                    levelItem.on('click', function() {
						setLevel(id, i)
                    });
                    submenu.append(levelItem);
                }
                menuItem.append(submenu);

                menuItem.hover(function() {
                    submenu.show();
					var viewportHeight = window.innerHeight // $(window).height();  // 获取窗口高度
					var menuBottom = submenu.offset().top + submenu.outerHeight();  // 获取菜单底部相对于document的位置
					// 判断是否超出窗口底部
					if (menuBottom > viewportHeight) {
						var overflowHeight = menuBottom - viewportHeight;  // 超出的高度
						var currentTop = parseInt(submenu.css('top'), 10);  // 获取当前top值
						submenu.css('top', currentTop - overflowHeight);  // 调整top值使菜单不超出窗口底部
					}
					var viewportWidth = window.innerWidth
					var menuRight = submenu.offset().left + submenu.outerWidth();  // 获取菜单右边相对于document的位置
					// 判断是否超出窗口右边
					if (menuRight > viewportWidth) {
						var overflowWidth = menuRight - viewportWidth;  // 超出的宽度
						var currentLeft = parseInt(submenu.css('left'), 10);  // 获取当前left值
						submenu.css('left', currentLeft - overflowWidth);  // 调整left值使菜单不超出窗口右边
					}
                }, function() {
                    submenu.hide();
                });
            }
            menuContent.append(menuItem);
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
function setLevel(id, level) {
	return setNumberingLevel([id], level).then((res) => {
		return generateTree()
	})
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
	clickTreeLock,
	initTreeListener
}