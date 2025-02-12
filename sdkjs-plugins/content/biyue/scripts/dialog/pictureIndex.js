import ComponentSelect from '../../components/Select.js';
import NumberInput from '../../components/NumberInput.js';
import { addClickEvent, updateText, showCom, updateHintById, setBtnLoading, isLoading } from '../model/util.js';
;(function (window, undefined) {
	let list_doc = []
	let list_ignore = []
	let question_map = {}
	let ques_indexs = []
	let ques_visible_id = 0
	let select_image_link = null
	let link_coverage_percent = ''
	let link_type = ''
	let isDragging = false;
	window.Asc.plugin.init = function () {
		console.log('picture index init')
		window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'initDialog', initmsg: 'pictureIndexMessage' })
	}
	function updateProgress(percent) {
		if (!percent) return;
		
		percent = Math.round(percent);
		// 限制在50-100之间
		percent = Math.max(50, Math.min(100, percent));
		
		link_coverage_percent = percent
		$('.progress-bar').css('width', percent + '%');
		$('.progress-text').text(percent + '%');
	}
	function init() {
		select_image_link = new ComponentSelect({
			id: 'imageLinkSelect',
			options: [
				{ value: '0', label: '关闭就近关联' },
				{ value: '1', label: '开启就近关联' }
			],
			value_select: '1',
			callback_item: (data) => {
				changeImageLink(data)
			},
			width: '45%',
			pop_width: '100%'
		})
		addClickEvent('#areaLink', onAreaLink)
		addClickEvent('#allLink', onAllLink)
		enableBtnImageLink(true)
		showCom('#imageLinkTip')
		addClickEvent('#btnImageLink', onImageLink)
		addClickEvent('#confirmAutoLink', onConfirmAutoLink)
		addClickEvent('#cancelAutoLink', onCancelAutoLink)
		addClickEvent('#btnRefresh', onRefresh)
		showCom('.check-wrapper', false)
		renderList(list_doc, 'list')
		renderList(list_ignore, 'list-ignore')
		$('.list-ques').hide()
		showCom('.progress-container', link_type == 'area')
		showCom('.progress-text-wrapper', link_type == 'area')
		showCom('#imageLinkTip', false)
		updateListIgnore()
		$('.box1').eq(link_type == 'area' ? 0 : 1).addClass('selected')
		updateProgress(link_coverage_percent)

		const $progressBar = $('.progress-bar');
		const $container = $('.progress-container');
		const $handle = $('.slider-handle');
		const $progressText = $('.progress-text');

		// 先移除旧的事件监听器
		$(document).off('mousemove mouseenter mouseleave mouseup');
		$(window).off('blur');

		$handle.on('mousedown', function(e) {
			isDragging = true;
			e.preventDefault(); // 防止选中文本
		});
		
		// 再添加新的事件监听器
		$(document).on('mousemove', function(e) {
			if (!isDragging) return;

			const containerOffset = $container.offset();
			let x = e.pageX - containerOffset.left;
			const containerWidth = $container.width();

			// 限制拖动范围
			x = Math.max(0, Math.min(x, containerWidth));
			
			// 计算百分比
			const percent = (x / containerWidth) * 100;
			updateProgress(percent);
		});

		 // 处理鼠标离开窗口的情况
		 $(document).on('mouseleave', function() {
			if (isDragging) isDragging = false;
		});
		// 如果鼠标重新进入窗口，取消超时
		$(document).on('mouseenter', function(e) {
			if (e.buttons === 1) isDragging = true;
		});

		// 修改：同时监听 mouseup 和 mouseleave 事件
		$(document).on('mouseup', function() {
			if (isDragging) {
				isDragging = false;
			}
		});
		// 窗口失去焦点时
		$(window).on('blur', function() {
			console.log('window blur')
			if (isDragging) {
				isDragging = false;
			}
		});

		// 点击进度条容器更新进度
		$container.on('click', function(e) {
			// 只有在非拖动状态下才处理点击事件
			if (!isDragging) {
				const containerOffset = $(this).offset();
				const x = e.pageX - containerOffset.left;
				const percent = (x / $(this).width()) * 100;
				updateProgress(percent);
			}
		});
	}

	function updateListIgnore() {
		showCom('.ignore-none', $('.list-ignore').children().length === 0)
	}

	function onConfirmAutoLink() {
		showCom('.check-wrapper', false)
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'pictureIndexMessage',
			cmd: 'autoLink',
			data: {
				link_type: link_type,
				link_coverage_percent: link_coverage_percent
			}
		})
	}

	function onCancelAutoLink() {
		showCom('.check-wrapper', false)
		showCom('#btnImageLink', true)
	}
	 
	function renderList(list, listId) {
		var html = ''
		if (list) {
			list.forEach((e, index) => {
				var uid = e.sort_id
				var dataset = `data-id=${e.uid} data-type=${e.type} data-sort-id="${e.sort_id}"`
				if (listId == 'list-ignore') {
					dataset += ' data-ignore="1"'
				}
				html += `<div class="item-pic clicked row-between" id=${e.uid} ${dataset}>
					<div>
						<span>${e.type == 'table' ? '表格' : '图片'} ${uid}</span>
						<i class="iconfont dot icon-puma"></i>
					</div>
					<div class="row-align-center">
						<div ${dataset} class="ignore">${listId == 'list' ? '忽略' : '恢复'}</div>
						<div ${dataset} class="row-align-center link2 clrb">
							<div class="link1" ${dataset}>关联</div>
							<i ${dataset} class="iconfont idown icon-angeldown1 clrb"></i>
						</div>
					</div>
				</div>`
			})
		}
		updateText(`.${listId}`, html)
		if (list) {
			list.forEach(e => {
				addClickEvent(`#${e.uid}`, locatePicture)
				addClickEvent(`#${e.uid} .ignore`, onIgnore)
				// addClickEvent(`#${e.uid} .link1`, onLink)
				addClickEvent(`#${e.uid} .link2`, onLink)
				addClickEvent(`#${e.uid} .nodot`, onNoDot)
				addClickEvent(`#${e.uid} .dot`, onOpenDot)
				showDotButton(e.uid, e.partical_no_dot)
				if (listId == 'list-ignore') {
					$(`#${e.uid} .link2`).hide()
				}
			})
		}
	}

	function locatePicture(event) {
		var dataset = handleClick(event)
		if (!dataset) return;
		
		var list = dataset.ignore ? list_ignore : list_doc
		var index = list.findIndex(e => e.uid == dataset.id)
		if (index === -1) {
			console.warn('Picture not found:', dataset.id);
			return;
		}
		
		var data = list[index]
		if (data) {
			window.Asc.plugin.sendToPlugin('onWindowMessage', {
				type: 'elementLinkedMessage',
				cmd: 'locate',
				data: {
					target_id: data.id,
					type: data.type
				}
			})
		}
		updateSelected(dataset.id, dataset.ignore)
		// if (!dataset.ignore) {
		// 	renderQuesList(data, false)
		// }
	}

	function updateSelected(id, ignore) {
		$(`.pic-selected`).removeClass('pic-selected')
		if (ignore) {
			$(`.list-ignore #${id}`).addClass('pic-selected')
		} else {
			$(`.list #${id}`).addClass('pic-selected')
		}
	}

	function onIgnore(event) {
		var dataset = handleClick(event)
		if (dataset) {
			var targetList = dataset.ignore ? list_ignore : list_doc
			var index = targetList.findIndex(e => {
				return e.uid == dataset.id
			})
			if (index >= 0 && ques_visible_id && targetList[index].id == ques_visible_id) {
				hideQuesList(true, true)
			}
			var id = dataset.id
			var target_ignore = false
			if (dataset.ignore) {
				moveAndSort($(`#${id}`), '.list')
				$(`#${id}`).removeAttr("data-ignore")
				$(`#${id} .ignore`).attr("title", "加入忽略").removeAttr("data-ignore").html('忽略');
				// $(`#${id} .link`).removeAttr("data-ignore");
				$(`#${id} .link2`).show()
				list_doc.push(targetList[index])
				// list_doc.splice(targetIndex, 0, targetList[index])
				list_ignore.splice(index, 1)
			} else {
				moveAndSort($(`#${id}`), '.list-ignore')
				$(`#${id}`).attr("data-ignore", '1')
				$(`#${id} .ignore`).attr("title", "移除忽略").attr("data-ignore", '1').html('恢复');
				// $(`#${id} .link`).attr("data-ignore", '1');
				$(`#${id} .link2`).hide()
				list_ignore.push(targetList[index])
				// list_ignore.splice(targetIndex, 0, targetList[index])
				list_doc.splice(index, 1)
				target_ignore = true
			}
			updateListIgnore()
			window.Asc.plugin.sendToPlugin('onWindowMessage', {
				type: 'pictureIndexMessage',
				cmd: 'ignore',
				data: {
					...targetList[index],
					ignore: target_ignore
				}
			})
		}
	}
	function moveAndSort(item, targetListSelector) {
		let type = item.data('type');
		let sortId = item.data('sort-id');
		let inserted = false;
	
		// 在目标列表中找到正确的位置，并插入
		$(targetListSelector).children('.item-pic').each(function() {
			let thisType = $(this).data('type');
			let thisSortId = $(this).data('sort-id');
			if ((type === thisType && sortId < thisSortId) || (type === 'drawing' && thisType === 'table')) {
				item.insertBefore($(this));
				inserted = true;
				return false; // 跳出循环
			}
		});
	
		// 如果没有合适的位置，则添加到末尾
		if (!inserted) {
			$(targetListSelector).append(item);
		}
	}

	function onLink(event) {
		var dataset = handleClick(event)
		if (dataset) {
			var data = list_doc.find(e => {
				return e.uid == dataset.id
			})
			renderQuesList(data, false)
		}
	}

	function onNoDot(event) {
		handleParticalNoDot(event, true)
	}

	function onOpenDot(event) {
		handleParticalNoDot(event, false)
	}

	function handleParticalNoDot(event, partical_no_dot) {
		var dataset = handleClick(event)
		var targetList = dataset.ignore ? list_ignore : list_doc
		var data = targetList.find(e => {
			return e.uid == dataset.id
		})
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'pictureIndexMessage',
			cmd: 'updateDot',
			data: {
				target_id: data.id,
				type: data.type,
				partical_no_dot: partical_no_dot
			}
		})
		showDotButton(dataset.id, partical_no_dot)
	}

	function handleClick(event) {
		if (!event) {
			return
		}
		event.cancelBubble = true
		event.preventDefault()
		event.stopPropagation()
		var target = event.currentTarget || event.target
		if (!target) {
			return null
		}
		return target.dataset
	}

	function renderQuesList(data, ignore) {
		if (ignore) {
			hideQuesList(true, false)
		} else {
			if (ques_visible_id == data.id || !data) {
				hideQuesList(true, false)
				return
			}
			ques_visible_id = data.id
			var html = '<div>'
			for (var quesId of ques_indexs) {
				var quesData = question_map[quesId]
				if (!quesData || quesData.level_type != 'question') {
					continue
				}
				html += `<div class="row-between item-ques clicked" id="q${quesId}">
					<div>
						<div class="ques-pre">
							<i class="iconfont icon-suozhu ft12 clrb"></i>
						</div>
						<div class="ques-text" data-id="${quesId}" title="${quesData.text}">${quesData.text}</div>
					</div>
					<div>
						<i class="iconfont icon-suozhu ft12 clrb ilink" data-id="${quesId}" title="关联"></i>
					</div>
				</div>`
			}
			if (ques_indexs.length == 0) {
				html += '<div class="none">尚未切题</div>'
			}
			html += '</div>'
			$('.list-ques').show().insertAfter(`#${data.uid}`).html(html)
			for (var quesId of ques_indexs) {
				showLinkButton(quesId, data.ques_use.includes(quesId))
				addClickEvent(`#q${quesId} .ques-text`, locateQues)
				addClickEvent(`#q${quesId} .ilink`, onLinkQues)
			}
		}
	}

	function locateQues(event) {
		var dataset = handleClick(event)
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'pictureIndexMessage',
			cmd: 'locateQues',
			data: {
				ques_id: dataset.id
			}
		})
	}

	function onLinkQues(event) {
		var dataset = handleClick(event)
		var data = list_doc.find(e => {
			return e.id == ques_visible_id
		})
		if (data) {
			var targetQuesUse = data.ques_use || []
			var index = targetQuesUse.findIndex(e => {
				return e == dataset.id
			})
			var isLink = index >= 0
			if (isLink) {
				targetQuesUse.splice(index, 1)
			} else {
				targetQuesUse.push(dataset.id)
			}
			window.Asc.plugin.sendToPlugin('onWindowMessage', {
				type: 'pictureIndexMessage',
				cmd: 'link',
				data: {
					target_type: data.type,
					target_id: data.id,
					ques_use: targetQuesUse,
					ques_id: dataset.id
				}
			})
			showLinkButton(dataset.id, !isLink)
		}
	}

	function showLinkButton(qid, linked) {
		if (linked) {
			$(`#q${qid} .ilink`).removeClass('icon-suozhu').removeClass('clrb').addClass('icon-jiesuo').addClass('clrr').attr("title", "取消关联")
			$(`#q${qid} .ques-pre .iconfont`).show()
			$(`#q${qid} .ques-text`).addClass('clrb')
		} else {
			$(`#q${qid} .ilink`).removeClass('icon-jiesuo').removeClass('clrr').addClass('icon-suozhu').addClass('clrb').attr("title", "关联")
			$(`#q${qid} .ques-pre .iconfont`).hide()
			$(`#q${qid} .ques-text`).removeClass('clrb')
		}
	}

	function showDotButton(id, partical_no_dot) {
		if (partical_no_dot) {
			$(`#${id} .dot`).removeClass('icon-puma').removeClass('clrb').addClass('icon-bupuma').addClass('clrr').attr("title", "不铺码")
		} else {
			$(`#${id} .dot`).removeClass('icon-bupuma').removeClass('clrr').addClass('icon-puma').addClass('clrb').attr("title", "铺码")
		}
	}
	
	function changeImageLink(data) {
		enableBtnImageLink(data.value * 1)
		showCom('#linkwrapper', data.value * 1)
	}
	
	function onAreaLink() {
		link_type = 'area'
		showCom('.progress-container', true)
		showCom('.progress-text-wrapper', true)
		$('.selected').removeClass('selected')
		$('.box1').eq(0).addClass('selected')
		updateProgress(link_coverage_percent || 70)
	}

	function onAllLink() {
		link_type = 'all'
		showCom('.progress-container', false)
		showCom('.progress-text-wrapper', false)
		$('.selected').removeClass('selected')
		$('.box1').eq(1).addClass('selected')
	}
	
	function enableBtnImageLink(v) {
		var btnlink = $('#btnImageLink')
		if (!btnlink) {
			return
		}
		if (v) {
			btnlink.removeClass('btn-unable')
		} else {
			btnlink.addClass('btn-unable')
		}
		// window.auto_image_link = v
	}
	// 图片就近关联
	function onImageLink() {
		if ($('#btnImageLink').hasClass('btn-unable')) {
			return
		}
		showCom('.check-wrapper', true)
		showCom('#btnImageLink', false)
	}

	function hideQuesList(resetVisibleId = true, resetPos = true) {
		if (resetPos) {
			$('.list-ques').appendTo('.info')
		}
		$('.list-ques').hide()
		if (resetVisibleId) {
			ques_visible_id = 0
		}
	}

	function onRefresh() {
		if (isLoading('btnRefresh')) { 
			return
		}
		setBtnLoading('btnRefresh', true)
		hideQuesList()
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'pictureIndexMessage',
			cmd: 'refresh'
		})
	}

	window.Asc.plugin.attachEvent('pictureIndexMessage', function (message) {
		console.log('pictureIndexMessage 接收的消息', message)
		if (message) {
			list_doc = message.list || []
			list_ignore = message.list_ignore || []
			if (message.BiyueCustomData && message.BiyueCustomData.question_map) {
				question_map = message.BiyueCustomData.question_map
				var node_list = message.BiyueCustomData.node_list || []
				ques_indexs = []
				for (var key in question_map) {
					if (question_map[key].level_type == 'struct') {
						continue
					}
					var qid = question_map[key].is_merge ? question_map[key].ids[0] : key
					ques_indexs.push({
						id: key,
						index: node_list.findIndex(e=>{return e.id == qid})
					})
				}
				ques_indexs = ques_indexs.sort((a, b) => {
					return a.index - b.index
				})
				ques_indexs = ques_indexs.map(e => {
					return e.id
				})
				link_coverage_percent = message.BiyueCustomData.link_coverage_percent
				link_type = message.BiyueCustomData.link_type || 'all'
			} else {
				question_map = {}
			}
		}
		if (message.type == 'autolink') {
			updateHintById('#imageLinkTip', '就近关联完成', '#4caf50')
			renderList(list_doc, 'list')
			renderList(list_ignore, 'list-ignore')
			hideQuesList()
			showCom('#btnImageLink', true)
		} else {
			init()
		}
		setBtnLoading('btnRefresh', false)
	})
	window.Asc.plugin.attachEvent('locatePicture', function (message) {
		console.log(' locatePicture接收的消息', message)
		if (message && message.uid) {
			var picdata = null
			var ignore = false
			picdata = list_doc.find(e => {
				return e.uid == message.uid
			})
			if (!picdata) {
				picdata = list_ignore.find(e => {
					return e.uid == message.uid
				})
				ignore = true
			}
			$(`.pic-selected`).removeClass('pic-selected')
			if (picdata) {
				updateSelected(picdata.uid, ignore)
				hideQuesList(true, false)
			}
		}
	})
	window.Asc.plugin.attachEvent('updateUse', function (message) {
		console.log('updateUse 接收的消息', message)
		if (message && message.list) {
			for (var item of message.list) {
				var picdata = null
				var ignore = false
				picdata = list_doc.find(e => {
					return e.uid == item.uid
				})
				if (!picdata) {
					picdata = list_ignore.find(e => {
						return e.uid == item.uid
					})
					ignore = true
				}
				if (message.from == 'particalNoDot') {
					if (picdata) {
						picdata.partical_no_dot = item.partical_no_dot
						showDotButton(item.uid, item.partical_no_dot)
					}
				} else {
					if (picdata) {
						picdata.ques_use = item.ques_use
					} else {
						list_doc.push(item)
					}
				}
			}
		}
	})

	// 建议在适当的时机清理这些全局变量
	function cleanup() {
		list_doc = []
		list_ignore = []
		question_map = {}
		ques_indexs = []
		ques_visible_id = 0
		link_coverage_percent = ''
		link_type = ''
		isDragging = false
	}
})(window, undefined)
