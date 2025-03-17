import { addClickEvent, updateText, showCom, updateHintById, setBtnLoading, isLoading } from '../model/util.js';
;(function (window, undefined) {
	let list_doc = []
	window.Asc.plugin.init = function () {
		console.log('picture index init')
		window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'initDialog', initmsg: 'pictureListMessage' })
	}
	function init() {
		addClickEvent('#btnRefresh', onRefresh)
		renderList(list_doc, 'list')
	}
	 
	function renderList(list, listId) {
		var html = ''
		if (list) {
			list.forEach((e, index) => {
				var uid = e.sort_id
				var dataset = `data-id=${e.uid} data-type=${e.type} data-sort-id="${e.sort_id}"`
				html += `<div class="item-pic clicked row-between" id=${e.uid} ${dataset}>
					<div>
						<span>${e.type == 'table' ? '表格' : '图片'} ${uid}</span>
					</div>
				</div>`
			})
		}
		updateText(`.${listId}`, html)
		if (list) {
			list.forEach(e => {
				addClickEvent(`#${e.uid}`, locatePicture)
			})
		}
	}

	function locatePicture(event) {
		var dataset = handleClick(event)
		if (!dataset) return;
		
		var list = list_doc
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
		updateSelected(dataset.id)
	}

	function updateSelected(id) {
		$(`.pic-selected`).removeClass('pic-selected')
		$(`.list #${id}`).addClass('pic-selected')
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

	function onRefresh() {
		if (isLoading('btnRefresh')) {
			return
		}
		setBtnLoading('btnRefresh', true)
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'pictureIndexMessage',
			cmd: 'refresh',
			data: 'pictureList'
		})
	}

	window.Asc.plugin.attachEvent('pictureListMessage', function (message) {
		console.log('pictureListMessage 接收的消息', message)
		if (message) {
			list_doc = message.list || []
		}
		init()
		setBtnLoading('btnRefresh', false)
	})
	window.Asc.plugin.attachEvent('locatePicture', function (message) {
		console.log(' locatePicture接收的消息', message)
		if (message && message.uid) {
			var picdata = null
			picdata = list_doc.find(e => {
				return e.uid == message.uid
			})
			$(`.pic-selected`).removeClass('pic-selected')
			if (picdata) {
				updateSelected(picdata.uid)
			}
		}
	})

	// 建议在适当的时机清理这些全局变量
	function cleanup() {
		list_doc = []
	}
})(window, undefined)
