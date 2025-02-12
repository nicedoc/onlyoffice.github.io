import { addClickEvent, updateText, showCom, updateHintById, setBtnLoading, isLoading } from '../model/util.js';
;(function (window, undefined) {
	const UnderlineType = [{
		id: 11,
		value: 'None',
		name: '无',
		valid: true
	}, {
		id: 12,
		value: 'Single',
		name: '单线',
		valid: true
	}, {
		id: 0,
		value: 'Dash',
		name: '破折号',
		valid: true
	}, {
		id: 1,
		value: 'DashDotDotHeavy',
		name: '粗重的点划线',
		valid: false
	}, {
		id: 2,
		value: 'DashDotHeavy',
		name: '粗重的划点线',
		valid: false
	}, {
		id: 3,
		value: 'DashedHeavy',
		name: '粗重的虚线',
		valid: false
	}, {
		id: 4,
		value: 'DashLong',
		name: '长破折号',
		valid: false
	}, {
		id: 5,
		value: 'DashLongHeavy',
		name: '长且粗重的破折号',
		valid: false
	}, {
		id: 6,
		value: 'DotDash',
		name: '点划线',
		valid: true
	}, {
		id: 7,
		value: 'DotDotDash',
		name: '点点划线',
		valid: true
	}, {
		id: 8,
		value: 'Dotted',
		name: '点状线',
		valid: true
	}, {
		id: 9,
		value: 'DottedHeavy',
		name: '粗重的点状线',
		valid: false
	}, {
		id: 10,
		value: 'Double',
		name: '双线',
		valid: true
	}, {
		id: 13,
		value: 'Thick',
		name: '粗线',
		valid: true
	}, {
		id: 14,
		value: 'Wave',
		name: '波浪线',
		valid: true
	}, {
		id: 15,
		value: 'WavyDouble',
		name: '双层波浪线',
		valid: true
	}, {
		id: 16,
		value: 'WavyHeavy',
		name: '粗重的波浪线',
		valid: false
	}
	// , {
	// 	id: 17,
	// 	value: 'Words',
	// 	name: '单词线',
	// 	valid: true
	// }
	]
	window.Asc.plugin.init = function () {
		console.log('set underline init')
		window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'initDialog', initmsg: 'setUnderLineMessage' })
	}
	function init() {
		renderList(UnderlineType, 'list')
		addClickEvent('.cancel', onCancel)
		addClickEvent('.confirm', onConfirm)
	}
	 
	function renderList(list, listId) {
		var html = ''
		if (list) {
			list.forEach((e, index) => {
				var str = e.value=='None' ? '无' : `<img src="resources/underline/${e.value}.png" alt="${e.name}" onerror="this.style.display='none'" />`
				html += `<div class="item ${e.valid ? 'enabled' : 'disabled'}" id=${e.id} data-id="${e.id}" data-value="${e.value}" title="${e.name}">
					${str}
				</div>`
			})
		}
		updateText(`.${listId}`, html)
		if (list) {
			list.filter(e => e.valid).forEach(e => {
				addClickEvent(`#${e.id}`, onItem)
			})
		}
	}

	function onItem(event) {
		var dataset = handleClick(event)
		if (!dataset) return;
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

	function onCancel() {
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'cancelDialog'
		});
	}

	function onConfirm() {
		const selected = $('.pic-selected');
		if (selected.length === 0) {
			return;
		}
		$(`.pic-selected`).removeClass('pic-selected')
		const value = selected.data('id');
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'setUnderLineMessage',
			data: value
		});
	}

	window.Asc.plugin.attachEvent('setUnderLineMessage', function (message) {
		init()
	})
})(window, undefined)
