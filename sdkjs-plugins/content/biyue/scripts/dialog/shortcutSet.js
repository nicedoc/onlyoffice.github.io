import ComponentSelect from '../../components/Select.js'
;(function (window, undefined) {
	var select_ask_shortcut = null
	var vShortcut = '0'
	window.Asc.plugin.init = function () {
		console.log('shortcut set init')
		window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'initDialog', initmsg: 'shortcutMessage' })
	}
	function init() {
		select_ask_shortcut = new ComponentSelect({
			id: 'shortcutKeyDiv',
			options: [
				{ value: '0', label: '未定义' },
				{ value: 'ctrl', label: '双击 + ctrl' },
				{ value: 'alt', label: '双击 + alt' },
				{ value: 'shift', label: '双击 + shift' }
			],
			value_select: vShortcut,
			callback_item: (data) => {
				changeAskShortcut(data)
			},
			width: '100%',
			pop_width: '100%'
		})
		$("#myIcon").on('click', function(){
			$("#tooltip").toggle();
		});
		$("#myIcon").on('mouseover', function(){
			$("#tooltip").show();
		});
		$("#myIcon").on('mouseout', function(){
			$("#tooltip").hide();
		});
	}

	function changeAskShortcut(data) {
		console.log('changeAskShortcut', data)
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'shortcutMessage',
			cmd: 'update',
			data: {
				ask_shortcut: data.value
			}
		})
	}

	window.Asc.plugin.attachEvent('shortcutMessage', function (message) {
		console.log('接收的消息', message)
		if (message && message.BiyueCustomData) {
			vShortcut = message.BiyueCustomData.ask_shortcut || '0'
		} else {
			vShortcut = '0'
		}
		init()
	})
})(window, undefined)
