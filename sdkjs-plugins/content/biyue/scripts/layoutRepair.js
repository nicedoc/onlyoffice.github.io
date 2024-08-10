;(function (window, undefined) {
	var layout = 'layout_repair'
	window.Asc.plugin.init = function () {
		console.log('layoutRepair init')
		window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'LayoutRepairMessage' })
	}

	function init(info) {
		if (!info || !info.has32) {
			$('#underSpace').hide()
		}
		if (!info || !info.has160) {
			$('#nbsp').hide()
		}
		if (!info || !info.has65307) {
			$('#chineseSemicolon').hide()
		}
		$('#cancel').on('click', onCancel)
		$('#confirm').on('click', onConfirm)
		$('#underSpace .replace').on('click', () => {
			onReplaceOrIgnore(1, 32, 95)
		})
		$('#nbsp .replace').on('click', () => {
			onReplaceOrIgnore(1, 160, 32)
		})
		$('#chineseSemicolon .replace').on('click', () => {
			onReplaceOrIgnore(1, 65307, 59)
		})
		$('#chineseSemicolon .ignore').on('click', () => {
			onReplaceOrIgnore(0, 65307, 59)
		})
		$('#chineseSpace .replace').on('click', () => {
			onReplaceOrIgnore(1, 12288, 32)
		})
	}

	function onReplaceOrIgnore(type, val, newValue) {
		console.log('onReplaceOrIgnore', type, val, newValue)
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'LayoutRepairMessage',
			cmd: {
				type: type,
				value: val,
				newValue: newValue
			}
		})
	}

	function onConfirm() {
		console.log('onConfirm')
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'confirmDialog'
		})
	}

	function onCancel() {
		console.log('onCancel')
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'cancelDialog',
		})
	}

	window.Asc.plugin.attachEvent('initLayoutRepair', function (message) {
		console.log('接收的消息', message)
		init(message)
	})
})(window, undefined)
