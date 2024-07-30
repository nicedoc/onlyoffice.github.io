;(function (window, undefined) {
	window.Asc.plugin.init = function () {
		console.log('message init')
		window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'showMessageBox' })
	}
	var MESSAGE_INFO = {
		showCancel: true,
		cancelText: '取消',
		confirmText: '确定',
		content: '',
		extra_data: null
	}
	var message_info = null

	function init(info) {
		$('#content').html(info.content || '')
		var btnCancel = $('#cancel')
		if (btnCancel) {
			if (info.showCancel) {
				btnCancel.show()
				btnCancel.html(info.cancelText || MESSAGE_INFO.cancelText)
			} else {
				btnCancel.hide()
			}
			btnCancel.on('click', onCancel)
		}
		var btnConfirm = $('#confirm')
		if (btnConfirm) {
			btnConfirm.html(info.confirmText || MESSAGE_INFO.confirmText)
			btnConfirm.on('click', onConfirm)
		}
	}

	function onCancel() {
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'onMessageDialog',
			cmd: 'cancel',
			extra_data: message_info.extra_data
		})
	}

	function onConfirm() {
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'onMessageDialog',
			cmd: 'confirm',
			extra_data: message_info.extra_data
		})
	}

	window.Asc.plugin.attachEvent('initMessageBox', function (message) {
		console.log('接收的消息', message)
		message_info = Object.assign({}, MESSAGE_INFO, message)
		init(message_info)
	})
})(window, undefined)
