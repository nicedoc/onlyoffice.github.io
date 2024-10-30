;(function (window, undefined) {
	var layout = 'layout_repair'
	var detect_list = [{
		idname: 'smallImg',
		delete: true,
		type: 'error',
		keyname: 'hasSmallImage',
		text1: '宽高过小的图片，',
		text2: '建议删除，',
		value: 'smallimg'
	}, {
		idname: 'nbsp',
		replace: true,
		ignore: false,
		type: 'error',
		keyname: 'has160',
		text1: '不间断空格，',
		text2: '建议替换为空格，',
		value: 160,
		newValue: 32
	}, {
		idname: 'tab',
		replace: true,
		ignore: false,
		type: 'error',
		keyname: 'hasTab',
		text1: '括号里使用了tab，',
		text2: '建议替换为空格，',
		value: 'tab',
		newValue: 32
	}, {
		idname: 'underSpace',
		replace: true,
		ignore: false,
		type: 'error',
		keyname: 'has32',
		text1: '空格下划线，',
		text2: '建议替换为等长的下划线，',
		value: 32,
		newValue: 95
	}, {
		idname: 'chineseSemicolon',
		replace: true,
		ignore: true,
		type: 'warning',
		keyname: 'has65307',
		text1: '中文分号，',
		text2: '建议替换为英文，',
		value: 65307,
		newValue: 59
	}, {
		idname: 'chineseSpace',
		replace: true,
		ignore: true,
		type: 'warning',
		keyname: 'has12288',
		text1: '中文空格，',
		text2: '建议替换为英文，',
		value: 12288,
		newValue: 32
	}, {
		idname: 'whiteBg',
		replace: true,
		type: 'warning',
		keyname: 'hasWhiteBg',
		text1: '段落背景为白色，',
		text2: '建议替换为透明色，',
		value: 'whitebg'
	}]
	window.Asc.plugin.init = function () {
		console.log('layoutRepair init')
		window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'LayoutRepairMessage' })
	}

	function init(info) {
		var str = ``
		for (var i = 0; i < detect_list.length; ++i) {
			var item = detect_list[i]
			if (info && info[item.keyname]) {
				str += `<div id=${item.idname} style="margin: 4px 0">`
				if (item.type == 'error') {
					str += `<span style="color: #F56C6C;">Error：</span>`
				} else if (item.type == 'warning') {
					str += `<span style="color: #E6A23C;">Warning：</span>`
				}
				str +='<span>检查到</span>'
				str += `<span style="font-weight: bold;">${item.text1}</span>`
				str += `<span>${item.text2}</span>`
				if (item.ignore) {
					str +='<span class="ignore">忽略></span>'
				}
				if (item.replace) {
					str +='<span class="replace">替换></span>'
				}
				if (item.delete) {
					str +='<span class="delete">删除></span>'
				}
				str += '</div>'
			}
		}
		$('#list').html(str)
		$('#cancel').on('click', onCancel)
		$('#confirm').on('click', onConfirm)
		for (var i = 0; i < detect_list.length; ++i){
			addItemClick(i)	
		}
	}

	function addCmdEvent(idname, classname, cmdType, value, newValue, idName, keyname) {
		var com = $(`#${idname} ${classname}`)
		if (com) {
			com.on('click', () => {
				onCommand(cmdType, value, newValue, idName, keyname)
			})
		}
	}

	function addItemClick(i) {
		var value = detect_list[i].value
		var newValue = detect_list[i].newValue
		var idName = `#${detect_list[i].idname}`
		var classlist = ['.ignore', '.replace', '.delete']
		for (var j = 0; j < classlist.length; ++j) {
			addCmdEvent(detect_list[i].idname, classlist[j], j, value, newValue, idName, detect_list[i].keyname)
		}
	}

	function onCommand(type, val, newValue, comId, keyname) {
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'LayoutRepairMessage',
			cmd: {
				type: type,
				value: val,
				newValue: newValue,
				keyname: keyname
			}
		})
		$(comId).hide()
	}

	function onConfirm() {
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'confirmDialog'
		})
	}

	function onCancel() {
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'cancelDialog',
		})
	}

	window.Asc.plugin.attachEvent('initLayoutRepair', function (message) {
		console.log('接收的消息', message)
		init(message)
	})
})(window, undefined)
