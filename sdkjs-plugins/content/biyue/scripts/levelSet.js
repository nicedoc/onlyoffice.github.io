;(function (window, undefined) {
	var layout = 'level_set'
	window.Asc.plugin.init = function () {
		console.log('levelSet init')
		window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'levelMessage' })
	}

	function init() {
		var listEl = document.getElementById('levelList')
		if (!listEl) return
		for (let i = 0; i < 10; ++i) {
			const level = document.createElement('div');
			level.className = 'item'
			level.textContent = `${i + 1}级: `;
			
			const select = document.createElement('select');
			select.innerHTML = `
				<option value="question">题目</option>
				<option value="struct">结构</option>
			`;
			if (i < 2) {
				select.value == 'struct'
			} else {
				select.value = 'question'
			}
			level.appendChild(select);
			listEl.appendChild(level);
		}
		$('#cancel').on('click', onCancel)
		$('#confirm').on('click', onConfirm)
	}

	function onConfirm() {
		const options = [];
    	const levels = document.querySelectorAll('#levelList select');
		var levelMap = {}
		levels.forEach((select, index) => {
			const selectedValue = select.value;
			levelMap[index] = selectedValue
		});
		
		console.log(options);
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'LevelSetConfirm',
			levels: levelMap
		})
	}

	function onCancel() {
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'cancelDialog',
		})
	}

	window.Asc.plugin.attachEvent('initInfo', function (message) {
		console.log('接收的消息', message)
		init()
	})
})(window, undefined)
