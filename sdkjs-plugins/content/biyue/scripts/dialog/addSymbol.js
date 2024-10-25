;(function (window, undefined) {
	var select_value = null
	window.Asc.plugin.init = function () {
		window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'showSymbols' })
	}

	async function loadIcons() {
		const response = await fetch('./resources/iconfont_word/iconfont.json'); // 相对路径指向数据文件
		const data = await response.json();
		const glyphs = data.glyphs;
		const container = document.getElementById('icon-container');
		const confirmButton = document.getElementById('confirm');
		if (confirmButton) {
			confirmButton.classList.add('disabled');
		}
		glyphs.forEach(glyph => {
			const iconBox = document.createElement('div');
			iconBox.className = 'icon-item';
			iconBox.title = glyph.unicode
			iconBox.innerHTML = `
				<i class="iconfont icon-${glyph.font_class}" aria-hidden="true"></i>
				<div>${glyph.name}</div>
			`;
			
			iconBox.addEventListener('click', function() {
				const allIconItems = document.querySelectorAll('.icon-item');
                allIconItems.forEach(item => item.classList.remove('selected'));
                    
                // Add selected class to the current item
            	iconBox.classList.add('selected');
				// Enable the confirm button
				if (confirmButton) {
					confirmButton.classList.remove('disabled');
					confirmButton.classList.add('enabled');
				}
				select_value = glyph.unicode
			});

			container.appendChild(iconBox);
		});
		confirmButton.addEventListener('click', onConfirm)
		const cancelButton = document.getElementById('cancel');
		if (cancelButton) {
			cancelButton.addEventListener('click', onCancel)
		}
	}

	function onCancel() {
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'cancelDialog'
		})
	}

	function onConfirm() {
		if (!select_value) {
			return
		}
		console.log('onConfirm', select_value)
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'insertSymbol',
			data: select_value
		})
	}

	window.Asc.plugin.attachEvent('initSymbols', function (message) {
		console.log('接收的消息', message)
		loadIcons()
	})
})(window, undefined)
