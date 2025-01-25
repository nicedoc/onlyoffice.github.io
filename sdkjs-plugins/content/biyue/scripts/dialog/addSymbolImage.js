;(function (window, undefined) {
	var list = [{
		label: '田字格',
		value: 'tianzige',
		type1: 'tianzige_1',
		type2: 'tianzige_2',
	}, {
		label: '文字格',
		value: 'tianzige2',
		type1: 'tianzige2_1',
		type2: 'tianzige2_2',
	}, {
		label: '拼音格',
		value: 'tianzige3',
		type1: 'tianzige3',
		type2: 'tianzige3',
	}, {
		label: '英语四线三格',
		value: 'sixian'
	}]
	var insert_type = ''
	window.Asc.plugin.init = function () {
		window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'showSymbols' })
	}
	function init() {
		initList()
		$('.input-box').hide()
		$('.preview1').hide()
		$('#copy').hide()
		$('#confirm').on('click', onInert)
		$('#cancel').on('click', onCancel)
		$('#generate').on('click', generateImage)
		$('#copy').on('click', onCopy)
	}
	function initList() {
		var content = ''
		for (var item of list) {
			var imgUrl = `https://by-base.oss-cn-shenzhen.aliyuncs.com/${item.type1 || item.value}.png`
			content += `<div class="item" id="${item.value}" title="${item.label}"><img class="itemimg" src="${imgUrl}" /></div>`
		}
		$('.list').html(content)
		for (var item of list) {
			$('#' + item.value).on('click', clickItem)
		}
	}

	function clickItem(e) {
		if (!e || !e.currentTarget) {
			return
		}
		var id = e.currentTarget.id
		var text = ''
		var item = list.find(item => {
			return item.value == id
		})
		if (!item) {
			return
		}
		if (id == 'sixian') {
			text = '请输入大于0小于300字符的长度'
		} else {
			text = `请输入${item.label}头部内容，或数量。`
		}
		$('.tip').html(text)
		$('.input-box').show()
		var selectedItem = $('.selected')
		if (selectedItem && selectedItem.length > 0) {
			selectedItem.removeClass('selected')
		}
		insert_type = id
		$(`#${insert_type}`).addClass('selected')
		onCancel()
	}

	function onInert() {
		var contentBefore = $('#input').val()
		var num = contentBefore * 1
		var isNum = !isNaN(num)
		var imgWidth = 50
		var item = list.find(item => {
			return item.value == insert_type
		})
		if (!item) {
			return
		}
		$('.preview1').show()
		var content = ''
		content = `<p>`
		var img_url = ''
		if (insert_type == 'sixian') {
			var html = ''
			img_url = `https://by-base.oss-cn-shenzhen.aliyuncs.com/${item.value}.png`
			if (num > 0 && num<=300){
				let space = ''
				for (let index = 0; index < num; index++) {
				  space += '&nbsp;'
				}
				html = '<span class="fourwire-region fixed-size" style="background-image: url(https://by-base.oss-cn-shenzhen.aliyuncs.com/sixian.png) !important; background-size: 2.6em; background-repeat:repeat-x;background-position:0% 50%; line-height:3em;height:3em;width:auto;display: inline-block;">'+ space +'</span>'
				$('.preview').addClass('width-fit')
			}
			content += html
		} else {
			var imgName = isNum ? item.type1 : item.type2
			var img_url = `https://by-base.oss-cn-shenzhen.aliyuncs.com/${imgName}.png`
			if (isNum) {
				for (var i = 0; i < num; ++i) {
					content += `<strong class="pinyin tianzige" data-content-before=""><img class="fixed-size tianzige-region" width="${imgWidth}" src="${img_url}"/></strong>`
				}
			} else {
				content += `<strong class="pinyin tianzige" data-content-before=${contentBefore}><img class="fixed-size tianzige-region" width="${imgWidth}" src="${img_url}"/></strong>`
			}
			$('.preview').addClass('width-fit')
		}
		content += `</p>`
		$('.preview').html(content)
		$('#copy').show()
	}

	function onCancel() {
		$('#input').val('')
		$('.preview').html('')
		$('.preview').removeClass('width-fit')
		$('#image-preview').html('')
		$('#copy').hide()
		$('.preview1').hide()
	}

	function generateImage() {
		var element = document.querySelector('.preview');  // 指定你想要转换的元素
		$('#image-preview').html('')
		html2canvas(element, {
			useCORS: true
		}).then(function(canvas){
			// 这部分代码将转换出的canvas转化为图片并添加到body中
			var img = document.createElement('img');
			img.src = canvas.toDataURL('image/png');
			var imgPreview = document.getElementById('image-preview');
			if (imgPreview) {
				imgPreview.appendChild(img)
			}
			window.Asc.plugin.sendToPlugin('onWindowMessage', {
				type: 'insertSymbolImage',
				data: {
					src: img.src,
					width: canvas.width * (25.4 / 96), // mm
					height: canvas.height * (25.4 / 96) // mm
				}
			})
		});
	}
	function onCopy() {
		html2canvas(document.querySelector('.preview'), {useCORS: true}).then(canvas => {
			canvas.toBlob(function(blob) {
				const item = new ClipboardItem({ "image/png": blob });
				navigator.clipboard.write([item]).then(function() {
					console.log('Image copied to clipboard');
				}, function(error) {
					console.error('Unable to write to clipboard. Error:');
					console.error(error);
				});
			});
		});
	}
	window.Asc.plugin.attachEvent('initSymbols', function (message) {
		init()
	})
})(window, undefined)