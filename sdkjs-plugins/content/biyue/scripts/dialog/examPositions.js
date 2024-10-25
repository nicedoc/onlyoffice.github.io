;(function (window, undefined) {
	let info = null
	let position_list = []
	let pages = {}
	let ZONE_TYPE = {
		AGAIN: 15, // 再练区域
		STATISTICS: 14, // 统计
		PASS: 11, // 通过
		SELF_EVALUATION: 9, // 自定义评价
		THER_EVALUATION: 10, // 教师评价
		END: 16, // 完成批改
		IGNORE: 18, // 忽略区
		CHECK: 28, // 检查区
		SEALING_LINE: 100, // 密封线
	}
	window.Asc.plugin.init = function () {
		console.log('examPositions init')
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'positionsMessage',
		})
	}

	function addType(zone_type, name, x, y, w, h, draw_type) {
		var groupDiv = $('<div>', { class: 'group' })
		var groupNameDiv = $('<div>', { class: 'group-name' }).text(name)
		groupDiv.append(groupNameDiv)
		switch (zone_type) {
			case ZONE_TYPE.AGAIN:
				groupDiv.append(addField(zone_type, 1, 1, x, y, w, h))
				addToList(zone_type, 1, 1, x, y, w, h, 'shape')
				break
			case ZONE_TYPE.END:
			case ZONE_TYPE.PASS:
			case ZONE_TYPE.SELF_EVALUATION:
			case ZONE_TYPE.THER_EVALUATION:
				groupDiv.append(addField(zone_type, 1, info.pages.length, x, y, w, h))
				addToList(zone_type, 1, info.pages.length, x, y, w, h, 'image')
				break
			case ZONE_TYPE.STATISTICS:
				for (var i = 1; i <= info.pages.length; i++) {
					groupDiv.append(addField(zone_type, i, i, x, y, w, h))
					addToList(zone_type, i, i, x, y, w, h, 'shape')
				}
				break
			case ZONE_TYPE.CHECK:
				groupDiv.append(addField(zone_type, 1, 1, x, y, w, h))
				addToList(zone_type, 1, 1, x, y, w, h, 'image')
				break
			default:
				break
		}
		$('.list').append(groupDiv)
	}

	function addToList(zone_type, v, pagenum, x, y, w, h, draw_type) {
		position_list.push({
			zone_type: zone_type * 1,
			page_num: pagenum * 1 - 1,
			v: v,
			x: x * 1,
			y: y * 1,
			w: w * 1,
			h: h * 1,
			draw_type: draw_type,
		})
	}

	function addField(
		zone_type,
		zone_index,
		page_num,
		initx,
		inity,
		initw,
		inith
	) {
		var zoneDiv = $('<div>', {
			class: 'zone-item',
		})
		var html = `<span class="label-title">Page:</span><span id=${zone_type}_${zone_index}_page>${page_num}</span><span class="label-title">X:</span><div class="el-input"><input id=${zone_type}_${zone_index}_x value=${initx} type="text" autocomplete="off" class="el-input__inner"></input>
    </div><span class="label-title">Y:</span><div class="el-input"><input id=${zone_type}_${zone_index}_y value=${inity} type="text" autocomplete="off" class="el-input__inner"></input></div><span class="label-title">W:</span><div class="el-input"><input id=${zone_type}_${zone_index}_w value=${initw} type="text" autocomplete="off" class="el-input__inner"></input></div><span class="label-title">H:</span><div class="el-input"><input id=${zone_type}_${zone_index}_h value=${inith} type="text" autocomplete="off" class="el-input__inner"></input></div><div id=${zone_type}_${zone_index}_unuse class="btn-cancel">删除</div><div id=${zone_type}_${zone_index}_use class="btn-confirm">应用</div>`
		zoneDiv.html(html)
		$(`#${zone_type}_${zone_index}_unuse`).on('click', function () {
			console.log('删除')
		})
		$(`#${zone_type}_${zone_index}_use`).on('click', function () {
			console.log('应用')
		})
		return zoneDiv
	}

	function onBtnClick(e) {
		if (!e) {
			return
		}
		var id
		if (e.target) {
			id = e.target.id
		} else if (e.currentTarget) {
			id = e.currentTarget.id
		}
		console.log('click', id)
		if (id == 'cancel') {
			window.Asc.plugin.sendToPlugin('onWindowMessage', {
				type: 'cancelDialog',
			})
		} else if (id == 'confirm') {
			window.Asc.plugin.sendToPlugin('onWindowMessage', {
				type: 'drawPosition',
				data: position_list,
			})
		} else if (id == 'allDelete') {
			window.Asc.plugin.sendToPlugin('onWindowMessage', {
				type: 'deletePosition',
				data: position_list,
			})
		} else {
			var strs = id.split('_')
			var zone_type = strs[0]
			var zone_index = strs[1]
			var fieldname = strs[2]
			if (fieldname == 'use') {
				// 应用
				var pagenum = $(`#${zone_type}_${zone_index}_page`).text()
				var x = $(`#${zone_type}_${zone_index}_x`).val()
				var y = $(`#${zone_type}_${zone_index}_y`).val()
				var w = $(`#${zone_type}_${zone_index}_w`).val()
				var h = $(`#${zone_type}_${zone_index}_h`).val()
				window.Asc.plugin.sendToPlugin('onWindowMessage', {
					type: 'drawPosition',
					data: [
						{
							zone_type: zone_type * 1,
							page_num: pagenum * 1 - 1,
							v: zone_index,
							x: x * 1,
							y: y * 1,
							w: w * 1,
							h: h * 1,
							draw_type: 'image',
						},
					],
				})
			} else if (fieldname == 'unuse') {
				// 删除
				console.log('删除')
				window.Asc.plugin.sendToPlugin('onWindowMessage', {
					type: 'deletePosition',
					data: [
						{
							zone_type: zone_type * 1,
							v: zone_index,
						},
					],
				})
			}
		}
	}

	function init() {
		addType(ZONE_TYPE.AGAIN, '再练', 29, 20, 20, 10)
		addType(
			ZONE_TYPE.END,
			'完成',
			parseInt(info.pages[info.pages.length - 1].XLimit - 35),
			parseInt(info.pages[info.pages.length - 1].YLimit - 20),
			20,
			10
		)
		addType(
			ZONE_TYPE.PASS,
			'通过',
			parseInt(info.pages[info.pages.length - 1].XLimit - 95),
			parseInt(info.pages[info.pages.length - 1].YLimit - 20),
			20,
			10
		)
		addType(
			ZONE_TYPE.STATISTICS,
			'统计',
			parseInt(info.pages[0].XLimit - 8),
			parseInt(info.pages[0].YLimit - 8),
			2,
			2
		)
		addType(
			ZONE_TYPE.CHECK,
			'检查',
			parseInt(info.pages[info.pages.length - 1].XLimit - 35),
			10,
			20,
			10
		)
		$('.btn-cancel').on('click', onBtnClick)
		$('.btn-confirm').on('click', onBtnClick)
		$('#allDelete').on('click', onBtnClick)
	}

	window.Asc.plugin.attachEvent('initInfo', function (message) {
		info = message
		console.log('examPosdata', message)
		init()
	})
})(window, undefined)
