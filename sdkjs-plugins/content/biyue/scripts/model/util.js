function showCom(comname, v) {
	var com = $(comname)
	if (!com) {
		return
	}
	if (v) {
		com.show()
	} else {
		com.hide()
	}
}

function updateText(comname, v) {
	var com = $(comname)
	if (!com) {
		return
	}
	com.html(v)
}

function addClickEvent(comname, func) {
	var com = $(comname)
	if (!com) {
		return
	}
	com.on('click', func)
}

// 关闭其他select弹窗
function closeOtherSelect(exceptId) {
	var list = $('.open')
	if (list.length == 0) {
		return
	}
	for (var i = 0; i < list.length; ++i) {
		if (!list[i]) {
			continue
		}
		if (list[i].id == exceptId) {
			continue
		}
		$(`#${list[i].id}`).removeClass('open')
	}
}

export { showCom, updateText, addClickEvent, closeOtherSelect }