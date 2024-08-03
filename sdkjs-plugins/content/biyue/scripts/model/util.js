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

export { showCom, updateText, addClickEvent }