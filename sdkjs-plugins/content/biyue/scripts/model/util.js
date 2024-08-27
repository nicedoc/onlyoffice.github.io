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
function getListByMap(map, keyname = 'value', labelname = 'label') {
	if (!map || typeof map !== 'object') {
		return []
	}
	var list = []
	Object.keys(map).forEach(e => {
		list.push({
			[`${keyname}`]: e,
			[`${labelname}`]: map[e]
		})
	})
	return list
}

function getInfoForServerSave() {
	var quesmap = window.BiyueCustomData.question_map || {}
	var treemap = {}
	Object.keys(quesmap).forEach(id => {
		treemap[id] = Object.assign({}, quesmap[id])
		delete treemap[id].text
		delete treemap[id].ques_default_name
	})
	var info = {
		node_list: window.BiyueCustomData.node_list || [],
		question_map: treemap,
		client_node_id: window.BiyueCustomData.client_node_id,
		time: window.BiyueCustomData.time
	}
	return JSON.stringify(info)
}

export { showCom, updateText, addClickEvent, closeOtherSelect, getListByMap, getInfoForServerSave }