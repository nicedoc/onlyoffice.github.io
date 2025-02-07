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
	addTypeEvent(comname, 'click', func)
}

function addTypeEvent(comname, eventType, func) {
	var com = $(comname)
	if (!com) {
		return
	}
	com.off(eventType, func)
	com.on(eventType, func)
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
	if (!window.BiyueCustomData) {
		return
	}
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

function setBtnLoading(elementId, isLoading) {
	var element = $(`#${elementId}`)
	if (!element) {
		return
	}
	if (isLoading) {
		element.append('<span class="loading-spinner"></span>')
		element.addClass('btn-unable')
	} else {
		element.removeClass('btn-unable')
		var children = element.find('.loading-spinner')
		if (children) {
			children.remove()
		}
 	}
}

function isLoading(elementId) {
	var element = $(`#${elementId}`)
	if (!element) {
		return
	}
	var loading = element.find('.loading-spinner')
	return loading && loading.length
}

function getYYMMDDHHMMSS() {
	var date = new Date()
	var year = date.getFullYear()
	var month = date.getMonth() + 1
	var day = date.getDate()
	var hour = date.getHours()
	var minute = date.getMinutes()
	var second = date.getSeconds()
	return `${year}_${month}_${day}_${hour}:${minute}:${second}`	
}

function getFixedValue(v, fractionDigits = 1) {
	var v = v || 0
	v *= 1
	v = v.toFixed(fractionDigits) * 1
	return v + ''
}

function handleCommandError(window, params, error) {
	console.error(params, error)
	if (params) {
		if (!window.error_list) {
			window.error_list = []
		}
		window.error_list.push({ params: params, error: error })
		updateText('#printStack', 'error数量：' + window.error_list.length)
		$('#printStack').css('color', "#ff0000")
	}
}
function updateHintById(id, message, color, duration = 1500) {
    var $tooltip = $(id);
    if ($tooltip.length === 0) {
        return;
    }
    $tooltip.text(message);
    $tooltip.css({
        'color': color || '#999',
        'display': 'block'
    });
    setTimeout(function() {
        $tooltip.css('display', 'none');
    }, duration);
}
export { showCom, updateText, addClickEvent, closeOtherSelect, getListByMap, getInfoForServerSave, setBtnLoading, isLoading, getYYMMDDHHMMSS, addTypeEvent, getFixedValue, handleCommandError, updateHintById }