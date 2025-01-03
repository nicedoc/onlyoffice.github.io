function getDataByParams(params) {
	if(!params) {
		return null
	}
	if (!params.client_id) {
		if (params.parentTag) {
			params.client_id = params.parentTag.client_id
		}
		if (!params.client_id) {
			return null
		}
	}
	var client_id = params.client_id
	var question_map = window.BiyueCustomData.question_map
	if (!question_map) {
		return null
	}
	var quesData = question_map ? question_map[client_id] : null
	if (client_id) {
		var node_list = window.BiyueCustomData.node_list || []
		var nodeData = node_list.find(e => {
			return e.id == client_id
		})
		if (nodeData) {
			if (nodeData.level_type == 'struct') {
				return {
					client_id: client_id,
					level_type: nodeData.level_type,
					data: question_map[nodeData.id]
				}
			} else if (nodeData.level_type == 'text') {
				return {
					client_id: client_id,
					level_type: nodeData.level_type,
				}
			} else {
				if (nodeData.merge_id && question_map[nodeData.merge_id]) {
					var ids = question_map[nodeData.merge_id].ids || []
					if (ids.find(e => {
						return e == client_id
					})) {
						quesData = question_map[nodeData.merge_id]
						client_id = nodeData.merge_id
					}
					if (nodeData.cell_ask && nodeData.write_list && nodeData.write_list.length) {
						params.regionType = 'write'
						client_id = nodeData.write_list[0].id
					}
				}
			}
		}
	}
	var ques_client_id = 0
	var findIndex = -1
	if (params.regionType == 'write') {
		var keys = Object.keys(question_map)
		for (var i = 0; i < keys.length; ++i) {
			var ask_list = question_map[keys[i]].ask_list || []
			findIndex = ask_list.findIndex(e => {
				if (e.other_fields && e.other_fields.includes(client_id)) {
					return true
				}
				return e.id == client_id
			})
			if (findIndex >= 0) {
				ques_client_id = keys[i]
				break
			}
		}
	} else if (quesData && quesData.level_type == 'question') {
		ques_client_id = client_id
	} else if (params.regionType == 'question') {
		ques_client_id = client_id
	}
	if (!ques_client_id) {
		return null
	}
	quesData = question_map[ques_client_id]
	if (!quesData) {
		return null
	} else {
		return {
			ques_client_id: ques_client_id,
			client_id: client_id,
			level_type: quesData.level_type,
			data: quesData,
			findIndex: findIndex
		}
	}
}
function getFocusAskData(quesId, index) {
	var quesData = window.BiyueCustomData.question_map[quesId]
	if (!quesData || !quesData.ask_list || index >= quesData.ask_list.length) {
		return null
	}
	var nlist = window.BiyueCustomData.node_list || []
	var ids = quesData.is_merge ? quesData.ids : [quesId]
	var focusList = []
	var targetIds = [quesData.ask_list[index].id]
	if (quesData.ask_list[index].other_fields) {
		targetIds = targetIds.concat(quesData.ask_list[index].other_fields)
	}
	for (var qid of ids) {
		var nodeData = nlist.find(e => {
			return e.id == qid
		})
		if (nodeData) {
			for (var wData of nodeData.write_list) {
				var targetIndex = targetIds.findIndex(e => {
					return e == wData.id
				})
				if (targetIndex >= 0) {
					focusList.push(wData)
					targetIds.splice(targetIndex, 1)
					if (targetIds.length == 0) {
						return focusList
					}
				}
			}
		}
	}
	return null
}
// 是否文本题
function isTextMode(quesMode) {
	return quesMode == 6
}
// 作答模式是否选择题
function isChoiceMode(quesMode) {
	return quesMode == 1 || quesMode == 5
}

export {
	getDataByParams,
	getFocusAskData,
	isTextMode,
	isChoiceMode
}