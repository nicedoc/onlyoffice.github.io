// 获取试卷树
function getExamTree(node_list, question_map) {
	if (!node_list || !question_map) {
		return
	}
	var list = []
	var handled_ids = []
	var pre_struct = ''
	for (var key = 0; key < node_list.length; key++) {
		var quesId = node_list[key].id
      	let item = question_map[quesId] || ''
		if (node_list[key].level_type == 'question') {
			if (node_list[key].merge_id) {
				quesId = node_list[key].merge_id
				if (handled_ids.includes(quesId)) {
					continue
				}
			}
		}
		item = question_map[quesId]
		if (!item) {
			continue
		}
		handled_ids.push(quesId)
		if (item.level_type == 'struct') {
			list.push({
				level_type: 'struct',
				text: item.text.split('\r\n')[0],
				id: item.id,
				items: []
			})
		} else if (item.level_type == 'question') {

		}
	}
}

export {
	getExamTree
}