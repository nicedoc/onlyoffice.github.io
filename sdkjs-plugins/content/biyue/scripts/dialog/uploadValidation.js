import { addClickEvent, isLoading, setBtnLoading, showCom, updateText } from "../model/util.js"
import { UPLOAD_VALIDATE_RESULT } from '../model/uploadValidateEnum.js'
(function (window, undefined) {
	var validate_info = {}
	window.Asc.plugin.init = function () {
		window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'initDialog', initmsg: 'uploadValidationMessage' })
	}
	function init() {
		var msgCodes = [UPLOAD_VALIDATE_RESULT.NOT_WORKBOOK_LAYOUT, UPLOAD_VALIDATE_RESULT.NOT_QUESTION, UPLOAD_VALIDATE_RESULT.MISS_QUESTION_UUID]
		console.log('validate_info', validate_info)
		var basictext = ''
		if (validate_info.exam_info) {
			basictext += '<div>基础信息：</div>'
			basictext += `<div>结构${validate_info.exam_info.struct_count}个、${validate_info.exam_info.type_info}</div>`
		}
		updateText('#exambasic', basictext)
		var message = ''
		if (msgCodes.includes(validate_info.code)) {
			showCom('#message', true)
			switch (validate_info.code) {
				case UPLOAD_VALIDATE_RESULT.NOT_WORKBOOK_LAYOUT:
					message += `<div style="color:#f00">未获取到练习册配置的尺寸</div>`
					break
				case UPLOAD_VALIDATE_RESULT.NOT_QUESTION:
					message += `<div style="color:#f00">没有符合条件的题目，请先检查题目类型</div>`
					message += `<div class="clicked all-update batchbtn" data-id="toBatchType">前往批量题型</div>`
					break
				case UPLOAD_VALIDATE_RESULT.MISS_QUESTION_UUID:
					message += `<div style="color:#f00">有题目没有分配uuid，请先进行全量更新</div>`
					message += `<div class="clicked all-update" id="toAllUpdate">全量更新</div>`
					break
				default: break
			}
		} else if (validate_info.code == UPLOAD_VALIDATE_RESULT.QUESTION_ERROR) {
			if (validate_info.detail) {
				var detail = validate_info.detail
				message += addQues(detail.ques_score_err, '题目未设置分数', 'score')
				message += addQues(detail.ask_score_err, '小问总分和题目总分不一致', 'score')
				message += addQues(detail.ask_score_empty_err, '小问分数为空', 'score')
				message += addQues(detail.ques_type_err, '题目缺少题型', 'type')
				message += addQues(detail.title_region_err, '缺少题干区')
				message += addQues(detail.write_ask_region_err, '缺少作答区')
				message += addQues(detail.correct_ask_region_err, '小问订正框未设置')
			}
			if (validate_info.feature && validate_info.feature.length) {
				message += `<div>未找到${validate_info.feature.join('、')}，请检查<span class="clicked clr1" id="toFeature">智批元素</span></div>`
			}
		}
		updateText('#message', message)
		addClickEvent('#toAllUpdate', toAllUpdate)
		addClickEvent('#reCheck', onRecheck)
		addClickEvent('#toFeature', toFeature)
		$('body').off('click', '.batchbtn', onBatchButton);
		$('body').on('click', '.batchbtn', onBatchButton);
		$('body').off('click', '.locate', onLocate);
		$('body').on('click', '.locate', onLocate);
	}

	function onBatchButton() {
		var id = $(this).data('id');
		// 执行你的代码
		onCommand(id)
	}

	function onLocate() {
		var id = $(this).data('id');
		// 执行你的代码
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'uploadValidationMessage',
			cmd: 'locate',
			data: id
		})
	}

	function onCommand(cmd) {
		window.Asc.plugin.sendToPlugin('onWindowMessage', {
			type: 'uploadValidationMessage',
			cmd: cmd
		})
	}
	function toAllUpdate() {
		onCommand('toAllUpdate')
	}

	function toFeature() {
		onCommand('toFeature')
	}

	function addQues(list, title, type) {
		if (!list || list.length == 0) {
			return ''
		}
		var strBatch = ''
		if (type == 'score') {
			strBatch = `<div class="clicked clr1 clicked batchbtn" data-id="toBatchScore">批量分数</div>`
		} else if (type == 'type') {
			strBatch = `<div class="clicked clr1 clicked batchbtn" data-id="toBatchType">批量题型</div>`
		}
		var message = `<div class="row-between">
						<div>${title}</div>
						${strBatch}
					</div>`
		message += '<div class="ques-wrapper">'
		list.forEach((item, index) => {
			message += `<div class="row-between">
				<div class="ques clicked" data-id=${item.ques_id} title="${item.text}">${item.ques_name}</div>
				<div class="clicked clr1 locate" data-id=${item.ques_id}>定位</div>
			</div>`
		});
		message += '</div>'
		return message
	}

	function onRecheck() {
		if (isLoading('reCheck')) {
			return
		}
		setBtnLoading('reCheck', true)
		onCommand('reCheck')
	}

	window.Asc.plugin.attachEvent('uploadValidationMessage', function (message) {
		console.log('接收的消息', message)
		validate_info = message.validate_info || {}
		setBtnLoading('reCheck', false)
		init()
	})
})(window, undefined)
