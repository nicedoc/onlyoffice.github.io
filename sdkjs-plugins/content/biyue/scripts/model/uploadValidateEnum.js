var UPLOAD_VALIDATE_RESULT = {
	OK: 1,
	NOT_WORKBOOK_LAYOUT: -1, // 未获取到练习册配置的尺寸
	NOT_QUESTION: -2, // 没有符合条件的题目，请先进行全量更新!
	MISS_QUESTION_UUID: -3 , // 有题目没有分配uuid，请先进行全量更新。
	QUESTION_ERROR: -4, // 题目不符合规范
	FEATURE_ERROR: -5, // 功能区不符合规范
}

var MARK_METHOD_TYPE = {
	TAG_WRONG: 1, // '标错',
	SCORE: 2, // '打分',
	HALF_PAIR_REGION: 3, // 半对
	CHOICE_OPTION: 4 // 答题卡里的选择题选项ABCD，用于服务端识别
}

export {
	UPLOAD_VALIDATE_RESULT,
	MARK_METHOD_TYPE
}