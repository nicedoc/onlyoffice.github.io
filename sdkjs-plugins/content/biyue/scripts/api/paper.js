import request from '../request.js'
import authRequest from '../request_auth.js'

// * 如果涉及到post类型的请求,需要告知接口是否需要使用application/json的方式来请求,避免因为特殊符号导致接口无法接收到参数 *
// 属于正式在用的接口 start
// ------------------------------START------------------------------------
function getQuesType(paper_uuid, content_list) {
	return request({
		url: '/oodoc/get/ques_type',
		method: 'POST',
		isJsonData: true,
		headers: {
			'Content-Type': 'application/json;charset=UTF-8',
		},
		data: {
			paper_uuid,
			content_list,
		},
	})
}

function reqComplete(tree, version) {
	return request({
		url: '/oodoc/complete',
		method: 'POST',
		isJsonData: true,
		headers: {
			'Content-Type': 'application/json;charset=UTF-8',
		},
		data: {
			tree,
			version
		},
	})
}

function reqPaperInfo(paper_uuid) {
	return request({
		url: '/oodoc/get/info',
		method: 'get',
		params: {
			paper_uuid
		},
	})
}

function reqSaveQuestion(paper_uuid, question_uuid, content_type, content_number, scores) {
	return request({
		url: '/oodoc/save/question',
		method: 'POST',
		isJsonData: true,
		headers: {
			'Content-Type': 'application/json;charset=UTF-8',
		},
		data: {
			paper_uuid,	// 试卷UUID	是	[string]	933f5fa3-3f1c-4020-887f-206d67d6741f	查看
			question_uuid,	// 题目UUID	是	[string]	8bd18eee-478e-4d0b-975b-d4d224443d92	查看
			content_type,	// 题型	是	[int]	2	查看
			content_number, //	题号（题型为6时非必填）		[string]	a	查看
			scores
		},
	})
}

function reqSubjectMarkTypes(subject_value) {
	return authRequest({
		url: '/person/subject/mark_types',
		method: 'get',
		params: {
			subject_value
		},
	})
}

function reqSaveInfo(paper_uuid, info) {
	return request({
		url: '/oodoc/save/info',
		method: 'POST',
		isJsonData: true,
		headers: {
			'Content-Type': 'application/json;charset=UTF-8',
		},
		data: {
			paper_uuid,	// 试卷UUID	是	[string]	933f5fa3-3f1c-4020-887f-206d67d6741f	查看
			info,
		},
	})
}


/**
 * 上传试卷预览图
 * http://api.dcx.com/#/home/project/inside/api/detail?groupID=479&childGroupID=498&apiID=1931&projectName=%E7%AD%86%E6%9B%B0%20-%20%E9%A2%98%E5%BA%93&projectID=39
 */
function paperUploadPreview(data) {
  return request({
    url: '/paper/upload/preview',
    method: 'post',
    isFileData: true,
    data: data
  })
}

/**
 * http://api.dcx.com/#/home/project/inside/api/detail?groupID=-1&apiID=6343&projectName=%E7%AD%86%E6%9B%B0%20-%20%E9%A2%98%E5%BA%93&projectID=39
 * 保存试卷坐标--以 json 格式
 */
function paperSavePosition(paper_uuid, position, extra_info, comment_custom) {
	position = position ? JSON.stringify(position) : ''
	extra_info = extra_info ? JSON.stringify(extra_info) : ''
	return request({
		url: '/paper/save/position/json',
		method: 'post',
		isJsonData: true,
		headers: {
			'Content-Type': 'application/json',
		},
		data: {
			paper_uuid,
			position,
			extra_info,
			comment_custom,
		},
	})
}

// 属于正式在用的接口 END
// ------------------------------END------------------------------------




// -----------------------下列属于实验测试用的旧接口----------------------

/**
 * 创建试卷结构
 * 链接：http://api.dcx.com/#/home/project/inside/api/detail?groupID=719&childGroupID=884&apiID=4598&projectName=%E7%AD%86%E6%9B%B0%20-%20%E9%A2%98%E5%BA%93&projectID=39
 */
function structAdd({ paper_uuid, name, rich_name }) {
	return request({
		url: '/online_v2/struct/add',
		method: 'post',
		data: {
			paper_uuid,
			name,
			rich_name,
		},
	})
}

/**
 * 删除试卷结构
 * 链接：http://api.dcx.com/#/home/project/inside/api/detail?groupID=719&childGroupID=884&apiID=4608&projectName=%E7%AD%86%E6%9B%B0%20-%20%E9%A2%98%E5%BA%93&projectID=39
 */
function structDelete(paper_uuid, struct_id) {
	return request({
		url: '/online_v2/struct/delete',
		method: 'post',
		data: {
			paper_uuid,
			struct_id,
		},
	})
}

/**
 * 清空试卷结构
 * 链接：http://api.dcx.com/#/home/project/inside/api/detail?groupID=-1&apiID=4613&projectName=%E7%AD%86%E6%9B%B0%20-%20%E9%A2%98%E5%BA%93&projectID=39
 */
function structEmpty({ paper_uuid }) {
	return request({
		url: '/online_v2/struct/empty',
		method: 'post',
		data: {
			paper_uuid,
		},
	})
}

/**
 * 编辑试卷结构名称
 * 链接：http://api.dcx.com/#/home/project/inside/api/detail?groupID=719&childGroupID=884&apiID=4612&projectName=%E7%AD%86%E6%9B%B0%20-%20%E9%A2%98%E5%BA%93&projectID=39
 */
function structRename(paper_uuid, struct_id, name) {
	return request({
		url: '/online_v2/struct/rename',
		method: 'post',
		data: {
			paper_uuid,
			struct_id,
			name,
		},
	})
}

function questionCreate(formData) {
	return request({
		url: '/online_v2/question/create',
		method: 'post',
		data: {
			paper_uuid: formData.paper_uuid,
			content: formData.content,
			blank: formData.blank,
			type: formData.type,
			score: formData.score,
			no: formData.no,
			struct_id: formData.struct_id,
		},
	})
}

/**
 * 删除一道题
 * 链接：http://api.dcx.com/#/home/project/inside/api/detail?groupID=719&childGroupID=888&apiID=4675&projectName=%E7%AD%86%E6%9B%B0%20-%20%E9%A2%98%E5%BA%93&projectID=39
 */
function questionDelete(paper_uuid, question_uuid, confirm) {
	return request({
		url: '/online_v2/question/delete',
		method: 'post',
		data: {
			paper_uuid,
			question_uuid,
			confirm,
		},
	})
}

/**
 * http://api.dcx.com/#/home/project/inside/api/detail?groupID=-1&apiID=4941&projectName=%E7%AD%86%E6%9B%B0%20-%20%E9%A2%98%E5%BA%93&projectID=39
 * 在线编辑-编辑题目题干
 */
function questionUpdateContent(data) {
	return request({
		url: '/paper/online/question/update/content',
		method: 'POST',
		data: data,
	})
}

function questionBatchEditScore({ paper_uuid, ques_score }) {
	// ques_score = ques_score ? JSON.stringify(ques_score) : ''
	return request({
		url: '/question/batch_edit/score',
		method: 'POST',
		isJsonData: true,
		headers: {
			'Content-Type': 'application/json;charset=UTF-8',
		},
		data: {
			paper_uuid,
			ques_score,
		},
	})
}

/**
 * http://api.dcx.com/#/home/project/inside/api/detail?groupID=-1&apiID=2022&projectName=%E7%AD%86%E6%9B%B0%20-%20%E9%A2%98%E5%BA%93&projectID=39
 * 试卷是否可定稿
 */
function paperCanConfirm(paper_uuid) {
	return request({
		url: '/paper/can/confirm',
		method: 'get',
		params: {
			paper_uuid,
		},
	})
}

/**
 * 添加试卷 - 从题库中导入
 * 链接：http://api.dcx.com/#/home/project/inside/api/detail?groupID=315&childGroupID=323&apiID=1793&projectName=%E7%AD%86%E6%9B%B0&projectID=35
 */
function examImportQuestionBank({
	exam_title,
	subject_code,
	subject_value,
	grade_code,
	page_num,
	ques_num,
	exam_score,
	ref_id,
	print_paper_type,
	paper_num,
	for_mask,
	exam_begin,
	exam_end,
	participants_num,
	is_for_grade,
	is_update,
	exam_type,
	scope_id,
	light_code,
}) {
	return authRequest({
		url: '/teacher/exam/import/question_bank',
		method: 'post',
		data: {
			exam_title, // 	试卷标题
			subject_code, // 	科目代号
			subject_value, // 科目ID，如果传入，以ID为准
			grade_code, // 	年级编号
			page_num, // 	试卷页数
			ques_num, // 	题目数量
			exam_score, // 	试卷分数
			ref_id, // 试卷在题库的ID
			print_paper_type, // 	纸张打印类型
			paper_num, // 	卷面页数
			for_mask, // 	需要批改的试卷
			exam_begin, // 	考试开始时间
			exam_end, // 	考试结束时间
			participants_num, // 	参与人数
			is_for_grade, // 	是否是年段卷
			is_update, // 	是否更新旧试卷
			exam_type, // 	试卷类型，默认是quize
			scope_id, // 	用途
			light_code, // 	淡码方案，默认非淡码
		},
	})
}

/**
 * 步骤2：上传试卷图片
 * 链接：http://api.dcx.com/#/home/project/inside/api/detail?groupID=315&childGroupID=323&grandSonGroupID=475&apiID=1794&projectName=%E7%AD%86%E6%9B%B0&projectID=35
 */
function examPageUpload(data) {
	return authRequest({
		url: '/teacher/exam/page/upload',
		method: 'post',
		isFileData: true,
		data: data,
	})
}

/**
 * 步骤3：上传试卷切题信息
 * 链接：http://api.dcx.com/#/home/project/inside/api/detail?groupID=315&childGroupID=323&grandSonGroupID=475&apiID=1798&projectName=%E7%AD%86%E6%9B%B0&projectID=35
 */

function examQuestionsUpdate(data) {
	return authRequest({
		url: '/teacher/exam/questions/update',
		method: 'post',
		isJsonData: true,
		headers: {
			'Content-Type': 'application/json',
		},
		data: data,
	})
}

/**
 * 获取我的年段（打印用）
 * 链接：http://api.dcx.com/#/home/project/inside/api/detail?groupID=-1&apiID=5564&projectName=%E7%AD%86%E6%9B%B0&projectID=35
 */
function teacherClassMyGradePrint() {
	return authRequest({
		url: '/teacher/class/my/grade/print',
		method: 'get',
	})
}
// 打印纸类型枚举
function teacherExamPrintPaperSelections() {
	return authRequest({
		url: '/teacher/exam/print_paper/selection',
		method: 'get',
	})
}

/**
 * http://api.dcx.com/#/home/project/inside/api/detail?groupID=-1&apiID=6156&projectName=%E7%AD%86%E6%9B%B0&projectID=35
 * 验证切图信息
 */
function paperValidatePosition(ques_workbook_id, paper_uuid, position, extra_info) {
  position = position ? JSON.stringify(position) : ''
  extra_info = extra_info ? JSON.stringify(extra_info) : ''
  return authRequest({
    url: '/quesbank/exam/validate/position',
    method: 'post',
    data: {
      ques_workbook_id,
      paper_uuid,
      position,
      extra_info
    }
  })
}
/**
 * 题干latex转图片
 * 链接：http://api.dcx.com/#/home/project/inside/api/detail?groupID=719&childGroupID=888&apiID=5737&projectName=%E7%AD%86%E6%9B%B0%20-%20%E9%A2%98%E5%BA%93&projectID=39
 */
function onLatexToImg(content) {
	return request({
	  url: '/online_v2/latex/to/img',
	  method: 'post',
	  data: {
		content
	  }
	})
}

export {
	getQuesType,
	reqComplete,
	reqPaperInfo,
	reqSaveQuestion,
	reqSubjectMarkTypes,
	reqSaveInfo,
	structAdd,
	structDelete,
	structEmpty,
	structRename,
	questionCreate,
	questionDelete,
	questionUpdateContent,
	questionBatchEditScore,
	paperCanConfirm,
	examImportQuestionBank,
	examPageUpload,
	paperSavePosition,
	examQuestionsUpdate,
	teacherClassMyGradePrint,
	teacherExamPrintPaperSelections,
  	paperUploadPreview,
  	paperValidatePosition,
	onLatexToImg
}
