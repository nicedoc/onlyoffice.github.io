import {
	paperCanConfirm,
	teacherClassMyGradePrint,
	examImportQuestionBank,
	teacherExamPrintPaperSelections,
	examPageUpload,
} from './api/paper.js'
import { setXToken } from './auth.js'

;(function (window, undefined) {
	let source_data = {}
	let paper_info = null
	let export_grade_code = ''
	let print_paper_type = ''
	let exam_import_res = {}
	window.Asc.plugin.init = function () {
		console.log('examExport init')
		window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'PaperMessage' })
	}

	function init() {
		$('#examtitle').html(paper_info.paper.title)
		getGradeSelections()
		getPrintSelections()
		$('#btnCanConfirm').on('click', confirmPaper)
		$('#confirm').on('click', onConfirm)
		$('#uploadImages').on('click', uploadImages)
		const fileInput = document.getElementById('fileInput')
		const preview = document.getElementById('preview')
		fileInput.addEventListener('change', () => {
			if (preview) {
				preview.innerHTML = '' // 清空预览区域
			}
			const files = fileInput.files
			Array.from(files).forEach((file) => {
				const reader = new FileReader()
				reader.onload = () => {
					const img = document.createElement('img')
					img.src = reader.result
					img.style.maxWidth = '100px' // 设置预览图片的最大宽度
					preview.appendChild(img)
				}
				reader.readAsDataURL(file)
			})
		})
	}

	function confirmPaper() {
		paperCanConfirm(paper_info.paper.paper_uuid)
			.then((res) => {
				console.log('定稿成功')
			})
			.catch((error) => {
				console.log(error)
			})
	}

	function getGradeSelections() {
		teacherClassMyGradePrint()
			.then((res) => {
				const subjectValue = paper_info.paper.subject_id
				const list = res.data.grade_list || []
				const gradeOptions = {}
				let grade_content = ''
				list.forEach((item) => {
					if (item.subject_value === subjectValue) {
						if (!gradeOptions[item.grade_code]) {
							gradeOptions[item.grade_code] = {
								name: item.grade_name,
								code: item.grade_code,
							}
							if (export_grade_code == '') {
								export_grade_code = item.grade_code
							}
							grade_content +=
								'<input type="radio" name="grade" value="' +
								item.grade_code +
								'" /> ' +
								item.grade_name
						}
					}
				})
				$('#grades').html(grade_content)
				$('#grades').on(
					'change',
					'input[type="radio"][name="grade"]',
					function () {
						export_grade_code = $(this).val()
						console.log('export_grade_code', export_grade_code)
					}
				)
			})
			.catch((res) => {
				console.log(res)
			})
	}
	function getPrintSelections() {
		teacherExamPrintPaperSelections()
			.then((res) => {
				if (res.data) {
					let print_content = ''
					for (var key in res.data) {
						print_content +=
							'<input type="radio" name="print" value="' +
							res.data[key].type +
							'" /> ' +
							res.data[key].name
					}
					print_paper_type = 'k16'
					$('#printType').html(print_content)
					$(
						`#printType input[type="radio"][name="print"][value=${print_paper_type}]`
					)
						.prop('checked', true)
						.change()
					$('#printType').on(
						'change',
						'input[type="radio"][name="print"]',
						function () {
							print_paper_type = $(this).val()
						}
					)
				}
			})
			.catch((err) => {
				console.log(err)
			})
	}

	function onConfirm() {
		if (!paper_info.paper) {
			console.log('paper_info.paper is null')
			return
		}
		if (paper_info.paper.total_score == 0) {
			var score = 0
			if (paper_info.info && paper_info.info.questions) {
				for (var i = 0; i < paper_info.info.questions.length; ++i) {
					score += paper_info.info.questions[i].score * 1
				}
				paper_info.paper.total_score = score
			}
		}
		let exam_type =
			paper_info.page_type === 'exam_week' ? 'weekly_quiz' : 'dtk_quiz'
		if (paper_info.page_type === 'exam_exercise') {
			exam_type = 'exercise'
		}
		examImportQuestionBank({
			exam_title: paper_info.paper.title, // 试卷标题
			subject_code: paper_info.paper.subject_code, // 科目代号
			subject_value: paper_info.paper.subject_id, // 科目ID，如果传入，以ID为准
			grade_code: export_grade_code, // 	年级编号
			page_num: 8, // 	试卷页数 临时写死
			ques_num: paper_info.paper.question_num, // 	题目数量
			exam_score: paper_info.paper.total_score, // 	试卷分数
			ref_id: paper_info.paper.paper_uuid, // 试卷在题库的ID
			print_paper_type: print_paper_type, // 	纸张打印类型
			paper_num: 8, // 临时写死
			exam_type: exam_type,
		})
			.then((res) => {
				console.log('importExam res', res)
				exam_import_res = res.data
				window.Asc.plugin.sendToPlugin('onWindowMessage', {
					type: 'exportExamSuccess',
					data: {
						grade_code: export_grade_code,
						print_paper_type: print_paper_type,
						exam_no: res.data.exam_no,
						exam_id: res.data.exam_id,
					},
				})
			})
			.catch((error) => {
				console.log(error)
			})
	}

	async function uploadImages() {
		console.log('uploadImages')
		const fileInput = document.getElementById('fileInput')
		if (!fileInput) {
			return
		}
		var pid = 1
		for (var file of fileInput.files) {
			console.log('file', file)
			const data = new FormData()
			// data.append('exam_no',  exam_import_res.exam_no)
			data.append('exam_no', 'math41711qzsnayz20240520201848Kaf')
			data.append('pid', pid)
			data.append('page_image', file)
			data.append('page_type', 'paper')

			await examPageUpload(data)
				.then((res) => {
					console.log(res)
					pid++
				})
				.catch((error) => {
					console.log(error)
				})
		}
		console.log('所有底图上传成功')
	}

	window.Asc.plugin.attachEvent('initPaper', function (message) {
		console.log('examExport 接收的消息', message)
		setXToken(message.xtoken)
		paper_info = message.paper_info
		source_data = message
		init()
	})
})(window, undefined)
