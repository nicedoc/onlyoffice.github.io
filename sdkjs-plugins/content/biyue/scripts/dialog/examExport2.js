import {
	paperCanConfirm,
	teacherClassMyGradePrint,
	examImportQuestionBank,
	teacherExamPrintPaperSelections,
	examPageUpload,
	paperUploadPreview,
	paperSavePosition
	// paperValidatePosition
  } from '../api/paper.js'
  import { setXToken } from '../auth.js'
  import { isLoading, setBtnLoading } from '../model/util.js';
  ;(function (window, undefined) {
	let source_data = {}
	let paper_info = null
	let export_grade_code = ''
	let print_paper_type = ''
	let exam_import_res = {}
	let questionPositions = {}
	let preQuestionPositions = {} // 准备上传的题目坐标
	let preEvaluationPosition = {} // 准备上传的元素坐标
	let biyueCustomData = {}
	let img_file_list = []
	let page_size = ''
	let numericFields = ['x', 'y', 'w', 'h', 'page']
	let g_times = null
	window.Asc.plugin.init = function () {
		console.log('examExport init')
		window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'exportMessage' })
	}
	function init() {
		$('#examtitle').html(paper_info.paper.title)
		$('#btnCanConfirm').on('click', confirmPaper)
		$('#confirm').on('click', onConfirm)
		$('#uploadImages').on('click', uploadImages)
		$('#uploadPreview').on('click', uploadPreview)
		let fileInput = ''
		let preview = ''
		if (paper_info.workbook && paper_info.workbook.id > 0) {
			// 练习册试卷
			$('.info').hide()
			$('.workbook-info').show()
			$('.paper-title').html(paper_info.paper.title)

			fileInput = document.getElementById('workbook_fileInput')
			preview = document.getElementById('workbook_preview')
			onCheckPositions()
		} else {
			getGradeSelections()
			getPrintSelections()
			fileInput = document.getElementById('fileInput')
			preview = document.getElementById('preview')
		}
		if (fileInput){
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
		teacherExamPrintPaperSelections().then((res) => {
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
				$(`#printType input[type="radio"][name="print"][value=${print_paper_type}]`
				).prop('checked', true)
				.change()
				$('#printType').on('change', 'input[type="radio"][name="print"]', function () {
					print_paper_type = $(this).val()
				})
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

	// 上传试卷预览图
	async function uploadPreview() {
		g_times = []
		printTime('[点击开始上传]')
		const file_list = img_file_list || []
		if (!file_list) {
			return
		}
		if (isLoading('uploadPreview')) {
			return
		}
		setBtnLoading('uploadPreview', true)
		var pid = 1
		printTime('[开始上传预览图]')
		for (var file of file_list) {
			console.log('file', file)
			const data = new FormData()

			data.append('paper_uuid', biyueCustomData.paper_uuid)
			data.append('page_num', questionPositions.paper_info.pageCount)
			data.append('page_no', pid)
			data.append('file', file)

			await paperUploadPreview(data)
				.then((res) => {
					console.log(res)
					printTime(`上传预览图${pid}成功`)
					pid++
				})
				.catch((error) => {
					console.log(error)
				})
		}
		printTime('所有底图上传成功')
		onUpdatePostions()
	}
// 更新试卷切题信息
function onUpdatePostions() {
	let positions = preQuestionPositions
	var evaluationPosition = preEvaluationPosition
	console.log(
		'positions and evaluationPosition:',
		positions,
		evaluationPosition
	)
	paperSavePosition(
		biyueCustomData.paper_uuid,
		positions,
		evaluationPosition,
		''
	).then((res) => {
			printTime('保存位置成功')
			// 将窗口的信息传递出去
			window.Asc.plugin.sendToPlugin('onWindowMessage', {
				type: 'positionSaveSuccess',
				data: source_data,
			})
			setBtnLoading('uploadPreview', false)
		})
		.catch((error) => {
			setBtnLoading('uploadPreview', false)
			alert(error && error.message ? `上传失败:${error.message}` : '上传失败')
			console.log(error)
		})
	}

	function onCheckPositions() {
		page_size = ''
		if (paper_info && paper_info.workbook && paper_info.workbook.layout) {
			page_size = paper_info.workbook.layout
		} else {
			alert('未获取到练习册配置的尺寸')
			return
		}
		console.log('本地检查通过')
		initImg()
	}
	function initImg() {
		printTime('[开始下载预览图]')
		$('#workbookImgLoading').show()
		window.Asc.plugin.executeMethod('GetFileToDownload', ['PNG'], (res) => {
			console.log(res)
			const zipUrl = res
			let preview = document.getElementById('workbook_preview')
			if (preview) {
				preview.innerHTML = '' // 清空预览区域
			}

			img_file_list = [] // 清空上传用的file

			fetch(zipUrl)
				.then((response) => {
					if (!response.ok) {
						throw new Error('Failed to fetch zip file')
					}
					return response.arrayBuffer()
				})
				.then((arrayBuffer) => {
					return JSZip.loadAsync(arrayBuffer) // 使用 JSZip 加载 ArrayBuffer
				})
				.then(async (zip) => {
					printTime('[预览图下载完成]')
					try {
						let files = Object.entries(zip.files)
						// 保证图片的顺序
						files.sort((a, b) => {
							const aName = a[0]
							const bName = b[0]
							var aNum = aName.replace('image', '').split('.')[0] * 1
							var bNum = bName.replace('image', '').split('.')[0] * 1
							return aNum - bNum
						})
						console.log('========== 排序后')
						files.forEach((e) => {
							console.log('===== ', e[0])
						})

						for (const [filename, file] of files) {
							if (!file.dir) {
								// 只处理文件，忽略文件夹
								// 读取文件内容为 Blob
								const fileBlob = await file.async('blob')

								// 创建一个新的 File 对象并将其存储到数组中 用于后续上传
								const fileObj = new File([fileBlob], filename, {
									type: fileBlob.type,
								})
								img_file_list.push(fileObj)

								// 使用 FileReader 将 Blob 转换为 Data URL
								const reader = new FileReader()
								reader.onloadend = function () {
									// 创建并添加 img 元素到页面
									const img = document.createElement('img')
									img.src = reader.result
									img.style.maxWidth = '500px' // 设置预览图片的最大宽度
									preview.appendChild(img)
								}

								if (fileBlob instanceof Blob) {
									reader.readAsDataURL(fileBlob)
								} else {
									console.error('fileBlob 不是有效的 Blob 对象')
								}
							}
						}
						if (img_file_list && img_file_list.length > 0) {
							// 显示上传按钮
							$('#preview-label').show()
							$('#uploadPreview').show()
							$('#workbookImgLoading').hide()
						}
					} catch (error) {
						console.error('Error fetching or processing ZIP file:', error)
					}
				})
		})
	}
	function printTime(text) {
		if (!g_times) {
			g_times = []
		}
		g_times.push(Date.now())
		if (g_times.length > 1) {
			console.log(g_times[g_times.length - 1], text, (g_times[g_times.length - 1] - g_times[g_times.length - 2]) + 'ms')
		} else {
			console.log(g_times[g_times.length - 1], text)
		}
	}
	window.Asc.plugin.attachEvent('initPaper', function (message) {
		g_times = null
		printTime('[上传卷面] 初始化')
		console.log('examExport 接收的消息', message)
		setXToken(message.xtoken)
		paper_info = message.paper_info
		source_data = message
		biyueCustomData = message.biyueCustomData
		questionPositions = message.questionPositions || {}
		preQuestionPositions = message.preQuestionPositions
		preEvaluationPosition = message.preEvaluationPosition
		console.log('questionPositions---->', preQuestionPositions, preEvaluationPosition)
		init()
	  })
  })(window, undefined)