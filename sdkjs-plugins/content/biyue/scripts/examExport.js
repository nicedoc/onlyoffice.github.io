import {
  paperCanConfirm,
  teacherClassMyGradePrint,
  examImportQuestionBank,
  teacherExamPrintPaperSelections,
  examPageUpload,
  paperUploadPreview,
  paperSavePosition
  // paperValidatePosition
} from './api/paper.js'
import { setXToken } from './auth.js'
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

  window.Asc.plugin.init = function () {
    console.log('examExport init')
    window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'PaperMessage' })
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

  // 上传试卷预览图
  async function uploadPreview() {
    console.log('uploadPreview')
    const file_list = img_file_list || []
    if (!file_list) {
      return
    }
    var pid = 1
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
          pid++
        })
        .catch((error) => {
          console.log(error)
        })
    }
    console.log('所有底图上传成功')
    onUpdatePostions()
  }

  // 更新试卷切题信息
  function onUpdatePostions() {
    let positions = preQuestionPositions
    var evaluationPosition = preEvaluationPosition

    console.log('positions and evaluationPosition:', positions, evaluationPosition)
    paperSavePosition(biyueCustomData.paper_uuid, positions, evaluationPosition, '')
    .then((res) => {
      console.log('保存位置成功')
      // 将窗口的信息传递出去
      window.Asc.plugin.sendToPlugin('onWindowMessage', {
        type: 'positionSaveSuccess',
        data: source_data,
      })
    })
    .catch((error) => {
      if (error.message) {
        alert('上传失败:', error.message)
      }
      console.log(error)
    })
  }

  function getFieldsByZoneType(zoneType) {
    // 遍历 feature_list 查找匹配的 zone_type
    let result = zoneType === 'pass' ? {} : []

    for (const feature of questionPositions.feature_list) {
      if (zoneType === 'pass' && feature.zone_type === zoneType) {
        // pass区域需要传对象的结构
        return feature.fields && feature.fields.length > 0 ? feature.fields[0] : {}
      } else if (feature.zone_type === zoneType) {
        return feature.fields
      }
    }
    return result
  }

  function getPositions () {
    let positions = {}
    let ques_list = questionPositions.ques_list || []
    let index = 1
    for (const key in ques_list) {
      let item = ques_list[key] || ""
      if (item && item.ref_id && item.ques_type !== 6) {
        if (!positions[item.ref_id]) {
          positions[item.ref_id] = {}
        }
        item.ques_no = index++
        positions[item.ref_id] = item
      }
    }
    return positions
  }

  function onCheckPositions() {
    page_size = ''
    if (paper_info && paper_info.workbook && paper_info.workbook.layout) {
      page_size = paper_info.workbook.layout
    } else {
      alert('未获取到练习册配置的尺寸')
      return
    }

    preQuestionPositions = getPositions()
    preEvaluationPosition = getEvaluationPosition()
    if (Object.keys(preQuestionPositions).length === 0) {
        $('#error_message').html('没有符合条件的题目，请先进行全量更新!')
        return
    }
    const haveError = checkPositions(preQuestionPositions)
    if (haveError) {
        return
    }
    // 过滤题目选区里的字段类型,如果是数值类型则还会去除小数位
    const need_formatted = ['title_region', 'write_ask_region', 'mark_ask_region', 'correct_region', 'correct_ask_region']
    formattedPositions(preQuestionPositions, need_formatted)

    console.log('本地检查通过')
    // console.log('开始进行服务端检查...')
    // onPaperValidatePosition()
    initImg()
  }

  function formattedPositions(positions, fieldsToCheck) {
    for (const key in positions) {
      const item = positions[key]
      fieldsToCheck.forEach(field => {
        if (field in item) {
          positions[key][field] = convertDataToNumber(positions[key][field], numericFields)
        }
      })
    }
  }

  function convertDataToNumber(data, numericFields) {
    // 判断是否为数字的函数
    const isNumber = (value) => !isNaN(parseFloat(value)) && isFinite(value)

    // 转换对象数组中的数据类型
    const convertObjectArray = (arr) => {
      arr.forEach(item => {
        for (const prop in item) {
          if (numericFields.includes(prop)) {
            item[prop] = isNumber(item[prop]) && item[prop] > 0 ? Math.floor(item[prop]) : 0
          } else {
            item[prop] = String(item[prop])
          }
        }
      })
    }

      // 遍历数据对象或数组
    const convert = (item) => {
      for (const key in item) {
        if (Array.isArray(item[key]) && item[key].length > 0 && typeof item[key][0] === 'object') {
          convertObjectArray(item[key])
        } else {
          if (numericFields.includes(key)) {
            item[key] = isNumber(item[key]) && item[key] > 0 ? Math.floor(item[key]) : 0
          } else {
            item[key] = String(item[key])
          }
        }
      }
    }

    if (Array.isArray(data)) {
      data.forEach(item => convert(item))
    } else {
      convert(data)
    }

    return data
  }

  function checkPositions(postions) {
    let error = false
    let error_message = ''
    const title_region_err = []
    const write_ask_region_err = []
    const correct_ask_region_err = []
    const need_update = false
    const ques_score_err = []
    const ask_score_err = []
    const ask_score_empty_err = []
    const ques_type_err = []

    for (const key in postions) {
      if (!postions[key].ref_id) {
        need_update = true
      }
      let score = parseFloat(postions[key].score) || 0
      if (!score) {
        ques_score_err.push(postions[key].ques_name)
      }
      if (!postions[key].ques_type) {
        ques_type_err.push(postions[key].ques_name)
      }
      if (checkRegionEmpty(postions[key], 'title_region')) {
        title_region_err.push(postions[key].ques_name)
      }
      if (checkRegionEmpty(postions[key], 'write_ask_region', true) || !postions[key]['write_ask_region'] || postions[key]['write_ask_region'].length === 0) {
        write_ask_region_err.push(postions[key].ques_name)
      }
      if (checkRegionEmpty(postions[key], 'correct_ask_region', true)) {
        correct_ask_region_err.push(postions[key].ques_name)
      }
      postions[key].score = parseFloat(postions[key].score) || 0
      if (postions[key].mark_method === '4') {
        postions[key].ask_num = 1
      } else if (postions[key].mark_ask_region && postions[key].mark_method !== '4') {
        const mark_ask_region = postions[key].mark_ask_region || {}
        postions[key].ask_num = Object.keys(mark_ask_region).length

        if (postions[key].ask_num > 1) {
          const question_ask = {}
          if (postions[key].mark_method === '2') {
            // 主观题
          } else {
            const positionsObj = postions[key].mark_ask_region
            let ask_score_sum = 0
            let has_ask_score_empty_err = false
            for (const k in positionsObj) {
              if (positionsObj[k] && positionsObj[k][0]) {
                question_ask[positionsObj[k][0].order] = {
                  order: positionsObj[k][0].order + '',
                  score: positionsObj[k][0].v + '' // 因为以前的历史原因 小问的分数需要字符串类型
                }
                if (postions[key].score) {
                    //  检查小问的分数和题目的分数是否一致
                    let ask_score = parseFloat(positionsObj[k][0].v) || 0
                    ask_score_sum += ask_score
                    if (!ask_score) {
                      has_ask_score_empty_err = true
                    }
                }
              }
            }
            if ((ask_score_sum * 1 !== postions[key].score * 1) && postions[key].score) {
                //  小问分数和题目分数不一致
                ask_score_err.push(postions[key].ques_name)
            }
            if (has_ask_score_empty_err) {
                // 题目有分数的情况下，有小问的分数为空
                ask_score_empty_err.push(postions[key].ques_name)
            }
            postions[key].question_ask = question_ask
          }
        }
      }
    }
    let message = ''

    if (ques_score_err.length > 0) {
        if (message !== '') {
          message += `<br/><br/>`
        }
        message += `有题目未设置分数,请检查题目:` + ques_score_err.join(',')
    }
    if (ask_score_err.length > 0) {
        if (message !== '') {
          message += `<br/><br/>`
        }
        message += `有题目小问分数和题目分数不一致,请检查题目:` + ask_score_err.join(',')
    }
    if (ask_score_empty_err.length > 0) {
        if (message !== '') {
          message += `<br/><br/>`
        }
        message += `有题目的小问分数出现为空,请检查题目:` + ask_score_empty_err.join(',')
    }

    if (ques_type_err.length > 0) {
        if (message !== '') {
            message += `<br/><br/>`
        }
        message += `有题目未设置题目类型,请检查题目:` + ques_type_err.join(',')
    }

    if (title_region_err.length > 0) {
        message = `有题干区域未设置,请检查题目:` + title_region_err.join(',')
    }
    if (write_ask_region_err.length > 0) {
        if (message !== '') {
            message += `<br/><br/>`
        }
        message += `有作答区域未设置,请检查题目:` + write_ask_region_err.join(',')
    }
    if (correct_ask_region_err.length > 0) {
        if (message !== '') {
            message += `<br/><br/>`
        }
        message += `有小问订正框区域未设置,请检查题目:` + correct_ask_region_err.join(',')
    }


    if (need_update) {
        message = '有题目没有分配uuid，请先进行全量更新。'
    }
    if (message) {
        error = true
        error_message = message
        $('#error_message').html(error_message)
    }
    return error
  }

  function onPaperValidatePosition() {
    // 服务端的坐标检查
    var questionPositions = preQuestionPositions
    var evaluationPosition = preEvaluationPosition
    let error_message = ''

    // 由于需要避免特殊符号导致的服务端获取不到内容的问题，接口需要使用application/json的方式传暂时跳过服务端的检查步骤

    // paperValidatePosition(paper_info.workbook.id, paper_info.paper.paper_uuid, questionPositions, evaluationPosition).then(res => {
    //   const data = res.data || {}
    //   if (res.code === 1 && !data.error) {
    //     // 检查通过 继续下一步去上传
    //   } else {
    //     error_message = '【服务端检查未通过】' + (data.error || '')
    //   }
    // }).catch(err => {
    //   console.log(err.message)
    //   error_message = '【服务端检查未通过】' + (err.message || '')
    // })
  }

  function getEvaluationPosition() {
    var evaluationPosition = {
      self_evaluation: getFieldsByZoneType('self_evaluation'),
      teacher_evaluation: getFieldsByZoneType('teacher_evaluation'),
      pass_regional: getFieldsByZoneType('pass'),
      ignore_region: getFieldsByZoneType('ignore'),
      end_regional: getFieldsByZoneType('end'),
      again_regional: getFieldsByZoneType('again'),
      stat_regional: getFieldsByZoneType('statistics'),
      partial_no_dot_regional: questionPositions.partical_no_dot_list || []
    }
    // 处理选区的数据类型
    for (const key in evaluationPosition) {
      evaluationPosition[key] = convertDataToNumber(evaluationPosition[key], numericFields)
    }

    evaluationPosition['page_size'] = page_size
    evaluationPosition['exam_type'] = 'exercise'
    return evaluationPosition
  }

  function checkRegionEmpty(postions, name, checkValue) {
    const ret = false
    if (!postions[name]) {
      return true
    }
    const arr = postions[name] || []
    for (const key in arr) {
      const w = arr[key].w || 0
      const h = arr[key].h || 0
      const v = arr[key].h || 0
      if (w * 1 + h * 1 === 0 || checkValue && v * 1 === 0) {
        return true
      }
    }
    return ret
  }

  function initImg() {
    $('#workbookImgLoading').show()
    window.Asc.plugin.executeMethod("GetFileToDownload", ["PNG"], res=> {
      console.log(res)
      const zipUrl = res;
      let preview = document.getElementById('workbook_preview')
      if (preview) {
        preview.innerHTML = '' // 清空预览区域
      }

      img_file_list = []  // 清空上传用的file

      fetch(zipUrl).then(response => {
        if (!response.ok) {
          throw new Error('Failed to fetch zip file');
        }
        return response.arrayBuffer();
      }).then(arrayBuffer => {
        return JSZip.loadAsync(arrayBuffer); // 使用 JSZip 加载 ArrayBuffer
      }).then(async zip => {
        try {
          for (const [filename, file] of Object.entries(zip.files)) {
            if (!file.dir) { // 只处理文件，忽略文件夹
              // 读取文件内容为 Blob
              const fileBlob = await file.async('blob');

              // 创建一个新的 File 对象并将其存储到数组中 用于后续上传
              const fileObj = new File([fileBlob], filename, { type: fileBlob.type });
              img_file_list.push(fileObj)

              // 使用 FileReader 将 Blob 转换为 Data URL
              const reader = new FileReader()
              reader.onloadend = function () {
                // 创建并添加 img 元素到页面
                const img = document.createElement('img');
                img.src = reader.result;
                img.style.maxWidth = '500px' // 设置预览图片的最大宽度
                preview.appendChild(img)
              }

              if (fileBlob instanceof Blob) {
                reader.readAsDataURL(fileBlob);
              } else {
                console.error('fileBlob 不是有效的 Blob 对象');
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
          console.error('Error fetching or processing ZIP file:', error);
        }
      })

    })
  }

  window.Asc.plugin.attachEvent('initPaper', function (message) {
    console.log('examExport 接收的消息', message)
    setXToken(message.xtoken)
    paper_info = message.paper_info
    source_data = message
    biyueCustomData = message.biyueCustomData
    questionPositions = message.questionPositions || {}
    console.log('questionPositions---->', questionPositions, getPositions())
    init()
  })
})(window, undefined)
