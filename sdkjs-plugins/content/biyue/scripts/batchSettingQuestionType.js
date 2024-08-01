;(function (window, undefined) {
  let biyueCustomData = {}
  let question_map = {}
  let questionList = []
  let tree_map = {}
  let question_type_options = []
  let hidden_empty_struct = false
  window.Asc.plugin.init = function () {
    console.log('examExport init')
    window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'PaperMessage' })
  }

  function init() {
    getOptions()
    $('#confirm').on('click', onConfirm)
    $('#hidden_empty_struct').on('click', onSwitchStruct)
  }

  function getOptions() {
    question_type_options = []
    if (biyueCustomData.paper_options) {
      question_type_options = biyueCustomData.paper_options.question_type || []
    }
    renderData()
  }

  function renderData() {
    let node_list = biyueCustomData.node_list || []
    question_map = biyueCustomData.question_map || {}

    let html = ''
    let question_list = []
    let tree = {}
    let pre_struct = ''
    for (const key in node_list) {
      let item = question_map[node_list[key].id] || ''
      if (node_list[key].level_type == 'question') {
        if (item && item.question_type !== 6) {
          html += `<span class="question">${(item.ques_default_name ? item.ques_default_name : '')}`
          html += `<select class="type-item ques-${ node_list[key].id }">`
          for (const key in question_type_options) {
            let selected = item.question_type * 1 === question_type_options[key].value * 1 ? 'selected' : ''
            html += `<option value="${question_type_options[key].value}" ${selected}>${question_type_options[key].label}</option>`
          }

          html += `</select></span>`

          if (!question_list.includes(node_list[key].id)) {
            question_list.push(node_list[key].id)
          }
          if (pre_struct && tree[pre_struct]) {
            tree[pre_struct].push(node_list[key].id)
          }
        }
      } else if (node_list[key].level_type == 'struct' && item){
        pre_struct = node_list[key].id

        html += `<div class="group" id="group-id-${ pre_struct }"><span>${ item.text.split('\r\n')[0] }</span></div>`
        if (!tree[pre_struct]) {
            tree[pre_struct] = []
        }
      }
    }
    tree_map = tree
    questionList = question_list || []
    $('.batch-setting-type-info').html(html)

    // 处理题目的下拉选项事件
    for (const key in questionList) {
      let id = questionList[key]
      let doms = document.querySelectorAll('.ques-' + id) || []
      doms.forEach(function(dom) {
          dom.addEventListener('change', function() {
            getQuestionTypeNum()
          })
      })
    }

    for (const key in tree_map) {
        if (tree_map[key].length > 0) {
            // 对有题的结构增加批量设置题型的下拉框
            let dom = document.querySelector('#group-id-' + key)
            if (dom) {
                let html = dom.innerHTML // 取出当前的题组内容
                let selectHtml = `<select id="bat-type-group-${key}" class="type-item">`
                selectHtml += `<option value="" style="display: none;"></option>`
                for (const key in question_type_options) {
                  selectHtml += `<option value="${question_type_options[key].value}">${question_type_options[key].label}</option>`
                }
                selectHtml += "</select>"

                let inputHtml = `<div class="bat-type-set">设置为${selectHtml}<span class="bat-type-set-btn" id="bat-type-set-btn-${ key }" data-id="${ key }">设置</span></div>`
                dom.innerHTML = html + inputHtml

                let btnDom = document.querySelector(`#bat-type-set-btn-${ key }`)
                if (btnDom) {
                  btnDom.addEventListener('click', function() {
                      let id = btnDom.dataset.id || ''
                      let inputDom = document.querySelector(`#bat-type-group-${ id }`)
                      if (inputDom.value > 0) {
                        batchSetStructQuestionType(id, inputDom.value || 0)
                        getQuestionTypeNum()
                      }
                  })
                }
            }
        }
    }

    getQuestionTypeNum()
  }

  function batchSetStructQuestionType(struct_id, value) {
    // 根据结构批量写入题目类型
    let arr = tree_map[struct_id] || []
    for (const key in arr) {
        let id = arr[key]
        question_map[id].question_type = value
    }

    renderData()
  }

  function getQuestionTypeNum() {
    let question_type_map = {}
    for (const key in question_type_options) {
      if (question_type_options[key]) {
        let item = question_type_options[key] || {}
        question_type_map[item.value] = {
          label: item.label,
          number: 0
        }
      }
      question_type_map[null] = {
        label: '未知题型',
        number: 0
      }
    }
    for (const key in questionList) {
      let id = questionList[key]
      let doms = document.querySelectorAll('.ques-' + id) || []
      doms.forEach(function(dom) {
        let val = dom.value || ''
        if (question_type_map[val]) {
          question_type_map[val].number += 1
        } else {
          question_type_map[null].number += 1
        }
      })
    }
    let html = ''
    for (const key in question_type_map) {
        if (question_type_map[key].number > 0) {
          html += `<span style="margin: 0 8px;">${question_type_map[key].label}${question_type_map[key].number}题</span>`
        }
    }
    $('#getQuestionTypeNum').html(html)
  }

  function onSwitchStruct() {
    //  开启/隐藏 无题目的结构
    hidden_empty_struct = !hidden_empty_struct
    $('#hidden_empty_struct').prop('checked', hidden_empty_struct)

    for (const key in tree_map) {
        let arr = tree_map[key] || []
        if (arr.length === 0) {
            if (hidden_empty_struct) {
                $('#group-id-'+ key).hide()
            } else {
                $('#group-id-'+ key).show()
            }
        }
    }
  }

  function onConfirm() {
    for (const key in questionList) {
      let id = questionList[key] || ''
      let dom = $('.ques-'+id)
      if (id && question_map[id] && dom) {
        let params = {
          id: id,
          question_type: dom.val()
        }
        changeQuestionType(params)
      }
    }
    // 将窗口的信息传递出去
    window.Asc.plugin.sendToPlugin('onWindowMessage', {
      type: 'changeQuestionMap',
      data: question_map,
    })
  }

  function changeQuestionType({id, question_type}) {
    question_map[id].question_type = parseFloat(question_type) || 0
  }

  window.Asc.plugin.attachEvent('initPaper', function (message) {
    console.log('batchSettingQuestionType 接收的消息', message)
    biyueCustomData = message.biyueCustomData || {}
    init()
  })
})(window, undefined)
