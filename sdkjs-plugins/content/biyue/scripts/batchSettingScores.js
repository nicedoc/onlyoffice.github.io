;(function (window, undefined) {
  let biyueCustomData = {}
  let question_map = {}
  let questionList = []
  let tree_map = {}
  let hidden_empty_struct = false
  window.Asc.plugin.init = function () {
    console.log('examExport init')
    window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'PaperMessage' })
  }

  function init() {
    renderData()
    $('#confirm').on('click', onConfirm)
    $('#hidden_empty_struct').on('click', onSwitchStruct)
  }

  function renderData() {
    let node_list = biyueCustomData.node_list || []
    question_map = biyueCustomData.question_map || {}
    let choice_display = biyueCustomData.choice_display || {}
    let html = ''
    let question_list = []
    let tree = {}
    let pre_struct = ''
    for (const key in node_list) {
      let item = question_map[node_list[key].id] || ''
      if (node_list[key].level_type == 'question') {
        if (item && item.question_type !== 6) {
          html += `<span class="question" title="${ item.text }">${(item.ques_name || item.ques_default_name  || '')}`
          let show_choice_region = choice_display.style == 'show_choice_region' // 判断是否为开启集中作答区
          if (item.ask_list && item.ask_list.length > 0 && (!show_choice_region || show_choice_region && !node_list[key].use_gather)) {
            for (const ask_k in item.ask_list) {
              html += `<input type="text" class="score ques-${ node_list[key].id } ask-index-${ask_k} ask-${item.ask_list[ask_k].id}" value="${item.ask_list[ask_k].score || 0}">`
            }
          } else {
            html += `<input type="text" class="score ques-${ node_list[key].id }" value="${ item.score || 0 }">`
          }
          html += `</span>`

          if (!question_list.includes(node_list[key].id)) {
            question_list.push(node_list[key].id)
          }
          if (!pre_struct && !tree[pre_struct]) {
            tree[pre_struct] = []
          }
          if (tree[pre_struct]) {
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
    $('.batch-setting-score-info').html(html)

    // 处理批量设置所有未设置分数的输入框事件
    $('#bat-score-group-all').on('input', function() {
      let inputValue = $('#bat-score-group-all').val()
      $('#bat-score-group-all').val(checkInputValue(inputValue, 100))
    })

    // 处理批量设置所有未设置分数的输入框事件
    $('#bat-score-set-btn-all').on('click', function() {
      let inputValue = $('#bat-score-group-all').val()
      batchSetEmptyScore(inputValue)
      $('#bat-score-group-all').val('')
    })

    // 处理题目的分数输入框事件
    for (const key in questionList) {
        let id = questionList[key]
        let doms = document.querySelectorAll('.ques-' + id) || []
        doms.forEach(function(dom) {
            dom.addEventListener('change', function() {
              getScoreSum()
            })
            dom.addEventListener('input', function() {
              let inputValue = dom.value
              dom.value = checkInputValue(inputValue, 100)
            })
        })
    }

    for (const key in tree_map) {
        if (tree_map[key].length > 0) {
            // 对有题的结构增加批量设置分数的输入框
            let dom = document.querySelector('#group-id-' + key)
            if (dom) {
                let html = dom.innerHTML // 取出当前的题组内容
                let inputHtml = `<div class="bat-score-set">设置为<input class="input" type="text" id="bat-score-group-${ key }">分<span class="bat-score-set-btn" id="bat-score-set-btn-${ key }" data-id="${ key }">设置</span></div>`
                dom.innerHTML = html + inputHtml

                let inputDom = document.querySelector(`#bat-score-group-${ key }`)
                if (inputDom) {
                  inputDom.addEventListener('input', function() {
                    let inputValue = inputDom.value
                    inputDom.value = checkInputValue(inputValue, 100)
                  })
                }

                let btnDom = document.querySelector(`#bat-score-set-btn-${ key }`)
                if (btnDom) {
                  btnDom.addEventListener('click', function() {
                      let id = btnDom.dataset.id || ''
                      let inputDom = document.querySelector(`#bat-score-group-${ id }`)
                      batchSetStructScore(id, inputDom.value || 0)
                      getScoreSum()
                  })
                }
            }
        }
    }
    getScoreSum()
  }

  function checkInputValue(val = '', max) {
    // 只允许数字和一个小数点，并且小数点后最多一位数字
    var sanitizedValue = val.replace(/[^0-9.]/g, ''); // 移除非数字和小数点的字符

    // 处理多个小数点的情况
    var parts = sanitizedValue.split('.');
    if (parts.length > 2) {
        sanitizedValue = parts[0] + '.' + parts.slice(1).join('');
    }

    // 处理最多一位小数的情况
    if (parts.length === 2) {
        sanitizedValue = parts[0] + '.' + (parts[1].length > 1 ? parts[1].substring(0, 1) : parts[1]);
    }
    if (max) {
      sanitizedValue = sanitizedValue > max ? max : sanitizedValue
    }
    return sanitizedValue
  }

  function batchSetStructScore(struct_id, value) {
    // 根据结构批量写入分数
    let arr = tree_map[struct_id] || []
    for (const key in arr) {
        let id = arr[key]
        if (question_map[id].ask_list.length > 0) {
            let ask_list = question_map[id].ask_list || []
            let sum = 0
            for (const k in ask_list) {
                ask_list[k].score = value
                sum += parseFloat(value) || 0
            }
            question_map[id].score = sum
        } else {
          question_map[id].score = value
        }
    }

    renderData()
  }

  function batchSetEmptyScore(value) {
    // 根据找到所有没有设置分数的空写入分数
    for (const key in tree_map) {
      if (tree_map[key].length > 0) {
        for (const k in tree_map[key]) {
          // 题目的id
          let id = tree_map[key][k]
          if (question_map[id] && question_map[id].ask_list && question_map[id].ask_list.length > 0) {
            let ask_list = question_map[id].ask_list || []
            let sum = 0
            let hasChange = false
            for (const k in ask_list) {
              let score = parseFloat(ask_list[k].score) || 0
              if (!score || score * 1 == 0) {
                ask_list[k].score = value
                hasChange = true
              }
              sum += parseFloat(ask_list[k].score) || 0
            }
            if (hasChange) {
              question_map[id].score = sum
            }
          } else if (!question_map[id].score) {
            question_map[id].score = value
          }
        }
      }
    }
    renderData()
  }

  function getScoreSum() {
    let score_sum = 0
    for (const key in questionList) {
      let id = questionList[key]
      let doms = document.querySelectorAll('.ques-' + id) || []
      doms.forEach(function(dom) {
        let val = parseFloat(dom.value) || 0
        if (val == 0) {
          dom.value = 0
          dom.style.color='#ff0000'
        } else {
          dom.style.color=''
          dom.value = val * 1
        }
        score_sum += val
      })
    }

    $('#score_sum').html(score_sum)
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
    let choice_display = biyueCustomData.choice_display || {}
    let show_choice_region = choice_display.style == 'show_choice_region' // 判断是否为开启集中作答区
    let node_list = biyueCustomData.node_list || []
    for (const key in questionList) {
      let id = questionList[key] || ''
      let dom = $('.ques-'+id)
      if (id && question_map[id] && dom) {
        let ask_list = question_map[id].ask_list || []
        var nodeData = node_list.find(e => {
          return e.id == id
        })
        if (ask_list.length > 0 && (!show_choice_region || show_choice_region && !nodeData.use_gather)) {
          // 有小问区的题
          let sumScore = 0
          for (const k in ask_list) {
            let ask_dom = $(`.ques-${ id }.ask-index-${ k }`)
            if (ask_dom) {
              let params = {
                id: id,
                type: 'ask',
                index: k,
                score: ask_dom.val()
              }
              sumScore += params.score * 1
              changeScore(params)
            }
          }
          // 题目的分数设置为小问的分数的和
          let params = {
            id: id,
            type: 'question',
            score: sumScore
          }
          changeScore(params)
        } else {
          // 没有小问区的题目
          let params = {
            id: id,
            type: 'question',
            score: dom.val()
          }
          changeScore(params)
        }
      }
    }
    // 将窗口的信息传递出去
    window.Asc.plugin.sendToPlugin('onWindowMessage', {
      type: 'changeQuestionMap',
      data: question_map,
    })
  }

  function changeScore({id, type, index, score}) {
    if (type == 'ask') {
      question_map[id].ask_list[index].score = score
    } else {
      question_map[id].score = score
    }
  }

  window.Asc.plugin.attachEvent('initPaper', function (message) {
    console.log('batchSettingScores 接收的消息', message)
    biyueCustomData = message.biyueCustomData || {}
    init()
  })
})(window, undefined)
