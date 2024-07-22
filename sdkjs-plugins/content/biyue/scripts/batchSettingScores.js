;(function (window, undefined) {
  let biyueCustomData = {}
  let question_map = {}
  let questionList = []
  window.Asc.plugin.init = function () {
    console.log('examExport init')
    window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'PaperMessage' })
  }

  function init() {
    let node_list = biyueCustomData.node_list || []
    question_map = biyueCustomData.question_map || {}
    let html = ''
    let question_list = []
    for (const key in node_list) {
      let item = question_map[node_list[key].id] || ''
      if (node_list[key].level_type == 'question') {
        if (item && item.question_type !== 6) {
          html += `<span class="question">${(item.ques_default_name ? item.ques_default_name : '')}`
          if (item.ask_list.length > 0) {
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
        }
      } else if (node_list[key].level_type == 'struct' && item){
        html += `<div class="group">${ item.text.split('\r\n')[0] }</div>`
      }
    }
    $('.batch-setting-score-info').html(html)
    $('#confirm').on('click', onConfirm)
    questionList = question_list || []
  }

  function onConfirm() {
    for (const key in questionList) {
      let id = questionList[key] || ''
      let dom = $('.ques-'+id)
      if (id && question_map[id] && dom) {
        let ask_list = question_map[id].ask_list || []
        if (ask_list.length > 0) {
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
