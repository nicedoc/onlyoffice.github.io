;(function (window, undefined) {
  let biyueCustomData = {}
  window.Asc.plugin.init = function () {
    console.log('examExport init')
    window.Asc.plugin.sendToPlugin('onWindowMessage', { type: 'PaperMessage' })
  }

  function init() {
    let node_list = biyueCustomData.node_list || []
    let question_map = biyueCustomData.question_map || {}
    let html = ''
    for (const key in node_list) {
      let item = question_map[node_list[key].id] || ''
      if (node_list[key].level_type == 'question') {
        if (item && item.question_type !== 6) {
          html += `<span class="question">${(item.ques_default_name ? item.ques_default_name : '') + '.'}`
          if (item.ask_list.length > 0) {
            for (const key in item.ask_list) {
              html += `<input type="text" class="score score-${item.ask_list[key].id}" value="${item.ask_list[key].score || 0}">`
            }
          } else {
            html += `<input type="text" class="score score-${ node_list[key].id }" value="${ node_list[key].score || 0 }">`
          }
          html += `</span>`
        }
      } else if (node_list[key].level_type == 'struct' && item){
        html += `<div class="group">${ item.text.split('\r\n')[0] }</div>`
      }
    }
    $('.batch-setting-score-info').html(html)
    console.log('html:', html)
  }

  window.Asc.plugin.attachEvent('initPaper', function (message) {
    console.log('batchSettingScores 接收的消息', message)
    biyueCustomData = message.biyueCustomData || {}
    init()
  })
})(window, undefined)
