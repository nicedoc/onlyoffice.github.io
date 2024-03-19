(function (window, undefined) {

    // 题目的类型
    var questionTypes = [
      { value: '0', label: '未定义' },
      { value: '1', label: '单选' },
      { value: '2', label: '填空' },
      { value: '3', label: '作答' },
      { value: '4', label: '判断' },
      { value: '5', label: '多选' },
      { value: '6', label: '文本' },
      { value: '7', label: '单选组合' },
      { value: '8', label: '作文' }
    ];
    var activeQuesItem = ''

    window.Asc.plugin.init = function() {
      console.log("settings init");
		  window.Asc.plugin.sendToPlugin("onBiyueMessage");
      // 在页面加载时生成下拉框选项
      var selectElement = document.getElementById("questionType");
      for (var i = 0; i < questionTypes.length; i++) {
        var option = document.createElement("option");
        if (questionTypes[i]) {
          option.text = questionTypes[i].label || '';
          option.value = questionTypes[i].value || '';
          selectElement.add(option);
        }
      }
      document.getElementById("score").oninput = function () {
        sanitizeInput(this)
      }
    }

    function sanitizeInput(input) {
      let value = input.value
      // 过滤非数字和小数点的字符
      let filteredInput = value.replace(/[^0-9.]/g, '');

      // 检查小数点的个数，如果多于一个，则进行处理
      if (filteredInput.split('.').length > 2) {
        let parts = filteredInput.split('.');
        filteredInput = parts.shift() + '.' + parts.join('');
      }
      // 如果开头是小数点，则加上 0
      if (filteredInput.startsWith('.')) {
        filteredInput = '0' + filteredInput;
      }
      // 限制小数位最多为1位
      let decimalIndex = filteredInput.indexOf('.');
      if (decimalIndex !== -1) {
          let decimalPart = filteredInput.substring(decimalIndex + 1);
          if (decimalPart.length > 1) {
              filteredInput = filteredInput.substring(0, decimalIndex + 2);
          }
      }
      // 去除数字前面多余的零
      let parts = filteredInput.split('.');
      parts[0] = String(Number(parts[0])); // 转换成数值然后转换回字符串，去除前面的零
      filteredInput = parts.join('.');
      // 将大写数字转换为小写
      filteredInput = filteredInput.replace(/[０-９]/g, function(match) {
          return String.fromCharCode(match.charCodeAt(0) - 65248);
      });
      if (filteredInput > 100) {
        // 限制单题分数上限
        filteredInput = 100
      }
      input.value = filteredInput;
    }
    function getDialogForm () {
      let form = {}
      let questionTypeDom = document.getElementById('questionType')
      if (questionTypeDom) {
        form.mode = questionTypeDom.value || ''
      }
      let scoreDom = document.getElementById('score')
      if (scoreDom) {
        form.score = scoreDom.value || 0
      }
      return form
    }

  window.Asc.plugin.attachEvent("onParams", function(message) {
    let result = message || {};
    console.log('接收的消息', result);
    activeQuesItem = result || ''
    let tagObj = ''
    if (result.Tag) {
      tagObj = JSON.parse(result.Tag)
    }
    let tagDom = document.getElementById('tag')
    if (tagDom) {
      tagDom.innerHTML = result.Tag
    }
    let questionTypeDom = document.getElementById('questionType')
    if (questionTypeDom && tagObj) {
      questionTypeDom.value = tagObj.mode || ''
    }
    let scoreDom = document.getElementById('score')
    if (scoreDom && tagObj) {
      scoreDom.value = tagObj.score || ''
    }

  });
  window.Asc.plugin.attachEvent("getParams", function() {
    let params = {
      form: getDialogForm(),
      activeQuesItem: activeQuesItem
    }

    // 将窗口的信息传递出去
    window.Asc.plugin.sendToPlugin("getSettingsMessage", params);
  });
})(window, undefined);

