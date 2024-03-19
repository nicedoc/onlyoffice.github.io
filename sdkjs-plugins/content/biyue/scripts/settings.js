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
      // 只允许输入数字
      input.value = input.value.replace(/[^\d]/g, '');
      // 将大写数字转换为小写
      input.value = input.value.toLowerCase();
      if (input.value > 100) {
        // 限制单题分数上限
        input.value = 100
      }
    }

  window.Asc.plugin.attachEvent("onParams", function(message) {
    let result = message || {};
    console.log('接收的消息', result);
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
})(window, undefined);

