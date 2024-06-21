class NumberInput{
  constructor(id, options = {}) {
    this.id = id
    this.options = options
    this.width = options.width || '85px'
    this.render() 
  }
  render() {
    var content = ''  
    content = `
    <div id=${this.id} class="spinner" style="width: ${this.width};">
      <input type="text" class="form-control" spellcheck="false" data-hint="1" data-hint-direction="bottom" data-hint-offset="big">
      <div class="spinner-buttons">
        <button type="button" class="spinner-up">
          <i class="arrow"></i>
        </button>
        <button type="button" class="spinner-down">
          <i class="arrow"></i>
        </button>
      </div>
    </div>`
    $(`#${this.id}`).replaceWith(content)
    $(`#${this.id} .spinner-up`).on('click', () => {
      this.onInputBtn(1)
    })
    $(`#${this.id} .spinner-down`).on('click', () => {
      this.onInputBtn(-1)
    })
    $(`#${this.id} input`).on('input', () => {
      this.valueChange($(`#${this.id} input`).val())
    })
  }

  hide() {
    $(`#${this.id}`).hide()
  }

  show() {
    $(`#${this.id}`).show()
  }

  setValue(value) {
    $(`#${this.id} input`).val(value)
  }

  onInputBtn(offset) {
    var inputEl = $(`#${this.id} input`)
    var v = inputEl.val()
    var vv = parseInt(v)
    if (isNaN(vv)) {
      vv = 1
    }
    inputEl.val(offset + vv)
    this.valueChange(inputEl.val())
  }

  valueChange(value) {
    if (this.options.change) {
      this.options.change(this.id, value)
    }
  }
}
export default NumberInput