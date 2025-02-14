import { JSONPath } from '../vendor/jsonpath-plus/dist/index-browser-esm.js';
import { biyueCallCommand } from './command.js';

// "type": "instrText",
// "instr": 'EQ \\* jc0 \\* "Font:SimSun" \\* hps14 \\o \\ad(\\s \\up 13(tūn),吞)'
let parseEQ = function(eqText)
{   
    const patt = / EQ \\\* (?<align>\w+\d+) \\\* \"Font:(?<font>.+)\" \\\* hps(?<hps>\d+) \\o \\ad\(\\s \\up (?<up>\d+)\((?<ruby>.+)\),(?<rubyBase>.)\)/gs;
    var res = patt.exec(eqText);
    
    if (res === null)
        return null;

    var HpsRaise = 2*parseInt(res.groups.up); 
    return {
        Align: parseJustify(res.groups.align),
        Hps: parseInt(res.groups.hps),
        HpsRaise: HpsRaise,
        HpsBaseText: HpsRaise + 2,
        FontName: res.groups.font,
        Ruby: res.groups.ruby,
        RubyBase: res.groups.rubyBase
    }
}

function parseJustify(str)
{
    return 0;
}

// 将域公式实现的汉语拼音转换为用Ruby标签形式
let SearchRubyField = function (text) {
    var k = JSON.parse(text);

    var ranges = [];
    var cur_range = null;
    // 获取结构区域    
    // $..content[?(typeof(@) == "string"  && @.match('^\\d+\\.'))] 
    const fieldPatt = `$..content[?(@.type=="fldChar" || @.type=="instrText" )]`;
    JSONPath({
        path: fieldPatt, json: k, resultType: "all", callback: function (res) {
			if (res && res.value) {
				var node = res.value;    
				if (node.type == "fldChar" && node.fldCharType == "begin")
				{
					cur_range = {
						beg: res.path,
					}
				}
				else if (node.type == "fldChar" && node.fldCharType == "end")
				{
					if (cur_range.Ruby !== null)
					{
						cur_range.end = res.path;
						ranges.push(cur_range);
					}
					cur_range = null;
				}
				else if (node.type == "instrText")
				{
					cur_range.Ruby= parseEQ(node.instr);
				}
			}
        }
    });    

    // filter
    ranges = ranges.filter((item) => { return item.ruby !== null; })
    return ranges;
}

let InsertRuby = function(ranges) {
    Asc.scope.ranges = ranges;
    return biyueCallCommand(window, function(){
        let PosInParagraph = function(PosArray) 
        {
            var idx = PosArray.length - 2;
            return PosArray[idx].Position;
        }
        var ranges = Asc.scope.ranges;
        var complete = 0;
        for (var i = ranges.length - 1; i >=0; i--)
        {
            var range = ranges[i];
            // 找到paragraph
            var oRange = Api.asc_MakeRangeByPath(range.beg, range.end);
            if (!oRange)
            {
                console.log("make range failed ", range.beg, range.end);
                continue;
            }
            var oParagraphs = oRange.GetAllParagraphs();
            if (oParagraphs.length !== 1)
            {
                console.log("skip range ", range.beg, range.end);
                continue;
            }
            
            var oParagraph = oParagraphs[0];
            // 删除fldChar begin 到 fldChar end
            var begPos = PosInParagraph(oRange.StartPos);
            var endPos = PosInParagraph(oRange.EndPos);

            if (begPos < oParagraph.Paragraph.Content.length &&                
                endPos < oParagraph.Paragraph.Content.length &&
                begPos <= endPos)
            {
                oParagraph.Paragraph.RemoveFromContent(begPos, endPos - begPos + 1);
            }
            else
            {
                continue;
            }

            var apiRun = editor.CreateRun();                       

            apiRun.AddRuby(range.Ruby.Ruby, range.Ruby.RubyBase, range.Ruby)
            
            // 插入ruby
            oParagraph.AddElement(apiRun, begPos);
            complete++;
        }

        return complete;
    }, false, true);        
}


// 替换字段
let ReplaceRubyField = function()
{
    return biyueCallCommand(window, function() {
        var oDocument = Api.GetDocument();
        var text_json = oDocument.ToJSON(false, false, false, false, true, true);
        return text_json;
    }, false, false)
    .then((result) => {        
        var ranges = SearchRubyField(result);        
        console.log("1. 查找拼音域, 字数", ranges.length)
        return InsertRuby(ranges);})
    .then((result) => {
        console.log("完成替换", result)
    })
    .catch((err)=>{
        console.log("command error", err);
    });
}

export { ReplaceRubyField };


