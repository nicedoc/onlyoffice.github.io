/*
    convert api
        toPdf
        toJpg
        toXml
 */

// 嵌入式编辑单题目有三个功能点
// 1 文档部份内容转为xml
//  sdkjs:serialize2.js 控件 -> bin 
//  docservice:convert  bin -> xml
// 2 编辑器打开xml
// 3 编辑器保存xml
function toXml(window, controlId, callback) {
    if (controlId !== undefined || controlId !== null) {
        window.Asc.plugin.executeMethod("SelectContentControl", [controlId])
    }

    window.Asc.plugin.executeMethod("GetSelectionToDownload", ["docx"], function (data) {
        console.log(data);
        // 假设这是你的 ZIP 文件的 URL  
        const zipFileUrl = data;        
        fetch(zipFileUrl)  
        .then(response => {  
            if (!response.ok) {  
            throw new Error('Failed to fetch zip file');  
            }  
            return response.arrayBuffer(); // 获取 ArrayBuffer 而不是 Blob，因为 JSZip 需要它  
        })  
        .then(arrayBuffer => {  
            return JSZip.loadAsync(arrayBuffer); // 使用 JSZip 加载 ArrayBuffer  
        })  
        .then(zip => {  
            // 现在你可以操作 zip 对象了  
            zip.forEach(function(relativePath, file) {  
                if (relativePath.indexOf('word/document.xml') === -1) {
                    return;
                }
                // 这里可以遍历 ZIP 文件中的所有文件  
                file.async("text").then(function(content) {  
                    // 假设文件是文本文件，打印文件内容和相对路径  
                    console.log(relativePath, content);  
                });  
            });  
        })  
        .catch(error => {  
            console.error('Error:', error);  
        });
        
    });
}


function downloadAs(window, format, callback) {
    window.Asc.plugin.executeMethod("GetFileToDownload", [format], function (data) {
        console.log(data);
        const url = data;
        callback(url);
    });
}
export { toXml, downloadAs };