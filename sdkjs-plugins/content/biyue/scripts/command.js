/**
 * Generates a random UUID.
 * @returns {string} A random UUID.
 */
function uuid() {
    var s = [];
    var hexDigits = "0123456789abcdef";
    for (var i = 0; i < 36; i++) {
        s[i] = hexDigits.substr(Math.floor(Math.random() * 0x10), 1);
    }
    s[14] = "4"; 
    s[19] = hexDigits.substr((s[19] & 0x3) | 0x8, 1); 
    s[8] = s[13] = s[18] = s[23] = "-";
    var uuid = s.join("");
    return uuid;
}

function commandLog(args) {
    //  console.log("commandLog", args);
}   

/**
 * Executes a command in the Biyue plugin.
 * @param {Window} window - The window object.
 * @param {Function} func - The function to be executed as a command.
 * @param {boolean} isClose - Indicates whether the command is a close command.
 * @param {boolean} isCalc - Indicates whether the command is a calculation command.
 * @returns {Promise} A promise that resolves with the result of the command execution.
 */
async function biyueCallCommand(window, func, isClose, isCalc) {
    if (window.commandCallbackMap === undefined) {
        window.commandCallbackMap = {};
    }
    if (window.commandStack == undefined)  {
        window.commandStack = [];
    }

    return new Promise((resolve, reject) => {        
        var token = uuid();
        var l = func;
        var g = isClose;
        var h = isCalc;
        var b = undefined;
        
        window.commandCallbackMap[token] = resolve;
        commandLog("pluginFrame send command, token=" + token);

        var f = function() {
            l = "var Asc = {};"+
            "Asc.scope = " + JSON.stringify(window.Asc.scope) + ";"+
            //'console.log("editorFrame exec:'+ token+'");' +
            "var scope = Asc.scope; (function() { return { token:\"" + token + "\", ret: (" + l.toString() + ")()}})();";
            window.Asc.plugin.windowID ? (window.Asc.plugin.windowID._pushWindowMethodCommandCallback(b),
            window.Asc.plugin.windowID.sendToPlugin("private_window_command", {
                code: l,
                isCalc: h
            })) : (window.Asc.plugin.info.recalculate = !1 === h ? !1 : !0,        
            window.Asc.plugin.executeCommand(!0 === g ? "close" : "command", l, b))
        }

        if (window.commandStack.length > 0) {
            window.commandStack.push(f);            
        } else {
            window.commandStack.push(f);
            f();
        }        
    });
}

/**
 * Dispatches the result of a command execution.
 * @param {Window} window - The window object.
 * @param {Object} result - The result of the command execution.
 */
function dispatchCommandResult(window, result) {
    if (window.commandCallbackMap === undefined) {
        return;
    }

    if (result === undefined) {
        return;
    }

    var token = result.token;
    if (token === undefined) {
        return;
    }
    commandLog("pluginFrame recv command result, token=" + token)

    var callback = window.commandCallbackMap[token];
    if (callback !== undefined) {
        delete window.commandCallbackMap[token];
        callback(result.ret);
    }
    var f = window.commandStack.shift();
    if (window.commandStack.length > 0) {
        window.commandStack[0]();
    }
}

export { biyueCallCommand, dispatchCommandResult };