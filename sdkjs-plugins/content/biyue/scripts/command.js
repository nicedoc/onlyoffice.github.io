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
    console.log("commandLog", args);
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
    return new Promise((resolve, reject) => {        
        var token = uuid();       
        commandLog("pluginFrame send command, token=" + token);        
        window.Asc.plugin.callReliableCommand(token, func, isClose, isCalc, function(token, error, retData)
        {
            commandLog("callback" + token);
            if (!error)
            {
                resolve(retData)
            }
            else
            {
                reject(error);
            }
        });
    });
}

export { biyueCallCommand };
