/*Copyright (c) 2016 Flynet Ltd.

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"),
to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
and/or sell copies of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
IN THE SOFTWARE.
*/

// Connects to the simulated host in FVTerm.
function connectFVTerm() {

    var requestData = {};

    // These values assume you've set up your web.config and hosts as per the .md file.
    // If your application is called something other than InsureRecog or your host is
    // not called Host1 then edit the correct values in here.
    requestData.application = "InsureRecog";
    requestData.autoStart = true;
    requestData.hostName = "Host1";

    ConnectFVTerm('FVTerm', 'http://localhost/FVTerm/SCTerm.html', requestData, function () {

        fvt = fvTermSession;

        fvt.onnewscreen = onNewScreen;
    });
}

// Returns not more than length chars from the current screen starting at row, column,
// trimming them of any whitespace at the start or end.
function GetTrimmedText(row, column, length) {

    var text = fvt.GetScnText(row, column, length);

    if (text == undefined || text == null) {
        return "";
    }

    return $.trim(text);
}