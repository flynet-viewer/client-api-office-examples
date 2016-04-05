var FVCtlMarker = true,
	fnRestartEmulator = null,
	fnStopEmulator = null,
	fvTermSession = null;

function FVTermSession(fvWin)
{
	this.fvWin = fvWin;
	this.emSess = fvWin.em;
	this.serverInfo = fvWin.serverInfo;
	this.screenCount = 0;
	this.onsend = this.onrestart = this.fnNextScreen = this.onnewscreen = this.onclose = this.onchange = this.onsave = null;
	fvWin.emEvents.RegisterClient(this);
}
var FVT = FVTermSession.prototype;

// Called for each screen being displayed.  If the FVTerm web.config
// file has an application and definition file defined, will include
// the recognized screen's name or default if not recognized.
// return value is ignored
FVT.OnLoad = function (emSess, screenName)
{
	this.emSess = emSess;
	if (this.fnNextScreen)
	{
		this.fnNextScreen(this.emSess.GetScnRows(1, emSess.rows), screenName);
		this.fnNextScreen = null;
	}
	if (this.onnewscreen)
		this.onnewscreen(this.emSess.GetScnRows(1, emSess.rows), screenName);
}

// Called when the session is closed
FVT.OnClose = function ()
{
	var callClose = (this.emSess != null);
	this.emSess = null;
	if (callClose && this.onclose)
		this.onclose();
}

// Called on each screen being sent to host
// Return false to prevent screen from being sent
FVT.OnSend = function (emSess, key)
{
	if (this.onsend)
		return this.onsend(key);
	return true;
}

// Called when user clicks the Connect...
// return value ignored
FVT.OnRestart = function ()
{
	this.emSess = null;
	if (this.onrestart)
		this.onrestart();
}

// Called each time focus moves off a field
// emSess = active session object
// field = active field object
// aidKey = if present means this has been called at the start of
//		a screen send.  
// Can return false to cancel the send
FVT.OnChange = function (emSess, field, aidKey)
{
	if (this.onchange)
		return this.onchange(field.val, field.row, field.col, field.len, field.id, aidKey);
	return true;
}

// The screen contents (user entries) have been saved to the session.
// This is a callback which will result from calling the emSess.fnSaveChanges(emSess)
FVT.OnSave = function ()
{
	if (this.onsave)
		this.onsave();
}

FVT.StartMacro = function (macroName, isPub)
{
	if (!this.emSess)
		return false;
	this.fvWin.fvmCtl.StartMacro(macroName, isPub);
	return true;
}

FVT.SetCursor = function (row, column, readyFunc, timeOut)
{
	if (!this.emSess)
		return false;
	var em = this.emSess,
		 targCsrO = (row - 1) * em.cols + column - 1,
		 atRow, atCol, keys, key, downCount, rightCount,
		 me = this, readyId;

	if (targCsrO != em.csrO)
	{
		if (em.formMode)
		{
			em.csrO = targCsrO;
			this.fvWin.SetCursor(this.emSess);
		}
		else
		{
			atRow = Math.floor(em.csrO / em.cols);
			atCol = Math.floor(em.csrO % em.cols);
			keys = '';
			downCount = row - atRow;
			rightCount = column - atCol;
			if (downCount >= 0)
				key = '[down]';
			else
			{
				key = '[up]';
				downCount *= -1;
			}
			while (downCount--)
			{
				keys += key;
			}
			if (rightCount >= 0)
				key = '[right]';
			else
			{
				key = '[left]';
				rightCount *= -1;
			}
			while (rightCount--)
			{
				keys += key;
			}
			this.fvWin.SendKeys(keys, this.emSess.hostWriteCount, false);
		}
	}
	if (readyFunc)
	{
		if (!timeOut)
			timeOut = 10; // seconds
		timeOut *= 10;   // # 100 millisecond intervals
		readyId = window.setInterval(function ()
		{
			if (em.csrO == targCsrO)
			{
				window.clearInterval(readyId);
				readyFunc(true);
			}
			else
			{
				timeOut--;
				if (timeOut <= 0)
				{
					window.clearInterval(readyId);
					readyFunc(false);
				}
			}
		}, 100);
	}
	return true;
}

FVT.SendKeys = function (keys, onNextScreen)
{
	if (!this.emSess)
		return;
	if (onNextScreen)
		this.fnNextScreen = onNextScreen;
	this.fvWin.SendKeys(keys, this.emSess.hostWriteCount, false);
}

FVT.GetField = function (row, column)
{
	if (!this.emSess)
		return null;
	return this.fvWin.SCGetField(this.emSess, row, column);
}

FVT.SetField = function (row, column, text)
{
	if (!this.emSess)
		return;
	this.fvWin.SCSetField(this.emSess, row, column, text);
}

FVT.GetScnText = function (row, column, length)
{
	if (!this.emSess)
		return null;
	return this.fvWin.SCScreenVal(this.emSess, row, column, length);
}

FVT.GetScnRows = function (startRow, endRow)
{
	if (!this.emSess)
		return null;
	return this.emSess.GetScnRows(startRow, endRow);
}

var retries = 0;
function ConnectFVTerm(frameID, pageUri, options, fnReady)
{
	var fvTerm = document.getElementById((frameID) ? frameID : 'FVTerm'),
		 sessionKey = (options) ? options.sessionKey : null,
		 hostName = (options) ? options.hostName : null,
		 autoStart = (options) ? options.autoStart : null,
		 userLoc = (options) ? options.userLocation : null,
		 keepAlive = (options) ? options.keepAliveOnClose : null,
		 connectTo = (options) ? options.connectTo : null,
		 application = (options && options.application) ? options.application : '',
		 userID = (options && options.userID) ? options.userID : null,
		 connectToText = '';
	
	retries = 0;

	if (connectTo)
	{
		connectToText = '*connect*'
		+ ((connectTo.hostName) ? connectTo.hostName : '') + '*'
		+ ((connectTo.ipAddress) ? connectTo.ipAddress : '') + '*'
		+ ((connectTo.port) ? connectTo.port : '') + '*'
		+ ((connectTo.connection) ? connectTo.connection : '') + '*'
		+ ((connectTo.termType) ? connectTo.termType : '') + '*'
	}

	if (!pageUri)
		pageUri='/FVTerm/SCTerm.html';
	if ((application!='') || sessionKey || hostName || autoStart || connectTo)
		pageUri+='?Application='+application;
	if (sessionKey)
		pageUri+='&sessionKey='+sessionKey;
	if (connectTo)
		pageUri += '&connectTo=' + connectToText;
	if (hostName)
		pageUri += '&hostName=' + encodeURIComponent(hostName);
	if (keepAlive)
		pageUri += '&keepAlive=' + ((keepAlive)? 'true':'false');
	if (userLoc)
		pageUri += '&userLoc=' + encodeURIComponent(userLoc);
	if (userID)
		pageUri += '&userID=' + encodeURIComponent(userID);
	if (hostName || autoStart || connectTo)
		pageUri+='&AutoStart=true';

	if(fvTerm.src)
		fvTerm.src = pageUri;
	else if(fvTerm.contentWindow !== null && fvTerm.contentWindow.location !== null)
		 fvTerm.contentWindow.location = pageUri;
	else
		fvTerm.setAttribute('src', pageUri); 
	window.setTimeout(function () { ConnectFVTerm2(fvTerm, fnReady) }, 100);
}

function ConnectFVTerm2(fvTerm, fnReady)
{
	var fvWin = (fvTerm.contentWindow) ? fvTerm.contentWindow : fvTerm.contentDocument.parentWindow,
		 fvt;

	try
	{
		if (fvWin && fvWin.emEvents && fvWin.main && fvWin.profileReady)
		{
			fvTermSession = fvt = new FVTermSession(fvWin);
			if (fnReady)
				fnReady(fvt);
		}
		else
		{
			throw { message: 'Not Ready Yet' };
		}
	}
	catch (e)
	{
		retries++;
		if (retries > 20)
			alert("Could not connect to FVTerm!");
		else
			window.setTimeout(function () { ConnectFVTerm2(fvTerm, fnReady) }, 500);
	}
}
