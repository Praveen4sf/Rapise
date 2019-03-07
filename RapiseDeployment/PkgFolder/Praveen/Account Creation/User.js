//Put your custom functions and variables in this file



/**
 * Navigates to the specified URL and performs login at https://login.microsoftonline.com/
 * Opens a browser if necessary.
 * @param url
 * @param userName
 * @param password
 */
function LoginMicrosoftOnline(/**string*/ url, /**string*/ userName, /**string*/ password)
{
	var o = {
		"UseAnotherAccount": "//div[@id='otherTileText']",
		"UserName": "//input[@name='loginfmt']",
		"Sumbit": "//input[@type='submit']",
		"Password": "//input[@name='passwd' and @type='password']",
		"DontShowAgain": "//input[@name='DontShowAgain']",
		"No": "//input[@type='button' and @id='idBtn_Back']"
	};

	Navigator.Open(url);
	
	if (IsSeleniumTest() && !IsMobileTest())		
	{
		WebDriver.SetBrowserSize(300, 800);
	}
	
	Global.DoSleep(5000);
	
	Tester.SuppressReport(true);

	try
	{
		if (Navigator.Find(o["UseAnotherAccount"]))
		{
			Navigator.Find(o["UseAnotherAccount"]).DoClick();	
		}
		
		Navigator.Find(o["UserName"]).DoSetText(userName);
		Navigator.Find(o["Sumbit"]).DoClick();
		Navigator.Find(o["Password"]).DoSetText(password);
		Navigator.Find(o["Sumbit"]).DoClick();
		
		if (Navigator.Find(o["DontShowAgain"]))
		{
			Navigator.Find(o["No"]).DoClick();	
		}
		
		Tester.SuppressReport(false);
		Tester.Message("Logged in as " + userName);
	}
	catch(e)
	{
		Tester.SuppressReport(false);	
		Tester.Message(e.message);
	}
}

function IsMobileTest()
{
	return g_browserProfile.indexOf("Android") != -1 || g_browserProfile.indexOf("iPhone") != -1;
}

function IsSeleniumTest()
{
	return (typeof(WebDriver) != "undefined" && WebDriver);
}

function IsAndroidTest()
{
	return g_browserProfile.indexOf("Android") != -1;
}

function GetWebDriverNonProfileCapabilities(profile)
{
    var caps = {};

// set capabilities based on profile name
    if (profile == "Android")
    {
        caps["platformName"] = "Android";
        caps["platformVersion"] = "8.1";
        caps["deviceName"] = "Android Emulator";
        caps["browserName"] = "Chrome";
    }
    else if (profile == "Android Device")
    {
        caps["platformName"] = "Android";
        caps["platformVersion"] = "6.0.1";
        caps["deviceName"] = "Nexus";
        caps["browserName"] = "Chrome";
        caps["udid"] = "0af5f98b02b5ead0";
    }    
    else if (profile == "iPhone")
    {
        caps["platformName"] = "iOS";
        caps["platformVersion"] = "11.4";
        caps["deviceName"] = "iPhone X";
        caps["browserName"] = "Safari";
        caps["automationName"] = "XCUITest";
        caps["newCommandTimeout"] = "300";
        caps["unexpectedAlertBehaviour"] = "dismiss";
    }
    else if (profile == "iPhone Device")
    {
        caps["platformName"] = "iOS";
        caps["platformVersion"] = "10.3.3";
        caps["deviceName"] = "iPhone 5s";
        caps["browserName"] = "Safari";
        caps["automationName"] = "XCUITest";
        caps["newCommandTimeout"] = "300";
        caps["udid"] = "b6789598c42703429379d147a6f81ecea95edb66";
        caps["xcodeOrgId"] = "3DD36US3JF";
        caps["xcodeSigningId"] = "iPhone Developer";
    }    
    else if (profile == "iPad")
    {
        caps["platformName"] = "iOS";
        caps["platformVersion"] = "10.3";
        caps["deviceName"] = "iPad Air 2";
        caps["browserName"] = "Safari";
        caps["automationName"] = "XCUITest";
        caps["newCommandTimeout"] = "300";
    }    
    return caps;
}

function KillBrowser()
{
	if (g_browserLibrary == "Chrome HTML")
	{
		Global.DoKillByName('chrome.exe');
		Global.DoKillByName('RapiseChromeProxy.exe');
	}
	
	if (g_browserLibrary == "Firefox HTML")
	{
		Global.DoKillByName('firefox.exe');
		Global.DoKillByName('RapiseChromeProxy.exe');
	}
	
	if (g_browserLibrary == "Internet Explorer HTML")
	{
		Global.DoKillByName('iexplore.exe');
	}
}

/** @scenario Login*/
function Login(/**string*/ appId)
{
	KillBrowser();
	var url = "https://ustgsandbox.crm.dynamics.com/main.aspx?appid=" + appId;	//"https://inflectra365.crm.dynamics.com/main.aspx?appid=" + appId;
	var userName = "tester1@mroedge.com";										//"adamsandman@inflectra365.onmicrosoft.com";
	var password = "TeamIndia1!"; 												//Global.DoDecrypt(Global.GetProperty('Password'));
	LoginMicrosoftOnline(url, userName, password);
	
	// iOS prompt to increase database size
	//Global.DoSleep(5000);
	//WebDriver.SwitchToAlert().Accept();	
}

function LoginFieldService(/**string*/ appId)
{
	KillBrowser();
	var url = "https://ustgsandbox.crm.dynamics.com/main.aspx?appid=" + appId + "&pagetype=entitylist&etn=msdyn_workorder" ;	//"https://inflectra365.crm.dynamics.com/main.aspx?appid=" + appId;
	var userName = "tester1@mroedge.com";										//"adamsandman@inflectra365.onmicrosoft.com";
	var password = "TeamIndia1!"; 												//Global.DoDecrypt(Global.GetProperty('Password'));
	LoginMicrosoftOnline(url, userName, password);
	
	// iOS prompt to increase database size
	//Global.DoSleep(5000);
	//WebDriver.SwitchToAlert().Accept();	
}



function ClickMenu(/**objectId*/ openButton, /**objectId*/ menuItem)
{
	var maxCount = 3;
	var timeout = 5000;
	
	var item = false;
	var count = 0;
	while(item == false)
	{
		if (count == maxCount)
		{
			break;
		}
		var button = SeS(openButton);
		button.DoClick();
		item = Global.DoWaitFor(menuItem, timeout);
		count++;
	}

	if (item)
	{
		item.DoClick();
	}
	else
	{
		Tester.Assert(menuItem + " found", false);
	}
}

function ClickListItem(/**objectId*/ objectId, /**string|number*/ item)
{
	SeS(objectId).DoClickItem(item);
	if (IsMobileTest() || IsSeleniumTest())
	{
		return;
	}

	Global.DoSleep(1000);
	var obj = Global.DoWaitFor(objectId, 1);
	if (obj)
	{
		obj.DoClickItem(item);
	}
}

function ClickWhilePresent(/**objectId*/ objectId)
{
	SeS(objectId).DoClick();
	
	if (IsAndroidTest())
	{
		Global.DoSleep(3000);
		var obj = Global.DoWaitFor(objectId, 1);
		if (obj)
		{
			obj.DoClick();
		}
	}
}

function SwitchToLastWindow()
{
	if (IsSeleniumTest())
	{
		WebDriver.SwitchToLastWindow();
	}
}

var count = 1;
function SaveDomTree()
{
	try
	{
		var baseName = "Snapshot" + count;
		count++;
		var dom = Navigator.GetDomTree(false);
		Tester.Assert("DOM loaded: " + baseName, dom != null);
		if (dom)
		{
			Navigator.SaveDom(baseName + '.json', dom);
			Navigator.SaveDomToXml(baseName + '.xml', dom, true);
			Navigator.DoScreenshot(baseName + ".png");
		}
	}
	catch(e)
	{
		Tester.Assert(e.message, false);	
	}
}


