// JavaScript Document
var ManifestFile = "skype2.xml"
var ManifestURL_Local_dir = "%ALLUSERSPROFILE%\\Application Data\\cdp\\skype2\\"
var ManifestURL_Local = ManifestURL_Local_dir + ManifestFile;
var StartUrl_Local = "index.html";
var StartUrl_Remote = "http://snapback-apps.com/skype2/index.html";
var ManifestURL_Remote = "http://snapback-apps.com/skype2/" + ManifestFile;
var browserEXE = "";
var browserParameters = "";
var oShell = new ActiveXObject("WScript.Shell");
findBrowser();

requestcommand();

oShell = null;

function requestcommand()
{
    var isLocal = 1;

    if ( isLocal )
    {
        //WScript.Echo("Skype will load a little while after clicking the OK button!\n\nA 2nd Skype session will start if a previous session is already started.\n\nExpect a delay!");
        output = doCommand2( "openskype" );
                
        if( output.substring( 0, 6 ) == "2,OK,{" )
        {
            //openBrowser( ManifestURL_Local_dir + "results.html" );
            //Stay on page instead of showing results page.
        }
        else if( output.substring( 0, 6 ) == "2,PIPE" )
        {
            openBrowser( ManifestURL_Local_dir + "pipedown.html" );
        }
        else if( output.substring( 0, 6 ) == "6,Upda" )
        {
            openBrowser( ManifestURL_Local_dir + "updating.html" );
        }
        else if( output.substring( 0, 6 ) == "3,OK,{" )
        {
            //openBrowser( ManifestURL_Local_dir + "results.html" );
        }
        else if( output.substring( 0, 6 ) == "4,OK,{" )
        {
            //openBrowser( ManifestURL_Local_dir + "results.html" );
        }
        else if( output.substring( 0, 6 ) == "2,Deny" )
        {
            openBrowser( ManifestURL_Local_dir + "deny.html" );
        }
        else
        {
            openBrowser( ManifestURL_Local_dir + "unknown.html" );
        }
    }
}

function doCommand2( command )
{
    var output = "";
    var UpdObj;
    var command2 = ManifestURL_Local + "," + command;
    UpdObj = new ActiveXObject( "CDPUpdater.Updater" );

    try
    {
        output = UpdObj.ExpandEnvVar( command2 );
        output = UpdObj.RequestCommand( output );
    }
    catch ( e )
    {
        WScript.Echo(e.message);
    }
    finally
    {
        UpdObj = null;
    }

    return output;
}

function GetManifestFile()
{
    return ManifestFile;
}

function GetManifestURL_Remote()
{
    return ManifestURL_Remote;
}

function findBrowser() {
    var browserClass = oShell.RegRead("HKCR\\.html\\");
    var browserShell = oShell.RegRead("HKCR\\" + browserClass + "\\shell\\open\\command\\");
    var exeSplit = String(browserShell).toLowerCase().indexOf(".exe") + 5;
    browserEXE = String(browserShell).substring(0, exeSplit);
    browserParameters = String(browserShell).substr(exeSplit);
    browserParameters = browserParameters.substr(browserParameters.indexOf('"'))
}


function openBrowser(url) {
    var newParameters = '"' + url + '"';
    var newCommand = browserEXE + " " + newParameters;
    oShell.run( newCommand );
}

