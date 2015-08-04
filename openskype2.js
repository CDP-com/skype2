// JavaScript Document: openskype2.js
// Open second instance of skype
var shell = new ActiveXObject("WScript.Shell");
var myProgramFiles = shell.ExpandEnvironmentStrings( "%ProgramFiles%" );
var ef = "/secondary";
var eff = fnToLocalURL( escape( myProgramFiles + '\\Skype\\Phone\\Skype.exe') );
eff += " " + ef;

try
{
    shell.Run(eff,1,false);
}
catch (e)
{
	var wsh = new ActiveXObject("WScript.Shell");
	wsh.popup( "Unable to start Skype!", 5, "Skype2 App Message" );

	wsh = null;
}

shell = null;

function fnToLocalURL( strEscaped )
{
    var sReturn = "";
    if ( strEscaped.indexOf( 'file' ) == 0 )
    {
        sReturn += strEscaped;
    } else {
        sReturn = 'file:///';
        sReturn += strEscaped;
    }

    //for (i=0; i<sReturn.length; i++)
    //{
        sReturn = sReturn.replace(/%3A/g, ':');
        sReturn = sReturn.replace(/%5C/g, '/');
        sReturn = sReturn.replace(/%22/g, '');
        sReturn = sReturn.replace(/%5B/g, '[');
        sReturn = sReturn.replace(/%5D/g, ']');
    //}
    var nPos1 = sReturn.lastIndexOf( "//" );
    if ( nPos1 > 18 )
    {
        var sReturn1 = ( sReturn.substr( 0, nPos1 ) + sReturn.substring( nPos1 + 1 ) );
        sReturn = sReturn1;
        //WScript.Echo(sReturn);
    }

    return sReturn;
}

