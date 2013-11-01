var CONFIG_FILE = "biogen.conf";
var CONFIG;
var DB;
var MAX_RECS = 1000;
var DSN = "";
var RETRY_DELAY_TIME = 15; //number of minutes to wait before retrying a message


if((WScript.Arguments.length >= 2) && (WScript.Arguments.length <= 3)) {
try {
	main();
} catch (e) { WScript.echo(e.description); throw e; }
} else {
	usage();
}


function usage()
{
	WScript.Echo("Usage: cscript sendmail.js <SPAM|QUEUE|ALL|'source'> <monitor_name> [Subject Line]\r\n"+
		     "Sends emails stored in the database.  Mode options:\r\n"+
		     "  ALL - pull from all possible messages, even always excluded source\r\n"+
		     "  QUEUE - pull from non-spam, non-excluded sources\r\n"+
		     "  SPAM - pull from the non-excluded spammers list\r\n"+
		     "  Other - pull from the specified source, even if it is excluded\r\n"+
		     "You must specify a Ganglia monitor name.  msgs_sent_ will be pre-pended.\r\n"+
		     "You may optionally specify a subject name.  If you do, only exact matches will be sent.\r\n");
}

function main()
{
	CONFIG = get_config(CONFIG_FILE);

	if(! CONFIG) {
		WScript.Echo("ERROR: Could not load configuration, aborting");
		return;
	}

	var mode = WScript.Arguments(0);
	var queue_name = "" + WScript.Arguments(1);
	var monitor_id = "msgs_sent_" + queue_name;
	var subject_line = (WScript.Arguments.length) == 3 ? WScript.Arguments(2) : false;

WScript.Echo("queue_name: " + queue_name);
WScript.Echo("mode: " + mode);
WScript.Echo("DB: " + CONFIG["EMAIL DB"]);
WScript.Echo("SMTP server: "+CONFIG["EMAIL SERVER"]);

	DB = new ActiveXObject("ADODB.Connection");
	DB.Open(CONFIG["EMAIL DB"]);

	var rs_guid = DB.Execute("select newid()");
	var numericstamp = String(rs_guid(0));
//WScript.Echo("numericstamp: " + numericstamp);

	/* tablockx forces an exclusive table lock Use tablock to do a non-exclusive lock holdlock tells SQL Server to hold locks for duration of the transaction */
	sql="begin transaction \n"+
			"select count(*) from messages with (tablockx holdlock) where 0=1 \n"+ 
			"update Messages "+
				"SET numericStamp='"+numericstamp+"',dateStamp=GETDATE() "+
				"WHERE MessageID IN (select top "+MAX_RECS+" messageID from messages where (dateStamp < DATEADD(n, - "+RETRY_DELAY_TIME+", GETDATE()) OR dateStamp is NULL) and (errorCount<5 or errorCount is null) "+get_spammer_clause(mode)+" " + get_always_exclude_clause(mode) + " " + get_subject_line_clause(subject_line) + " order by errorcount, received)\n" +
		"commit";
WScript.Echo("sql: " + sql);
	DB.Execute(sql);

	sql ="SELECT * FROM Messages WHERE numericStamp='"+numericstamp+"'";
	var rs = new ActiveXObject("ADODB.Recordset");

	try
	{	
WScript.Echo("Start processing email queue");
		var attempted 	= 0;
		var sent 	= 0;
		var save = new ActiveXObject("ADODB.Recordset");
		var status="";
		for (rs.Open(sql,DB,3,3);!rs.EOF;rs.MoveNext())
		{			
			attempted++;				
//WScript.Echo("Start processing email queue 4");
			save.Open("select * from sentMessages where messageID="+rs.Fields("messageID"),DB,3,3);
			if(!save.EOF) {
				WScript.Echo("Message has already been sent");
				save.Close();
				try
				{	rs.Delete();
				}catch(e)
				{	WScript.Echo("Error deleting record -- during duplicate messge blocking: "+e.description);
				}
				attempted--;
				continue;
			}

			var toNames=String(rs.Fields("toName")).split(/[,;]/gi);
			var toAddresses=String(rs.Fields("toAddress")).split(/[,;]/gi);
			var name;
			
			//reference manual is at http://support.softartisans.com/docs/smtpmail/default.htm

			WScript.Echo("Using server: "+CONFIG["EMAIL SERVER"]);
			var smtpConnection = new ActiveXObject("SoftArtisans.SMTPMail");	//this being inside the loop will make it run

			//smtpConnection.SMTPLog ="Mailer"

			smtpConnection.RemoteHost = CONFIG["EMAIL SERVER"]; 		//a little slower, but it'll be more stable
			if (CONFIG["SMTP USER"] != null && CONFIG["SMTP PASSWORD"] != null) {
				smtpConnection.UserName = CONFIG["SMTP USER"];
				smtpConnection.Password = CONFIG["SMTP PASSWORD"];
				WScript.Echo("User name = '" + CONFIG["SMTP USER"] + "' password = '" + CONFIG["SMTP PASSWORD"] + "'");
			}
			for (var i=0;i<toAddresses.length;i++)
			{	// names with commas get treated as extra email addresses.
				// -- Wes, 2005/02/28
				name = (toNames[i]) ? "\"" + toNames[i].replace(/,/g, "").replace(/"/g, "'") + "\"" : ""; //"
				smtpConnection.AddRecipient(name,String(toAddresses[i]).replace(/\s/g, ""));
			}

			// Support for CC
			// Added by Wes 23/03/2006

			var ccName;
			var ccNames=String(rs.Fields("ccName")).split(/[,;]/gi);
			var ccAddresses=String(rs.Fields("ccAddress")).split(/[,;]/gi);
			for (var i = 0; i < ccAddresses.length; i++)
			{	ccName = (ccNames[i]) ? "\"" + ccNames[i].replace(/,/g, "").replace(/"/g, "'") + "\"" : ""; //"
				if(! ccName.match(/NULL$/i)) {
					smtpConnection.AddCC(ccName,String(ccAddresses[i]).replace(/\s/g, ""));
				}
			}


			// filter content
			//
			var content = String(rs.Fields("content")).replace( /’/g, "'");
				content = String(content).replace( /“/g, "\"");
				content = String(content).replace( /”/g, "\"");
				content = String(content).replace( /–/g, "-" );
				content = String(content).replace( /…/g, "..." );

			smtpConnection.FromName    = "\""+String(rs.Fields("fromName")).replace(/"/g, "'") + "\""; //"
			smtpConnection.FromAddress = String(rs.Fields("fromAddress"));
			smtpConnection.Subject     = String(rs.Fields("Subject"));
			smtpConnection.ContentType = String(rs.Fields("ContentType"));
			smtpConnection.BodyText    = content;
			smtpConnection.IgnoreRecipientErrors = false;
			
			WScript.echo((new Date()).toString() + " To: " + String(rs.Fields("toAddress")) + " X-Kci-ProcessorID: " + numericstamp + " processing X-Kci-EmailID: " + String(rs.Fields("messageid")));

			// Add extra headers
			//
			var headers = String(rs.Fields("extraHeaders")).split(/\n/);

			for(var i = 0; i < headers.length; i++)
			{	if(headers[i] != "null")
				{	smtpConnection.AddExtraHeader(headers[i]);
				}
			}

			try {
				// add attachments
				// 1) get the attachment from the DB
				// 2) write it to a temp file
				// 3) send it
				// 4) delete the temp file

				var messageID = rs.Fields("messageID");
				var attachmentRS = DB.Execute("select * from attachments where messageID = "+messageID);
				var fs = new ActiveXObject("Scripting.FileSystemObject");
				var fileStream = new ActiveXObject("ADODB.Stream");
				fileStream.Type = 1; // Binary
				var tempFolder, tempFilename;
				var tempFiles = Array();

				for(; ! attachmentRS.EOF; attachmentRS.MoveNext()) {
					tempFolder = fs.GetSpecialFolder(2);  // 2 = TemporaryFolder
					tempFilename = tempFolder.Path + "\\"+attachmentRS("filename").value;
					fileStream.Open();
					fileStream.Write(attachmentRS("data").value);
					//Perhaps adding a flush will finish the writing of the file before it is saved
					fileStream.Flush();
					fileStream.SaveToFile( tempFilename, 2); // 2 = overwrite
					fileStream.Close();
					smtpConnection.AddAttachment(tempFilename);
					tempFiles.push(tempFilename);
				}
				attachmentRS.Close();
				// Changed to use a properly formatted RFC 2822 date string
				smtpConnection.DateTime = formatDateSMTP(new Date());

				smtpConnection.CharSet = Number(rs.Fields("charSet").value);
				smtpConnection.ContentTransferEncoding = 4
				status = String(smtpConnection.SendMail());

				if (status != "true")
					WScript.Echo("Failed to send: " + smtpConnection.Response);


				var i;
				for (i in tempFiles) {
					try {
						fs.DeleteFile(tempFiles[i]);
					} catch(e) {
						WScript.Echo("Warning: could not delete " + tempFilename);
					}
				}
			} catch(e) {
				status = false;
				WScript.Echo("Error processing attachments.  MessageID = " + messageID + ", Error description: " + e.description);
			}

			if(status == "true")
			{	sent++;
				try
				{	save.AddNew();
					save.Fields("messageID")     = Number(rs.Fields("messageID"));
					save.Fields("toName")        = String(rs.Fields("toName"));
					save.Fields("toAddress")     = String(rs.Fields("toAddress"));
					save.Fields("fromName")	     = String(rs.Fields("fromName"));
					save.Fields("fromAddress")   = String(rs.Fields("fromAddress"));
					save.Fields("subject")	     = String(rs.Fields("Subject"));
					save.Fields("ContentType")   = String(rs.Fields("ContentType"));
					save.Fields("content")	     = String(content);
					save.Fields("notifyOnError") = String(rs.Fields("notifyOnError"));
					save.Fields("source")	     = String(rs.Fields("source"));
					save.Fields("received")	     = rs.Fields("received").value;
					save.Fields("ccName")	     = rs.Fields("ccName").value;
					save.Fields("ccAddress")     = rs.Fields("ccAddress").value;
					save.Fields("ErrorAppended") = rs.Fields("ErrorAppended").value;
					save.Update();
				}catch(e)
				{	WScript.Echo("Could not save: "+ e.description); throw e;
				}
					// By design, calling delete or update may randomly throw an error
					// Think that I'm kidding?  -> http://support.microsoft.com/default.aspx?scid=KB;en-us;q195491
					// Or search for 'Cursor operation conflict' on google
				try
				{	rs.Delete();
				}catch(e)
				{	WScript.Echo("Error deleting record: "+e.description);
				}
			}else
			{	try
				{	WScript.Echo( "ERROR: status is "+status );
					if(rs("errorCount") == null)
					{	var eCount = 0;
					} else
					{	var eCount = Number(rs("errorCount").value);
					}
					eCount++;
					rs("errorCount") = eCount;
					rs.Update();
				}catch(e)
				{	WScript.Echo("Error updating error count: "+e.description);
				}
			}
			save.Close();
		}
	}catch( e )
	{ 	WScript.Echo("Error processing queue: "+e.description);
	}
	try
	{	rs.Close();
	}catch( e )
	{ 	WScript.Echo("Error closing recordset: "+e.description);
	}

	WScript.Echo("Emails attempted: "+attempted+"\n"+"Emails sent: "+sent+"\n"+"Emails queued: "+get_queued());

	DB.Close();
}

function get_spammer_clause(mode)
{
	if (mode == "ALL") { 
		return ""; 
	} else if (mode != "SPAM" && mode != "QUEUE"  && mode != "!SPAM") {
		return " AND (source = '"+mode+"')";
	} else if (mode == "QUEUE") {
		return get_spammer_clause("!SPAM");
	} else {
		sources = [];
		var rs = DB.Execute("select sourcename from spammers");
		for (; !rs.EOF; rs.MoveNext()) {
			sources.push("'" + rs("sourcename").value + "'");
		}
		rs.Close();
	
		var not = (mode == "SPAM") ? "" : "NOT";
		return sources.length == 0 ? "" :  " AND source " + not + " IN (" + sources.join(",")+")";	
	}
}

function get_always_exclude_clause(mode)
{
	if (mode != "SPAM" && mode != "QUEUE") {
		return ""; 
	} else {
		sources = [];
		var rs = DB.Execute("select source from AlwaysExclude");
		for (; !rs.EOF; rs.MoveNext()) {
			sources.push("'" + rs("source").value + "'");
		}
		rs.Close();
		return sources.length == 0 ? "" :  " AND source NOT IN (" + sources.join(",")+")";	
	}
}

function get_subject_line_clause(subject_line) {
	if (subject_line) {
		return "AND Subject LIKE '" + subject_line + "'";
	} else {
		return "";
	}
}

function get_queued(unsendable)
{	var error_clause = (unsendable) ? "(errorcount >= 5)" : "((errorcount < 5) or errorcount is null)";
	var sql = "SELECT count(*) as cnt FROM Messages WHERE toaddress != '' and "+error_clause;
	var rs = new ActiveXObject("adodb.recordset");
	rs.Open(sql, DB);
	var cnt = rs("cnt").value;
	rs.Close();
	return cnt;
}

function get_config(configfile)
{
	var fs = new ActiveXObject("Scripting.FileSystemObject");

	try {
		var file = fs.GetFile(configfile || "monitor.conf");
		var ts = file.OpenAsTextStream(1, -2);
	} catch(e) {
		WScript.Echo("ERROR: Could not open configuration file monitor, aborting");
		return null;
	}

	var config = {};
	var line;
	while (!ts.AtEndOfStream) {
		line = ts.ReadLine();
		if(! line.match(/(.*)::(.*)/)) {
			WScript.Echo("WARNING: Could not parse configuration file line: "+line);
			continue;
		}

		config[RegExp.$1] = RegExp.$2;
	}

	return config;
}



	//-----------------------------------------------------------------
	// formatDate()
	// takes in a date object, and formats it according to the
	// formatStr parameter.  The tokens that formatStr recognizes:
	// 			mmmm - long month name
	//			mmm  - short month name
	//			mm   - 2 digit month number (02 = feb)
	//			m    - 1 digit month number (2 = feb)
	//			yyyy - 4 digit year
	//			yy   - 2 digit year (please don't use this)
	//			dd   - 2 digit day (02 for the second)
	//			d    - 1 digit day (2 for the second)
	// All other characters are ignored, allowing you to do something like:
	//	"mmmm-dd, yyyy" to get "January-02, 2002"\
	// ADDITION: French date formatting can be handled by using the last param,
	// 	and sending it "fr"
	//-----------------------------------------------------------------
	function formatDate(inputDate,formatStr,lang)
	{	var day=inputDate.getDate();
		var mon=inputDate.getMonth();
		var yr =inputDate.getFullYear();
		var rtn=String(formatStr);

		var hrs=(inputDate.getHours() % 12);
		if (hrs==0)															//get the number of hours into non-24 hour time
			hrs=12;
		var ampm=" AM";
		if (inputDate.getHours()>=12)												//figure out if it's am or pm
			ampm=" PM"
		var mins=String("0"+inputDate.getMinutes());								//and get the minutes into a 2 character string
		mins=mins.substr(mins.length-2);
		var timeStr=hrs+":"+mins+ampm;
		var milTimeStr=inputDate.getHours()+":"+mins;

		var shortMonth=new Array(	"Jan","Feb","Mar",
									"Apr","May","Jun",
									"Jul","Aug","Sep",
									"Oct","Nov","Dec");
		var longMonth =new Array(	"January","February","March",
									"April",  "May",     "June",
									"July",   "August",  "September",
									"October","November","December");
		var longDaysOfWeek=new Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday");
		if (arguments.length>2 && String(lang).toLowerCase().substr(0,2)=="fr")
		{	shortMonth=new Array(	"janv",	"févr",	"mars",
									"avr",	"mai",	"juin",
									"juil",	"août",	"sept",
									"oct",	"nov",	"déc");
			longMonth=new Array(	"janvier",	"février",	"mars",
									"avril",	"mai",		"juin",
									"juillet",	"août",		"septembre",
									"octobre",	"novembre",	"décembre");
			longDaysOfWeek=new Array("Dimanche",
									"Lundi",
									"Mardi",
									"Mercredi",
									"Jeudi",
									"Vendredi",
									"Samedi");
		}
		var orig=rtn;
		rtn=rtn.replace("dd",(day<10?"0":"")+day);
		if (rtn==orig)
		{	rtn=rtn.replace("d",day);
		}
		rtn=rtn.replace("yyyy",yr);
		rtn=rtn.replace("yy",String(yr).substr(2));
		orig=rtn;
		rtn=rtn.replace("mmmm",longMonth[mon]);
		rtn=rtn.replace("mmm",shortMonth[mon]);
		rtn=rtn.replace("mm",(mon<9?"0":"")+(mon+1));

		rtn=rtn.replace("HH:MM AMPM",timeStr);
		rtn=rtn.replace("HH:MM MIL",milTimeStr);

		rtn=rtn.replace("HH:MM:SS AMPM",hrs+":"+mins+":"+inputDate.getSeconds()+ampm);
		rtn=rtn.replace("HH:MM:SS MIL",milTimeStr+":"+inputDate.getSeconds());

		rtn=rtn.replace("DDDD",longDaysOfWeek[inputDate.getDay()]);
		if (rtn==orig)
		{	rtn=rtn.replace("m",(mon+1));
		}
		return(rtn);
	}

// Pads a number with zeros.
// --- Alf, 2007-08-14
function padZero(integer, length)
{
	natural = String(integer);
	s = "";
	var i;
	for (i=0; i < (length-natural.length); i++)
	{
		s += "0";
	}
	s += String(integer);
	return s;
}

// Takes a date object.
// Returns an RFC 2822 formatted date string, including timezone information.
// Uses local timezone.  Convert to UTC first, if that's what you want.
// --- Alf, 2007-08-14
function formatDateSMTP(d)
{
	var s = "";
	var days = new Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat");
	var months = new Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec");
	s += days[d.getDay()] + ", ";
	s += d.getDate() + " ";
	s += months[d.getMonth()] + " ";
	s += d.getFullYear() + " ";
	s += padZero(d.getHours(),2) + ":";
	s += padZero(d.getMinutes(),2) + ":";
	s += padZero(d.getSeconds(),2) + " ";
	// Javascript uses the opposite sign convention to the rest of the world
	var sign = (d.getTimezoneOffset() > 0) ? "-" : "+";
	var tz = Math.floor(d.getTimezoneOffset() / 60)*100 + (d.getTimezoneOffset() % 60);
	s += sign + padZero(tz,4);
	return s;
}
