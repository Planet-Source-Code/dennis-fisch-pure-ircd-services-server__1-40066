				*********************************************************
				*     First Linkable pure Visual Basic IRCd ever        *
				*	It works Linked and Standalone                  *
				*		Version 3.0.272			        *
				*   Suggestions/Bugs to pure@pure-irc.tk.de please      *
				*    	    WebPage: http://pure-irc.tk/                *
				*********************************************************

Features:
	Server Links (no limits on this, but no server must ever be - directly or indirectly - be linked twice with another server or they will end
	up sending each other the same message all the time.

	The 4 most important Services as external server, so they will act global(needs to be linked), ChanServ, NickServ, MemoServ, OperServ.
	Vhost cloak for Operators.
	Detects Netsplits mostly correctly, like...it knows which User's it cant reach anymore after two server have split.
	Very fast while using very few system resources.
	Extensive Server logs
	very detailed .conf file.
	Clone Control (Session Limit)
	Maximum Topic/Kick reason/PrivMsg/Notice and Nickname length (customizable in the .conf file)
	Maximum length of packet customizable through .conf file.
	Highly customizable logs in .conf file.
	
	Generic Commands:
		Nick (bad characters will be filtered)
		User
		Join (Key Support)
		Part (with reason)
		PrivMsg (to Channel and to Nick)
		Notice (to Channel and to Nick)
		Kick (with reason)
		Mode (Channel modes: bceiIklmnpst, usermodes: osw (to be expanded))
		UserHost
		Quit (with Quit Message)
		Topic
		Invite
		Pong
		Away

	Client Queris:
		Ping
		Ison
		Whois
		Motd
		Version
		Time
		Info
		Lusers
		Stats (currently only the parameter "u" is supported, to be expanded)
		Info
		Links
		Names
		List
		Admin
		Who (to be improved)

	Administration:
		Oper
		Restart
		Die
		Kline
		kill
		Akill
		Rehash
		WriteHash (out of date, will be removed soon)
		ClientInfo
		Delete (to maintain a corrupted Server's User Database)

	Links:
		Connect
		Squit

	Services:
		NickServ (also available through "/msg ns" and "/ns")
		ChanServ (also available through "/msg cs" and "/cs")
		OperServ (also available through "/msg os" and "/os")
		MemoServ (also available through "/msg ms" and "/ms")


Inbuilt Services:
	NickServ:
		Register
		Drop
		Kill
		Identifiy
		Info
		ChangeInfo
		Help
	ChanServ:
		Register
		Drop
		Identify
		AddToList
		RemoveFromList
		List
		Clear
	MemoServ:
		Read
		Send
		List
		Delete
	OperServ:
		Stats
		Addstaff
		Delstaff
		Kill
		Akill
		Clear
		Global
		Logon News


Known issues:
	"Hangs" sometimes, no idea why, seems to be a non-VB issue.
	Some netsplits arent resolved correctly, but thats easy for an operator to maintain. Often ghosts remain in the channel.
	Services are only Locally at this moment, this will be changed soon.
	

Linking:
	There is a Setting in the .conf file called "LinkPort", set it to something accessible by the internet.
	now other Pure-IRCd can link to you. To link to another server yourself, use this command:
		/connect [ip or dns] [port]
	because you'd have to be oper you will recieve Servermessages, and one will read this:
		[Your ServerName] -- linked -- [RemoteServerName]
	you can terminate a link connection only if you know the other's server name, you can look that up with this command:
		/os stats
	the use this command to terminate the Link:
		/squit [ServerName]
	You should never ever ever Link two Server's directly or indirectly twice because they will end up sending each other the same message back and forth all the time.