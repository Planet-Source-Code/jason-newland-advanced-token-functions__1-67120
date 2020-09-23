Token Functions

by Jason James Newland

From time to time when using VB one needs a Token identifier of some description to do simple string manipulation without the need to do big loops and InStr. There are 8 total token functions within the module AddTok, DelTok, PutTok, RepTok, InsTok, GetTok, IsTok and FindTok.

<string> = AddTok(<string>, <string to add>, <token/delimiter>)
<string> = DelTok(<string>, <string to delete>, <tokenposition>, <token/delimiter>)
<string> = PutTok(<string>, <string to add>, <tokenposition>, <token/delimiter>)
<string> = RepTok(<string>, <string to replace>, <tokenposition>, <token/delimiter>)
<string> = InsTok(<string>, <string to insert>, <tokenposition>, <token/delimiter>)
Debug.Print IsTok(<string>, <string to check>, <token/delimiter>) [returns True or False]
Debug.Print FindTok(<string>, <string to find>, <startposition>, <token/delimiter>) [returns 	token position as Integer]
	an example usage would be FindTok("hi,there", "there", 1, 44) = 2 as token two in 	the string
and finally
Debug.Print GetTok(<string>, <tokenposition>[-], <token/delimiter>, [<toposition>])

<toposition> is optional and is used in combination with <tokenposition> followed by a -

example

GetTok("hi this is a test", "2-", 32, 4) would return "this is a" tokens 2 to 4
GetTok("hi this is a test", "2-", 32) would return "this is a test" from token 2 the whole 		rest of string
GetTok("hi,i,a,separated,by,commas", "4", 44) would return "separated" token 4 and only 	token 4

Hope this has proved useful for someone and I wrote it over 6 months ago as the sources on the net i found for GetTok just did not provide the functionality of string manipulation that I required. A good use for this module would be in an IRC client or such other chat program (or anything for that matter) for parsing and splitting the incoming socket data.

:) have fun!
