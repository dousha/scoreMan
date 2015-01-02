<!--
DOCTYPE: asp/asp: modifier.asp
Author: dousha@github.com
Encoding: UTF-8
Tabstop: 4
-->
<!DOCTYPE html>
<html>
	<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	</head>
	<body>
		<div class="container">
		<%
			codepage="65001"
			response.charset = "utf-8"
			set namee = request.form("name")
			set id = request.form("id")
			set subject = request.form("subject")
			set newScore = request.form("newScore")
			sql = "UPDATE unidata SET " & subject & "=" & newScore
			sql = sql & " WHERE namee='" & namee & "' AND ID=" & id
			set conn = CreateObject("ADODB.Connection")
			conn.provider = "Microsoft.JET.OLEDB.4.0"
			conn.open(server.mappath("database\unidata.mdb")) ' Student database
			conn.execute(sql)
			response.write("更改已提交")
			set fo = server.createobject("Scripting.FileSystemObject") 'FSO...Not sure if it works
			set obj = fo.opentextfile(server.mappath("log\log.log"), 8, true)
			obj.write "[MODI] " & sql & vbcrlf
			obj.close()
		%>
		</div>
	</body>
</html>
