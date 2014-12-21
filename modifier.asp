<!--
DOCTYPE: asp/asp: modifier.asp
Author: dousha@github.com
Encoding: UTF-8
Tabstop: 4
-->
<!DOCTYPE html>
<html>
	<head>
			<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.1/css/bootstrap.min.css">
			<script src="http://cdn.bootcss.com/jquery/1.11.1/jquery.min.js"></script>
			<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.1/js/bootstrap.min.js"></script>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	</head>
	<body>
		<div class="container">
		<%
			codepage="65001"
			response.charset = "utf-8"
			set name = request.form("name")
			set id = request.form("id")
			set subject = request.form("subject")
			set newScore = request.form("newScore")
			sql = "UPDATE unidata SET " & subject & "=" & newScore
			sql = sql & " WHERE namee='" & name & "' AND ID=" & id
			set conn = CreateObject("ADODB.Connection")
			conn.provider = "Microsoft.JET.OLEDB.4.0"
			conn.open(server.mappath("database\unidata.mdb")) ' Student database
			conn.execute(sql)
			response.write("更改已提交")
		%>
		</div>
	</body>
</html>
