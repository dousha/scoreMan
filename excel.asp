<!--
DOCTYPE: asp/asp: modifier.asp
Author: dousha@github.com
Encoding: UTF-8
Tabstop: 4
-->
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
			''' DBG
			on error resume next
			response.charset = "utf-8"
			' reso
			response.write(request.querystring("path"))
			conn = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source="
			conn = conn & request.querystring("path")
			set rs = server.createobject("ADODB.Recordset")
			rs.activeconnection = conn
			rs.source = "SELECT * FROM [Sheet1$]"
			rs.open
			response.write(rs.Cells(1, 1).value)
		%>
		</div>
	</body>
</html>