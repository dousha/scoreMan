﻿<!--
DOCTYPE: asp/asp: uploader.asp
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
			' You are NOT expected to understand this.
			codepage="65001"
			function bin2asi(byval varstr)
				asi = ""
				for i = 1 to 3
					asi = asi&chr(ascb(midb(varstr, i, 1)))
				next
				bin2asi = asi
			end function
			' don't know how is it working, copied from the Internet
			response.charset = "utf-8"
			formsize = request.totalbytes
			formdata = request.binaryread(request.totalbytes)
			bcrlf = chrB(13) & chrB(10)
			divider = leftB(formdata, clng(instrb(formdata, bcrlf)) - 1)
			position = instrb(formdata, bcrlf & bcrlf) + 4
			filesize = instrb(position + 1, formdata, divider) - position - 4
			exnamestart = instrb(1, formdata, chrb(46), 1) + 1       
			exnameend = instrb(exnamestart, formdata, chrb(34), 1)
			exname = midb(formdata, exnamestart, exnameend - exnamestart)
			set dr = CreateObject("Adodb.Stream")
			dr.Mode = 3 : dr.Type = 1 : dr.Open
			set dr1 = CreateObject("Adodb.Stream")
			dr1.Mode = 3 : dr1.Type = 1 : dr1.Open
			dr.write formdata
			dr.position = position - 1
			dr.copyto dr1, filesize
			dr1.savetofile server.mappath("database\u") & formatdatetime(now, 1) & ".xls", 2
			set dr = nothing
			set dr1 = nothing
			' I have to use this stupid way to fix errors (missing 00 00)
			set fo = server.createobject("Scripting.FileSystemObject") 'FSO...Not sure if it works
			set obj = fo.opentextfile(server.mappath("database\u") & formatdatetime(now, 1) & ".xls", 8, true)
			obj.write chrB(00) & chrB(00) & chrB(00) & chrB(00)
			obj.close()
			' reso
			conn = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source="
			conn = conn & server.mappath("database\u") & formatdatetime(now, 1) & ".xls"
			set rs = server.createobject("ADODB.Recordset") ' Updated excel
			sql = "SELECT * FROM [Sheet1$]"
			rs.open sql, conn, 1, 1
			' sdb
			set conn2 = CreateObject("ADODB.Connection")
			conn2.provider = "Microsoft.JET.OLEDB.4.0"
			conn2.open(server.mappath("database\unidata.mdb")) ' Student database
			' DBG
			rs.movefirst
			response.write("<div class=""panel panel-primary""><div class=""panel-heading"">")
			response.write("您提交了以下数据</div><div class=""panel-body"">")
			response.write("<table class=""table"">")
			set fo = server.createobject("Scripting.FileSystemObject") 'FSO...Not sure if it works
			set obj = fo.opentextfile(server.mappath("log\log.log"), 8, true)
			do while not rs.eof
				usql = "INSERT INTO unidata "
				inserts = "(ID,class,namee,"
				values = "(" & rs("考号") & ",'" & rs("班级") & "','" & rs("姓名") & "',"
				response.write("<tr>")
				response.write("<td>" & rs("考号") & "</td>")
				response.write("<td>" & rs("班级") & "</td>")
				response.write("<td>" & rs("姓名") & "</td>")
				' if null ignore
				if isnull(rs("总分")) = false then
					inserts = inserts & "total,"
					values = values & rs("总分") & ","
					response.write("<td>" & rs("总分") & "</td>")
				end if
				if isnull(rs("数学")) = false then
					inserts = inserts & "math,"
					values = values & rs("数学") & ","
					response.write("<td>" & rs("数学") & "</td>")
				end if
				if isnull(rs("语文")) = false then
					inserts = inserts & "chi,"
					values = values & rs("语文") & ","
					response.write("<td>" & rs("语文") & "</td>")
				end if
				if isnull(rs("外语")) = false then
					inserts = inserts & "lang,"
					values = values & rs("外语") & ","
					response.write("<td>" & rs("外语") & "</td>")
				end if
				if isnull(rs("政治")) = false then
					inserts = inserts & "soc,"
					values = values & rs("政治") & ","
					response.write("<td>" & rs("政治") & "</td>")
				end if
				if isnull(rs("历史")) = false then
					inserts = inserts & "hist,"
					values = values & rs("历史") & ","
					response.write("<td>" & rs("历史") & "</td>")
				end if
				if isnull(rs("地理")) = false then
					inserts = inserts & "geo,"
					values = values & rs("地理") & ","
					response.write("<td>" & rs("地理") & "</td>")
				end if
				if isnull(rs("物理")) = false then
					inserts = inserts & "phy,"
					values = values & rs("物理") & ","
					response.write("<td>" & rs("物理") & "</td>")
				end if
				if isnull(rs("化学")) = false then
					inserts = inserts & "chem,"
					values = values & rs("化学") & ","
					response.write("<td>" & rs("化学") & "</td>")
				end if
				if isnull(rs("生物")) = false then
					inserts = inserts & "bio,"
					values = values & rs("生物") & ","
					response.write("<td>" & rs("生物") & "</td>")
				end if
				response.write("</tr>")
				' trim
				inserts = left(inserts, len(inserts) - 1)
				inserts = inserts & ")"
				values = left(values, len(values) - 1)
				values = values & ")"
				usql = usql & inserts & " VALUES " & values
				conn2.execute(usql)
				obj.write "[MODI] " & usql & vbcrlf
				obj.close()
				rs.movenext
			loop
			response.write("</table><div align=""right"">")
			response.write("</div>")
		%>
		</div>
	</body>
</html>