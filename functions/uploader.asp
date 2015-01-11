<!--
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
			if isnull(formsize) then
				response.write("未提交文件")
				response.end()
			end if
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
			dr1.savetofile server.mappath("..\database\u") & formatdatetime(now, 1) & ".xls", 2
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
			set rs2 = server.createobject("ADODB.Recordset") ' and it's record set
			qsql = "SELECT * FROM unidata"
			rs2.open qsql, conn2, 1, 1
			' DBG
			rs.movefirst
			response.write("<div class=""panel panel-primary""><div class=""panel-heading"">")
			response.write("您提交了以下数据</div><div class=""panel-body"">")
			response.write("<table class=""table"">")
			set fo = server.createobject("Scripting.FileSystemObject") 'FSO...Not sure if it works
			set obj = fo.opentextfile(server.mappath("log\log.log"), 8, true)
			do while not rs.eof
				' if dup change to update
				needInsert = true
				if not rs2.bof and rs2.eof then
					rs2.movefirst
				end if
				do while not rs2.eof and needInsert = true
					if rs2("ID") = rs("考号") then
						needInsert = false
					end if
					rs2.movenext
				loop
				if needInsert = true then
					' INSERT MODE
					usql = "INSERT INTO unidata "
					inserts = "(ID,class,namee,"
					values = "(" & rs("考号") & ",'" & rs("班级") & "','" & rs("姓名") & "',"
				else
					' UPDATE MODE
					usql = "UPDATE unidata SET "
				end if
				response.write("<tr>")
				response.write("<td>" & rs("考号") & "</td>")
				response.write("<td>" & rs("班级") & "</td>")
				response.write("<td>" & rs("姓名") & "</td>")
				' if null ignore
				if isnull(rs("总分")) = false then
					if needInsert = true then
						inserts = inserts & "total,"
						values = values & rs("总分") & ","
					else
						usql = usql & "total=" & rs("总分") & ","
					end if
					response.write("<td>" & rs("总分") & "</td>")
				end if
				if isnull(rs("数学")) = false then
					if needInsert = true then
						inserts = inserts & "math,"
						values = values & rs("数学") & ","
					else
						usql = usql & "math=" & rs("数学") & ","
					end if
					response.write("<td>" & rs("数学") & "</td>")
				end if
				if isnull(rs("语文")) = false then
					if needInsert = true then
						inserts = inserts & "chi,"
						values = values & rs("语文") & ","
					else
						usql = usql & "chi=" & rs("语文") & ","
					end if
					response.write("<td>" & rs("语文") & "</td>")
				end if
				if isnull(rs("外语")) = false then
					if needInsert = true then
						inserts = inserts & "lang,"
						values = values & rs("外语") & ","
					else
						usql = usql & "lang=" & rs("外语") & ","
					end if
					response.write("<td>" & rs("外语") & "</td>")
				end if
				if isnull(rs("政治")) = false then
					if needInsert = true then
						inserts = inserts & "soc,"
						values = values & rs("政治") & ","
					else
						usql = usql & "soc=" & rs("政治") & ","
					end if
					response.write("<td>" & rs("政治") & "</td>")
				end if
				if isnull(rs("历史")) = false then
					if needInsert = true then
						inserts = inserts & "hist,"
						values = values & rs("历史") & ","
					else
						usql = usql & "hist=" & rs("历史") & ","
					end if
					response.write("<td>" & rs("历史") & "</td>")
				end if
				if isnull(rs("地理")) = false then
					if needInsert = true then
						inserts = inserts & "geo,"
						values = values & rs("地理") & ","
					else
						usql = usql & "geo=" & rs("地理") & ","
					end if
					response.write("<td>" & rs("地理") & "</td>")
				end if
				if isnull(rs("物理")) = false then
					if needInsert = true then
						inserts = inserts & "phy,"
						values = values & rs("物理") & ","
					else
						usql = usql & "phy=" & rs("物理") & ","
					end if
					response.write("<td>" & rs("物理") & "</td>")
				end if
				if isnull(rs("化学")) = false then
					if needInsert = true then
						inserts = inserts & "chem,"
						values = values & rs("化学") & ","
					else
						usql = usql & "chem=" & rs("化学") & ","
					end if
					response.write("<td>" & rs("化学") & "</td>")
				end if
				if isnull(rs("生物")) = false then
					if needInsert = true then
						inserts = inserts & "bio,"
						values = values & rs("生物") & ","
					else
						usql = usql & "bio=" & rs("生物") & ","
					end if
					response.write("<td>" & rs("生物") & "</td>")
				end if
				response.write("</tr>")
				' trim
				if needInsert = true then
					inserts = left(inserts, len(inserts) - 1)
					inserts = inserts & ")"
					values = left(values, len(values) - 1)
					values = values & ")"
					usql = usql & inserts & " VALUES " & values
				else
					usql = left(usql, len(usql) - 1)
					usql = usql & " WHERE ID=" & rs("考号")
				end if
				conn2.execute(usql)
				if not isnull(usql) then
					obj.write "[MODI] " & usql & vbcrlf
				end if
				rs.movenext
			loop
			response.write("</table><div align=""right"">")
			response.write("您可以单击“返回”回到管理页面")
			response.write("</div>")
			set conn = nothing
			set conn2 = nothing
			obj.close
		%>
		</div>
	</body>
</html>