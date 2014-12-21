<!--
DOCTYPE: asp/asp: query.asp
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
			<script>
				function modify(){
					var name = document.getElementById("name").value;
					var id = document.getElementById("id").value;
					var subject = document.getElementById("subject").value;
					var newScore = document.getElementById("newScore").value;
					var ajaxHandler;
					if(window.XMLHttpRequest){
						ajaxHandler = new XMLHttpRequest(); // general ajax object
					}
					else{
						ajaxHandler = new ActiveXObject("Microsoft.XMLHttp"); // micros**t's IE5/6 support
					}
					ajaxHandler.onreadystatechange = function(){
						if(ajaxHandler.readyState == 4 && ajaxHandler.status == 200){
							document.getElementById("message").innerHTML = ajaxHandler.responseText;
						}
					}
					ajaxHandler.open("POST", "modifier.asp", true);
					ajaxHandler.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
					ajaxHandler.send();
				}
			</script>
	</head>
	<body>
		<div class="container">
		<%
			codepage="65001"
			response.charset = "utf-8"
			dim name
			dim id
			dim pwd
			name = request.form("name")
			id = request.form("id")
			pwd = request.form("pwd")
			' jud if stu
			if pwd = empty then
				' Student querying mode
				set rs = nothing
				response.expires = -1
				sql = "SELECT * FROM unidata WHERE namee = "
				sql = sql & "'" & name & "' AND ID = " & id
				' mv cdb \v
				set conn = CreateObject("ADODB.Connection")
				conn.provider = "Microsoft.JET.OLEDB.4.0"
				conn.open(server.mappath("database\unidata.mdb")) ' Student database
				set rs = CreateObject("ADODB.recordset")
				rs.Open sql, conn
				' Now query should be completed
				' and print out the panel
				response.write("<div class=""panel panel-primary""><div class=""panel-heading"">")
				response.write(rs("class") & "班的" & name & "同学, 最近一次考试的成绩如下")
				response.write("</div><div class=""panel-body"">")
				if rs.bof AND rs.eof then
					response.write("没有查询到成绩，名字或者学号是不是输错了？")
				else
					response.write("<table class=""table"">")
					response.write("<tr><td>项目</td><td>值</td></tr>")
					response.write("<tr><td>总分</td><td>" & rs("total") & "</td>")
					response.write("<tr><td>语文</td><td>" & rs("chi") & "</td>")
					response.write("<tr><td>数学</td><td>" & rs("math") & "</td>")
					response.write("<tr><td>语言</td><td>" & rs("lang") & "</td>")
					if isnull(rs("soc")) = true then
						response.write("<tr><td>物理</td><td>" & rs("phy") & "</td>")
						response.write("<tr><td>化学</td><td>" & rs("chem") & "</td>")
						response.write("<tr><td>生物</td><td>" & rs("bio") & "</td>")
						' DBG
						response.write("理科模式")
					elseif isnull(rs("phy")) = true then
						response.write("<tr><td>政治</td><td>" & rs("soc") & "</td>")
						response.write("<tr><td>历史</td><td>" & rs("hist") & "</td>")
						response.write("<tr><td>地理</td><td>" & rs("geo") & "</td>")
						response.write("文科模式")
					else
						response.write("<tr><td>物理</td><td>" & rs("phy") & "</td>")
						response.write("<tr><td>化学</td><td>" & rs("chem") & "</td>")
						response.write("<tr><td>生物</td><td>" & rs("bio") & "</td>")
						response.write("<tr><td>政治</td><td>" & rs("soc") & "</td>")
						response.write("<tr><td>历史</td><td>" & rs("hist") & "</td>")
						response.write("<tr><td>地理</td><td>" & rs("geo") & "</td>")
						response.write("综合模式")
					end if
					response.write("<tr><td>体育</td><td>" & rs("pe") & "</td>")
					response.write("<tr><td>杂项</td><td>" & rs("misc") & "</td>")
					response.write("</table>")
				end if
				response.write("</div><div class=""panel-footer""><a href=""report.html"">反馈登分错误</a></div>")
				rs.close()
			else
				' Teacher managing mode
				response.expires = 5
				sql = "SELECT * FROM passwd WHERE namee = '" & name & "' AND pwd = '" & pwd & "'"
				set conn = CreateObject("ADODB.Connection")
				conn.open "driver={microsoft access driver (*.mdb)};dbq=" & server.MapPath("database\login.mdb")
				set rs  = CreateObject("ADODB.recordset")
				rs.Open sql, conn, 1, 1
				' Now query should completed
				if rs.bof AND rs.eof then
					response.write("登入失败！请检查您的姓名和密码是否正确！")
				else
					response.write("<div class=""panel panel-primary""><div class=""panel-heading"">")
					response.write("敬爱的" & name & "老师，欢迎登入学生成绩管理页面")
					response.write("</div><div class=""panel panel-default"">")
					response.write("<div class=""panel-heading"">上传Excel表(*.xls)</div>")
					response.write("<div class=""panel-body"">")
					response.write("<form name=""upload"" method=""post"" action=""uploader.asp"" enctype=""multipart/form-data"">")
					response.write("<input type=""file"" name=""uploader"">")
					response.write("<button type=""submit"">上传</button></form></div></div>")
					response.write("<hr>")
					response.write("<div class=""panel panel-default""><div class=""panel-heading"">")
					response.write("更改学生成绩</div><div class=""panel-body"">")
					response.write("<form method=""post"" action=""modifier.asp""><span>学号:</span>")
					response.write("<input id=""id"" name=""id""><br><span>")
					response.write("姓名:</span><input id=""name"" name=""name""><br>")
					response.write("<select id=""subject"" name=""subject"">")
					response.write("<option value=""chi"">语文</option>")
					response.write("<option value=""math"">数学</option>")
					response.write("<option value=""lang"">外语</option>")
					response.write("<option value=""soc"">政治</option>")
					response.write("<option value=""hist"">历史</option>")
					response.write("<option value=""geo"">地理</option>")
					response.write("<option value=""phy"">物理</option>")
					response.write("<option value=""chem"">化学</option>")
					response.write("<option value=""bio"">生物</option>")
					response.write("<option value=""pe"">体育</option>")
					response.write("<option value=""misc"">杂项</option>")
					response.write("</select>更改为:<input id=""newScore"" name=""newScore"">")
					response.write("<button onclick=""modify()"">更改</button></form>")
					response.write("<div id=""message""></div>")
					response.write("</div>")
					response.write("</div><div class=""panel-footer"">")
					response.write("<a href=""../example.xls"">下载示范Excel表(example.xls)</a><br>")
					response.write("<a href=""./help.html"">使用帮助</a><br>")
					response.write("<p class=""text-danger"">上传过程中请勿关闭页面。操作不可撤销</p></div>")
				end if
				rs.close()
			end if
		%>
		</div>
	</body>
</html>
