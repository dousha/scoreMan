<%
	dim id = request.form("id");
	dim name = request.form("name");
	dim passwd = request.form("password");
	set conn = server.CreateObject("adodb.connection")
	conn.open "driver={microsoft access driver (*.mdb)};dbq=" & server.MapPath("database/login.mdb")
	exec = "select * from passwd where id = '" + id + "'"
	set rs = server.createobject("adodb.recordset")
	rs.open exec, conn, 1, 1
	if passwd = nothing then
		' student mode
		rs = nothing
		conn = nothing
		conn = server.createobject("adodb.connection")
		conn.open "driver={microsoft access driver (*.mdb)};dbq=" & server.MapPath("database/unidata.mdb")
		exec = "select * form unidata where id='" + id + "' and namee = '" + name + "'"
		set rs = server.createobject("adodb.recordset")
		rs.open exec, conn, 1, 1
		'' FIXME: sort has not completed yet
		response.write rs("namee") + ":" + rs("class") + ":" + rs("ID") + ":" + rs("total") + ":" + rs("chi") + ":" + rs("math") + ":" + rs("lang") + ":" + rs("phy") + ":" + rs("chem") + ":" + rs("soc") + ":" + rs("hist") + ":" + rs("geo") + ":" + rs("bio") + ":" + rs("pe") + ":" + rs("misc")
	else
		if rs("password") = passwd then
			response.write "isAdmin"
		else
			response.write "bad login"
		end if
	end if
%>