<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<%
if session("studentname")="" then
  Response.Redirect "default.asp"
end if

if Request.Form("submit")="确认" then  '如果选择了考试科目，则进入考试界面
  if Request.Form("selectsubject")="" then
	  response.write " <center>你没有选择考试科目，请选择考试科目！</center>"
  else
    dim rs,sql    
	  session("selectsubjectname")=Request.Form("selectsubject")
		set rs = server.createobject("adodb.recordset")
		sql="select * from subject where subjectname='"&session("selectsubjectname")&"'"
		rs.open sql,conn,1,1	
	  '保存单选试题数量
	  'Response.Cookies("singlenumber") = rs("singlenumber")
	  session("singlenumber")=rs("singlenumber")
	  '保存多选试题数量
	  'Response.Cookies("multinumber") = rs("multinumber")
	  session("multinumber")=rs("multinumber")
	  '保存单选试题分值
	  'Response.Cookies("singleper") = rs("singleper")
	  session("singleper")=rs("singleper")
	  '保存多选试题分值
	  'Response.Cookies("multiper") = rs("multiper")
	  session("multiper")=rs("multiper")
	  '保存考试时间
	  'Response.Cookies("testime") = rs("testtime")
  	session("testtime")=rs("testtime")	
	  '保存考试科目名称
	  'Response.Cookies("selectsubjectname") = request.Form("selectsubject")
	  session("selectsubjectname")=request.form("selectsubject")
	  rs.close
		set rs=nothing
		
	 '进入考试界面
	  Response.Redirect "test.asp"
  end if  
end if  

%>
<html>
<head>
<title>考试科目选择-----在线考试系统</title>
</head>
<body bgcolor="#66CCCC">
<table border="0" cellspacing="0" cellpadding="0" width="500" height="156" align="center" border=1 bordercolor=lightgreen>
<tr> 
<td><FONT size=4 color=red face=隶书>
<%Response.Write session("studentname")%></FONT><font face ="华文行楷" size="5" color=blue>,欢迎您进入在线考试系统</font>
</td></tr>
<tr>
<td>
<br>
<form action="selectsubject.asp" method="post" id="form" name="form">
<p align=left><FONT color=green face=宋体 size=4>首先请您选择要考试的科目</FONT> 
<br>
<%
set rs = server.createobject("adodb.recordset")
sql="select * from subject"
rs.open sql,conn,1,1
if err.number<>0 then 
	response.write "数据库操作失败："&err.description
elseif rs.bof and rs.eof then
	response.write "<center>对不起,暂时没有任何考试科目。</center>"
	rs.close		    
else                          
	do while not rs.eof
		Response.Write( "<input name=selectsubject type=radio value=" & rs("subjectname") & ">" & rs("subjectname") & "," & rs("testtime") & "分钟<br>")
		rs.movenext         
	loop   
end if
rs.close
set rs=nothing
call endConnection()
%>
<p align=center ><input  name="submit" type="submit" value="确认"></p>
</form>
</td>
</tr>
</table>
</body>
</html>




