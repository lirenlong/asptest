<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<%

if session("studentname")="" then
  Response.Redirect "index.asp"
end if
%>
<html>
<head>
<title>考试界面----再线考试系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"><style type="text/css">
<!--
body {
	background-image: url(images/backgrand.jpg);
}
-->
</style></head>
<script language="javascript">
function attention()
{
  alert('时间到了,请交卷!');
  document.getElementById('submit').click();
}
setTimeout("attention()",<%=session("testtime")*60*1000%>)
</script>
<body bgcolor="#66CCCC">
<p align="center"><b><font face="宋体" size="5" color="#FF0000"><%=session("selectsubjectname")%>微波技术与天线考试</font></b></p>
<%
if request.form("submit1")="开始考试"  then
%>
<form name="testform" method="post" action="result.asp">
  <table border="0" cellspacing="0"  bordercolor="#111111" width="80%" align="center" >
    <tr>
      <td width="100%" height="25"><b><font size="3" color="#000080">一、单项选择题(每题<%=session("singleper")%>分,共<%=session("singlenumber")%>题)</font></b></td>
    </tr>
  </table>
<%

  dim i,sql,rs,count,temp,strid1,strid2
  strid1=""
  strid2=""
  randomize
 for i=1 to session("singlenumber")
 'for i=1 to CInt(CStr(Request.Cookies("singlenumber")))
    set rs=server.createobject("adodb.recordset")
	  'sql="select * from question where subjectname='"& Request.Cookies("selectsubjectname") & "'and type='单选题' and haveselect=0 "
    sql="select * from question where subjectname='"&session("selectsubjectname") & "'and type='单选题' and haveselect=0 "
    rs.open sql,conn,3,2
    count=rs.recordcount
    temp=fix(count*rnd(10))
    rs.move temp
    rs("haveselect")=1

    strid1=strid1 & rs("ID") & ","

%>
  <table border="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#C0C0C0" width="80%"  cellpadding="0" bgcolor="#FFFFFF" align="center">
    <tr>
      <td width="100%" bgcolor="#EFEFEF" height="20">&nbsp;&nbsp;<b><%=i%>、<%=rs("question")%></b></td>
    </tr>
<%
    if rs("A")<>"" then
%>
    <tr>
      <td width="100%">&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="NO<%=rs("id")%>" value="A">A、<%=rs("A")%></td>
    </tr>
<%
    end if
    if rs("B")<>"" then
%>
    <tr>
      <td width="100%">&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="NO<%=rs("id")%>" value="B">B、<%=rs("B")%></td>
    </tr>
<%
    end if
    if rs("C")<>"" then
%>
    <tr>
      <td width="100%">&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="NO<%=rs("id")%>" value="C">C、<%=rs("C")%></td>
    </tr>
<%
    end if
    if rs("D")<>"" then
%>
    <tr>
      <td width="100%">&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="NO<%=rs("id")%>" value="D">D、<%=rs("D")%></td>
    </tr>
<%
    end if
%>   
  </table>
<%  
    rs.update
	next
 ' rs.close
 ' set rs=nothing
%>

  <table border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber3" align="center">
    <tr>
      <td width="100%" height="25"><b><font color="#000080" size="3">二、多项选择题(每题<%=session("multiper")%>分,共<%=session("multinumber")%>题。每题至少有1项正确的答案，少选、多选、错选均不得分)</font></b></td>
    </tr>
  </table>
  <%
  randomize
  'for i=1 to CInt(Request.Cookies("multinumber"))
  for i=1 to session("multinumber")
    set rs=server.createobject("adodb.recordset")
	  'sql="select * from question where subjectname='" & Request.Cookies("selectsubjectname") & "'and type='多选题' and haveselect=0 "
    sql="select * from question where subjectname='"&session("selectsubjectname") & "'and type='多选题' and haveselect=0 "
    rs.open sql,conn,3,2
    count=rs.recordcount
    temp=fix(count*rnd(10))
    rs.move temp
    rs("haveselect")=1

    strid2=strid2 & rs("ID") & ","
%>
  <table border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="80%" align="center" id="AutoNumber2" bgcolor="#FFFFFF">
    <tr>
      <td width="100%" bgcolor="#EFEFEF" height="20">&nbsp;&nbsp;<b><%=i%>、<%=rs("question")%>一导体回路位于磁场力线垂直的平面内，欲使回路中产生感&nbsp; 应电动势应使</b></td>
    </tr>
<%
    if rs("A")<>"" then
%>
    <tr>
      <td width="100%">&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="NO<%=rs("id")%>" value="A">        
      A
		、<%=rs("A")%></td>
    </tr>
<%
    end if
    if rs("B")<>"" then
%>
    <tr>
      <td width="100%">&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="NO<%=rs("id")%>" value="B">B、<%=rs("B")%></td>
    </tr>
<%
    end if
    if rs("C")<>"" then
%>
    <tr>
      <td width="100%">&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="NO<%=rs("id")%>" value="C">C、<%=rs("C")%></td>
    </tr>
<%
    end if
    if rs("D")<>"" then
%>
    <tr>
      <td width="100%">&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="NO<%=rs("id")%>" value="D">D、<%=rs("D")%></td>
    </tr>
<%
    end if
%> 
  </table>
<% 
    rs.update
  next
  'rs.close

  response.write("<input type='hidden' name='hidQuestID1' value=" & strID1 & ">")
  response.write("<input type='hidden' name='hidQuestID2' value=" & strID2 & ">")

  set rs=nothing
  set rs=server.createobject("adodb.recordset")
  sql="select * from question where haveselect=1 "
  rs.open sql,conn,3,2
  rs.movefirst
  do while  not rs.eof  
    rs("haveselect")=0
    rs.update
    rs.movenext
  loop
  rs.close
  set rs=nothing
  call endConnection()
'response.write(strid1)
'response.write(strid2)
%> 
<p align=center><input type="submit" value="交卷" name="submit" ></p>
</form>
<%
else 
%>
<form method="POST" action="test.asp"  name="form">
<p align=center><input type="submit" value="开始考试" id='submit' name="submit1" ></p>
</form>
<%
  response.write "<center>请选择开始考试！</center>"
end if
%>
</body>
</html>
