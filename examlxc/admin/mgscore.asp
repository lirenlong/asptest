<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<% 
dim id'����������û���id
dim sql,rs,rsc
    if request("action")="del" then   'ɾ����¼
	sql="delete from score where id=" &request("id")
	conn.execute sql
	if err.number <> 0 then
		response.write "���ݿ��������" + err.description
		err.clear
	else %>
        <script language=vbscript>
		msgbox "�����ɹ�!����Ϊ<%=trim(request("id"))%>�Ŀ��Լ�¼��ɾ��!"
		</script>
<%  end if
end if
%>
<html>
<head>
<title>�����Գɼ�----���߿���ϵͳ</title>
<script language=javascript>
function SureDel(id)
{
    if ( confirm("��ȷ��Ҫɾ���ÿ��Լ�¼��"))
        {
            window.location.href = "mgscore.asp?action=del&id=" +id
        }
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body background=file:///C|/Documents and Settings/Administrator/����/examlxc/images/88.jpg > 
<center><font color="lime" size=+3>���Գɼ�����</font></center>
<table width=624 border="1" cellspacing="0" cellpadding="0" align="center" bordercolor=lightgreen>
  <tr> 
    <td width="20%"> 
      <div align="center">ѧ������</div>
    </td>
    <td width="20%"> 
      <div align="center">���Կ�Ŀ</div>
    </td>
    <td width="20%"> 
      <div align="center">����ʱ��</div>
    </td>
    <td width="20%"> 
      <div align="center">���Է���</div>
    </td>
    <td width="20%"> 
      <div align="center">����</div>
    </td>
  </tr>
  <%
  set rs=server.createobject("adodb.recordset")
  rs.open "select * from score order by studentname ",conn,1,1
  if err.number <> 0 then
	           response.write "���ݿ����"
           else
	           if rs.bof and rs.eof then
		           rs.close
		           response.write "Ŀǰû�п��Լ�¼"
	           else
			       do while not rs.eof %>
  <tr> 
    <td width="20%" height="126" > 
      <div align="center"><%=rs("studentname")%></div>
    </td>
    <td width="20%" height="126" > 
      <div align="center"><%=rs("subjectname")%></div>
    </td>
    <td width="20%" height="126" > 
      <div align="center"><%=rs("endtime")%></div>
    </td>
    <td width="20%" height="126" > 
      <div align="center"><%=rs("score")%></div>
    </td>
    <td width="20%" height="126" >
      <div align="center">
        <% 					 
	      response.write "<a href='javascript:SureDel(" & cstr(rs("id")) & ")'>ɾ��</a>"		     
	 %>
      </div>
    </td>
  </tr>
  <% rs.movenext				   
	loop			   
	end if	       
       end if 
	'rs.close
	'set rs=nothing %>
</table>
 <p align=center><a href="primarypage.asp"><font color=red size=+0 face=����>���ع������</font></a></p>
</body>
</html>