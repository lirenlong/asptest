<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<% 
dim isedit '�Ƿ��ڱ༭״̬
dim color  '�����ɫ
dim id,studentname'����������û���id
dim sql,rs,rsc
color=1
isedit=false
if request("action")="edit" then
    isedit=true
end if
if request("action")="edit" then   '�޸��û�
    if trim(request("studentpassword"))="" then
	    response.write "����!���벻��Ϊ��! <a href=mgstudent.asp>����</a>"
        response.end
    end if
	sql="update student set studentname='" & cstr(trim(request("studentname"))) & "',studentpassword='" & cstr(trim(request("studentpassword")))
	conn.execute sql
	if err.number <> 0 then
	    response.write "���ݿ��������:" + err.description
	else %>
	    <script language=vbscript>
			msgbox "�����ɹ�!�û� <%=trim(request("studentname"))%> ����Ϣ�Ѿ�����!"
		</script>
  <%end if
end if
if request("action")="add" then   '������û�
    if trim(request("studentname"))="" or trim(request("studentpassword"))="" then
	    response.write "����!�û��������벻��Ϊ��! <a href=# onclick='javascript:window.history.go(-1)'>����</a>"
        response.end
    end if
	set rs=server.createobject("adodb.recordset")   '���ѧ���Ƿ�����
    rs.open "select * from student where studentname='" & cstr(trim(request("studentname"))) & "'",conn,1,1
    if err.number <> 0 then
	          response.write "���ݿ����"
    else  if not rs.bof and not rs.eof then
	          response.write "����!����ѧ������! <a href=# onclick='javascript:window.history.go(-1)'>����</a>"
              response.end
          end if
    end if
	rs.close
	set rs=nothing
	sql="insert into student(studentname,studentpassword) values('" & cstr(trim(request("studentname"))) & "','" & cstr(trim(request("studentpassword"))) & "')"
	conn.execute sql
	if err.number <> 0 then
	    response.write "���ݿ��������:" + err.description
	else %>
	    <script language=vbscript>
			msgbox "�����ɹ�!���û� <%=trim(request("studentname"))%> ����Ϣ��ӳɹ�!"
		</script>
  <%end if
end if
if request("action")="del" then   'ɾ���û�
	sql="delete from student where id=" &request("id")
	conn.execute sql
	if err.number <> 0 then
		response.write "���ݿ��������" + err.description
		err.clear
	else %>
        <script language=vbscript>
		msgbox "�����ɹ�!�û� <%=trim(request("studentname"))%> ����Ϣ��ɾ��!"
		</script>
<%  end if
end if
%>
<html>
<head>
<title>����ѧ��----���߿���ϵͳ</title>
<script language=javascript>
function SureDel(id)
{
    if ( confirm("��ȷ��Ҫɾ�����û���"))
        {
            window.location.href = "mgstudent.asp?action=del&id=" +id
        }
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body background=file:///C|/Documents and Settings/Administrator/����/examlxc/images/88.jpg > 
<center><font color="lime" size=+3>ѧ������</font></center>
<table width=528 border="1" cellspacing="0" cellpadding="0" align="center" bordercolor=blue>
  <tr> 
    <td width="25%"> 
      <div align="center">ѧ������</div>
    </td>
    <td width="20%"> 
      <div align="center">����</div>
    </td>
    <td width="20%"> 
      <div align="center">����</div>
    </td>
  </tr>
  <%
  set rs=server.createobject("adodb.recordset")
  rs.open "select * from student ",conn,1,1
  if err.number <> 0 then
	           response.write "���ݿ����"
           else
	           if rs.bof and rs.eof then
		           rs.close
		           response.write "Ŀǰû��ѧ��"
	           else
			       do while not rs.eof %>
  <tr> 
    <td width="25%" height="78" > 
      <div align="center"><%=rs("studentname")%></div>
    </td>
    <td width="20%" height="78" > 
      <div align="center"><%=rs("studentpassword")%></div>
    </td>
    <td width="20%" height="78" >
      <div align="center">
        <% 					 
	      response.write "<a href='javascript:SureDel(" & cstr(rs("id")) & ")'>ɾ��</a>"		     
	 %>
      </div>
    </td>
  </tr>
  <% rs.movenext
	color=color+1				   
	loop			   
	end if	       
       end if 
	'rs.close
	'set rs=nothing %>
</table>
 <p align="center"> 
    <%  response.write "<font size=3>�� �� �� �� ѧ ��</font><br>" %> 
<form action="mgstudent.asp" method="post">
	    <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else  %>add<% End If %>'>
		<%If isedit then%>
              <input type="Hidden" name="studentname" value='<%=cstr(request("studentname"))%>'>
        <%End If%>
	    �û�����:<input type="text" name="studentname" class=input maxlength=14 size="16"><br>
	    �û�����:<input type="password" name="studentpassword"  class=input maxlength=12 size="16"><br>
	            <input type="submit" name="submit" value="ȷ ��" class=button>
</form>
     <p align=center><a href="primarypage.asp"><font color=red size=+0 face=����>���ع������</font></a></p>
 </p>
</body>
</html>