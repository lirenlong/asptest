<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<% 
dim isedit '�Ƿ��ڱ༭״̬
dim color  '�����ɫ
dim id,name'����������û���id
dim sql,rs,rsc
color=1
isedit=false
if request("action")="edit" then
    isedit=true
end if
if request("action")="edit" then   '�޸Ĺ���Ա
    if trim(request("password"))="" then
	    response.write "����!���벻��Ϊ��! <a href=mgadmin.asp>����</a>"
        response.end
    end if
	sql="update admin set name='" & cstr(trim(request("name"))) & "',password='" & cstr(trim(request("password")))
	conn.execute sql
	if err.number <> 0 then
	    response.write "���ݿ��������:" + err.description
	else %>
	    <script language=vbscript>
			msgbox "�����ɹ�!����Ա<%=trim(request("name"))%>����Ϣ�Ѿ�����!"
		</script>
  <%end if
end if
if request("action")="add" then   '����¹���Ա
    if trim(request("name"))="" or trim(request("password"))="" then
	    response.write "����!�û��������벻��Ϊ��! <a href=# onclick='javascript:window.history.go(-1)'>����</a>"
        response.end
    end if
	set rs=server.createobject("adodb.recordset")   '���ѧ���Ƿ�����
    rs.open "select * from admin where name='" & cstr(trim(request("name"))) & "'",conn,1,1
    if err.number <> 0 then
	          response.write "���ݿ����"
    else  if not rs.bof and not rs.eof then
	          response.write "����!�ù���Ա����! <a href=# onclick='javascript:window.history.go(-1)'>����</a>"
              response.end
          end if
    end if
	rs.close
	set rs=nothing
	sql="insert into admin(name,password) values('" & cstr(trim(request("name"))) & "','" & cstr(trim(request("password"))) & "')"
	conn.execute sql
	if err.number <> 0 then
	    response.write "���ݿ��������:" + err.description
	else %>
	    <script language=vbscript>
			msgbox "�����ɹ�!�¹���Ա<%=trim(request("name"))%>����Ϣ��ӳɹ�!"
		</script>
  <%end if
end if
if request("action")="del" then   'ɾ������Ա
	sql="delete from admin where id=" &request("id")
	conn.execute sql
	if err.number <> 0 then
		response.write "���ݿ��������" + err.description
		err.clear
	else %>
        <script language=vbscript>
		msgbox "�����ɹ�!����Ա<%=trim(request("name"))%> ����Ϣ��ɾ��!"
		</script>
<%  end if
end if
%>
<html>
<head>
<title>�������Ա----���߿���ϵͳ</title>
<script language=javascript>
function SureDel(id)
{
    if ( confirm("��ȷ��Ҫɾ���ù���Ա��"))
        {
            window.location.href = "mgadmin.asp?action=del&id=" +id
        }
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body background=file:///C|/Documents and Settings/Administrator/����/examlxc/images/88.jpg > 
<center>
  <font color="lime" size=+3>��ʦ����</font>
</center>
<table width=442 height="86" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor=lightgreen>
  <tr> 
    <td width="25%"> 
      <div align="center">����Ա����</div>
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
  rs.open "select * from admin ",conn,1,1
  if err.number <> 0 then
	           response.write "���ݿ����"
           else
	           if rs.bof and rs.eof then
		           rs.close
		           response.write "Ŀǰû�й���Ա"
	           else
			       do while not rs.eof %>
  <tr> 
    <td width="25%" height="21" > 
      <div align="center"><%=rs("name")%></div>
    </td>
    <td width="20%" height="21" > 
      <div align="center"><%=rs("password")%></div>
    </td>
    <td width="20%" height="21" >
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
    <%  response.write "<font size=3>����µĹ���Ա</font><br>" %> 
<form action="mgadmin.asp" method="post">
	    <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else  %>add<% End If %>'>
		<%If isedit then%>
              <input type="Hidden" name="name" value='<%=cstr(request("name"))%>'>
        <%End If%>
	    �û�����:<input type="text" name="name" class=input maxlength=14 size="16"><br>
	    �û�����:
	    <input type="password" name="password"  class=input maxlength=12 size="16"><br>
	            <input type=submit value="ȷ ��" class=button>
                <p align=center><a href="primarypage.asp"><font color=red size=+0 face=����></font></a></p>
                <div align="center"><a href="primarypage.asp"><font color=red size=+0 face=����>���ع������</font></a>
                </div>
</form>
     </p>
</body>
</html>