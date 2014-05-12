<!--#include file="conn.asp"-->
<html>
<head>
<title>考试界面</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#FFFFFF">
<%
subjectname=session("selectsubjectname")
studentname=session("studentname")
singlenumber=session("singlenumber")
singleper=session("singleper")
multinumber=session("multinumber")
multiper=session("multiper")
endtime=now()
score=0
selectstr1=request.form("hidQuestID1")
selectstr2=request.form("hidQuestID2")
len1=len(selectstr1)
len2=len(selectstr2)
str1=left(selectstr1,len1-1)
str2=left(selectstr2,len2-1)
dim id1,id2
id1=split(str1,",")
id2=split(str2,",")

for i=1 to singlenumber
 result=request.form("no"&id1(i-1))
 if  not isempty(result) then
      sql="select * from question where id="& clng(id1(i-1))
      set rs=server.createobject("adodb.recordset")
      rs.open sql,conn,3,2     
        if result=rs("answer") then
          score=score+cint(singleper)
        end if
        rs.close
        set rs=nothing 
  else
  end if
next
        
for i=1 to multinumber
 result=request.form("no"&id2(i-1))
  if  not isempty(result) then
      sql="select * from question where id="& clng(id2(i-1))
      set rs=server.createobject("adodb.recordset")
      rs.open sql,conn,3,2     
        if result=rs("answer") then
          score=score+cint(multiper)
        end if   
        rs.close
        set rs=nothing 
  else
  end if
next 
sql="select * from score"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,3,2
rs.addnew
rs("studentname")=studentname
rs("subjectname")=subjectname
rs("endtime")=endtime
rs("score")=score
rs.update
rs.close
set rs=nothing
call endConnection()
total=singlenumber*singleper+multinumber*multiper
response.write("<center>"&studentname&"你好！你的考试成绩为："&score&"分，总分为"&total&"分</center><br>")
%>
<p align=center><a href="lo.asp"><font color=red size=+0 face=楷体>返回登录界面</font></a></p>
<p align=center><a href="selectsubject.asp"><font color=red size=+0 face=楷体>返回考试界面继续考试</font></a></p>
</body>
</html>

