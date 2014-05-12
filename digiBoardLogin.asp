<%@LANGUAGE="VBSCRIPT"%>

<!--#include file="Connections/llxxcc.asp" -->
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_llxxcc_STRING
Recordset1.Source = "SELECT * FROM admin"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Request.QueryString
MM_valUsername=CStr(Request.Form("username"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization=""
  MM_redirectLoginSuccess="digiBoardAdmin.asp"
  MM_redirectLoginFailed="digiBoardLogin.asp"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_llxxcc_STRING
  MM_rsUser.Source = "SELECT username, passwd"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM admin WHERE username='" & Replace(MM_valUsername,"'","''") &"' AND passwd='" & Replace(Request.Form("passwd"),"'","''") & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<html>
<head>
<title>管理员登录界面</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.bluebox {
	border: 1px solid #0066CC;
}
-->
</style>
</head>

<body>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td>
<form ACTION="<%=MM_LoginAction%>" name="form1" method="POST">
        <table width="300" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="5"><img src="images/winxp_r1_c1.jpg" width="5" height="30"></td>
            <td background="images/winxp_r1_c3.jpg"><font color="#FFFFFF" size="2"><img src="images/winxp_r1_c9.jpg" width="18" height="30" align="absmiddle">管理员登录界面</font></td>
            <td width="28"><a href="1.asp"><img src="images/winxp_r1_c5.jpg" alt="离开登录画面" width="28" height="30" border="0"></a></td>
          </tr>
        </table>
        <table width="300" border="0" align="center" cellpadding="4" cellspacing="2" class="bluebox">
          <tr>
            <td width="60"> <div align="right"><font size="2">账号</font></div></td>
            <td><font size="2">
              <input name="username" type="text" id="username" size="20">
              </font></td>
          </tr>
          <tr>
            <td width="60"> <div align="right"><font size="2">密码</font></div></td>
            <td><font size="2">
              <input name="passwd" type="password" id="passwd" size="20">
              </font></td>
          </tr>
          <tr>
            <td colspan="2"> <div align="center"><font size="2"></font> <font size="2">
                <input type="submit" name="Submit" value="登录管理界面">
                <input type="button" name="Submit2" value="离开登录界面" onClick="window.location.href='digiBoard.asp';">
                </font></div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
