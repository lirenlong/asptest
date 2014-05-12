<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/llxxcc.asp" -->
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form") Then

  MM_editConnection = MM_llxxcc_STRING
  MM_editTable = "denglu"
  MM_editRedirectUrl = "personal.asp"
  MM_fieldsStr  = "denglu_n_f|value|denglu_mi_f|value|denglu_tr_n_f|value|denglu_mail_f|value|denglu_tel_f|value"
  MM_columnsStr = "denglu_name|',none,''|denglu_mina|',none,''|denglu_true_name|',none,''|denglu_mail|',none,''|denglu_tel|none,none,NULL"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_llxxcc_STRING
Recordset1.Source = "SELECT * FROM denglu"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>用户注册页面</title>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
.STYLE2 {
	font-family: "宋体";
	font-size: x-large;
}
body {
	background-image: url(images/backgrand.jpg);
}
-->
</style>
</head>
<body>
<form id="form" name="form" method="POST" action="<%=MM_editAction%>">
  <p align="center" class="STYLE2">用户注册表</p>
  <table width="890" border="1" align="center">
    <tr>
      <td width="276"><div align="center">用户名：</div></td>
      <td width="598"><label>
        <input name="denglu_n_f" type="text" id="denglu_n_f" size="25" />
        <span class="STYLE1"> *</span>（此项为必填选项）</label></td>
    </tr>
    <tr>
      <td><div align="center">密码：</div></td>
      <td><label>
        <input name="denglu_mi_f" type="password" id="denglu_mi_f" size="25" />
        <span class="STYLE1">*</span>（为保证密码安全，请将密码设为6位以上）</label></td>
    </tr>
    <tr>
      <td><div align="center">用户真实姓名：</div></td>
      <td><label>
        <input name="denglu_tr_n_f" type="text" id="denglu_tr_n_f" size="25" />
        <span class="STYLE1">*</span>(此项为必填选项，请输入真实姓名)</label></td>
    </tr>
    <tr>
      <td><div align="center">E-MAIL:</div></td>
      <td><label>
        <input name="denglu_mail_f" type="text" id="denglu_mail_f" size="25" />
      </label></td>
    </tr>
    <tr>
      <td><div align="center">用户电话：</div></td>
      <td><label>
        <input name="denglu_tel_f" type="text" id="denglu_tel_f" size="25" />
        <span class="STYLE1">*</span>(为方便我们联系您，请务必正确填写)</label></td>
    </tr>
    <tr>
      <td><div align="center">操作：</div></td>
      <td><label>
        <input type="submit" name="Submit" value="提交" />
     &nbsp;&nbsp;&nbsp; 
     <input type="submit" name="Submit2" value="重填" />
      </label></td>
    </tr>
  </table>

  <input type="hidden" name="MM_insert" value="form">
</form>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
