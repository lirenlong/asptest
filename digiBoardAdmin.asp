<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>



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
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_llxxcc_STRING
  MM_editTable = "board"
  MM_editColumn = "digiB_id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "digiBoardAdmin.asp"
  MM_fieldsStr  = "digiB_subject|value|digiB_content|value"
  MM_columnsStr = "digiB_subject|',none,''|digiB_content|',none,''"

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
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
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
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
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
Recordset1.Source = "SELECT * FROM board ORDER BY digiB_id DESC"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim Recordset1_total
Dim Recordset1_first
Dim Recordset1_last

' set the record count
Recordset1_total = Recordset1.RecordCount

' set the number of rows displayed on this page
If (Recordset1_numRows < 0) Then
  Recordset1_numRows = Recordset1_total
Elseif (Recordset1_numRows = 0) Then
  Recordset1_numRows = 1
End If

' set the first and last displayed record
Recordset1_first = 1
Recordset1_last  = Recordset1_first + Recordset1_numRows - 1

' if we have the correct record count, check the other stats
If (Recordset1_total <> -1) Then
  If (Recordset1_first > Recordset1_total) Then
    Recordset1_first = Recordset1_total
  End If
  If (Recordset1_last > Recordset1_total) Then
    Recordset1_last = Recordset1_total
  End If
  If (Recordset1_numRows > Recordset1_total) Then
    Recordset1_numRows = Recordset1_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (Recordset1_total = -1) Then

  ' count the total records by iterating through the recordset
  Recordset1_total=0
  While (Not Recordset1.EOF)
    Recordset1_total = Recordset1_total + 1
    Recordset1.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (Recordset1.CursorType > 0) Then
    Recordset1.MoveFirst
  Else
    Recordset1.Requery
  End If

  ' set the number of rows displayed on this page
  If (Recordset1_numRows < 0 Or Recordset1_numRows > Recordset1_total) Then
    Recordset1_numRows = Recordset1_total
  End If

  ' set the first and last displayed record
  Recordset1_first = 1
  Recordset1_last = Recordset1_first + Recordset1_numRows - 1

  If (Recordset1_first > Recordset1_total) Then
    Recordset1_first = Recordset1_total
  End If
  If (Recordset1_last > Recordset1_total) Then
    Recordset1_last = Recordset1_total
  End If

End If
%>
<%
Dim MM_paramName
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = Recordset1
MM_rsCount   = Recordset1_total
MM_size      = Recordset1_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
Recordset1_first = MM_offset + 1
Recordset1_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (Recordset1_first > MM_rsCount) Then
    Recordset1_first = MM_rsCount
  End If
  If (Recordset1_last > MM_rsCount) Then
    Recordset1_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then
  MM_keepMove = MM_keepMove & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<html>
<head>
<title>数字留言板</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!-- Fireworks MX Dreamweaver MX target.  Created Tue Jan 28 14:10:55 GMT+0800 (￥x￥_?D·CRE?!) 2003-->
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
 var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
   var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
   if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

//-->
</script>
<style type="text/css">
<!--
.lineBox {
	border: 1px dotted #CCCCCC;
}
.Box {
	border: 1px solid #666666;
}
a:hover.ImgButton {
	left: 2px;
	top: 2px;
	right: 0px;
	bottom: 0px;
	position: relative;
}
.backIMG {
	background-image: url('images/backImg3.jpg');
	background-repeat: no-repeat;
	background-position: left top;
	background-attachment: fixed
}
-->
</style>
<style type="text/css">
<!--
form {
	margin: 0px;
}
-->
</style>
</head>
<body bgcolor="#ffffff" onLoad="MM_preloadImages('images/digiBoard_r2_c2_f2.gif','images/digiBoard_r2_c4_f2.gif','images/digiBoard_r2_c3_f2.gif')">
<table width="700" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <!-- fwtable fwsrc="easyBoard.png" fwbase="digiBoard.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="1" -->
  <tr>
    <td height="91"><img name="digiBoard_r1_c1" src="images/digiBoard_r1_c1.gif" width="700" height="91" border="0" alt=""></td>
  </tr>
  <tr>
    <td height="23"> <table border="0" cellpadding="0" cellspacing="0" width="700">
        <tr>
          <td><img name="digiBoard_r2_c1" src="images/digiBoard_r2_c1.gif" width="367" height="23" border="0" alt=""></td>
          <td><a href="digiBoardPost.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('digiBoard_r2_c2','','images/digiBoard_r2_c2_f2.gif',1);"><img name="digiBoard_r2_c2" src="images/digiBoard_r2_c2.gif" width="108" height="23" border="0" alt=""></a></td>
          <td><a href="1.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('digiBoard_r2_c3','','images/digiBoard_r2_c3_f2.gif',1);"><img name="digiBoard_r2_c3" src="images/digiBoard_r2_c3.gif" width="108" height="23" border="0" alt=""></a></td>
          <td><a href="digiBoardLogin.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('digiBoard_r2_c4','','images/digiBoard_r2_c4_f2.gif',1);"><img name="digiBoard_r2_c4" src="images/digiBoard_r2_c4.gif" width="108" height="23" border="0" alt=""></a></td>
          <td><img name="digiBoard_r2_c5" src="images/digiBoard_r2_c5.gif" width="9" height="23" border="0" alt=""></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="36"><img name="digiBoard_r3_c1" src="images/digiBoard_r3_c1.gif" width="700" height="36" border="0" alt=""></td>
  </tr>
  <tr>
    <td valign="top" background="images/digiBoard_r4_c1.gif"> <% If Not Recordset1.EOF Or Not Recordset1.BOF Then %>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="14">　</td>
          <td> <%
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF))
%>
            <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1" id="form1">

              <table width="96%" border="0" align="center" cellpadding="4" cellspacing="0" bgcolor="#FFFFFF" class="Box">

              <tr valign="top">
                  <td width="120" align="center" class="backIMG"><font color="#000066" size="-1">&nbsp;<img src="images/face/<%=(Recordset1.Fields.Item("digiB_face").Value)%>">
                    <br>
                    <strong><%=(Recordset1.Fields.Item("digiB_name").Value)%></strong> <br>
                    </font><br>
                    <font size="2">[<A HREF="digiBoardDel.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "digiB_id=" & Recordset1.Fields.Item("digiB_id").Value %>">删除</A>]</font>
                  </td>
                <td> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td>
                        <input name="digiB_subject" type="text" id="digiB_subject" value="<%=(Recordset1.Fields.Item("digiB_subject").Value)%>" size="20"></td>
                        <td align="right"><strong><font color="#999999" size="-1"><%=(Recordset1.Fields.Item("digiB_potime").Value)%></font></strong></td>
                    </tr>
                  </table>
                  <hr size="1" noshade class="lineBox"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td valign="top" style="word-break:break-all"><font color="#666666" size="-1">
                          <textarea name="digiB_content" cols="50" rows="5" id="digiB_content"><%=(Recordset1.Fields.Item("digiB_content").Value)%></textarea>
                          <wbr></font></td>
                    </tr>
                  </table>
                  <hr size="1" noshade>
                    <input type="submit" name="Submit" value="更新">
                    <input type="reset" name="Submit2" value="重设">
                  </td>
              </tr>
            </table>

              <input type="hidden" name="MM_update" value="form1">
              <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("digiB_id").Value %>">
            </form>			
            <hr width="96%" size="1" noshade class="lineBox">
            <%
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%> <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#E5E5E5">
              <tr>
                <td width="50%"><font size="-1">&nbsp; 记录 <%=(Recordset1_first)%> 到 <%=(Recordset1_last)%> 共 <%=(Recordset1_total)%> </font></td>
                <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td align="right"> <font size="-1">
                        <% If MM_offset <> 0 Then %>
                        <a href="<%=MM_moveFirst%>">第一页</a>
                        <% End If ' end MM_offset <> 0 %>
                        </font> <font size="-1">
                        <% If MM_offset <> 0 Then %>
                        <a href="<%=MM_movePrev%>">上一页</a>
                        <% End If ' end MM_offset <> 0 %>
                        </font> <font size="-1">
                        <% If Not MM_atTotal Then %>
                        <a href="<%=MM_moveNext%>">下一页</a>
                        <% End If ' end Not MM_atTotal %>
                        </font> <font size="-1">
                        <% If Not MM_atTotal Then %>
                        <a href="<%=MM_moveLast%>">最后一页</a>
                        <% End If ' end Not MM_atTotal %>
                        </font></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="10">　</td>
        </tr>
      </table>
      <% End If ' end Not Recordset1.EOF Or NOT Recordset1.BOF %> <% If Recordset1.EOF And Recordset1.BOF Then %>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="14">　</td>
          <td align="center"><font size="2">目前数据库中并没有任何数据！ </font></td>
          <td width="10">　</td>
        </tr>
      </table>
      <% End If ' end Recordset1.EOF And Recordset1.BOF %> </td>
  </tr>
  <tr>
    <td height="21"><img name="digiBoard_r5_c1" src="images/digiBoard_r5_c1.gif" width="700" height="21" border="0" alt=""></td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>