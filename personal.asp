<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/llxxcc.asp" -->
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
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername=CStr(Request.Form("denglu_name_f"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization=""
  MM_redirectLoginSuccess="doexercises.html"
  MM_redirectLoginFailed="personal.asp"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_llxxcc_STRING
  MM_rsUser.Source = "SELECT denglu_name, denglu_mina"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM denglu WHERE denglu_name='" & Replace(MM_valUsername,"'","''") &"' AND denglu_mina='" & Replace(Request.Form("denglu_mina_f"),"'","''") & "'"
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ѧ��Դ</title>
<style type="text/css">
<!--
.STYLE2 {color: #CCCCCC}
body {
	background-image: url(images/backgrand.jpg);
	background-repeat: repeat;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.STYLE7 {
	font-family: "����";
	font-weight: bold;
}
-->
</style>
</head>

<body>
<table width="780%" border="0">
  <tr>
    <td width="335" height="1729" valign="top"><p><img src="images/biaoti.jpg" width="300" height="80" />
    </p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <h1 align="center">&nbsp;</h1>
      <h1 align="center"><span class="STYLE7"><a href="jxdg.doc">��ѧ���</a></span></h1>
      <h1 align="center" class="STYLE7"><a href="jxrl.doc">��ѧ����</a></h1>
      <h1 align="center"><a href="ja.zip">�̰�</a></h1>
      <h1 align="center"><strong><a href="syzds.doc">ʵ��ָ����</a></strong></h1>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center"><a href="default.asp">������ҳ</a> </p>
      <p>&nbsp;</p>
    <td width="537" height="1729" valign="top"><h1 align="center"><a name="_Toc176143865" id="_Toc176143865"><strong>΢�����������߿γ̽�ѧ���</strong></a> <br />
      </h1>
      <p>�γ̱��룺11211361 ѧʱ��32 ѧ�֣�2<br />
        ����רҵ��ͨ�Ź��� <br />
      </p>
      <p>һ���γ̵����ʺ����� <br />
        ΢�����������߿γ���ͨ�Ź���רҵ��רҵ������޿γ̣�����Ρ� <br />
        ����΢����Ƶ�ʺܸߣ������̣ܶ����ʹ��΢�������͵�·���������е����ԡ�ͨ�����γ�ѧϰ��ʹѧ����������΢�������и���΢�������ṹ�������Ĺ���ԭ�����Ӧ�ķ����ֶΡ�ʹѧ�������մ��������ۣ�΢���������ۣ�΢�������Ĺ���ԭ�����Ӧ�ķ����ֶε�ͬʱ������������Ҳ��һ�����˽⡣ <br />
        �����γ̽�ѧ���ݵĻ���Ҫ���ص���ѵ㼰ѧʱ���� <br />
        1.���۽�ѧ 24 ѧʱ <br />
        1.1���������� ��10ѧʱ�� <br />
        1.1.1�������������� <br />
        1.1.2�޺Ĵ����߷��̼���� <br />
        1.1.3�޺Ĵ����ߵĻ������� <br />
        1.1.4�����޺Ĵ����߹���״̬�ķ��� <br />
        1.1.5�кĴ����� <br />
        1.1.6�迹Բͼ����Ӧ�� <br />
        1.1.7�������迹ƥ�� <br />
        �������գ�΢���ķ�Χ���ص㣻���ߵĸ�������ߵķֲ������������߷ֲ�������·���г������߷��̣������߷��̵Ľ⣻�����ߵ������迹�����������������ߵ������迹�������߷���ϵ���������ߵ��迹ƥ�䣻��һ���迹�ĸ��Բͼ�Ĺ��ɼ�����ԭ�����ʹ��Բͼ�� <br />
        ���գ������ߵĸ�������޺Ĵ����ߵĹ���״̬�����е��в�״̬�������в�״̬�µĵ�ѹ�����������߷ֲ����в�״̬�µ������迹�ֲ��������޺Ĵ����ߵ�פ��״̬������פ��״̬�µ����ߵ�ѹ�������ֲ��������迹�ֲ��� <br />
        �ص㣺�����޺Ĵ����ߵĻ����ص�͹���״̬������ʹ���迹Բͼ�� <br />
        �ѵ㣺�����迹������ϵ����פ���ȡ� <br />
        1.2���򲨵��Ϳ�ǻг���� ��4ѧʱ�� <br />
        �������գ�г��ǻ�еĵ��������ϵ��Ʒ�����صĸ������г��ǻ��г�񲨳� <br />
        �ص㣺Ʒ�����صĸ������г��ǻ��г�񲨳� <br />
        1.3΢��������� ��4ѧʱ�� <br />
        1.3.1΢������������� <br />
        1.3.2������������˫�ߴ����ߵĵ�Ч <br />
        1.3.3΢��Ԫ����ЧΪ΢�������ԭ�� <br />
        1.3.4���˿�΢������ <br />
        �������գ� ����·���ַ���������������ЧΪ˫�ߵ�ԭ�򣻲����ĵ�Ч�����迹����һ���迹����һ����ѹ�͹�һ�����������˿�΢������������迹���������˿�΢�����������ɢ����������˿�΢�����������ɢ��������迹����֮��ı任�� <br />
        ���գ����˿�΢����������ĵ��ɲ��������˿�΢�����������ת�Ʋ�����ƥ��Ԫ��������Ԫ������·Ԫ���� <br />
        �ص㣺���˿�΢�����������ɢ������� <br />
        1.4���� (6ѧʱ)<br />
        �������գ��������䵥Ԫ���������ߵ���Ҫ������ <br />
        ���գ��������ߣ��Գ��������ߡ� <br />
        �ص㣺�Գ����ӵĽṹ�����䳡�����ߵ����Բ����� <br />
        2.ʵ�� 8ѧʱ <br />
        2.1΢������ϵͳ�ĵ��ԣ�2ѧʱ�� <br />
        2.2����������פ���ȵĲ�����2ѧʱ�� <br />
        2.3�迹������2ѧʱ�� <br />
        2.4�迹ƥ�䣨2ѧʱ�� <br />
        �����ο��̲ĺ���Ҫ�ο����� <br />
        �ο��̲ģ���е��ҵ�����硶΢�����������ߡ����������ı� <br />
        �ο����ף� </p>
      <ol>
        <ol>
          <ol>
            <li>�����ʵ��ѧ�����硶΢������������Ӧ�á�������������� </li>
            <li>���ӹ�ҵ�����硶��ų���΢�����̻�����������ë���ܣ� </li>
            <li>�������ӿƼ���ѧ�����硶��ų�΢�����������ߡ�������ʢ�� �� </li>
          </ol>
        </ol>
      </ol>
      <p>�ġ�������ʽ�Ϳ���Ҫ�� <br />
        ������ԣ���������Ϊ100�� ��ռ��ĩ�ܳɼ���70��90%������Ϊƽʱ�ɼ���ƽʱ�ɼ����ݲ��顢ʵ�顢������������˸�������85%���ϡ� <br />
        ����Ҫ�� </p>
      <ol>
        <li>���մ����ߵ������迹������������ </li>
        <li>�����޺Ĵ����ߵĻ����ص㣬����״̬�������迹����ϵ���ȴ�����������⣻ </li>
        <li>���մ����ߵ��迹ƥ��ĸ���޷���ƥ��ĸ��Ӧ�ã��ķ�֮һ����ƥ������ </li>
        <li>���յ��е�Ų��ķ��ࣻ���е�Ų��Ĵ��������� </li>
        <li>���վ��β�����Բ�����ĳ��ṹ����ģ�� </li>
        <li>�������˿�΢������������迹������ɢ������� </li>
        <li>���ջ������䵥Ԫ�����ߵ����Բ����� </li>
      </ol>
      <p>�塢�й�˵�� <br />
        ִ���ˣ���· <br />
        ����ˣ���� <br />
        ��׼�ˣ������� <br />
        �������ſγ���Ҫ�������Ӣ�Ķ��ձ� <br />
        ΢�� Microwave<br />
        ��ų� Electromagnetic field<br />
        ���Ų� Transverse  electro-magnetic wave<br />
        ���� Antenna<br />
        ��ѹ Voltage<br />
        ���� Electric current<br />
        Ƶ�� Frequency<br />
        ��· Circuit<br />
        ��· Open circuit<br />
        ��· Short circuit<br />
        ˥���� Attenuator<br />
        ��� Amplitude<br />
        ɢ�� Divergence<br />
        ���� Rotation or curl<br />
        �ݶ� Gradient<br />
        Ⱥ�� Group velocity<br />
        ���� Waveguide<br />
        �迹 Impedance<br />
        ���� Admittance<br />
        ��һ�� Normalization<br />
        г��ǻ Resonant cavity<br />
        פ���� Standing wave ratio<br />
        ���� Radiation<br />
        ��ֱ Vertical separation<br />
        �������� Propagation characteristic<br />
        ��ģ Dominant mode<br />
        ���� Wave length<br />
        ������� Longitudinal component<br />
        ������� Transverse component<br />
        ��λ�ͺ� Lagging phase<br />
        �߽����� Boundary condition<br />
        ƥ�� Matching<br />
        ���� Modulate<br />
        ȫ���� Total reflection<br />
        ����ϵ�� Reflectance<br />
        ��� Demodulation<br />
        ������ Transmission line<br />
        �Ŵ��� Amplifier<br />
        ��Ƶ�� Mixer<br />
        ΢�����ɵ�· Microwave Integrated Circuit (MIC)<br />
        ���� Oscillator<br />
        ��� Coupling<br />
        Ʒ������ Goodness<br />
        Ƶ�ָ��� Frequency-division multiplexing<br />
        ����� Signal-to-noise ratio</p>
      <p>&nbsp;</p>
      <p><a href="default.asp">������ҳ</a></p>
    <br clear="all" /></td>
    <td width="5073"><span class="STYLE2"></span></td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
