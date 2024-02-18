<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/callStatic/classes/callstaticCls.asp"-->

<%
	Dim vSessionID, vUserID, vDate, vCallDate
	vSessionID 	= session("ssBctId")
	vUserID		= Request("tenUserID")
	vDate		= Request("yyyymmdd")
	vCallDate	= Request("calldate")

	If vUserID = "" Then
		Response.Write "<script>alert('잘못된 접근입니다.');window.close();</script>"
		Response.End
	End IF
	
	If Not (vSessionID = "coolhas" OR C_CSPowerUser OR C_ADMIN_AUTH) Then
		vUserID = vSessionID
	End IF

	Dim cCallList, vWavFile
	Set cCallList = new ClsCall
	cCallList.FUserID = vUserID
	cCallList.FSDate = vDate
	cCallList.FEDate = vCallDate
	cCallList.FCallWavPlay
	vWavFile = cCallList.FOneItem.fwavlink
	set cCallList = nothing
	
	If vWavFile = "x" OR vWavFile = "" Then
		Response.Write "<center>wav 파일이 없거나 잘못된 경로입니다.</center><p><input type='button' value='닫기' class='button' onclick='window.close();'>"
	Else
%>
		<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
		<Embed src="<%=vWavFile%>" Volume="0" ShowPositionControls="1" showstatusbar="1" showaudiocontrols="1" AUTOSTART="0" Width="300" Height="200"></Embed>
		</body>
<%
	End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->