<%@ language=vbscript %>
<% option Explicit %>
<% Response.CharSet = "euc-kr" %>
<%
'####################################################
' Description :  [2016 S/S ����] Wedding Membership ����ó�� ������
' History : 2016.09.12 ���¿�
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode, evt_code, sub_idx, sqlStr, userid
dim wdcp1code, wdcp2code, wdcp3code, wdcp4code, wdcp5code
	mode = requestcheckvar(request("mode"),32)
	userid = requestcheckvar(request("userid"),32)
	evt_code = getNumeric(requestcheckvar(request("evt_code"),32))
	sub_idx = getNumeric(requestcheckvar(request("sub_idx"),10))

	If session("ssBctId") ="greenteenz" Or session("ssBctId") = "djjung" Then
	Else
		response.write "<script>alert('�����ڸ� �� �� �ִ� ������ �Դϴ�.');window.close();</script>"
		response.End
	End If
	
	dim refer
		refer = request.ServerVariables("HTTP_REFERER")
	if InStr(refer,"10x10.co.kr")<1 then
		Response.Write "<script type='text/javascript'>alert('�߸��� �����Դϴ�.');</script>"
		dbget.close() : Response.End
	end If

''mkt �����ڰ� ���ο��� ����
If mode = "gubunok" Then
	if evt_code="" then
		Response.Write "<script type='text/javascript'>alert('�̺�Ʈ�ڵ尡 �����ϴ�.');</script>"
		dbget.close() : Response.End
	end If
	if userid="" then
		Response.Write "<script type='text/javascript'>alert('������ ID�� �����ϴ�.');</script>"
		dbget.close() : Response.End
	end If
	if sub_idx="" then
		Response.Write "<script type='text/javascript'>alert('idx�� �����ϴ�.');</script>"
		dbget.close() : Response.End
	end If

'' �Ǽ� 899,900,901,902,903
'' �׼� 2809,2810,2811,2812,2813

	IF application("Svr_Info") = "Dev" THEN
		wdcp1code   =  2813		''20���� �̻� 2��������
		wdcp2code   =  2812		''50���� �̻� 6��������
		wdcp3code   =  2811		''100���� �̻� 15��������
		wdcp4code   =  2810		''�ٹ� ����������1 (1�����̻� ���Ž�)
		wdcp5code   =  2809		''�ٹ� ����������2 (1�����̻� ���Ž�)
	Else
		wdcp1code   =  903		''20���� �̻� 2��������
		wdcp2code   =  902		''50���� �̻� 6��������
		wdcp3code   =  901		''100���� �̻� 15��������
		wdcp4code   =  900		''�ٹ� ����������1 (1�����̻� ���Ž�)
		wdcp5code   =  899		''�ٹ� ����������2 (1�����̻� ���Ž�)
	End If

	dim CPdate
	CPdate = Year(now) &"-"& right("0"&Month(now)+3,2) &"-"& right("0"&Day(now),2) &" "& right("0"& Hour(now),2) &":"& right("0"&Minute(now),2) &":"& right("0"&Second(now),2)

	''���� 5�� �߱�
	sqlstr = "insert into [db_user].[dbo].tbl_user_coupon" + vbcrlf
	sqlstr = sqlstr & " (masteridx,userid,coupontype,couponvalue, couponname,minbuyprice,targetitemlist,startdate,expiredate,couponmeaipprice,validsitename)" + vbcrlf
	sqlstr = sqlstr & " 	SELECT idx, '"& userid &"',coupontype,couponvalue,couponname,minbuyprice,targetitemlist,startdate,expiredate,couponmeaipprice,validsitename" + vbcrlf
	sqlstr = sqlstr & " 	from [db_user].[dbo].tbl_user_coupon_master m" + vbcrlf
	sqlstr = sqlstr & " 	where idx in ("& wdcp1code &", "& wdcp2code &", "& wdcp3code &", "& wdcp4code &", "& wdcp5code &")"
'	response.write sqlstr
	dbget.execute sqlstr

	''sub_opt1 N�� Y�� ����
	sqlStr = "update db_event.dbo.tbl_event_subscript set" + vbcrlf
	sqlStr = sqlStr & " sub_opt1 = replace( sub_opt1, '/!/N', '/!/Y') where" + vbcrlf
	sqlStr = sqlStr & " evt_code='"& evt_code &"' and sub_idx='"& sub_idx &"'"
	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('OK');"
	Response.Write "	parent.top.location.replace('/admin/datamart/mkt/73007_manage.asp');"
	Response.Write "</script>"
	dbget.close() : Response.End

else
	Response.Write "<script type='text/javascript'>alert('�����ڰ� �����ϴ�.');</script>"
	dbget.close() : Response.End
end if

%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
