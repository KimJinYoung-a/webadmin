<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim sMode,adminid
Dim strSql 
dim arrempno,intLoop,arrmidx
dim areaDiv, content
 
sMode =requestCheckvar(request("hidM"),2)
adminid	= session("ssBctId")
Select Case sMode
	Case "A"	'��ü ���(�����)
		strSql = " exec [db_partner].[dbo].[usp_Ten_user_tenbyten_InsertAllYearAgitPoint] '"& adminid& "' " 
		dbget.Execute(strSql)
		response.write "<script>location.href='/admin/member/agit/';alert('��ϵǾ����ϴ�.');</script>"

	Case "I"	'�̵���� ���(������)
		arrempno = requestCheckvar(request("chki"),500)   
		arrempno = split(arrempno,",")
		 
		For intLoop = 0 To ubound(arrempno)
			strSql = " exec [db_partner].[dbo].[usp_ten_user_tenbyten_InsertMonthAgitPoint] '"&Trim(arrempno(intLoop))&"' ,'"& adminid& "' " 
		  dbget.Execute(strSql)
		Next			
		response.write "<script>self.close();opener.location.href='/admin/member/agit/';alert('��ϵǾ����ϴ�.');</script>"
	
	Case "M"	'�Ա�, Ű�ݳ� ó��
		dim isipkum, idx,arridx
		
		arridx =  requestCheckvar(request("chki"),1000)
		arridx = split(arridx,",")
		
		for intLoop = 0 To ubound(arridx)
		strSql = "update db_partner.dbo.tbl_TenAgit_Booking set isipkum = "&requestCheckvar(request("rdoin"&Trim(arridx(intLoop))),3)&" ,isreturnkey="&requestCheckvar(request("rdorek"&Trim(arridx(intLoop))),3)&", ipkumdate=getdate() , lastupdate=getdate() , adminid ='"&session("ssBctId")&"' where idx = "&Trim(arridx(intLoop))
		dbget.Execute(strSql)
		next
		
		response.write "<script>parent.location.href='/admin/member/agit/useList.asp';document.location.href = 'about:blank'; alert('ó���Ǿ����ϴ�.');</script>"

	Case "SI"	'����Ʈ ���ھȳ� ���
		areaDiv = requestCheckvar(request("areaDiv"),4)
		content = requestCheckvar(request("agitSmsCont"),1000)
		strSql = "INSERT INTO db_partner.dbo.tbl_TenAgit_smsInfo "
		strSql = strSql & " VALUES (" & areaDiv & ",N'" & content & "',getdate(),'" & adminid & "',getdate() )"
		dbget.Execute(strSql)
		response.write "<script>alert('��ϵǾ����ϴ�.');location.href='/admin/member/agit/popAgitInfoSms.asp';</script>"

	Case "SU"	'����Ʈ ���ھȳ� ����
		areaDiv = requestCheckvar(request("areaDiv"),4)
		content = requestCheckvar(request("agitSmsCont"),1000)
		strSql = "UPDATE db_partner.dbo.tbl_TenAgit_smsInfo SET "
		strSql = strSql & "contents= N'" & content & "', "
		strSql = strSql & "adminid='" & adminid & "', "
		strSql = strSql & "lastUpdate=getdate() "
		strSql = strSql & "WHERE AreaDiv=" & areaDiv
		dbget.Execute(strSql)
		response.write "<script>alert('����Ǿ����ϴ�.');location.href='/admin/member/agit/popAgitInfoSms.asp';</script>"
END Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->