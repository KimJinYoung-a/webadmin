<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station ����Ͽ� ǥ�����ļ��� �ϰ�����
' Hieditor : 2012.01.12 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
	Dim strSql, lp
	Dim evtCode, mSortNo, arrEvtCd, arrSrtNo, param

	param = "?evt_code_search=" & request("evt_code_search")
	param = param  & "&evt_type_searchbox=" & request("evt_type_searchbox")
	param = param  & "&isusing_searchbox=" & request("isusing_searchbox")
	param = param  & "&evt_code_countbox=" & request("evt_code_countbox")
	param = param  & "&evt_mobile_yn=" & request("evt_mobile_yn")
	param = param  & "&menupos=" & request("menupos")
	param = param  & "&page=" & request("page")

	evtCode = request("evt_code")
	mSortNo = request("m_sortNo")
	arrEvtCd = split(evtCode,",")
	arrSrtNo = split(mSortNo,",")

	if evtCode="" or mSortNo="" then
		Call Alert_Return("�߸��� �����Դϴ�.")
	end if

	strSql = ""
	for lp=0 to ubound(arrEvtCd)
		strSql = strSql & "Update db_culture_station.dbo.tbl_culturestation_event Set m_sortNo=" & arrSrtNo(lp) & " Where evt_code=" & arrEvtCd(lp) & vbCrLf
	next
	
	dbget.execute strSql

	Call Alert_Move("�ϰ� ó���Ǿ����ϴ�.","event_list.asp" & param)
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->