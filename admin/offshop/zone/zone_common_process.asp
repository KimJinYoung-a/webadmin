<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 삽별구역설정
' Hieditor : 2010.01.11 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone/zone_cls.asp"-->
<%
dim mode , zonegroup , zonegroup_name , isusing
	mode = requestCheckVar(request("mode"),32)
	zonegroup = requestCheckVar(request("zonegroup"),10)
	zonegroup_name = requestCheckVar(request("zonegroup_name"),32)
	isusing = requestCheckVar(request("isusing"),1)
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->