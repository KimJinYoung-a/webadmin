<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �����׼���
' Hieditor : 2011.11.25 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone2/zone_cls.asp"-->

<%
dim omanager ,j ,zoneidx ,divid , tmp
	zoneidx = requestCheckVar(request("zoneidx"),10)
	divid = requestCheckVar(request("divid"),32)

if divid = "" or zoneidx = "" then response.end

set omanager = new czone_list
	omanager.frectzoneidx = zoneidx
	omanager.Getshopzonemanager()

if omanager.FResultCount > 0 then
	for j=0 to omanager.FResultCount-1
		tmp = tmp & omanager.FItemList(j).fusername &"<Br>"
	next
end if

set omanager = nothing
%>
<script language="">
	var divid = '<%=divid%>';

	parent.eval("document.all."+divid).innerHTML = "<%=tmp%>";
</script>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
