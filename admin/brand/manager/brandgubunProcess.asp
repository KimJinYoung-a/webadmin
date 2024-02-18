<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.30 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/managerCls.asp"-->
<%
dim brandgubunarr, hello_ynarr, interview_ynarr, tenbytenand_ynarr, artistwork_ynarr,shop_collection_ynarr
dim shop_event_ynarr, lookbook_ynarr, mode, adminid, cnt, i, sqlStr
	adminid=session("ssBctId")
	brandgubunarr 	= Request("brandgubunarr")
	hello_ynarr 	= Request("hello_ynarr")
	interview_ynarr 	= Request("interview_ynarr")
	tenbytenand_ynarr 	= Request("tenbytenand_ynarr")
	artistwork_ynarr 	= Request("artistwork_ynarr")
	shop_collection_ynarr 	= Request("shop_collection_ynarr")
	shop_event_ynarr 	= Request("shop_event_ynarr")
	lookbook_ynarr 	= Request("lookbook_ynarr")	
	menupos	= request("menupos")
	mode	= request("mode")

If brandgubunarr="" THEN
	Response.Write "<script language='javascript'>history.back(-1);</script>"
	dbget.close()	:	response.End
end if

brandgubunarr = split(brandgubunarr,",")
hello_ynarr = split(hello_ynarr,",")
interview_ynarr = split(interview_ynarr,",")
tenbytenand_ynarr = split(tenbytenand_ynarr,",")
artistwork_ynarr = split(artistwork_ynarr,",")
shop_collection_ynarr = split(shop_collection_ynarr,",")
shop_event_ynarr = split(shop_event_ynarr,",")
lookbook_ynarr = split(lookbook_ynarr,",")

cnt = ubound(brandgubunarr)

For i = 0 to cnt-1
	sqlStr = "UPDATE db_brand.dbo.tbl_street_brandgubun SET " & VBCRLF
	sqlStr = sqlStr & " lastadminid = '"&adminid&"'" & VBCRLF
	sqlStr = sqlStr & " ,lastupdate=getdate()" & VBCRLF
	sqlStr = sqlStr & " ,hello_yn = '"&trim(hello_ynarr(i))&"'" & VBCRLF
	sqlStr = sqlStr & " ,interview_yn = '"&trim(interview_ynarr(i))&"'" & VBCRLF
	sqlStr = sqlStr & " ,tenbytenand_yn = '"&trim(tenbytenand_ynarr(i))&"'" & VBCRLF
	sqlStr = sqlStr & " ,artistwork_yn = '"&trim(artistwork_ynarr(i))&"'" & VBCRLF			
	sqlStr = sqlStr & " ,shop_collection_yn = '"&trim(shop_collection_ynarr(i))&"'" & VBCRLF	
	sqlStr = sqlStr & " ,shop_event_yn = '"&trim(shop_event_ynarr(i))&"'" & VBCRLF
	sqlStr = sqlStr & " ,lookbook_yn = '"&trim(lookbook_ynarr(i))&"'" & VBCRLF		
	sqlStr = sqlStr & " WHERE brandgubun =" & trim(brandgubunarr(i))
	
	'response.write sqlStr & "<br>"
	dbget.execute sqlStr
Next

%>

<script language='javascript'>
	alert('저장되었습니다');
	location.replace('/admin/brand/manager/brandgubun.asp?menupos=<%= menupos %>');
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->