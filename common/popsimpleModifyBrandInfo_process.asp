<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : CS정보
' History : 서동석 생성
'           2021.06.18 한용민 수정(담당자 휴대폰,이메일 인증정보 데이터쪽에도 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

dim mode, submode, makerid
dim returnName, returnPhone, returnhp, returnEmail, returnZipcode, returnZipaddr, returnEtcaddr
dim csName, csPhone, cshp, csEmail, groupid, sql
	groupid		= requestCheckvar(trim(request("groupid")),10)
mode  = Replace(request("mode"), "'", "")
submode  = Replace(request("submode"), "'", "")
makerid  = Replace(request("makerid"), "'", "")

returnName  = Replace(request("returnName"), "'", "")
returnPhone  = Replace(request("returnPhone"), "'", "")
returnhp  = Replace(request("returnhp"), "'", "")
returnEmail  = Replace(request("returnEmail"), "'", "")
returnZipcode  = Replace(request("returnZipcode"), "'", "")
returnZipaddr  = Replace(request("returnZipaddr"), "'", "")
returnEtcaddr  = Replace(request("returnEtcaddr"), "'", "")

csName  = Replace(request("csName"), "'", "")
csPhone  = Replace(request("csPhone"), "'", "")
cshp  = Replace(request("cshp"), "'", "")
csEmail  = Replace(request("csEmail"), "'", "")


dim sqlStr, i

if (mode = "modifyReturnCharge") then
	'// 반품주소+담당자
	sqlStr = " update [db_partner].[dbo].tbl_partner" & vbcrlf
	sqlStr = sqlStr & " set return_zipcode = '" + CStr(html2db(returnZipcode)) + "', return_address = '" + CStr(html2db(returnZipaddr)) + "'" & vbcrlf
	sqlStr = sqlStr & " , return_address2 = '" + CStr(html2db(returnEtcaddr)) + "', deliver_phone = '" + CStr(html2db(returnPhone)) + "'" & vbcrlf
	sqlStr = sqlStr & " , deliver_hp = '" + CStr(html2db(returnhp)) + "', deliver_name = '" + CStr(html2db(returnName)) + "'" & vbcrlf
	sqlStr = sqlStr & " , deliver_email = '" + CStr(html2db(returnEmail)) + "', lastInfoChgDT=getdate() where" & vbcrlf
	sqlStr = sqlStr & " id = '" + CStr(makerid) + "'" & vbcrlf

	'response.write sqlStr & "<br>"
	dbget.Execute sqlStr

elseif (mode = "modifyCSCharge") then
	'cshp = replace(cshp,"-","")

	'// CS담당자
	if (submode = "ins") then
		'
		sqlStr = " insert into [db_cs].[dbo].tbl_cs_brand_memo(brandid, cs_name, cs_phone, cs_hp, cs_email, cs_modifyday, cs_reguserid) "
		sqlStr = sqlStr + " values('" + CStr(makerid) + "', '" + CStr(html2db(csName)) + "', '" + CStr(html2db(csPhone)) + "', '" + CStr(html2db(cshp)) + "', '" + CStr(html2db(csEmail)) + "', getdate(), '" + CStr(html2db(session("ssBctId"))) + "') "
		dbget.Execute sqlStr

	elseif (submode = "mod") then
		sqlStr = " update [db_cs].[dbo].tbl_cs_brand_memo "
		sqlStr = sqlStr + " set cs_name = '" + CStr(html2db(csName)) + "', cs_phone = '" + CStr(html2db(csPhone)) + "', cs_hp = '" + CStr(html2db(cshp)) + "', cs_email = '" + CStr(html2db(csEmail)) + "', cs_modifyday = getdate(), cs_reguserid = '" + CStr(html2db(session("ssBctId"))) + "' "
		sqlStr = sqlStr + " where brandid = '" + CStr(makerid) + "' "
		dbget.Execute sqlStr
	end if

	sql ="if exists(select userid from db_partner.dbo.tbl_partner_user with (nolock) where isusing='Y' and groupid ='"& groupid &"' and gubun=4 and userid='"& makerid &"')"
	sql = sql & " begin"
	sql = sql & "   update db_partner.dbo.tbl_partner_user set" & vbcrlf
	sql = sql & "   lastUpdate=getdate(),name=N'"& html2db(csName) &"',Title=N'CS담당자'" & vbcrlf
	sql = sql & " ,hp=N'"& html2db(cshp) &"'" & vbcrlf
	sql = sql & " ,email=N'"& html2db(csEmail) &"'" & vbcrlf
	sql = sql & "   where isusing='Y' and groupid ='"& groupid &"' and gubun=4 and userid='"& makerid &"'"
	sql = sql & " end"
	sql = sql & " else"
	sql = sql & " begin"
	sql = sql & "   insert into db_partner.dbo.tbl_partner_user (groupid,userid,gubun,Title,name"
	sql = sql & "   ,hp,email"
	sql = sql & "   ,regdate,lastUpdate,isUsing)"
	sql = sql & "       select N'"& groupid &"',N'"& makerid &"',4,N'CS담당자',N'"& html2db(csName) &"'"
	sql = sql & "   	,N'"& html2db(cshp) &"',N'"& html2db(csEmail) &"'" & vbcrlf
	sql = sql & "       ,getdate(),getdate(),N'Y'"
	sql = sql & " end"

	'response.write sql & "<Br>"
	dbget.Execute sql
end if

%>

<script language="javascript">
alert("저장 되었습니다.");
opener.location.reload();
opener.focus();
window.close();
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
