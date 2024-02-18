<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 다이 상품 등록 대기 상품 
' Hieditor : 2010.10.20 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode,idx,rejectmsg ,sqlStr
	mode = RequestCheckvar(request("mode"),16)
	idx = RequestCheckvar(request("idx"),10)
	rejectmsg = html2Db(request("rejectmsg"))
  	if rejectmsg <> "" then
		if checkNotValidHTML(rejectmsg) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		response.End
		end if
	end if
if mode="waitstate" then
	sqlStr = "update [db_academy].dbo.tbl_diy_wait_item"
	sqlStr = sqlStr + " set currstate='1'"
	sqlStr = sqlStr + " where itemid=" + idx
	
	'response.write sqlStr &"<br>"
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	
elseif mode="delstate" then
    ''진행안함
	sqlStr = "update [db_academy].dbo.tbl_diy_wait_item"
	sqlStr = sqlStr + " set currstate='0'"
	sqlStr = sqlStr + " ,rejectmsg='" + rejectmsg + "'"
	sqlStr = sqlStr + " ,rejectDate=getdate() "
	sqlStr = sqlStr + " where itemid=" + idx
	
	'response.write sqlStr &"<br>"
	dbACADEMYget.execute(sqlStr)
	
	''2016/13/08 추가
	sqlStr = "exec db_academy.[dbo].[sp_ACA_sendPushMsgItemConfirm_Artist] "&idx&",NULL"
	dbACADEMYget.execute(sqlStr)
	
else
    ''등록보류(재요청)
	sqlStr = "update [db_academy].dbo.tbl_diy_wait_item"
	sqlStr = sqlStr + " set currstate='2'"
	sqlStr = sqlStr + " ,rejectmsg='" + rejectmsg + "'"
	sqlStr = sqlStr + " ,rejectDate=getdate() "
	sqlStr = sqlStr + " where itemid=" + idx

	'response.write sqlStr &"<br>"
	dbACADEMYget.execute(sqlStr)
	
	''2016/13/08 추가
	sqlStr = "exec db_academy.[dbo].[sp_ACA_sendPushMsgItemConfirm_Artist] "&idx&",NULL"
	dbACADEMYget.execute(sqlStr)
end if
%>

<script language="JavaScript">

<% if mode="waitstate" then %>

	alert("등록 대기로 변경 되었습니다.");
	history.go(-1);

<% elseif mode="delstate" then %>

	alert("진행안함 으로 변경 되었습니다.");
	opener.location.reload();
	window.close();

<% else %>

	alert("등록 보류(재요청) 되었습니다.");
	opener.location.reload();
	window.close();

<% end if %>
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->