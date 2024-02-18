<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 타블로이드
' Hieditor : 2009.11.17 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim tabloid, keyword, mainimage, regdate, isusing , detailimage
dim wordimage , ingimage , mode ,wordovimage
	tabloid = request("tabloid")
	keyword = request("keyword")
	mainimage = request("mainimg")
	isusing = request("isusing")
	detailimage = request("detailimg")				
	wordimage = request("wordimg")
	wordovimage = request("wordovimg")	
	ingimage = request("ingimg")	
	mode = request("mode")
dim sql

'// 삭제
if mode = "delete" then
	
	tabloid = left(tabloid,len(tabloid)-1)
	
	'//트랜젝션 
	dbget.begintrans
	
	sql = "update db_momo.dbo.tbl_tabloid set" + vbcrlf	
	sql = sql & " isusing='N'" + vbcrlf
	sql = sql & " where tabloid in("&tabloid&")" + vbcrlf	
	
	'response.write sql &"<br>"
	dbget.execute sql
	
	sql = ""	
	sql = "update db_momo.dbo.tbl_tabloid_item set" + vbcrlf	
	sql = sql & " isusing='N'" + vbcrlf
	sql = sql & " where tabloid in ("&tabloid&")" + vbcrlf	
	
	'response.write sql &"<br>"
	dbget.execute sql
	
	'오류검사 및 반영
	If Err.Number = 0 Then   
		dbget.CommitTrans				'커밋(정상)

	Else
	    dbget.RollBackTrans				'롤백(에러발생시)
				
		response.write "<script language='javascript'>"
		response.write "alert('정상적인 접근방식이 아니거나 오류가 발생되었습니다.');"
		response.write "self.close();"	
		response.write "</script>"
		rsget.close : resposne.end
	End If
	
elseif mode = "ing" then	

	tabloid = split(tabloid,",")

	if ubound(tabloid) <> "1" then
	response.write "<script>alert('한개만 선택해 주세요'); self.close();</script>"
	rsget.close() : response.end
	end if
		
	sql = "update db_momo.dbo.tbl_tabloid set" + vbcrlf
	sql = sql & " best = best + 50" + vbcrlf
	sql = sql & " where tabloid = "&tabloid(0)&"" + vbcrlf	
	
	'response.write sql &"<br>"
	dbget.execute sql	
end if	
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<script>
	opener.location.reload();
	alert('처리되었습니다');
	self.close();
</script>