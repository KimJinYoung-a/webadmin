<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 함께해요
' Hieditor : 2010.11.18 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim withid,startdate,enddate,isusing , i , mode  , winner , sql , regdate
dim idx , userid , comment , orderno , withgubun , withimage_large , withimage_small
	idx = request("idx")
	withgubun = request("withgubun")
	withid = request("withid")
	startdate = request("startdate")
	enddate = request("enddate")
	isusing = request("isusing")	
	mode = request("mode")
	userid = request("userid")
	comment = request("comment")
	orderno = request("orderno")
	withimage_large = request("withimage_large")
	withimage_small = request("withimage_small")
	regdate = request("regdate")

'//신규 & 수정 
if mode = "edit" then
			
	'//신규	
	if withid = "" then
		sql = "insert into db_momo.dbo.tbl_with (startdate,enddate,isusing)" + vbcrlf
		sql = sql & " values (" + vbcrlf
		sql = sql & " '"&html2db(startdate)&" 00:00:00'"		
		sql = sql & " ,'"&html2db(enddate)&" 23:59:59'"		
		sql = sql & " ,'"&isusing&"'"
		sql = sql & " )"		
	
	'response.write sql &"<br>"
	dbget.execute sql
		
	'//수정	
	else 
	
	if withid = "" then
		response.write "<script>alert('아이디 값이 없습니다.'); self.close();</script>"
		dbget.close() : response.end
	end if		

	sql = "update db_momo.dbo.tbl_with set" + vbcrlf	
	sql = sql & " startdate='"&html2db(startdate)&" 00:00:00'" + vbcrlf
	sql = sql & " ,enddate='"&html2db(enddate)&" 23:59:59'" + vbcrlf		
	sql = sql & " ,isusing='"&isusing&"'" + vbcrlf		
	sql = sql & " where withid = "&withid&"" + vbcrlf	
	
	'response.write sql &"<br>"
	dbget.execute sql
		
	end if			

elseif mode = "editsns" then
	
	if withid = "" then 
		response.write "<script>alert('ID 값이 없습니다'); self.close();</script>"
		dbget.close() : response.end
	end if
	
	'//신규	
	if idx = "" then
		sql = "insert into db_momo.dbo.tbl_with_comment (withid ,withgubun,userid,comment,isusing,withimage_small,withimage_large,orderno,regdate)" + vbcrlf
		sql = sql & " values (" + vbcrlf
		sql = sql & " "&withid&""
		sql = sql & " ,"&withgubun&""
		sql = sql & " ,'"&userid&"'"
		sql = sql & " ,'"&html2db(comment)&"'"
		sql = sql & " ,'"&isusing&"'"
		sql = sql & " ,'"&html2db(withimage_small)&"'"
		sql = sql & " ,'"&html2db(withimage_large)&"'"
		sql = sql & " ,"&orderno&""
		sql = sql & " ,'"&html2db(regdate)&"'"
		sql = sql & " )"		
	
	'response.write sql &"<br>"
	dbget.execute sql
		
	'//수정	
	else 
	
	if withid = "" then
		response.write "<script>alert('아이디 값이 없습니다.'); self.close();</script>"
		dbget.close() : response.end
	end if		

	sql = "update db_momo.dbo.tbl_with_comment set" + vbcrlf	
	sql = sql & " withgubun="&withgubun&"" + vbcrlf
	sql = sql & " ,userid='"&userid&"'" + vbcrlf
	sql = sql & " ,comment='"&html2db(comment)&"'" + vbcrlf
	sql = sql & " ,isusing='"&isusing&"'" + vbcrlf
	sql = sql & " ,withimage_small='"&html2db(withimage_small)&"'" + vbcrlf
	sql = sql & " ,withimage_large='"&html2db(withimage_large)&"'" + vbcrlf
	sql = sql & " ,orderno="&orderno&"" + vbcrlf
	sql = sql & " ,regdate='"&html2db(regdate)&"'" + vbcrlf
	sql = sql & " where withid = "&withid&" and idx = "&idx&"" + vbcrlf	
	
	'response.write sql &"<br>"
	dbget.execute sql
		
	end if		
	
end if	
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<script>
	opener.location.reload();
	alert('처리되었습니다');
	self.close();
</script>
