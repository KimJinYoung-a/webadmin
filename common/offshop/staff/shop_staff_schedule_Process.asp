<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  오프라인 매장근무관리
' History : 2011.03.17 한용민 생성
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/staff/staff_cls.asp"-->
<%
Dim mode, idx
Dim userid, username, posit_sn, part_sn, ChkStart, ChkEnd, etcComment
Dim SQL, strMsg , empno ,shopid
	mode 		= requestCheckvar(request("mode"),4)
	idx 		= requestCheckvar(request("idx"),8)	
	userid		= requestCheckvar(request("userid"),32)
	username	= requestCheckvar(request("username"),20)
	posit_sn	= requestCheckvar(request("posit_sn"),8)
	part_sn		= requestCheckvar(request("part_sn"),8)
	empno		= requestCheckvar(request("empno"),32)
	shopid		= requestCheckvar(request("shopid"),32)
	ChkStart	= requestCheckvar(request("ChkStart"),10) & " " & replace(requestCheckvar(request("ChkSTime"),8),"24:00:00","23:59:59")
	ChkEnd		= requestCheckvar(request("ChkEnd"),10) & " " & replace(requestCheckvar(request("ChkETime"),8),"24:00:00","23:59:59")
	etcComment	= html2db(request("etcComment"))

	'// 처리 분기 //
	Select Case mode
		
		Case "add"
			strMsg = "등록되었습니다."
			SQL =	"Insert into db_shop.dbo.tbl_shop_staff_schedule " &_
					" (userid,empno,shopid, username, posit_sn, part_sn, ChkStart, ChkEnd, etcComment) values " &_
					" ('" & userid & "'" &_
					" ,'" & empno & "'" &_
					" ,'" & shopid & "'" &_
					" ,'" & username & "'" &_
					" ,'" & posit_sn & "'" &_
					" ,'" & part_sn & "'" &_
					" ,'" & ChkStart & "'" &_
					" ,'" & ChkEnd & "'" &_					
					" ,'" & etcComment & "')"
			
			'response.write SQL &"<br>"
			dbget.Execute(SQL)
		
		Case "modi"
			strMsg = "수정되었습니다."
			SQL =	"Update db_shop.dbo.tbl_shop_staff_schedule Set " &_
					"	userid = '" & userid & "' " &_
					"	,empno = '" & empno & "' " &_
					"	,shopid = '" & shopid & "' " &_
					"	,username = '" & username & "' " &_
					"	,posit_sn = '" & posit_sn & "' " &_
					"	,part_sn = '" & part_sn & "' " &_
					"	,ChkStart = '" & ChkStart & "' " &_
					"	,ChkEnd = '" & ChkEnd & "' " &_					
					"	,etcComment = '" & etcComment & "' " &_
					"Where idx=" & idx
			
			'response.write SQL &"<br>"
			dbget.Execute(SQL)
		
		Case "del"
			strMsg = "처리가 완료되었습니다."
			SQL =	"Update db_shop.dbo.tbl_shop_staff_schedule Set " &_
					"	isUsing = 'N' " &_
					"Where idx=" & idx
			
			'response.write SQL &"<br>"
			dbget.Execute(SQL)
	End Select

	response.write	"<script language='javascript'>" &_
					"	alert('" & strMsg & "');" &_
					"	opener.history.go(0);" &_
					"	self.close();" &_
					"</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->