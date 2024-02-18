<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'공통
dim evt_code
dim mode, sqlStr

'콘텐츠
dim main_copy
dim sub_copy
dim main_color
dim main_content
dim background_img
dim content_order

'유닛 
dim content_idx
dim unit_class
dim unit_order
dim unit_main_copy
dim unit_main_content
dim tag
dim unitDeleteIdx

'아이템
dim itemid
dim itemIdx
dim unit_idx
dim item_img
dim item_name
dim item_order
dim itemDeleteIdx

'기타
dim adminName
dim unitIdx
dim contentIdx
dim unitModParam
'공통 파라미터
evt_code		= request("evt_code")	
unitModParam	= request("unitModParam")
'콘텐츠 파라미터
mode			= request("mode")
main_copy		= request("main_copy")	
sub_copy		= request("sub_copy")	
main_color		= request("main_color")	
main_content	= request("main_content")		
background_img	= request("background_img")		
content_order	= request("content_order")		

'유닛 파라미터
content_idx		 	= request("content_idx")
unit_class		 	= request("unit_class")
unit_order		 	= request("unit_order")
unit_main_copy		= requestCheckVar(request("unit_main_copy"),128)
unit_main_content	= requestCheckVar(request("unit_main_content"),256)
unitDeleteIdx		= request("unitDeleteIdx")
tag		 			= request("tag")

'아이템 파라미터
itemid			= request("itemid")		
unit_idx		= request("unit_idx")			
item_img		= request("item_img")			
item_name		= request("item_name")			
item_order		= request("item_order")			
itemDeleteIdx		= request("itemDeleteIdx")

'기타 파라미터
contentIdx 		= request("contentidx")	' 콘텐츠 idx
unitIdx 		= request("unitIdx")' 유닛 idx
itemIdx			= request("itemIdx")' 아이템 idx
public Function GetAdminName(adminid)	
	If IsNull(adminid) Or adminid="" Then Exit Function
	On Error Resume Next
	dim SqlStr

	sqlStr = " Select top 1 username "
	sqlStr = sqlStr & " From db_partner.dbo.tbl_user_tenbyten "
	sqlStr = sqlStr & " where userid = '"& adminid &"'"
	rsget.CursorLocation = adUseClient
	rsget.CursorType=adOpenStatic
	rsget.Locktype=adLockReadOnly
	rsget.Open sqlStr, dbget

	If Not(rsget.bof Or rsget.eof) Then
		GetAdminName = rsget("username")
	End If
	rsget.close
	On Error goto 0
End Function	

adminName = GetAdminName(session("ssBctId"))	

'// 모드에 따른 분기
Select Case mode
	'수정
	Case "mod"		
	'콘텐츠 수정
		sqlStr = " Update db_event.dbo.[tbl_multi3_contents] " &_
				" Set main_copy='" & html2db(main_copy) & "'" &_				
				" ,	sub_copy ='" & html2db(sub_copy) & "'" &_
				" ,	main_color ='" & main_color & "'" &_
				" ,	main_content ='" & html2db(main_content) & "'" &_				
				" ,	background_img ='" & background_img & "'" &_				
				" ,	mod_name ='" & adminName & "'" &_				
				" ,	moddate = getdate()" &_						
				" Where evt_code =" & evt_code
		'response.write sqlStr
		'response.end
		dbget.Execute(sqlStr)	
		
	'유닛 수정
		dim unitIdxArr, i
		unitIdxArr = split(unitIdx, ",")		
		'response.write unitIdx & "<br>"
		'response.write Ubound(unitIdxArr) & "<br>"
		'response.write unitIdxArr(1)
		'response.end		
		for i = 1 to Ubound(unitIdxArr) + 1
			sqlStr = " Update db_event.dbo.[tbl_multi3_content_units] " & vbcrlf
			sqlStr = sqlStr + " Set evt_code ='" & evt_code & "'" & vbcrlf
			sqlStr = sqlStr + " 	,content_idx ='" & content_idx & "'" & vbcrlf
			sqlStr = sqlStr + " 	,unit_class ='" & request("unit_class")(i) & "'" & vbcrlf
			sqlStr = sqlStr + " 	,unit_order ='" & request("unit_order")(i) & "'" & vbcrlf
			sqlStr = sqlStr + " 	,unit_main_copy ='" & html2db(requestCheckVar(request("unit_main_copy")(i),128)) & "'" & vbcrlf
			sqlStr = sqlStr + " 	,unit_main_content ='" & html2db(requestCheckVar(request("unit_main_content")(i),256)) & "'" 	& vbcrlf			
			sqlStr = sqlStr + " 	,tag ='" & request("tag")(i) & "'" & vbcrlf
			sqlStr = sqlStr + " 	,mod_name ='" & adminName & "'" & vbcrlf
			sqlStr = sqlStr + " 	,moddate = getdate()" 		& vbcrlf
			sqlStr = sqlStr + " Where idx=" & request("unitIdx")(i) & vbcrlf
			dbget.Execute(sqlStr)
		next
	'아이템 수정
		dim itemIdxArr
		itemIdxArr = split(itemIdx, ",")				

		for i = 1 to Ubound(itemIdxArr) + 1
			sqlStr = " Update db_event.dbo.[tbl_multi3_items] " & vbcrlf
			sqlStr = sqlStr + " Set evt_code ='" & evt_code & "'" & vbcrlf
			sqlStr = sqlStr + " ,itemid ='" & request("itemid")(i) & "'" & vbcrlf		
			sqlStr = sqlStr + " ,item_img ='" & request("item_img")(i) & "'" & vbcrlf
			sqlStr = sqlStr + " ,item_name ='" & request("item_name")(i) & "'" & vbcrlf
			sqlStr = sqlStr + " ,item_order ='" & request("item_order")(i) & "'" & vbcrlf				
			sqlStr = sqlStr + " ,mod_name = '" & request("adminName") & "'" & vbcrlf				
			sqlStr = sqlStr + " ,moddate = getdate()" & vbcrlf				
			sqlStr = sqlStr + " Where idx=" & request("itemIdx")(i) & vbcrlf				
			dbget.Execute(sqlStr)
		next		
	'유닛 추가
	Case "unitadd"			
		sqlStr = " Insert Into db_event.dbo.[tbl_multi3_content_units] " & vbcrlf
		sqlStr = sqlStr + " (evt_code , content_idx, unit_class , unit_order, unit_main_copy, unit_main_content, tag, reg_name ) values " & vbcrlf					
		sqlStr = sqlStr + " ('" & evt_code &"'" & vbcrlf
		sqlStr = sqlStr + " ,'" & content_idx &"'" & vbcrlf
		sqlStr = sqlStr + " ,'" & unit_class &"'" & vbcrlf
		sqlStr = sqlStr + " ,'" & unit_order &"'" & vbcrlf
		sqlStr = sqlStr + " ,'" & html2db(unit_main_copy) &"'" & vbcrlf
		sqlStr = sqlStr + " ,'" & html2db(unit_main_content) &"'" & vbcrlf
		sqlStr = sqlStr + " ,'" & tag &"'" & vbcrlf
		sqlStr = sqlStr + " ,'" & adminName &"'" & vbcrlf
		sqlStr = sqlStr + ")" & vbcrlf
		dbget.Execute(sqlStr)					
	'유닛 수정
	Case "unitdelete"
		sqlStr = "delete db_event.dbo.[tbl_multi3_content_units] " &_
				" Where idx=" & unitDeleteIdx
		dbget.Execute(sqlStr)		
	'아이템 추가
	Case "itemadd"			
		sqlStr = "Insert Into db_event.dbo.[tbl_multi3_items] " &_
					" (evt_code , itemid , unit_idx, item_img , item_name, item_order, reg_name ) values " &_					
					" ('" & evt_code &"'" &_
					" ,'" & itemid &"'" &_
					" ,'" & unitIdx &"'" &_
					" ,'" & item_img &"'" &_
					" ,'" & item_name &"'" &_
					" ,'" & item_order &"'" &_
					" ,'" & adminName &"'" &_
					")"
		'response.write sqlStr					
		'response.end
		dbget.Execute(sqlStr)					
	'아이템 수정
	Case "itemdelete"
		sqlStr = "delete db_event.dbo.[tbl_multi3_items] " &_
				" Where idx=" & itemDeleteIdx
		dbget.Execute(sqlStr)		
	Case "contentadd"
		'콘텐츠 등록
		sqlStr = "Insert Into db_event.dbo.[tbl_multi3_contents] " &_
					" (evt_code, main_copy , sub_copy , main_color, main_content , background_img, reg_name, content_order) values " &_					
					" ('" & evt_code &"'" &_
					" ,'" & html2db(main_copy) &"'" &_
					" ,'" & html2db(sub_copy) &"'" &_
					" ,'" & main_color &"'" &_
					" ,'" & html2db(main_content) & "'" &_
					" ,'" & background_img & "'" &_
					" ,'" & adminName & "'" &_
					" ,'" & content_order & "'" &_									
					")"		
		dbget.Execute(sqlStr)
End Select
%>
<script>
<% If mode = "" then%>
	// 목록으로 복귀
	alert("저장했습니다.");
	window.opener.document.location.href = window.opener.document.URL;    // 부모창 새로고침	 
<% elseif mode = "unitmodify" or mode = "itemmodify" or mode = "contentadd" or mode = "contentmodify" then%>
	// 페이지 새로고침
	alert("저장했습니다.");	
	location.href = "pop_manage_multi3.asp?evt_code=<%=evt_code%>&unitModParam=<%=unitModParam%>";
<% elseif mode = "unitdelete" or mode = "itemdelete" then %>
	// 페이지 새로고침
	alert("삭제했습니다.");
	location.href = document.referrer;	
<% elseif mode = "itemadd" then %>
	// 페이지 새로고침
	alert("저장했습니다.");	
	self.close();
	window.opener.document.location.href = "pop_manage_multi3.asp?evt_code=<%=evt_code%>&unitIdxAddPram=<%=unitIdx%>";    
<% elseif mode = "mod" then %>
	alert("저장했습니다.");	
	location.href = "pop_manage_multi3.asp?evt_code=<%=evt_code%>&unitModParam=<%=unitModParam%>";	
<% Else %>
	// 목록으로 복귀
	alert("저장했습니다.");	
	self.close();
	window.opener.document.location.href = window.opener.document.URL;    
<% End If %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
