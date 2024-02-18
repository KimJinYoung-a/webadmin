<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_rankingCls.asp"-->

<%
	Dim sql, vIdx, vDIdx, vTitle, vSDate, vEdate, vIsusing, vOrderNum, vItemID, vItemName, vItemImg1, vItemImg2, vDIsusing, vItemDetail, vTitleImg
	vIdx 		= Request("idx")
	vDIdx 		= Request("didx")
	
	vTitle 		= html2db(Request("title"))
	vTitleImg	= Request("title_img")
	
	vSDate 		= Request("sdate")
	vEdate 		= Request("edate")
	vIsusing 	= Request("isusing")
	vOrderNum 	= Request("ordernum")
	vItemID 	= Request("itemid")
	
	vItemName 	= html2db(Request("itemname"))
	
	vItemDetail = html2db(Request("itemdetail"))
	
	vItemImg1 	= Request("itemimg1")
	vItemImg2 	= Request("itemimg2")
	vDIsusing 	= Request("disusing")
	
	
	'On Error Resume Next
	dbget.beginTrans
	
	If vIdx = "" Then
		sql = "INSERT INTO [db_momo].[dbo].[tbl_ranking_master] " & _
			  "		(title, title_img, startdate, enddate, isusing) " & _
			  "		VALUES " & _
			  "		('" & vTitle & "', '" & vTitleImg & "', '" & vSDate & "', '" & vEdate & "', '" & vIsusing & "') "
		dbget.execute sql
		
		sql = " SELECT @@identity "
		rsget.Open sql,dbget
		
		IF not rsget.Eof THEN
			vIdx = rsget(0)
		END IF
		rsget.Close
		
		If vItemID <> "0" Then
			sql = "SELECT basicimage, smallimage FROM [db_item].[dbo].[tbl_item] WHERE itemid = '" & vItemID & "'"
			rsget.Open sql,dbget,1
			IF not rsget.Eof THEN
				vItemImg1 = webImgUrl & "/image/basic/" & GetImageSubFolderByItemid(vItemID) & "/" & rsget(0)
				vItemImg2 = webImgUrl & "/image/small/" & GetImageSubFolderByItemid(vItemID) & "/" & rsget(1)
				rsget.Close
			ELSE
				rsget.Close
		        dbget.RollBackTrans
		        dbget.close()
		        response.write "<script>alert('" & vItemID & " 없는 상품입니다.')</script>"
		        response.write "<script>history.back()</script>"
		        response.end
			END IF
		End If
		
		
		sql = "INSERT INTO [db_momo].[dbo].[tbl_ranking_detail] " & _
			  "		(masteridx, ordernum, itemid, itemname, itemdetail, itemimg1, itemimg2, isusing) " & _
			  "		VALUES " & _
			  "		('" & vIdx & "', '" & vOrderNum & "', '" & vItemID & "', '" & vItemName & "', '" & vItemDetail & "', '" & vItemImg1 & "', '" & vItemImg2 & "', '" & vDIsusing & "') "
		dbget.execute sql
		
	Else
		sql = "UPDATE [db_momo].[dbo].[tbl_ranking_master] SET " & _
			  "		title = '" & vTitle & "', " & _
			  "		title_img = '" & vTitleImg & "', " & _
			  "		startdate = '" & vSDate & "', " & _
			  "		enddate = '" & vEdate & "', " & _
			  "		isusing = '" & vIsusing & "' " & _
			  "	WHERE idx = '" & vIdx & "' "
		dbget.execute sql
		
		If vDIdx <> "" Then
			If vItemID <> "0" Then
				sql = "SELECT basicimage, smallimage FROM [db_item].[dbo].[tbl_item] WHERE itemid = '" & vItemID & "'"
				rsget.Open sql,dbget,1
				IF not rsget.Eof THEN
					vItemImg1 = webImgUrl & "/image/basic/" & GetImageSubFolderByItemid(vItemID) & "/" & rsget(0)
					vItemImg2 = webImgUrl & "/image/small/" & GetImageSubFolderByItemid(vItemID) & "/" & rsget(1)
					rsget.Close
				ELSE
					rsget.Close
			        dbget.RollBackTrans
			        dbget.close()
			        response.write "<script>alert('" & vItemID & " 없는 상품입니다.')</script>"
			        response.write "<script>history.back()</script>"
			        response.end
				END IF
			End If
			
			sql = "UPDATE [db_momo].[dbo].[tbl_ranking_detail] SET " & _
				  "		ordernum = '" & vOrderNum & "', " & _
				  "		itemid = '" & vItemID & "', " & _
				  "		itemname = '" & vItemName & "', " & _
				  "		itemdetail = '" & vItemDetail & "', " & _
				  "		itemimg1 = '" & vItemImg1 & "', " & _
				  "		itemimg2 = '" & vItemImg2 & "', " & _
				  "		isusing = '" & vDIsusing & "' " & _
				  "	WHERE idx = '" & vDIdx & "' "
			dbget.execute sql
		Else
			If vDIsusing <> "" Then
				If vItemID <> "0" Then
					sql = "SELECT basicimage, smallimage FROM [db_item].[dbo].[tbl_item] WHERE itemid = '" & vItemID & "'"
					rsget.Open sql,dbget,1
					IF not rsget.Eof THEN
						vItemImg1 = webImgUrl & "/image/basic/" & GetImageSubFolderByItemid(vItemID) & "/" & rsget(0)
						vItemImg2 = webImgUrl & "/image/small/" & GetImageSubFolderByItemid(vItemID) & "/" & rsget(1)
						rsget.Close
					ELSE
						rsget.Close
				        dbget.RollBackTrans
				        dbget.close()
				        response.write "<script>alert('" & vItemID & " 없는 상품입니다.')</script>"
				        response.write "<script>history.back()</script>"
				        response.end
					END IF
				End If
				
				sql = "INSERT INTO [db_momo].[dbo].[tbl_ranking_detail] " & _
					  "		(masteridx, ordernum, itemid, itemname, itemdetail, itemimg1, itemimg2, isusing) " & _
					  "		VALUES " & _
					  "		('" & vIdx & "', '" & vOrderNum & "', '" & vItemID & "', '" & vItemName & "', '" & vItemDetail & "', '" & vItemImg1 & "', '" & vItemImg2 & "', '" & vDIsusing & "') "
				dbget.execute sql
			End If
		End If
	End If
	
	
	If Err.Number = 0 Then
	        dbget.CommitTrans
	Else
	        dbget.RollBackTrans
	        dbget.close()
	        response.write "<script>alert('[" & vError & "]데이타를 저장하는 도중에 에러가 발생하였습니다.')</script>"
	        response.write "<script>history.back()</script>"
	        response.end
	End If
	
'on error Goto 0
	
	dbget.close()
	Response.Write "<script>alert('저장되었습니다.');opener.location.reload();location.href='/admin/momo/ranking/ranking_detail.asp?idx="&vIdx&"';</script>"
	Response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
