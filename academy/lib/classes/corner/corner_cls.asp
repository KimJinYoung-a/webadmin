<%
'###########################################################
' Description : 코너관리
' History : 2009.09.10 한용민 생성
'###########################################################

Class clife_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public flife_id
	public flife_title
	public flist_image
	public fmain_image1
	public fmain_image2
	public fimg_map1
	public fimg_map2
	public fregdate
	public fisusing
	public fplusitem
	public fcommentyn
	public fidx
	public fuserid
	public fcomment	
	public ftitle
	public fgubun
	public fstartdate
	public fenddate
end class

class clife_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public foneitem
	
	public frectlife_id
	public frectisusing
	public frectidx
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	'///academy/corner/life_reg.asp
    public Sub flife_edit()
        dim sqlStr , i
        
		sqlStr = "select top 1 "+ vbcrlf
		sqlStr = sqlStr & " life_id,life_title,list_image,main_image1" + vbcrlf
		sqlStr = sqlStr & " ,main_image2,regdate,isusing,plusitem,commentyn" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_corner_life" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf			
		
		if frectlife_id <> "" then 
			sqlStr = sqlStr & " and life_id ='"&frectlife_id&"'" + vbcrlf 		 
		end if
	
		
        'response.write sqlStr&"<br>"
        rsACADEMYget.Open SqlStr, dbACADEMYget, 1
        ftotalcount = rsACADEMYget.RecordCount
        
        set FOneItem = new clife_oneitem
        
        if Not rsACADEMYget.Eof then


			foneitem.flife_id = rsacademyget("life_id")
			foneitem.flife_title = db2html(rsacademyget("life_title"))
			foneitem.flist_image = db2html(rsacademyget("list_image"))
			foneitem.fmain_image1 = db2html(rsacademyget("main_image1"))
			foneitem.fmain_image2 = db2html(rsacademyget("main_image2"))
			foneitem.fregdate = rsacademyget("regdate")
			foneitem.fisusing = rsacademyget("isusing")
 			foneitem.fplusitem = rsacademyget("plusitem")
			foneitem.fcommentyn = rsacademyget("commentyn")
			   						   
        end if
        rsACADEMYget.Close
    end Sub

	'///academy/corner/life_list.asp
	public sub flife_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(life_id) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_corner_life" + vbcrlf 
		sqlStr = sqlStr & " where 1=1 " + vbcrlf 
		
		if frectisusing = "Y" then 
			sqlStr = sqlStr & " and isusing ='Y'" + vbcrlf 
		elseif 	frectisusing = "N" then 
			sqlStr = sqlStr & " and isusing ='N'" + vbcrlf 		 
		end if
		if frectlife_id <> "" then 
			sqlStr = sqlStr & " and life_id ='"&frectlife_id&"'" + vbcrlf 		 
		end if
			
	
		rsacademyget.Open sqlStr,dbacademyget,1
			FTotalCount = rsacademyget("cnt")
		rsacademyget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " life_id,life_title,list_image,main_image1" + vbcrlf
		sqlStr = sqlStr & " ,main_image2,regdate,isusing,plusitem,commentyn" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_corner_life" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " + vbcrlf 
		
		if frectisusing = "Y" then 
			sqlStr = sqlStr & " and isusing ='Y'" + vbcrlf 
		elseif 	frectisusing = "N" then 
			sqlStr = sqlStr & " and isusing ='N'" + vbcrlf 		 
		end if
		if frectlife_id <> "" then 
			sqlStr = sqlStr & " and life_id ='"&frectlife_id&"'" + vbcrlf 		 
		end if
		
			
	
		sqlStr = sqlStr & " order by life_id desc" + vbcrlf

		'response.write sqlStr &"<br>"
		rsacademyget.pagesize = FPageSize
		rsacademyget.Open sqlStr,dbacademyget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsacademyget.EOF  then
			rsacademyget.absolutepage = FCurrPage
			do until rsacademyget.EOF
				set FItemList(i) = new clife_oneitem

				FItemList(i).flife_id = rsacademyget("life_id")
				FItemList(i).flife_title = db2html(rsacademyget("life_title"))
				FItemList(i).flist_image = db2html(rsacademyget("list_image"))
				FItemList(i).fmain_image1 = db2html(rsacademyget("main_image1"))
				FItemList(i).fmain_image2 = db2html(rsacademyget("main_image2"))
				FItemList(i).fregdate = rsacademyget("regdate")
				FItemList(i).fisusing = rsacademyget("isusing")
				FItemList(i).fplusitem = rsacademyget("plusitem")
				FItemList(i).fcommentyn = rsacademyget("commentyn")
												
				rsacademyget.movenext
				i=i+1
			loop
		end if
		rsacademyget.Close
	end sub
	
end Class	

Class crumour_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public frumour_id
	public frumour_title
	public frumour_userid
	public fstartdate
	public fenddate
	public flist_image
	public fmain_image1
	public fmain_image2
	public fregdate
	public fisusing

end class

class crumour_one_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public foneitem
	
	public frectrumour_id
	public frectisusing
	public frectidx
	public frectgubun
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	'///academy/corner/rumour_reg.asp
    public Sub frumour_edit()
        dim sqlStr , i
        
		sqlStr = "select top 1 "+ vbcrlf
		sqlStr = sqlStr & " idx, gubun, title, userid, startdate, enddate, list_image, main_image1 " + vbcrlf
		sqlStr = sqlStr & " , main_image2, regdate, isusing, plusitem, commentyn, img_map1, img_map2 " + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_fingers_story " + vbcrlf
		sqlStr = sqlStr & " where isusing='Y'" + vbcrlf			
		
		if frectidx <> "" then 
			sqlStr = sqlStr & " and idx = '" & frectidx & "'" + vbcrlf 		 
		end if
		
		If frectgubun <> "" Then
			sqlStr = sqlStr & " and gubun = '" & frectgubun & "'" + vbcrlf
		End IF
	
		sqlStr = sqlStr & " order by idx desc" + vbcrlf	
		
        'response.write sqlStr&"<br>"
        rsACADEMYget.Open SqlStr, dbACADEMYget, 1
        ftotalcount = rsACADEMYget.RecordCount
        
        set FOneItem = new clife_oneitem
        
        if Not rsACADEMYget.Eof then

			foneitem.fidx = rsACADEMYget("idx")
			foneitem.fgubun = rsACADEMYget("gubun")
			foneitem.ftitle = db2html(rsACADEMYget("title"))
			foneitem.fuserid = rsACADEMYget("userid")
			foneitem.fstartdate = rsACADEMYget("startdate")
			foneitem.fenddate = rsACADEMYget("enddate")
			foneitem.flist_image = db2html(rsACADEMYget("list_image"))
			foneitem.fmain_image1 = db2html(rsACADEMYget("main_image1"))
			foneitem.fmain_image2 = db2html(rsACADEMYget("main_image2"))
			foneitem.fimg_map1 = db2html(rsACADEMYget("img_map1"))
			foneitem.fimg_map2 = db2html(rsACADEMYget("img_map2"))
			foneitem.fregdate = rsACADEMYget("regdate")
			foneitem.fisusing = rsACADEMYget("isusing")
			foneitem.fplusitem = rsACADEMYget("plusitem")
			foneitem.fcommentyn = rsACADEMYget("commentyn")
			   						   
        end if
        rsACADEMYget.Close
    end Sub

	'///academy/corner/rumour_list.asp
	public sub frumour_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_fingers_story" + vbcrlf 
		sqlStr = sqlStr & " where 1=1 " + vbcrlf 

		If frectisusing <> "" Then
			sqlStr = sqlStr & " and isusing = '" & frectisusing & "' "
		End IF
		
		If frectrumour_id <> "" Then
			sqlStr = sqlStr & " and idx = '" & frectrumour_id & "' "
		End IF
		
		If frectgubun <> "" Then
			sqlStr = sqlStr & " and gubun = '" & frectgubun & "'" + vbcrlf
		End IF

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx, gubun, title, userid, startdate, enddate, list_image, main_image1 " + vbcrlf
		sqlStr = sqlStr & " , main_image2, regdate, isusing, plusitem, commentyn " + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_fingers_story " + vbcrlf
		sqlStr = sqlStr & " where 1=1 " + vbcrlf 
	
		If frectisusing <> "" Then
			sqlStr = sqlStr & " and isusing = '" & frectisusing & "' "
		End IF
		
		If frectrumour_id <> "" Then
			sqlStr = sqlStr & " and idx = '" & frectrumour_id & "' "
		End IF
		
		If frectgubun <> "" Then
			sqlStr = sqlStr & " and gubun = '" & frectgubun & "'" + vbcrlf
		End IF
	
		sqlStr = sqlStr & " order by idx desc" + vbcrlf

		'response.write sqlStr &"<br>"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.EOF
				set FItemList(i) = new clife_oneitem

				FItemList(i).fidx = rsACADEMYget("idx")
				FItemList(i).fgubun = rsACADEMYget("gubun")
				FItemList(i).ftitle = db2html(rsACADEMYget("title"))
				FItemList(i).fuserid = rsACADEMYget("userid")
				FItemList(i).fstartdate = rsACADEMYget("startdate")
				FItemList(i).fenddate = rsACADEMYget("enddate")
				FItemList(i).flist_image = db2html(rsACADEMYget("list_image"))
				FItemList(i).fmain_image1 = db2html(rsACADEMYget("main_image1"))
				FItemList(i).fmain_image2 = db2html(rsACADEMYget("main_image2"))
				FItemList(i).fregdate = rsACADEMYget("regdate")
				FItemList(i).fisusing = rsACADEMYget("isusing")
				FItemList(i).fplusitem = rsACADEMYget("plusitem")
				FItemList(i).fcommentyn = rsACADEMYget("commentyn")
												
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub
	
end Class	


Class cgood_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public flecturer_id
	public flecturer_name
	public fhistory
	public fhistory_act
	public fcatecd2
	public fsocname
	public fsocname_kor
	public fimage_profile
	public fimage_top
	public fregdate
	public fisusing
	public fhomepage
	public fCateCD2_Name
	public fidx
	public fimage_400x400
	public fimage_50x50
	public fimage_80x80
	public fimage_list
	public fimage_profile_75x75
	public fnewImage_profile
	public fbest
	public fitem_count
	public ftwitter
	public FOnesentence
	public FGubun
	public FCompany_name
	public FLec_yn
	public FDiy_yn

	Public FvideoUrl
	Public FvideoWidth
	Public FvideoHeight
	Public Fvideogubun
	Public FvideoType
	Public FvideoFullUrl

	Public Function getMyJob()
		If (Flec_yn = "Y" AND FDiy_yn = "Y") OR (Flec_yn = "N" AND FDiy_yn = "Y") Then
			getMyJob = "D"
		ElseIf Flec_yn = "Y" AND FDiy_yn = "N" Then
			getMyJob = "L"
		Else
			getMyJob = ""
		End If
	End Function

end class

class cgood_onelist
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public foneitem
	
	public frectlecturer_id
	public frectisusing
	public frectidx

	Public FRectArtistid
	Public FRectItemVideoGubun
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	'///academy/corner/good_item_reg.asp
    public Sub fgood_item_edit()
        dim sqlStr
		sqlStr = "select top 1 "+ vbcrlf
		sqlStr = sqlStr & " idx,lecturer_id,image_400x400,image_50x50,image_80x80,regdate,isusing" + vbcrlf	
		sqlStr = sqlStr & " from db_academy.dbo.tbl_corner_good_item" + vbcrlf	
		sqlStr = sqlStr & " where 1=1" + vbcrlf			
		
		if frectlecturer_id <> "" then
			sqlStr = sqlStr & " and lecturer_id= '"&frectlecturer_id&"'" + vbcrlf				
		end if
		if frectidx <> "" then
			sqlStr = sqlStr & " and idx= '"&frectidx&"'" + vbcrlf				
		end if
			
	
        'response.write sqlStr&"<br>"
        rsACADEMYget.Open SqlStr, dbACADEMYget, 1
        ftotalcount = rsACADEMYget.RecordCount
        
        set FOneItem = new cgood_oneitem
        
        if Not rsACADEMYget.Eof then
    																	
			foneitem.fidx = rsacademyget("idx")
			foneitem.flecturer_id = db2html(rsacademyget("lecturer_id"))
			foneitem.fimage_400x400 = db2html(rsacademyget("image_400x400"))
			foneitem.fimage_50x50 = db2html(rsacademyget("image_50x50"))
			foneitem.fimage_80x80 = db2html(rsacademyget("image_80x80"))
			foneitem.fregdate = rsacademyget("regdate")
			foneitem.fisusing = rsacademyget("isusing")												   
        end if
        rsACADEMYget.Close
    end Sub

	'///academy/corner/good_reg.asp
    public Sub fgood_edit()
        dim sqlStr
		sqlStr = "select top 1 "+ vbcrlf
		sqlStr = sqlStr & " lecturer_id, lecturer_name, history, history_act, catecd2" + vbcrlf	
		sqlStr = sqlStr & " , socname, socname_kor, image_profile, image_top, regdate" + vbcrlf	
		sqlStr = sqlStr & " , isusing, homepage , image_list , image_profile_75x75 , best, twitter, onesentence, newImage_profile" + vbcrlf	
		sqlStr = sqlStr & " from db_academy.dbo.tbl_corner_good" + vbcrlf	
		sqlStr = sqlStr & " where 1=1" + vbcrlf			
		
		if frectlecturer_id <> "" then
			sqlStr = sqlStr & " and lecturer_id= '"&frectlecturer_id&"'" + vbcrlf				
		end if
		
        'response.write sqlStr&"<br>"
        rsACADEMYget.Open SqlStr, dbACADEMYget, 1
        ftotalcount = rsACADEMYget.RecordCount
        
        set FOneItem = new cgood_oneitem
        
        if Not rsACADEMYget.Eof then
    			
    		foneitem.fbest = rsacademyget("best")	
			foneitem.flecturer_id = rsacademyget("lecturer_id")
			foneitem.flecturer_name = db2html(rsacademyget("lecturer_name"))
			foneitem.fhistory = db2html(rsacademyget("history"))
			foneitem.fhistory_act = db2html(rsacademyget("history_act"))															
			foneitem.fcatecd2 = rsacademyget("catecd2")
			foneitem.fsocname = db2html(rsacademyget("socname"))
			foneitem.fsocname_kor = db2html(rsacademyget("socname_kor"))
			foneitem.fimage_profile = db2html(rsacademyget("image_profile"))
			foneitem.fimage_profile_75x75 = db2html(rsacademyget("image_profile_75x75"))
			foneitem.fnewImage_profile = db2html(rsacademyget("newImage_profile"))
			foneitem.fimage_top = db2html(rsacademyget("image_top"))
			foneitem.fimage_list = db2html(rsacademyget("image_list"))
			foneitem.fregdate = rsacademyget("regdate")
			foneitem.fisusing = rsacademyget("isusing")
			foneitem.fhomepage = db2html(rsacademyget("homepage"))
			foneitem.ftwitter = db2html(rsacademyget("twitter"))
			foneitem.FOnesentence = db2html(rsacademyget("onesentence"))
			   
        end if
        rsACADEMYget.Close
    end Sub

    Public Sub FGood_myInfo
		Dim sqlStr
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 "
		sqlStr = sqlStr & "	p.id, p.company_name, c.socname, c.socname_kor, U.lec_yn, U.diy_yn "
		sqlStr = sqlStr & " FROM [db_user].[dbo].tbl_user_c c "
		sqlStr = sqlStr & " LEFT JOIN [db_partner].[dbo].tbl_partner p on c.userid = p.id  "
		sqlStr = sqlStr & " LEFT JOIN [ACADEMYDB].[db_academy].[dbo].tbl_lec_user U on c.userid=U.lecturer_id  "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & " and p.userdiv='9999'  "
		sqlStr = sqlStr & " and c.userdiv='14' "
		sqlStr = sqlStr & " and c.isusing='Y' "
		sqlStr = sqlStr & " and isnull(U.lec_yn, '') <> '' "
		sqlStr = sqlStr & " and isnull(U.diy_yn, '') <> '' "
		sqlStr = sqlStr & " and lecturer_id = '"&session("ssBctId")&"' "
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        set FOneItem = new cgood_oneitem
        
        if Not rsget.Eof then
			foneitem.FCompany_name	= db2html(rsget("company_name"))
			foneitem.FSocname		= db2html(rsget("socname"))
			foneitem.FSocname_kor	= db2html(rsget("socname_kor"))
			foneitem.FLec_yn		= rsget("lec_yn")
			foneitem.FDiy_yn		= rsget("diy_yn")
        end if
        rsget.Close
	End Sub

	'2016-08-17 김진영 재작성
	public sub getProfileList()
		Dim sqlStr, i, addSql

		If FRectisusing = "Y" Then 
			addSql = addSql & " and isusing ='"&FRectisusing&"'"
		End if
		
		If FRectlecturer_id <> "" then 
			addSql = addSql & " and lecturer_id ='"&FRectlecturer_id&"'"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(lecturer_id) as cnt, CEILING(CAST(Count(lecturer_id) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_academy.dbo.tbl_corner_good "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & addSql
		rsacademyget.Open sqlStr, dbacademyget, 1
			FTotalCount = rsacademyget("cnt")
			FTotalPage	= rsacademyget("totPg")
		rsacademyget.Close

		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If
		
		'데이터 리스트 
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " a.lecturer_id, a.lecturer_name, a.history, a.history_act, a.catecd2"	
		sqlStr = sqlStr & " , a.socname, a.socname_kor, a.image_profile, a.image_top, a.regdate, a.newImage_profile "	
		sqlStr = sqlStr & " , a.isusing, a.homepage , a.best"	
		sqlStr = sqlStr & " ,(select count(lecturer_id) as item_count from db_academy.dbo.tbl_corner_good_item where lecturer_id = a.lecturer_id and isusing='Y') as item_count"		
		sqlStr = sqlStr & " FROM db_academy.dbo.tbl_corner_good as a"	
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY a.regdate DESC"
		rsacademyget.pagesize = FPageSize
		rsacademyget.Open sqlStr,dbacademyget,1
		FResultCount = rsacademyget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsacademyget.EOF Then
			rsacademyget.absolutepage = FCurrPage
			Do until rsacademyget.EOF
				set FItemList(i) = new cgood_oneitem
					FItemList(i).FItem_count		= rsacademyget("item_count")
					FItemList(i).FLecturer_id		= rsacademyget("lecturer_id")
					FItemList(i).FLecturer_name		= db2html(rsacademyget("lecturer_name"))
					FItemList(i).FHistory			= db2html(rsacademyget("history"))
					FItemList(i).FHistory_act		= db2html(rsacademyget("history_act"))																
					FItemList(i).FSocname			= db2html(rsacademyget("socname"))
					FItemList(i).FSocname_kor		= db2html(rsacademyget("socname_kor"))
				If db2html(rsacademyget("newImage_profile")) <> "" Then
					FItemList(i).FNewImage_profile = imgFingers & "/corner/newImage_profile/thumbimg3/t3_" & db2html(rsacademyget("newImage_profile"))
				Else
					FItemList(i).FNewImage_profile = ""
				End If
					FItemList(i).FRegdate			= rsacademyget("regdate")
					FItemList(i).Fisusing			= rsacademyget("isusing")
					FItemList(i).FHomepage			= db2html(rsacademyget("homepage"))
				rsacademyget.movenext
				i = i + 1
			Loop
		End If
		rsacademyget.Close
	End Sub

	'///academy/corner/good_list.asp
	public sub fgood_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(a.lecturer_id) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_corner_good as a" + vbcrlf 
		sqlStr = sqlStr & " where 1=1 " + vbcrlf 
		
		if frectisusing = "Y" then 
			sqlStr = sqlStr & " and a.isusing ='Y'" + vbcrlf 
		elseif 	frectisusing = "N" then 
			sqlStr = sqlStr & " and a.isusing ='N'" + vbcrlf 		 
		end if
		if frectlecturer_id <> "" then 
			sqlStr = sqlStr & " and a.lecturer_id ='"&frectlecturer_id&"'" + vbcrlf 		 
		end if
			
	
		rsacademyget.Open sqlStr,dbacademyget,1
			FTotalCount = rsacademyget("cnt")
		rsacademyget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " a.lecturer_id, a.lecturer_name, a.history, a.history_act, a.catecd2" + vbcrlf	
		sqlStr = sqlStr & " , a.socname, a.socname_kor, a.image_profile, a.image_top, a.regdate, a.newImage_profile " + vbcrlf	
		sqlStr = sqlStr & " , a.isusing, a.homepage , a.best" + vbcrlf	
		sqlStr = sqlStr & " , (select top 1 CateCD2_Name from db_academy.dbo.tbl_lec_Cate2 " + vbcrlf
		sqlStr = sqlStr & " where CateCD2 = a.CateCD2) as CateCD2_Name" + vbcrlf	
		sqlStr = sqlStr & " ,(select count(lecturer_id) as item_count from db_academy.dbo.tbl_corner_good_item" + vbcrlf
		sqlStr = sqlStr & " where lecturer_id = a.lecturer_id and isusing='Y') as item_count" + vbcrlf		
		sqlStr = sqlStr & " from db_academy.dbo.tbl_corner_good as a" + vbcrlf	
		sqlStr = sqlStr & " where 1=1 " + vbcrlf 
		
		if frectisusing = "Y" then 
			sqlStr = sqlStr & " and a.isusing ='Y'" + vbcrlf 
		elseif frectisusing = "N" then 
			sqlStr = sqlStr & " and a.isusing ='N'" + vbcrlf 		 
		end if
		if frectlecturer_id <> "" then 
			sqlStr = sqlStr & " and a.lecturer_id ='"&frectlecturer_id&"'" + vbcrlf 		 
		end if
		
			
	
		sqlStr = sqlStr & " order by a.regdate desc" + vbcrlf	

		'response.write sqlStr &"<br>"
		rsacademyget.pagesize = FPageSize
		rsacademyget.Open sqlStr,dbacademyget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsacademyget.EOF  then
			rsacademyget.absolutepage = FCurrPage
			do until rsacademyget.EOF
				set FItemList(i) = new cgood_oneitem
				
				FItemList(i).fitem_count = rsacademyget("item_count")
				FItemList(i).fbest = rsacademyget("best")
				FItemList(i).fCateCD2_Name = rsacademyget("CateCD2_Name")
				FItemList(i).flecturer_id = rsacademyget("lecturer_id")
				FItemList(i).flecturer_name = db2html(rsacademyget("lecturer_name"))
				FItemList(i).fhistory = db2html(rsacademyget("history"))
				FItemList(i).fhistory_act = db2html(rsacademyget("history_act"))																
				FItemList(i).fcatecd2 = rsacademyget("catecd2")
				FItemList(i).fsocname = db2html(rsacademyget("socname"))
				FItemList(i).fsocname_kor = db2html(rsacademyget("socname_kor"))
				FItemList(i).fimage_profile = db2html(rsacademyget("image_profile"))
				FItemList(i).fnewImage_profile = db2html(rsacademyget("newImage_profile"))
				FItemList(i).fimage_top = db2html(rsacademyget("image_top"))
				FItemList(i).fregdate = rsacademyget("regdate")
				FItemList(i).fisusing = rsacademyget("isusing")
				FItemList(i).fhomepage = db2html(rsacademyget("homepage"))
								
				rsacademyget.movenext
				i=i+1
			loop
		end if
		rsacademyget.Close
	end sub

	'///academy/corner/good_item_list.asp
	public sub fgood_item_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_corner_good_item" + vbcrlf 
		sqlStr = sqlStr & " where 1=1 " + vbcrlf 
		
		if frectisusing = "Y" then 
			sqlStr = sqlStr & " and isusing ='Y'" + vbcrlf 
		elseif 	frectisusing = "N" then 
			sqlStr = sqlStr & " and isusing ='N'" + vbcrlf 		 
		end if
		if frectlecturer_id <> "" then 
			sqlStr = sqlStr & " and lecturer_id ='"&frectlecturer_id&"'" + vbcrlf 		 
		end if
			
	
		rsacademyget.Open sqlStr,dbacademyget,1
			FTotalCount = rsacademyget("cnt")
		rsacademyget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx, lecturer_id, image_400x400, image_50x50, regdate, isusing" + vbcrlf		
		sqlStr = sqlStr & " from db_academy.dbo.tbl_corner_good_item" + vbcrlf	
		sqlStr = sqlStr & " where 1=1 " + vbcrlf 
		
		if frectisusing = "Y" then 
			sqlStr = sqlStr & " and isusing ='Y'" + vbcrlf 
		elseif 	frectisusing = "N" then 
			sqlStr = sqlStr & " and isusing ='N'" + vbcrlf 		 
		end if
		if frectlecturer_id <> "" then 
			sqlStr = sqlStr & " and lecturer_id ='"&frectlecturer_id&"'" + vbcrlf 		 
		end if
	
	
		sqlStr = sqlStr & " order by idx desc" + vbcrlf	

		'response.write sqlStr &"<br>"
		rsacademyget.pagesize = FPageSize
		rsacademyget.Open sqlStr,dbacademyget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsacademyget.EOF  then
			rsacademyget.absolutepage = FCurrPage
			do until rsacademyget.EOF
				set FItemList(i) = new cgood_oneitem

				FItemList(i).fidx = rsacademyget("idx")
				FItemList(i).flecturer_id = rsacademyget("lecturer_id")
				FItemList(i).fimage_400x400 = db2html(rsacademyget("image_400x400"))
				FItemList(i).fimage_50x50 = db2html(rsacademyget("image_50x50"))
				FItemList(i).fregdate = rsacademyget("regdate")																
				FItemList(i).fisusing = rsacademyget("isusing")
								
				rsacademyget.movenext
				i=i+1
			loop
		end if
		rsacademyget.Close
	end Sub
	
	'// 작가 강사 안내 동영상 수정(2017-03-03 이종화)
	public Sub GetArtistProfileVideo()
		dim sqlstr,i
		sqlstr = "select top 1 videogubun, videotype, videourl, videowidth, videoheight, videofullurl "
		sqlstr = sqlstr & " from db_academy.dbo.[tbl_corner_videos] "
		sqlstr = sqlstr & " where artistid='"& FRectArtistid &"'"
        sqlstr = sqlstr & " and videogubun='" & Trim(FRectItemVideoGubun) & "'"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount

		if Not rsACADEMYget.Eof then
			set FOneItem = new cgood_oneitem
			FOneItem.FvideoUrl     = rsACADEMYget("videourl")
			FOneItem.FvideoWidth     = rsACADEMYget("videowidth")
			FOneItem.FvideoHeight     = rsACADEMYget("videoheight")
			FOneItem.Fvideogubun     = rsACADEMYget("videogubun")
			FOneItem.FvideoType     = rsACADEMYget("videotype")
			FOneItem.FvideoFullUrl     = rsACADEMYget("videofullurl")
		Else
			set FOneItem = new cgood_oneitem
			FOneItem.FvideoUrl     = ""
			FOneItem.FvideoWidth     = ""
			FOneItem.FvideoHeight     = ""
			FOneItem.Fvideogubun     = ""
			FOneItem.FvideoType     = ""
			FOneItem.FvideoFullUrl     = ""
		end if
		rsACADEMYget.Close

	end Sub
	
end Class	

Class cposcode_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx
	public fposcode
	public fposname
	public fimagetype
	public fimagewidth
	public fimageheight
	public fisusing
	public fimagepath
	public flinkpath
	public fevt_code
	public fregdate
	public fimagecount
	public fimage_order
	public fitemid

end class

class cposcode_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	
	public FRectPoscode
	public FRectIsusing
	public FRectvaliddate
	public FRectIdx
	public frecttoplimit

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	'//academy/corner/imagemake_list.asp
	public sub fcontents_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(a.idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_corner_poscode_image a" & vbcrlf
		sqlStr = sqlStr & " left join db_academy.dbo.tbl_corner_poscode b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf	
        sqlStr = sqlStr & " where 1=1" & vbcrlf

			if FRectIsusing <> "" then
				sqlStr = sqlStr & " and a.isusing = '"& FRectIsusing &"'" & vbcrlf		
			end if	

			if FRectPosCode <> "" then
				sqlStr = sqlStr & " and a.poscode = "& FRectPosCode &"" & vbcrlf		
			end if	
				
		
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " b.posname,b.imagetype,b.imagewidth,b.imageheight,b.imagecount" & vbcrlf
		sqlStr = sqlStr & " ,a.idx,a.imagepath,a.linkpath,a.regdate,a.poscode,a.isusing,a.image_order" & vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_corner_poscode_image a" & vbcrlf
		sqlStr = sqlStr & " left join db_academy.dbo.tbl_corner_poscode b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf	
        sqlStr = sqlStr & " where 1=1" & vbcrlf

			if FRectIsusing <> "" then
				sqlStr = sqlStr & " and a.isusing = '"&FRectIsusing&"'" & vbcrlf		
			end if	
			if FRectPosCode <> "" then
				sqlStr = sqlStr & " and a.poscode = "& FRectPosCode &"" & vbcrlf		
			end if	

		sqlStr = sqlStr & " order by a.idx Desc" + vbcrlf

		'response.write sqlStr &"<br>"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.EOF
				set FItemList(i) = new cposcode_oneitem
				
				FItemList(i).fposcode = rsACADEMYget("poscode")
				FItemList(i).fposname = db2html(rsACADEMYget("posname"))
				FItemList(i).fimagetype = rsACADEMYget("imagetype")
				FItemList(i).fimagewidth = rsACADEMYget("imagewidth")
				FItemList(i).fimageheight = rsACADEMYget("imageheight")
				FItemList(i).fisusing = rsACADEMYget("isusing")
				FItemList(i).fidx = rsACADEMYget("idx")
				FItemList(i).fimagepath = rsACADEMYget("imagepath")
				FItemList(i).flinkpath = rsACADEMYget("linkpath")
				FItemList(i).fregdate = rsACADEMYget("regdate")		
				FItemList(i).fimagecount = rsACADEMYget("imagecount")
				FItemList(i).fimage_order = rsACADEMYget("image_order")													
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub

'//academy/corner/imagemake_contents.asp
    public Sub fcontents_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " a.posname,a.imagetype,a.imagewidth,a.imageheight,a.imagecount" & vbcrlf
		sqlStr = sqlStr & " ,b.idx,b.imagepath,b.linkpath,b.regdate,b.poscode,b.isusing,b.image_order" & vbcrlf		
		sqlStr = sqlStr & " from db_academy.dbo.tbl_corner_poscode a" & vbcrlf
		sqlStr = sqlStr & " left join db_academy.dbo.tbl_corner_poscode_image b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf	
        sqlStr = sqlStr & " where 1=1" & vbcrlf
        sqlStr = sqlStr & " and b.idx = "& FRectIdx&""

        'response.write sqlStr&"<br>"
        rsACADEMYget.Open SqlStr, dbACADEMYget, 1
        FResultCount = rsACADEMYget.RecordCount
        
        set FOneItem = new cposcode_oneitem
        
        if Not rsACADEMYget.Eof then
    
			FOneItem.fposcode = rsACADEMYget("poscode")
			FOneItem.fposname = db2html(rsACADEMYget("posname"))
			FOneItem.fimagetype = rsACADEMYget("imagetype")
			FOneItem.fimagewidth = rsACADEMYget("imagewidth")
			FOneItem.fimageheight = rsACADEMYget("imageheight")
			FOneItem.fisusing = rsACADEMYget("isusing")
			FOneItem.fidx = rsACADEMYget("idx")
			FOneItem.fimagepath = db2html(rsACADEMYget("imagepath"))
			FOneItem.flinkpath = db2html(rsACADEMYget("linkpath"))
			FOneItem.fregdate = rsACADEMYget("regdate")
			FOneItem.fimagecount = rsACADEMYget("imagecount") 
			FOneItem.fimage_order = rsACADEMYget("image_order") 
			           
        end if
        rsACADEMYget.Close
    end Sub
	
	'////academy/corner/imagemake_poscode.asp
    public Sub fposcode_oneitem()		
        dim SqlStr
        SqlStr = "select" + vbcrlf
		sqlStr = sqlStr & " poscode,posname,imagetype,imagewidth,imageheight,isusing,imagecount" + vbcrlf        
		sqlStr = sqlStr & " from db_academy.dbo.tbl_corner_poscode" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf
        SqlStr = SqlStr + " and poscode=" + CStr(FRectPoscode)
         
        rsACADEMYget.Open SqlStr, dbACADEMYget, 1
        FResultCount = rsACADEMYget.RecordCount
        
        set FOneItem = new cposcode_oneitem
        if Not rsACADEMYget.Eof then
            
            FOneItem.fposcode = rsACADEMYget("poscode")
            FOneItem.fposname = db2html(rsACADEMYget("posname"))
            FOneItem.fimagetype	= rsACADEMYget("imagetype")
            FOneItem.fimagewidth = rsACADEMYget("imagewidth")
            FOneItem.fimageheight = rsACADEMYget("imageheight")
            FOneItem.fisusing = rsACADEMYget("isusing")
            FOneItem.fimagecount = rsACADEMYget("imagecount")
                       
        end if
        rsACADEMYget.close
    end Sub

	'///academy/corner/imagemake_poscode.asp
	public sub fposcode_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select" + vbcrlf
		sqlStr = sqlStr & " count(poscode) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_corner_poscode" + vbcrlf
					
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " poscode,isusing,posname,imagetype,imagewidth,imageheight,imagecount" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_corner_poscode" + vbcrlf			
		sqlStr = sqlStr & " where 1=1" + vbcrlf

		'response.write sqlStr &"<br>"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.EOF
				set FItemList(i) = new cposcode_oneitem
				
				FItemList(i).fposcode = rsACADEMYget("poscode")
				FItemList(i).fposname = db2html(rsACADEMYget("posname"))
				FItemList(i).fimagetype = rsACADEMYget("imagetype")
				FItemList(i).fimagewidth = rsACADEMYget("imagewidth")
				FItemList(i).fimageheight = rsACADEMYget("imageheight")
				FItemList(i).fisusing = rsACADEMYget("isusing")
				FItemList(i).fimagecount = rsACADEMYget("imagecount")
														
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub
		
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class

'//적용구분 
function DrawMainPosCodeCombo(selectBoxName,selectedId,changeFlag)
   dim tmp_str,query1
   %>
   <select name="<%=selectBoxName%>" <%= changeFlag %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
   query1 = " select poscode,posname from db_academy.dbo.tbl_corner_poscode"
   rsACADEMYget.Open query1,dbACADEMYget,1

   if  not rsACADEMYget.EOF  then
       do until rsACADEMYget.EOF
           if Lcase(selectedId) = Lcase(rsACADEMYget("poscode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsACADEMYget("poscode")&"' "&tmp_str&">" + db2html(rsACADEMYget("posname")) + "</option>")
           tmp_str = ""
           rsACADEMYget.MoveNext
       loop
   end if
   rsACADEMYget.close
   response.write("</select>")
end function

public Sub SelectLecturerId(byval lecturer_id)
	dim sqlStr,i
	sqlStr = "select  c.userid,p.company_name,c.socname, c.socname_kor"
	sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
	sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on c.userid=p.id"
	sqlStr = sqlStr + " where c.userid<>''" + vbCrlf
	sqlStr = sqlStr + " and c.userdiv < 22" + vbcrlf
	sqlStr = sqlStr + " and c.userdiv='14'" + vbcrlf

	rsget.open sqlStr,dbget,1

	if not rsget.eof then
			response.write "<select name='temp_lec_id' onchange='javascript:FnLecturerApp(this.value);'>"
			response.write "<option value=''>선택</option>"
		for i=0 to rsget.recordcount-1
			if lecturer_id=db2html(rsget("userid")) then
			response.write "<option value='" & db2html(rsget("userid")) & "," & db2html(rsget("company_name")) & "," & rsget("socname") & "," & left(rsget("socname_kor"),10) & "' selected>" & db2html(rsget("userid")) & "(" & db2html(rsget("company_name")) & ")</option>"
			else
			response.write "<option value='" & db2html(rsget("userid")) & "," & db2html(rsget("company_name")) & "," & rsget("socname") & "," & left(rsget("socname_kor"),10) & "'>" & db2html(rsget("userid")) & "(" & db2html(rsget("company_name")) & ")</option>"
			end if
		rsget.movenext
		next
			response.write "</select>"
	end if
	rsget.close

end sub
%>