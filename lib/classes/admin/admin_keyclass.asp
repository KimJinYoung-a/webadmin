<%
'###########################################################
' Description : 어드민 USB 인증
' History : 2008.06.30 한용민 생성 
'           2008.09.25 허진원 수정- Key Int→char 변경
'           2008.09.25 한용민 추가
'			2009.02.02 허진원 수정- 필드 정렬기능 추가
'###########################################################
%>
<%

Class ckey_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fkey_idx
	public fteamname
	public fusername
	public fusername_detail
	public fdel_isusing
	public fidx
end class

class ckey_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	public FrectSort
	public FrectTnm
	public FrectUnm
	public FrectUse
	
	public frectidx
	public frectkey_idx
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

''''''''''''// /admin/member/admin_keylist.asp 
	public sub getkey_list()
		dim sqlStr, addSql, i

		'검색쿼리
		addSql = " Where 1=1 "
		if FrectTnm<>"" then
			addSql = addSql & " and teamname='" & FrectTnm & "' "
		end if
		if FrectUnm<>"" then
			addSql = addSql & " and username like '%" & FrectUnm & "%' "
		end if
		if FrectUse<>"" then
			addSql = addSql & " and del_isusing='" & FrectUse & "' "
		end if

		'총 갯수 구하기
		sqlStr = "select" + vbcrlf  
		sqlStr = sqlStr & " count(key_idx) as cnt" + vbcrlf 
		sqlStr = sqlStr & " from db_partner.dbo.tbl_admin_key" + addSql + vbcrlf 
				
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf 
		sqlStr = sqlStr & " idx, key_idx ,teamname ,username ,username_detail , del_isusing" + vbcrlf 
		sqlStr = sqlStr & " from db_partner.dbo.tbl_admin_key" + addSql + vbcrlf 
		
		'정렬순 소트
		Select Case FrectSort
			Case "no"
				sqlStr = sqlStr & " order by idx asc" + vbcrlf 
			Case "key"
				sqlStr = sqlStr & " order by key_idx asc" + vbcrlf 
			Case "team"
				sqlStr = sqlStr & " order by teamname asc" + vbcrlf 
			Case "name"
				sqlStr = sqlStr & " order by username asc" + vbcrlf 
			Case Else
				sqlStr = sqlStr & " order by idx asc" + vbcrlf 
		End Select

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

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
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new ckey_oneitem

			        fitemlist(i).fidx = rsget("idx")				
			        fitemlist(i).fkey_idx = rsget("key_idx")
			        fitemlist(i).fteamname = rsget("teamname")
			        fitemlist(i).fusername = rsget("username")			     
			        fitemlist(i).fusername_detail = rsget("username_detail")			     
			        fitemlist(i).fdel_isusing = rsget("del_isusing")	
			        			        
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
		

''''''''''''// /admin/member/admin_keylist.asp 
	public sub getkey_edit()
		dim sqlStr,i

		'데이터 리스트 
		sqlStr = "select top 1" + vbcrlf 
		sqlStr = sqlStr & " idx ,key_idx ,teamname ,username ,username_detail ,del_isusing" + vbcrlf 
		sqlStr = sqlStr & " from db_partner.dbo.tbl_admin_key" + vbcrlf 
		sqlStr = sqlStr & " where idx = '"& frectidx &"'" + vbcrlf 

		'response.write sqlStr &"<br>"

		rsget.Open sqlStr,dbget,1
		ftotalcount = rsget.recordcount

		i=0
		if  not rsget.EOF  then

			do until rsget.EOF
				set FOneItem = new ckey_oneitem

			        FOneItem.fidx = rsget("idx")				
			        FOneItem.fkey_idx = rsget("key_idx")
			        FOneItem.fteamname = rsget("teamname")
			        FOneItem.fusername = rsget("username")			     
			        FOneItem.fusername_detail = rsget("username_detail")			     
			        FOneItem.fdel_isusing = rsget("del_isusing")
			        			        
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
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

end class	
%>	