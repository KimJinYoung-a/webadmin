<%
'###########################################################
' Description :  인트라넷 개인정보 
' History : 2007.07.30 한용민 수정
'###########################################################

class CPartnerUser
	public fposit_name		
	public Fpart_name			
	public Fcompany_ip			
	public Fid
	public Fpassword
	public Fdiscountrate
	public Fcompany_name
	public Faddress
	public Ftel
	public Ffax
	public Fbigo
	public Furl
	public Fmanager_name
	public Fmanager_address
	public Fcommission
	public Femail
	public Fuserdiv
	public Fisusing
	public Fbuseo
	public Fpart
	public Fcposition
	public Fintro
	public Fmsn
	public Fbirthday
	public Fuserimg
	public FVatinclude
	public Fmaeipdiv
	public Fdefaultmargine
	public Fpid
	public Fcompany_no
	public Fzipcode
	public Fceoname
	public Fmanager_phone
	public Fmanager_hp
	public Fdeliver_name
	public Fdeliver_phone
	public Fdeliver_hp
	public Fdeliver_email
	public Fjungsan_name
	public Fjungsan_phone
	public Fjungsan_hp
	public Fjungsan_email
	public Fjungsan_gubun
	public Fjungsan_bank
	public Fjungsan_date
	public Fjungsan_acctname
	public Fjungsan_acctno
	public Fcompany_upjong
	public Fcompany_uptae
	public FGroupId
	public FSubId
	public Fppass
	public Fsocname
	public Fsocname_kor
	public Fisextusing
	public Fspecialbrand
	public FPrtIdx
	public Fstreetusing
	public Fextstreetusing
	public FTotalitemcount
	public Fsoclog
	public Ftitleimgurl
	public Fdgncomment
	public Fmduserid
	public Fregdate
	public fbirth_isSolar
	
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public Sub GetOnePartner(byval userid)
		dim sqlStr
		dim oneitem

		sqlStr = "select top 1 f.* , i.company_ip , p.part_name , o.posit_name from [db_partner].[dbo].tbl_partner f"		
		sqlstr = sqlstr & " left join [db_partner].[dbo].tbl_equipment_ip i on f.id = i.id"			
		sqlstr = sqlstr & " left join [db_partner].[dbo].tbl_partInfo p on f.part_sn = p.part_sn"		
		sqlstr = sqlstr & " left join [db_partner].[dbo].tbl_positInfo o on f.posit_sn = o.posit_sn"	
		sqlStr = sqlStr + " where f.id='" + CStr(userid) + "'"			

		rsget.Open sqlStr,dbget,1

		if Not rsget.Eof then
			Fid = rsget("id")
			Fpassword = rsget("password")
			Fdiscountrate = rsget("discountrate")
			Fcompany_name = rsget("company_name")
			Faddress = rsget("address")
			Ftel = rsget("tel")
			Fmanager_hp = rsget("manager_hp")
			Ffax = rsget("fax")
			Fbigo = rsget("bigo")
			Furl = rsget("url")
			Fmanager_name = rsget("manager_name")
			Fmanager_address = rsget("manager_address")
			Fcommission = rsget("commission")
			Femail = rsget("email")
			Fbirthday = rsget("birthday")
			Fmsn = rsget("msn")
			Fzipcode = rsget("zipcode")
			Fbuseo = rsget("buseo")
			Fpart = rsget("part")
			Fcposition = rsget("cposition")
			Fintro = rsget("intro")
			Fuserimg = rsget("userimg")
			Fuserdiv = rsget("userdiv")
			Fisusing = rsget("isusing")
			Fcompany_ip = rsget("company_ip")		
			fpart_name = rsget("part_name")			
			fposit_name = rsget("posit_name")
			fbirth_isSolar = rsget("birth_isSolar")
				
		end if
		rsget.close
	end Sub

end class

Class cip_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fid
	public fcompany_name
	public fpart_sn
	public fcompany_ip
	public findex_count
	public fpart_name
	public fgubuncd
	public fipidx
	public fisusing
end class

class cip_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	public frectcompany_ip
	public frectgubuncd
	public frectipidx
		
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

	'/admin/notice/ip_list.asp 
	public sub getip_list()
		dim sqlStr,i ,sqlsearch

		if frectgubuncd <> "" then
			sqlsearch = sqlsearch & " and gubuncd ='"&frectgubuncd&"'"
		end if
		
		'총 갯수 구하기
		sqlStr = "select" + vbcrlf  
		sqlStr = sqlStr & " count(*) as cnt" + vbcrlf 
		sqlStr = sqlStr & " from db_partner.dbo.tbl_equipment_ip ei"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partInfo pi"
		sqlStr = sqlStr & " 	on ei.part_sn = pi.part_sn"
		sqlStr = sqlStr & " where ei.isUsing='Y' " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf 
		sqlStr = sqlStr & " ei.id ,ei.company_name ,ei.part_sn ,ei.company_ip,ei.ipidx ,ei.isusing,pi.part_name" + vbcrlf 
		sqlStr = sqlStr & " from db_partner.dbo.tbl_equipment_ip ei"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partInfo pi"
		sqlStr = sqlStr & " 	on ei.part_sn = pi.part_sn"
		sqlStr = sqlStr & " where ei.isUsing='Y' " & sqlsearch
		sqlStr = sqlStr & " order by ei.company_ip asc" + vbcrlf 

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
				set FItemList(i) = new cip_oneitem
					
					fitemlist(i).fipidx = rsget("ipidx")
			        fitemlist(i).fid = rsget("id")
			        fitemlist(i).fcompany_name = rsget("company_name")
			        fitemlist(i).fpart_sn = rsget("part_sn")			     
			        fitemlist(i).fcompany_ip = rsget("company_ip")
			        fitemlist(i).fpart_name = db2html(rsget("part_name"))
			        fitemlist(i).fisusing = rsget("isusing")
			        
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
		
	'/admin/notice/ip_list.asp 
	public sub getip_edit()
		dim sqlStr,i , sqlsearch
		
		if frectgubuncd <> "" then
			sqlsearch = sqlsearch & " and gubuncd ='"&frectgubuncd&"'"
		end if
		if frectipidx <> "" then
			sqlsearch = sqlsearch & " and ipidx ="&frectipidx&""
		end if
		
		'데이터 리스트 
		sqlStr = "select top 1" + vbcrlf 
		sqlStr = sqlStr & " id ,company_name ,part_sn ,company_ip , gubuncd ,ipidx , isusing" + vbcrlf 
		sqlStr = sqlStr & " from db_partner.dbo.tbl_equipment_ip" + vbcrlf 
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		ftotalcount = rsget.recordcount

		i=0
		if  not rsget.EOF  then

			do until rsget.EOF
				set FOneItem = new cip_oneitem
					
					FOneItem.fisusing = rsget("isusing")
					FOneItem.fipidx = rsget("ipidx")
			        FOneItem.fid = rsget("id")
			        FOneItem.fcompany_name = rsget("company_name")
			        FOneItem.fpart_sn = rsget("part_sn")			     
			        FOneItem.fcompany_ip = rsget("company_ip")
					FOneItem.fgubuncd = rsget("gubuncd")
					
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

'셀렉트 옵션 생성 함수(장비구분, 사용구분)
Sub DrawEquipMentGubun(gubuntype,selectBoxName,selectedId , evtfg)		
   dim tmp_str,query1, qyery2									

	response.write "<select name='" & selectBoxName & "' "&evtfg&">"		
	response.write "<option value=''"							
		if selectedId="" then									
			response.write " selected"
		end if
	response.write ">선택</option>"								

	 '옵션 내용 DB에서 가져오기
   query1 = " select gubuncd,gubunname from [db_partner].[dbo].tbl_equipment_gubun where gubuntype='" + gubuntype + "'"
   query1 = query1 + " order by gubuncd"
   rsget.Open query1,dbget,1									

   if  not rsget.EOF  then										

       '도돌이 시작
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("gubuncd")) then 		
               tmp_str = " selected"					
           end if
           response.write("<option value='"&rsget("gubuncd")&"' "&tmp_str&">" + db2html(rsget("gubunname")) + "</option>")
           tmp_str = ""					
           rsget.MoveNext
       loop
   end if
   rsget.close

   '셀렉트 끝
   response.write("</select>")
End Sub
%>