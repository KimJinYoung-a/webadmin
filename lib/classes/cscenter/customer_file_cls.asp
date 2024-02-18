<%
'###########################################################
' Description : 고객파일전송관리 클래스
' History : 2019.11.25 한용민 생성
'###########################################################

Class ccsfileitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fauthidx
	public fuserhp
	public fuserid
	public forderserial
	public fcomment
	public ffileurl1
	public ffileurl2
	public ffileurl3
	public ffileurl4
	public ffileurl5
	public fsmsyn
	public fkakaotalkyn
	public fcertno
	public fisusing
	public fregdate
    public fstatus
    public fadminid
    public fcustomer_file_regdate
    public fasmasteridx
end class

class ccsfilelist
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem

    public frectauthidx
    public frectuserhp
    public frectuserid
    public frectorderserial
	public frectasmasteridx
	public frectstatus
	public frectisusing

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public sub getordermasterinfo()
		dim sqlStr,i , sqlsearch
		
		if frectorderserial <> "" then
			sqlsearch = sqlsearch & " and orderserial ='"& frectorderserial &"'"
		end if

		'데이터 리스트 
		sqlStr = "select top 1" & vbcrlf 
		sqlStr = sqlStr & " orderserial, userid, buyhp as userhp" & vbcrlf 
		sqlStr = sqlStr & " from db_order.dbo.tbl_order_master with (readuncommitted)" & vbcrlf 
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		ftotalcount = rsget.recordcount
		FResultCount = rsget.recordcount

		i=0
		if  not rsget.EOF  then

			do until rsget.EOF
				set FOneItem = new ccsfileitem
					FOneItem.forderserial = rsget("orderserial")
					FOneItem.fuserid = rsget("userid")
					FOneItem.fuserhp = db2html(rsget("userhp"))
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public sub getuserinfo()
		dim sqlStr,i , sqlsearch
		
		if frectuserid <> "" then
			sqlsearch = sqlsearch & " and userid ='"& frectuserid &"'"
		end if

		'데이터 리스트 
		sqlStr = "select top 1" & vbcrlf 
		sqlStr = sqlStr & " userid, usercell as userhp" & vbcrlf 
		sqlStr = sqlStr & " from db_user.dbo.tbl_user_n with (readuncommitted)" & vbcrlf 
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		ftotalcount = rsget.recordcount
		FResultCount = rsget.recordcount

		i=0
		if  not rsget.EOF  then

			do until rsget.EOF
				set FOneItem = new ccsfileitem
					FOneItem.fuserid = rsget("userid")
					FOneItem.fuserhp = db2html(rsget("userhp"))
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public sub getcsasinfo()
		dim sqlStr,i , sqlsearch
		
		if frectasmasteridx <> "" then
			sqlsearch = sqlsearch & " and id ='"& frectasmasteridx &"'"
		end if

		'데이터 리스트 
		sqlStr = "select top 1" & vbcrlf 
		sqlStr = sqlStr & " l.orderserial, l.userid, u.usercell as userhp, l.id as asmasteridx" & vbcrlf 
		sqlStr = sqlStr & " from db_cs.dbo.tbl_new_as_list l with (readuncommitted)" & vbcrlf 
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_n u with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on l.userid=u.userid" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		ftotalcount = rsget.recordcount
		FResultCount = rsget.recordcount

		i=0
		if  not rsget.EOF  then

			do until rsget.EOF
				set FOneItem = new ccsfileitem
					FOneItem.forderserial = rsget("orderserial")
					FOneItem.fuserid = rsget("userid")
					FOneItem.fuserhp = db2html(rsget("userhp"))
					FOneItem.fasmasteridx = db2html(rsget("asmasteridx"))
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

    ' /cscenter/action/pop_cs_file_send.asp
	public sub getcsfile_one()
		dim sqlStr,i , sqlsearch
		
		if frectauthidx <> "" then
			sqlsearch = sqlsearch & " and authidx ="& frectauthidx &""
		end if
		if frectasmasteridx <> "" then
			sqlsearch = sqlsearch & " and asmasteridx ="& frectasmasteridx &""
		end if
		if frectstatus <> "" then
			sqlsearch = sqlsearch & " and status ="& frectstatus &""
		end if
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and isusing ='"& frectisusing &"'"
		end if

		'데이터 리스트 
		sqlStr = "select top 1" & vbcrlf 
		sqlStr = sqlStr & " authidx,userhp,userid,orderserial,comment,fileurl1,fileurl2,fileurl3,fileurl4" & vbcrlf 
		sqlStr = sqlStr & " ,fileurl5,smsyn,kakaotalkyn,certno,status,isusing,regdate, adminid, customer_file_regdate, asmasteridx" & vbcrlf 
		sqlStr = sqlStr & " from db_cs.dbo.tbl_customer_filelist with (readuncommitted)" & vbcrlf 
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		ftotalcount = rsget.recordcount
		FResultCount = rsget.recordcount

		i=0
		if  not rsget.EOF  then

			do until rsget.EOF
				set FOneItem = new ccsfileitem

					FOneItem.fauthidx = rsget("authidx")
					FOneItem.fuserhp = rsget("userhp")
					FOneItem.fuserid = rsget("userid")
					FOneItem.forderserial = rsget("orderserial")
					FOneItem.fcomment = db2html(rsget("comment"))
					FOneItem.ffileurl1 = rsget("fileurl1")
					FOneItem.ffileurl2 = rsget("fileurl2")
					FOneItem.ffileurl3 = rsget("fileurl3")
					FOneItem.ffileurl4 = rsget("fileurl4")
					FOneItem.ffileurl5 = rsget("fileurl5")
					FOneItem.fsmsyn = rsget("smsyn")
					FOneItem.fkakaotalkyn = rsget("kakaotalkyn")
					FOneItem.fcertno = rsget("certno")
                    FOneItem.fstatus = rsget("status")
					FOneItem.fisusing = rsget("isusing")
					FOneItem.fregdate = rsget("regdate")
                    FOneItem.fadminid = rsget("adminid")
					FOneItem.fcustomer_file_regdate = rsget("customer_file_regdate")
					FOneItem.fasmasteridx = rsget("asmasteridx")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public sub getcsfile()
		dim sqlStr,i ,sqlsearch

		if frectuserid <> "" or frectorderserial <> "" or frectasmasteridx <> "" then 
			if frectuserid <> "" then
				if frectuserhp <> "" then
					sqlsearch = sqlsearch & " and (replace(userhp,'-','') ='"& replace(frectuserhp,"-","") &"' or userid ='"& frectuserid &"')"
				else
					sqlsearch = sqlsearch & " and userid ='"& frectuserid &"'"
				end if
			end if
			if frectorderserial <> "" then
				if frectuserhp <> "" then
					sqlsearch = sqlsearch & " and (replace(userhp,'-','') ='"& replace(frectuserhp,"-","") &"' or orderserial ='"& frectorderserial &"')"
				else
					sqlsearch = sqlsearch & " and orderserial ='"& frectorderserial &"'"
				end if
			end if
			if frectasmasteridx <> "" then
				if frectuserhp <> "" then
					sqlsearch = sqlsearch & " and (replace(userhp,'-','') ='"& replace(frectuserhp,"-","") &"' or asmasteridx ="& frectasmasteridx &")"
				else
					sqlsearch = sqlsearch & " and asmasteridx ="& frectasmasteridx &""
				end if
			end if
		else
			if frectuserhp <> "" then
				sqlsearch = sqlsearch & " and replace(userhp,'-','') ='"& replace(frectuserhp,"-","") &"'"
			end if
		end if
		if frectstatus <> "" then
			sqlsearch = sqlsearch & " and status ="& frectstatus &""
		end if
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and isusing ='"& frectisusing &"'"
		end if
		
		'총 갯수 구하기
		sqlStr = "select" + vbcrlf  
		sqlStr = sqlStr & " count(*) as cnt" + vbcrlf 
		sqlStr = sqlStr & " from db_cs.dbo.tbl_customer_filelist with (readuncommitted)" & vbcrlf 
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

        if FTotalCount<1 then exit sub

		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf 
		sqlStr = sqlStr & " authidx,userhp,userid,orderserial,comment,fileurl1,fileurl2,fileurl3,fileurl4" & vbcrlf 
		sqlStr = sqlStr & " ,fileurl5,smsyn,kakaotalkyn,status,certno,isusing,regdate, adminid, customer_file_regdate, asmasteridx" & vbcrlf 
		sqlStr = sqlStr & " from db_cs.dbo.tbl_customer_filelist with (readuncommitted)" & vbcrlf 
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by authidx desc" + vbcrlf 

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
				set FItemList(i) = new ccsfileitem
					fitemlist(i).fauthidx = rsget("authidx")
					fitemlist(i).fuserhp = rsget("userhp")
					fitemlist(i).fuserid = rsget("userid")
					fitemlist(i).forderserial = rsget("orderserial")
					fitemlist(i).fcomment = db2html(rsget("comment"))
					fitemlist(i).ffileurl1 = rsget("fileurl1")
					fitemlist(i).ffileurl2 = rsget("fileurl2")
					fitemlist(i).ffileurl3 = rsget("fileurl3")
					fitemlist(i).ffileurl4 = rsget("fileurl4")
					fitemlist(i).ffileurl5 = rsget("fileurl5")
					fitemlist(i).fsmsyn = rsget("smsyn")
					fitemlist(i).fkakaotalkyn = rsget("kakaotalkyn")
					fitemlist(i).fcertno = rsget("certno")
                    fitemlist(i).fstatus = rsget("status")
					fitemlist(i).fisusing = rsget("isusing")
					fitemlist(i).fregdate = rsget("regdate")
                    fitemlist(i).fadminid = rsget("adminid")
					fitemlist(i).fcustomer_file_regdate = rsget("customer_file_regdate")
					fitemlist(i).fasmasteridx = rsget("asmasteridx")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public sub getcsfilenotpaging()
		dim sqlStr,i ,sqlsearch

		if frectuserhp <> "" then
			sqlsearch = sqlsearch & " and replace(userhp,'-','') ='"& replace(frectuserhp,"-","") &"'"
		end if
		if frectuserid <> "" then
			sqlsearch = sqlsearch & " and userid ='"& frectuserid &"'"
		end if
		if frectorderserial <> "" then
			sqlsearch = sqlsearch & " and orderserial ='"& frectorderserial &"'"
		end if
		if frectasmasteridx <> "" then
			sqlsearch = sqlsearch & " and asmasteridx ="& frectasmasteridx &""
		end if
		if frectstatus <> "" then
			sqlsearch = sqlsearch & " and status ="& frectstatus &""
		end if
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and isusing ='"& frectisusing &"'"
		end if

		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf 
		sqlStr = sqlStr & " authidx,userhp,userid,orderserial,comment,fileurl1,fileurl2,fileurl3,fileurl4" & vbcrlf 
		sqlStr = sqlStr & " ,fileurl5,smsyn,kakaotalkyn,status,certno,isusing,regdate, adminid, customer_file_regdate, asmasteridx" & vbcrlf 
		sqlStr = sqlStr & " from db_cs.dbo.tbl_customer_filelist with (readuncommitted)" & vbcrlf 
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by authidx desc" + vbcrlf 

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.recordcount
		FResultCount = rsget.recordcount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new ccsfileitem
					fitemlist(i).fauthidx = rsget("authidx")
					fitemlist(i).fuserhp = rsget("userhp")
					fitemlist(i).fuserid = rsget("userid")
					fitemlist(i).forderserial = rsget("orderserial")
					fitemlist(i).fcomment = db2html(rsget("comment"))
					fitemlist(i).ffileurl1 = rsget("fileurl1")
					fitemlist(i).ffileurl2 = rsget("fileurl2")
					fitemlist(i).ffileurl3 = rsget("fileurl3")
					fitemlist(i).ffileurl4 = rsget("fileurl4")
					fitemlist(i).ffileurl5 = rsget("fileurl5")
					fitemlist(i).fsmsyn = rsget("smsyn")
					fitemlist(i).fkakaotalkyn = rsget("kakaotalkyn")
					fitemlist(i).fcertno = rsget("certno")
                    fitemlist(i).fstatus = rsget("status")
					fitemlist(i).fisusing = rsget("isusing")
					fitemlist(i).fregdate = rsget("regdate")
                    fitemlist(i).fadminid = rsget("adminid")
					fitemlist(i).fcustomer_file_regdate = rsget("customer_file_regdate")
					fitemlist(i).fasmasteridx = rsget("asmasteridx")
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

function getstatusname(vstatus)
	dim tmpstatus

	if vstatus="0" then
		tmpstatus="입력대기"
	elseif vstatus="1" then
		tmpstatus="입력완료"
	else
		tmpstatus=""
	end if

	getstatusname=tmpstatus
end function

'// 인증여부 선택
Function drawfilecertsendgubun(selectBoxName,selectedId,chplg,dispNotValue)
%>
	<select name="<%=selectBoxName%>" <%= chplg %>>
		<% if dispNotValue="Y" then %>
			<option value="" <% if selectedId="" then response.write "selected" %>>SELECT</option>
		<% end if %>
		<option value="KAKAOTALK" <% if selectedId="KAKAOTALK" then response.write "selected" %>>카카오톡 발송</option>
		<option value="SMS" <% if selectedId="SMS" then response.write "selected" %>>SMS 발송</option>
	</select>
<%
end function

Function GetcsFileName(filename)
	On Error Resume Next
	Dim vUrl			'/소스 경로저장 변수
	Dim FullFilename		'파일이름
	Dim strName			'확장자를 제외한 파일이름

	vUrl = filename
	FullFilename = mid(vUrl,instrrev(vUrl,"/")+1)
	strName = Mid(FullFilename, 1, Instr(FullFilename, ".") - 1)

	GetcsFileName = strName
End Function
%>