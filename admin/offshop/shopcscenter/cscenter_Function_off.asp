<%
'###########################################################
' Description : 매장 고객센터
' Hieditor : 2012.03.20 한용민 생성
'###########################################################

dim IsStatusRegister			'접수
dim IsStatusEdit				'수정
dim IsStatusFinishing			'처리완료 시도
dim IsStatusFinished			'처리완료
dim IsDisplayPreviousCSList		'이전 CS 내역
dim IsDisplayCSMaster			'CS 마스터정보
dim IsDisplayItemList			'상품목록
dim IsDisplayRefundInfo			'환불정보
dim IsDisplayButton				'버튼
dim IsPossibleModifyCSMaster
dim IsPossibleModifyItemList
dim IsPossibleModifyRefundInfo
dim ARR_ERROR_MSG()
dim MAX_ERROR_MSG_COUNT
dim ERROR_MSG_TRY_MODIFY

MAX_ERROR_MSG_COUNT = 10
ReDim Preserve ARR_ERROR_MSG(MAX_ERROR_MSG_COUNT)

'변수 설정
function SetCSVariable_off(mode, divcd)
	IsStatusRegister 			= false
	IsStatusEdit 				= false
	IsStatusFinishing 			= false
	IsStatusFinished 			= false
	IsDisplayPreviousCSList 	= true
	IsDisplayCSMaster 			= true
	IsDisplayItemList 			= true
	IsDisplayRefundInfo 		= true
	IsDisplayButton 			= true
	IsPossibleModifyCSMaster	= true
	IsPossibleModifyItemList	= true
	IsPossibleModifyRefundInfo	= true
	
	'CS 접수
    if (mode = "regcsas") then	
    	IsStatusRegister 	= true

	'CS 수정
    elseif (mode = "editreginfo") then
    	IsStatusEdit 		= true
		IsPossibleModifyItemList	= false
		IsPossibleModifyRefundInfo	= false

		ERROR_MSG_TRY_MODIFY = "CS 수정상태에서는 상품정보/환불정보를 수정할 수 없습니다. 삭제후 재작성하세요."
    
    '완료시도
    elseif (mode = "finishreginfo") then
    	IsStatusFinishing 	= true
		IsPossibleModifyCSMaster	= false
		IsPossibleModifyItemList	= false
		IsPossibleModifyRefundInfo	= false

		ERROR_MSG_TRY_MODIFY = "CS 완료처리 단계에서는 처리내용입력 외 수정할 수 없습니다. CS 정보수정을 이용하세요."
    
    '완료된 내역
    elseif (mode = "finished") then    	    	
    	IsStatusFinished 	= true
		IsPossibleModifyCSMaster	= false
		IsPossibleModifyItemList	= false
		IsPossibleModifyRefundInfo	= false
    	IsDisplayButton 	= false
    	
    	ERROR_MSG_TRY_MODIFY = "완료된 내역은 수정할 수 없습니다."
    end if
end function

function GetCSCommName_off(groupCode, divcd)
	dim tmp_str,sqlStr

	sqlStr = " select top 1 comm_cd,comm_name "
	sqlStr = sqlStr + " from  "
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_cs_comm_code_off "
	sqlStr = sqlStr + " where comm_group='" + groupCode + "' "
	sqlStr = sqlStr + " and comm_cd='" + CStr(divcd) + "' "
	sqlStr = sqlStr + " and comm_isDel='N' "
	
	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1

	tmp_str = ""
	if not rsget.EOF  then
		tmp_str = db2html(rsget("comm_name"))
	end if
	rsget.close

	GetCSCommName_off = tmp_str
End function

function GetDefaultTitle_off(divcd, InsertedId, orderno ,masteridx)
    dim opentitle, opencontents ,sqlStr
    dim ipkumdiv, cancelyn, comm_name, ipkumdivName    
	
	sqlStr = ""
	sqlStr = " select m.cancelyn, C.comm_name"
	sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_master m"
	sqlStr = sqlStr + " left join db_shop.dbo.tbl_shopjumun_cs_master A"
	sqlStr = sqlStr + "     on A.orderno='" + orderno + "'"

	if (masteridx<>"") then
		sqlStr = sqlStr + " and A.masteridx=" + CStr(masteridx)
	end if

    sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_cs_comm_code_off C"
    sqlStr = sqlStr + " on C.comm_cd='" + divcd + "'"

    sqlStr = sqlStr + " where m.idx='" + masteridx + "'"

	'response.write sqlStr &"<br>"
    rsget.Open sqlStr,dbget,1
    
    if Not rsget.Eof then
        cancelyn    = rsget("cancelyn")
        comm_name   = rsget("comm_name")
    end if
    
    rsget.close

	GetDefaultTitle_off = comm_name    
end function

'주문취소
public function IsCSCancelProcess_off(divcd)
	if (divcd = "A008") then
		IsCSCancelProcess_off = true
	else
		IsCSCancelProcess_off = false
	end if
end function

'/'맞교환회수(텐바이텐배송) ,a/s
public function IsCSReturnProcess_off(divcd)
	if (divcd = "A030") then
		IsCSReturnProcess_off = true
	else
		IsCSReturnProcess_off = false
	end if
end function

''접수시 상품별 체크 가능여부
public function IsPossibleCheckItem_off(divcd, ismastercanceled, isdetailcanceled)
	IsPossibleCheckItem_off = false
	if (ismastercanceled) then exit function
	if (isdetailcanceled) then exit function

	if (IsCSCancelProcess_off(divcd)) then
		IsPossibleCheckItem_off = true

	elseif (IsCSReturnProcess_off(divcd) = true) then
		IsPossibleCheckItem_off = false
		
		if _
			(divcd = "A030") _				
		then
			'a/s
			IsPossibleCheckItem_off = true
		end if
	else
		'기타
		IsPossibleCheckItem_off = true
	end if
end function

'' CS Master 저장
function RegCSMaster_off(divcd, orderno,reguserid, title, contents_jupsu,masteridx)
    dim sqlStr, InsertedId
	
	sqlStr = ""
    sqlStr = " select * from db_shop.dbo.tbl_shopjumun_cs_master where 1=0 "
    rsget.Open sqlStr,dbget,1,3
    rsget.AddNew
    
    	rsget("orgmasteridx") = masteridx
        rsget("divcd")          = divcd
    	rsget("orderno")    = orderno
    	rsget("customername")   = ""    	
    	rsget("writeuser")      = reguserid
    	rsget("title")          = title
    	rsget("contents_jupsu") = contents_jupsu
    	rsget("currstate")      = "B001"
    	rsget("deleteyn")       = "N"

        ''''''''''''''''''''''''''''''''''
    	''rsget("requireupche")   = "N"
    	''rsget("makerid")        = ""
    	''''''''''''''''''''''''''''''''''

    rsget.update
	    InsertedId = rsget("masteridx")
	rsget.close

	dim opentitle, opencontents

	opentitle = GetDefaultTitle_off(divcd, InsertedId, orderno ,masteridx)
	
	sqlStr = ""
	sqlStr = " update db_shop.dbo.tbl_shopjumun_cs_master"  + VbCrlf
	sqlStr = sqlStr + " set customername='" + html2db(reqname) + "'"   + VbCrlf
	sqlStr = sqlStr + " , opentitle='" + html2db(opentitle) + "'" + VbCrlf
	sqlStr = sqlStr + " , opencontents='" + html2db(opencontents) + "'" + VbCrlf
	sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_master T" + VbCrlf
    sqlStr = sqlStr + " where T.idx='" + masteridx + "'"  + VbCrlf
    sqlStr = sqlStr + " and db_shop.dbo.tbl_shopjumun_cs_master.masteridx=" + CStr(InsertedId)
		
	'response.write sqlStr &"<br>"
	dbget.Execute sqlStr
	
	RegCSMaster_off = InsertedId
end function

function AddOneCSDetail_off(csmasteridx, dorderdetailidx, orderno, dregitemno)
    dim sqlStr , jumundetailidx , jumunitemgubun

	if masteridx = "" then exit function    

	'/원래 판매테이블 의 detailidx
	sqlStr = "select "
	sqlStr = sqlStr & " idx,orderno , itemno,itemgubun"
	sqlStr = sqlStr & " from [db_shop].dbo.tbl_shopjumun_detail"
	sqlStr = sqlStr & " where idx = "&dorderdetailidx&""

    'response.write sqlStr &"<Br>"
    rsget.Open sqlStr,dbget,1    
	    if Not rsget.Eof then
			jumundetailidx = rsget("idx")
			jumunitemgubun = rsget("itemgubun")
	    end if    
    rsget.Close

	sqlStr = ""
    sqlStr = " insert into [db_shop].dbo.tbl_shopjumun_cs_detail"
    sqlStr = sqlStr + " (masteridx, orgdetailidx,orderno, itemid, itemoption,makerid"
    sqlStr = sqlStr + " , regitemno, confirmitemno,orderitemno ,itemgubun) "
    sqlStr = sqlStr + " values(" + CStr(csmasteridx) + ""
    sqlStr = sqlStr + " ," + CStr(dorderdetailidx) + ""
    sqlStr = sqlStr + " ,'" + CStr(orderno) + "'"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ," + CStr(dregitemno) + ""
    sqlStr = sqlStr + " ," + CStr(dregitemno) + ""
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,'"&jumunitemgubun&"'"    
    sqlStr = sqlStr + " )"
    
	'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
end function

function AddCSDetailByArrStr_off(byval detailitemlist, csmasteridx, orderno,masteridx ,isupchebeasong)
    dim sqlStr, tmp, buf, i ,dorderdetailidx, dregitemno    

    buf = split(detailitemlist, "|")

    for i = 0 to UBound(buf)
		if (TRIM(buf(i)) <> "") then
			tmp = split(buf(i), Chr(9))

			dorderdetailidx = tmp(0)
			dregitemno      = tmp(1)

	        call AddOneCSDetail_off(csmasteridx, dorderdetailidx,orderno , dregitemno)
		end if
	next

	sqlStr = " update [db_shop].dbo.tbl_shopjumun_cs_detail"
	sqlStr = sqlStr + " set itemid=T.itemid"
	sqlStr = sqlStr + " , itemoption=T.itemoption"
	sqlStr = sqlStr + " , itemgubun=T.itemgubun"
	sqlStr = sqlStr + " , makerid=T.makerid"
	sqlStr = sqlStr + " , orderitemno=T.itemno"
	
	if isupchebeasong <> "" then
		sqlStr = sqlStr + " , isupchebeasong='"& isupchebeasong &"'"
	end if
	
	sqlStr = sqlStr + " from [db_shop].dbo.tbl_shopjumun_detail T"
	sqlStr = sqlStr + " where T.orderno='" + orderno + "'"
	sqlStr = sqlStr + " and [db_shop].dbo.tbl_shopjumun_cs_detail.masteridx=" + CStr(csmasteridx)
	sqlStr = sqlStr + " and [db_shop].dbo.tbl_shopjumun_cs_detail.orgdetailidx=T.idx"
	
	'response.write sqlStr &"<Br>"
	dbget.Execute sqlStr
end function

'//업체처리
function RegCSMasterAddUpche_off(csmasteridx, imakerid)
    dim sqlStr
    sqlStr = " update db_shop.dbo.tbl_shopjumun_cs_master"    + VbCrlf
    sqlStr = sqlStr + " set makerid='" + imakerid + "'"   + VbCrlf
    sqlStr = sqlStr + " , requireupche='Y'"               + VbCrlf
    sqlStr = sqlStr + " where masteridx=" + CStr(csmasteridx)

	'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
  
end function

'//매장처리
function RegCSMasterAddmaejang_off(csmasteridx, imakerid)
    dim sqlStr
    sqlStr = " update db_shop.dbo.tbl_shopjumun_cs_master"    + VbCrlf
    sqlStr = sqlStr + " set makerid='" + imakerid + "'"   + VbCrlf
    sqlStr = sqlStr + " , requiremaejang='Y'"+ VbCrlf
    sqlStr = sqlStr + " where masteridx=" + CStr(csmasteridx)

	'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
end function

'/배송지 등록
function Regdelivery_off(csmasteridx, reqname ,reqphone ,reqhp ,reqemail,reqzipcode ,reqzipaddr ,reqaddress ,comment)
    dim sqlStr
    
	sqlStr = " if exists("     + VbCrlf
    sqlStr = sqlStr + " 	select top 1 * from db_shop.dbo.tbl_shopjumun_cs_delivery"     + VbCrlf
    sqlStr = sqlStr + " 	where asid = "&csmasteridx&""     + VbCrlf
    sqlStr = sqlStr + " )"     + VbCrlf
    sqlStr = sqlStr + " 	update db_shop.dbo.tbl_shopjumun_cs_delivery set"     + VbCrlf
    sqlStr = sqlStr + " 	reqname='" + html2db(reqname) + "'"   + VbCrlf
    sqlStr = sqlStr + " 	,reqphone = '" + CStr(reqphone) + "'"  + VbCrlf
    sqlStr = sqlStr + " 	,reqhp = '" + CStr(reqhp) + "'"        + VbCrlf
    sqlStr = sqlStr + " 	,reqzipcode = '" + CStr(reqzipcode) + "'"  + VbCrlf
    sqlStr = sqlStr + " 	,reqzipaddr = '" + CStr(reqzipaddr) + "'"    + VbCrlf
    sqlStr = sqlStr + " 	,reqetcaddr = '" + html2db(reqaddress) + "'"    + VbCrlf
    sqlStr = sqlStr + " 	,reqetcstr = '" + html2db(comment) + "'"    + VbCrlf
    sqlStr = sqlStr + " 	,reqemail = '" + html2db(reqemail) + "'"    + VbCrlf    
    sqlStr = sqlStr + " 	where asid='" + CStr(csmasteridx) + "'" + VbCrlf
    sqlStr = sqlStr + " else"     + VbCrlf
    sqlStr = sqlStr + " 	insert into db_shop.dbo.tbl_shopjumun_cs_delivery("     + VbCrlf
    sqlStr = sqlStr + " 	asid ,reqname ,reqphone ,reqhp ,reqzipcode ,reqzipaddr ,reqetcaddr ,reqetcstr "     + VbCrlf
    sqlStr = sqlStr + " 	,reqemail ,regdate) values"     + VbCrlf
    sqlStr = sqlStr + " 	("     + VbCrlf
    sqlStr = sqlStr + " 	"&csmasteridx&" ,'" + html2db(reqname) + "','" + CStr(reqphone) + "' ,'" + CStr(reqhp) + "'"     + VbCrlf
    sqlStr = sqlStr + " 	,'" + CStr(reqzipcode) + "','" + CStr(reqzipaddr) + "' ,'" + html2db(reqaddress) + "'"     + VbCrlf
    sqlStr = sqlStr + " 	,'" + html2db(comment) + "','" + html2db(reqemail) + "',getdate()"     + VbCrlf
    sqlStr = sqlStr + " 	)"

    'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr

end function

'/ 배송 수령지
function DrawdeliveryCombo(selectBoxName,selectedId,chplug,shopdiv,sudongyn,frm)
	dim tmp_str,query1
	%>
	<select class="select" name="<%=selectBoxName%>" <%= chplug %>>
		<% if sudongyn = "Y" then %>
			<option value='SUDONG' <%if selectedId="sudong" then response.write " selected"%>>직접입력</option>
		<% end if %>
	<%
		query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user"
		query1 = query1 & " where isusing='Y' "
		query1 = query1 & " and userid<>'streetshop000'"
		query1 = query1 & " and userid<>'streetshop800'"
		query1 = query1 & " and userid<>'streetshop870'"
		
		if shopdiv <> "" then
			query1 = query1 & " and shopdiv in ("&shopdiv&")"
		end if
					
		rsget.Open query1,dbget,1
		
		if  not rsget.EOF  then
		rsget.Movefirst
		
		do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("userid")) then
		tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&"/"&rsget("shopname")&"</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
		end if
		rsget.close
	response.write("</select>")
	'response.write query1 &"<Br>"
end function

'' CS Master 수정
function EditCSMaster_off(divcd, orderserial, modiuserid, title, contents_jupsu, csmasteridx)    
    dim sqlStr

    sqlStr = " update db_shop.dbo.tbl_shopjumun_cs_master"
    sqlStr = sqlStr + " set writeuser='" + modiuserid + "'"
    sqlStr = sqlStr + " ,title='" + title + "'"
    sqlStr = sqlStr + " ,contents_jupsu='" + contents_jupsu + "'"
    sqlStr = sqlStr + " where masteridx=" + CStr(csmasteridx)

	'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
end function

'' CS Detail 수정
function EditCSDetailByArrStr_off(byval detailitemlist, csmasteridx, orderno)
    dim sqlStr, tmp, buf, i
    dim dorderdetailidx, dregitemno, dcausecontent

    buf = split(detailitemlist, "|")

    for i = 0 to UBound(buf)
		if (TRIM(buf(i)) <> "") then
			tmp = split(buf(i), Chr(9))

			dorderdetailidx = tmp(0)
			dregitemno      = tmp(1)
			dcausecontent   = tmp(2)

	        call EditOneCSDetail(csmasteridx, dorderdetailidx, orderno, dregitemno, dcausecontent)
		end if
	next
end function

function EditOneCSDetail(csmasteridx, dorderdetailidx, orderno, dregitemno, dcausecontent)
    dim sqlStr

    sqlStr = " update db_shop.dbo.tbl_shopjumun_cs_detail set"
    sqlStr = sqlStr + " regitemno=" + dregitemno + ""
    sqlStr = sqlStr + " , confirmitemno=" + dregitemno + ""
    sqlStr = sqlStr + " where masterid=" + CStr(id)
    sqlStr = sqlStr + " and orgdetailidx=" + CStr(dorderdetailidx)

	'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
end function

function FinishCSMaster_off(iAsid, finishuser, contents_finish)
    dim sqlStr ,IsCsErrStockUpdateRequire

    IsCsErrStockUpdateRequire = False

    sqlStr = "select divcd, finishdate, currstate"
    sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_cs_master"
    sqlStr = sqlStr + " where masteridx=" + CStr(iAsid)
    
    'response.write sqlStr &"<br>"
    rsget.Open sqlStr,dbget,1
    
    if Not rsget.Eof then
        IsCsErrStockUpdateRequire = (rsget("divcd")="A011") and (IsNULL(rsget("finishdate"))) and (rsget("currstate")<>"B007")
    end if
    
    rsget.close

    sqlStr = " update db_shop.dbo.tbl_shopjumun_cs_master set"	+ VbCrlf
    sqlStr = sqlStr + " finishuser='" + finishuser + "'"            + VbCrlf
    sqlStr = sqlStr + " , contents_finish='" + contents_finish + "'"    + VbCrlf
    sqlStr = sqlStr + " , finishdate=getdate()"                         + VbCrlf
    sqlStr = sqlStr + " , currstate='B007'"                             + VbCrlf
    sqlStr = sqlStr + " where masteridx=" + CStr(iAsid)

    'response.write sqlStr &"<br>"
    dbget.Execute sqlStr

end function

function ValidDeleteCS_off(masteridx)
    dim sqlStr
    dim currstate

    ValidDeleteCS_off = false

    sqlStr = "select * from db_shop.dbo.tbl_shopjumun_cs_master"
    sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
	
	'response.write sqlStr &"<Br>"
    rsget.Open sqlStr,dbget,1
        currstate = rsget("currstate")
    rsget.Close

    If (currstate>="B006") then Exit function

    ValidDeleteCS_off = true
end function

function DeleteCSProcess_off(masteridx, finishuserid)
    dim sqlStr, resultCount

    sqlStr = " update db_shop.dbo.tbl_shopjumun_cs_master set" + VbCrlf
    sqlStr = sqlStr + "  deleteyn='Y'" + VbCrlf
    sqlStr = sqlStr + " ,finishuser = '" + finishuserid+ "'" + VbCrlf
    sqlStr = sqlStr + " ,finishdate = getdate()" + VbCrlf
    sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
    sqlStr = sqlStr + " and currstate<'B006'"

	'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr, resultCount

    DeleteCSProcess_off = (resultCount>0)
end function

function CsState2Name_off(byval v)
	if IsNull(v) or (v="") then
		Exit function
	end if

	if v="0" then

	elseif v="B001" then
		CsState2Name_off = "접수"
	elseif v="B004" then
		CsState2Name_off = "운송장입력"		
	elseif v="B006" then
		CsState2Name_off = "업체처리완료"
	elseif v="B007" then
		CsState2Name_off = "최종처리완료"
	elseif v="B008" then
		CsState2Name_off = "매장처리완료"		
	else
	end if
end function

'//2016년 어드민 신버전		'/2017.01.26 한용민 생성
function getcurrstate_div(currstate ,divcd)
%>
	<table class="tbType1 listTb">
	<!--<colgroup>
	<col width="20" />
	</colgroup>-->
		<thead>
		<tr>
			<% if currstate = "" then %>
				<% if currstate="" then %>
					<th><div>[0]등록</div></th>
				<% else %>
					<td>[0]등록</td>
				<% end if %>
			<% end if %>
			
			<% if currstate="B001" then %>
				<th><div>[1]접수</div></th>
			<% else %>
				<td>[1]접수</td>
			<% end if %>

			<%' if currstate="B002" then %>
				<!--<th><div>[2]미처리(사유)</div></th>-->
			<%' else %>
				<!--<td>[2]미처리(사유)</td>-->
			<%' end if %>

			<%' if currstate="B003" then %>
				<!--<th><div>[3]택배사전송</div></th>-->
			<%' else %>
				<!--<td>[3]택배사전송</td>-->
			<%' end if %>

			<% if currstate="B004" then %>
				<th><div>[2]운송장입력</div></th>
			<% else %>
				<td>[2]운송장입력</td>
			<% end if %>

			<%' if currstate="B005" then %>
				<!--<th><div>[5]확인요청</div></th>-->
			<%' else %>
				<!--<td>[5]확인요청</td>-->
			<%' end if %>

			<% if currstate="B006" or currstate="B008" then %>
				<th><div>[3]업체&매장처리완료</div></th>
			<% else %>
				<td>[3]업체&매장처리완료</td>
			<% end if %>

			<% if currstate="B007" then %>
				<th><div>[4]최종처리완료</div></th>
			<% else %>
				<td>[4]최종처리완료</td>
			<% end if %>
		</tr>
		</thead>
	</tbody>
	</table>
<%		
end function

function getcurrstate_table(currstate ,divcd)
%>
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="silver">
		<tr align="center"bgcolor="#E6E6E6">
			<% if currstate = "" then %>
				<td <% if currstate="" then %> bgcolor="pink" <% end if %> >[0]등록</td>
			<% end if %>
			
			<td <% if currstate="B001" then %> bgcolor="pink" <% end if %> >[1]접수</td>
			<!--<td <%' if currstate="B002" then %> bgcolor="pink" <%' end if %> >[2]미처리(사유)</td>-->
			<!--<td <%' if currstate="B003" then %> bgcolor="pink" <%' end if %> >[3]택배사전송</td>-->
			<td <% if currstate="B004" then %> bgcolor="pink" <% end if %> >[2]운송장입력</td>
			<!--<td <%' if currstate="B005" then %> bgcolor="pink" <%' end if %> >[5]확인요청</td>-->
			<td <% if currstate="B006" or currstate="B008" then %> bgcolor="pink" <% end if %> >[3]업체&매장처리완료</td>
			<td <% if currstate="B007" then %> bgcolor="pink" <% end if %> >[4]최종처리완료</td>
		</tr>
	</table>
<%		
end function

function drawcurrstate(boxname ,selectid, chplug)
%>
    <select name="<%=boxname%>" <%= chplug %>>
    	<option value="">전체</option>
		<option value="notfinish" <% if (selectid = "notfinish") then response.write "selected" end if %>>미처리전체</option>    	
		<option value="B001" <% if (selectid = "B001") then response.write "selected" end if %>>접수</option>
		<option value="B004" <% if (selectid = "B004") then response.write "selected" end if %>>운송장입력</option>
		<option value="B006" <% if (selectid = "B006") then response.write "selected" end if %>>업체처리완료</option>
		<option value="B006" <% if (selectid = "B008") then response.write "selected" end if %>>매장처리완료</option>
		<option value="notfinal" <% if (selectid = "notfinal") then response.write "selected" end if %>>업체&매장처리완료</option>		
		<option value="B007" <% if (selectid = "B007") then response.write "selected" end if %>>최종처리완료</option>
    </select>
<%		
end function
%>