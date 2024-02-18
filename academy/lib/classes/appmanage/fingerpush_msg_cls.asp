<%
'###########################################################
' Description : fingers apps push message
' Hieditor : 2016.11.30 이종화
'###########################################################

''디바이스별
function getNoticsPushTargetCount(idoc_idx)
    dim ret : ret = 0
    dim sqlStr
    '' A :전체수신 , C: 공지, P:상품
    
    sqlStr = " select count(*) as CNT from [db_academy].[dbo].[tbl_app_regInfo] R"
    sqlStr = sqlStr & " where isusing='Y'"
    sqlStr = sqlStr & " and isNULL(pushyn,'A') in ('A','C')"
    sqlStr = sqlStr & " and isNULL(userid,'')<>'' "
    
    rsACADEMYget.CursorLocation = adUseClient
    rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
    if  not rsACADEMYget.EOF  then
        ret = rsACADEMYget("CNT")
    end if    
    rsACADEMYget.close
    
    getNoticsPushTargetCount = ret
end function

''강사 아이디별 최종 (안드로이드,IOS)
function getNoticsPushTargetCountLastUser(idoc_idx)
    dim ret : ret = 0
    dim sqlStr
    '' A :전체수신 , C: 공지, P:상품
    
    sqlStr = " select count(*) as CNT"
    sqlStr = sqlStr & " from ("
    sqlStr = sqlStr & " select userid,deviceid,lastupdate,isusing,pushyn, row_number() over (partition by userid order by lastupdate desc) as RNk"
    sqlStr = sqlStr & "  from [db_academy].[dbo].[tbl_app_regInfo]"
    sqlStr = sqlStr & " where isNULL(userid,'')<>''"
    sqlStr = sqlStr & " ) T"
    sqlStr = sqlStr & " where T.RNk=1"
    sqlStr = sqlStr & " and isusing='Y' and isNULL(pushyn,'A') in ('A','C')"
    
    rsACADEMYget.CursorLocation = adUseClient
    rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
    if  not rsACADEMYget.EOF  then
        ret = rsACADEMYget("CNT")
    end if    
    rsACADEMYget.close
    
    getNoticsPushTargetCountLastUser = ret
end function


Sub drawSelectBoxTestTarget(selectBoxName, selectedId, addStr)
   dim tmp_str,sqlStr
   %><select class="select" name="<%=selectBoxName%>" <%=addStr%>>
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
   <%
    sqlStr = " select lecturer_id,lecturer_name from db_academy.dbo.tbl_lec_user  "
    sqlStr = sqlStr & " where lecturer_id in ('fingertest01','fingertest02','fingertest03','fingertest04','fingertest05','thefingers01') "
	sqlStr = sqlStr & " order by lecturer_id"


    rsACADEMYget.CursorLocation = adUseClient
    rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly

   if  not rsACADEMYget.EOF  then
       rsACADEMYget.Movefirst

       do until rsACADEMYget.EOF
           if Lcase(selectedId) = Lcase(rsACADEMYget("lecturer_id")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsACADEMYget("lecturer_id")&"' "&tmp_str&">"&rsACADEMYget("lecturer_id")&" ["&rsACADEMYget("lecturer_name")&"]</option>")
           tmp_str = ""
           rsACADEMYget.MoveNext
       loop
   end if
   rsACADEMYget.close
   
   response.write("</select>")
End Sub

Sub drawSelectBoxTestDevice(testlecid, selectBoxName, selectedId, addStr)
   dim tmp_str,sqlStr
   %><select class="select" name="<%=selectBoxName%>" <%=addStr%>>
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
   <%
    sqlStr = " select regIdx,deviceid,appVer,isusing,isNULL(pushyn,'') as pushyn from db_academy.dbo.tbl_app_regInfo  "
    sqlStr = sqlStr & " where isNULL(userid,'')<>''"
    sqlStr = sqlStr & " and isNULL(userid,'')='"&testlecid&"'"
	sqlStr = sqlStr & " order by regIdx desc"


    rsACADEMYget.CursorLocation = adUseClient
    rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly

   if  not rsACADEMYget.EOF  then
       rsACADEMYget.Movefirst

       do until rsACADEMYget.EOF
           if Lcase(selectedId) = Lcase(rsACADEMYget("regIdx")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsACADEMYget("regIdx")&"' "&tmp_str&">"&rsACADEMYget("deviceid")&" ["&rsACADEMYget("appVer")&"]</option>")
           tmp_str = ""
           rsACADEMYget.MoveNext
       loop
   end if
   rsACADEMYget.close
   
   response.write("</select>")
End Sub


Sub drawSelectBoxTarget(selectBoxName, selectedId, addStr)
   dim tmp_str,sqlStr
   %><select class="select" name="<%=selectBoxName%>" <%=addStr%>>
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
   <%
    sqlStr = " select targetKey,targetName from db_academy.dbo.tbl_app_academy_targetQuery  "
    sqlStr = sqlStr & " where (isusing='Y' "
    if selectedId<>"" then
        sqlStr = sqlStr & " or targetKey='"&selectedId&"'"
    end if
    sqlStr = sqlStr & " )"
	sqlStr = sqlStr & " order by targetKey"

    rsACADEMYget.CursorLocation = adUseClient
    rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly

   if  not rsACADEMYget.EOF  then
       rsACADEMYget.Movefirst

       do until rsACADEMYget.EOF
           if Lcase(selectedId) = Lcase(rsACADEMYget("targetKey")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsACADEMYget("targetKey")&"' "&tmp_str&">"&rsACADEMYget("targetName")&"</option>")
           tmp_str = ""
           rsACADEMYget.MoveNext
       loop
   end if
   rsACADEMYget.close
   
   response.write("</select>")
End Sub

Class cpush_item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx
	public fnotiGbn   ''구분 :1.공지, 2.상품승인, 3.상품등록보류..
	Public freservedate
	Public fpushtitle
	Public fpushurl
	Public fpushimg
	Public fstate
	Public ftestpush
	Public fisusing
	'// regInfo_log
	Public Fregidx
	Public fappkey
	Public fdeviceid
	Public fappVer
	Public flastact
	Public fregdate
	Public fttlCnt
	Public fwaitCnt
	Public fsentCnt
	Public fsuccCnt
	Public ffailCnt
	Public ffirstSentDate
	Public flastSentDate
	Public fdiffnomuts

    Public fistargetMsg
    Public fnoduppDate  '' 금일 발송건에 대해 중복발송 안함.
    Public ftargetKey   '' 
    Public fadmcomment  '' 코멘트
    Public fbaseIdx     '' 타케팅 마스터 키 - 안드로이드에서 중복 수신 방지하기위한 키값 pkey - 이값이 동일하면 메세지가 여러개 와도 받지 않는다.

    Public ftargetName
    Public FtargetState
    Public FmayTargetCnt
    
    public Function getTargetStateName()
        if (FtargetState=0) then
            getTargetStateName = "대기중"
        elseif (FtargetState=1) then
            getTargetStateName = "<font color='green'>타겟예약</font>"
        elseif (FtargetState=3) then
            getTargetStateName = "타겟중"
        elseif (FtargetState=7) then
            getTargetStateName = "<font color='red'>타겟완료</font>"
        else
            getTargetStateName = FtargetState
        end if
    end function

    public Function IsTargetActionValid()
        IsTargetActionValid = (fistargetMsg=1) and ((FtargetState=0) or (FtargetState=1)) and ((fstate=0) or (fstate=1))
    end function
end class

class cpush_msg_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	Public FOneItem

	public frectcomplete
	public Frectdate
	Public Frectidx
	Public FRectuserid
	Public Fstate
	Public Fisusing
	public Frectpushtitle
	public Frectpushurl
	public FrecttargetKey

	public sub fpushmsglist()
		dim sqlStr, i, sqlsearch

		if Frectidx <> "" then
			sqlsearch = sqlsearch & " and P.idx = '"& Frectidx &"'"
		end If
		if Frectdate <> "" then
			sqlsearch = sqlsearch & " and convert(varchar(10),P.reservedate,120)='"& Frectdate &"'"
		end If
		if Fstate <> "" then
			sqlsearch = sqlsearch & " and P.state='"& Fstate &"'"
		end If
		if Fisusing <> "" then
			sqlsearch = sqlsearch & " and P.isusing='"& Fisusing &"'"
		end if
		if Frectpushtitle <> "" then
			sqlsearch = sqlsearch & " and P.pushtitle like '%"& Frectpushtitle &"%'"
		end if
		if Frectpushurl <> "" then
			sqlsearch = sqlsearch & " and P.pushurl like '%"& Frectpushurl &"%'"
		end if
		if FrecttargetKey <> "" then
			sqlsearch = sqlsearch & " and P.targetKey = "& FrecttargetKey &""
		end If

		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_academy.dbo.tbl_app_Academy_push_reserve P"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		

		'response.write sqlStr &"<br>"
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close
		
		if FTotalCount < 1 then exit sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " P.idx , P.pushtitle , P.pushurl , P.pushimg , P.state , P.testpush , P.isusing , P.reservedate, P.istargetMsg"
		sqlStr = sqlStr & " , P.noduppDate, P.targetKey, P.baseIdx, P.targetState, isNULL(P.mayTargetCnt,0) as mayTargetCnt"
		sqlStr = sqlStr & " ,Q.targetName"
		sqlStr = sqlStr & " from db_academy.dbo.tbl_app_Academy_push_reserve P"
		sqlStr = sqlStr & "     left join db_academy.dbo.tbl_app_academy_targetQuery Q"
		sqlStr = sqlStr & "     on P.targetKey=Q.targetKey"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		
		sqlStr = sqlStr & " order by P.reservedate Desc,(CASE WHEN P.targetKey=9999 THEN 0 ELSE 1 END) desc, P.idx " ''idx=>reservedate
		
		'response.write sqlStr &"<br>"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly

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
				set FItemList(i) = new cpush_item

				FItemList(i).fidx			= rsACADEMYget("idx")
				FItemList(i).fnotiGbn       = rsACADEMYget("notiGbn")
				FItemList(i).fpushtitle		= rsACADEMYget("pushtitle")
				FItemList(i).fpushurl		= rsACADEMYget("pushurl")
				FItemList(i).fpushimg		= rsACADEMYget("pushimg")
				FItemList(i).fstate			= rsACADEMYget("state")
				FItemList(i).ftestpush		= rsACADEMYget("testpush")
				FItemList(i).fisusing		= rsACADEMYget("isusing")
				FItemList(i).freservedate	= rsACADEMYget("reservedate")
				FItemList(i).fistargetMsg   = rsACADEMYget("istargetMsg")
				FItemList(i).fnoduppDate    = rsACADEMYget("noduppDate")
				FItemList(i).ftargetKey     = rsACADEMYget("targetKey")
				FItemList(i).fbaseIdx       = rsACADEMYget("baseIdx")
				FItemList(i).ftargetState   = rsACADEMYget("targetState")
				FItemList(i).fmayTargetCnt   = rsACADEMYget("mayTargetCnt")
				
				FItemList(i).ftargetName    = rsACADEMYget("targetName")
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end Sub
	
	public sub pushmsgtest_getrow()
		dim sqlStr, i, sqlsearch

		if Frectidx <> "" then
			sqlsearch = sqlsearch & " and idx="&Frectidx&""
		end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " idx , pushtitle , pushurl , pushimg , state , testpush , isusing , reservedate, istargetMsg, noduppDate, targetKey, admcomment, baseIdx, targetState, isNULL(mayTargetCnt,0) as mayTargetCnt"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_app_push_reserve"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		
		sqlStr = sqlStr & " order by idx Desc"
		
		'response.write sqlStr &"<br>"
		rsget.Open SqlStr, dbget, 1
		FResultCount = rsget.RecordCount

	set FOneItem = new cpush_item

	if Not rsget.Eof then

		FOneItem.Fidx			= rsget("idx")
		FOneItem.fpushtitle		= rsget("pushtitle")
		FOneItem.fpushurl		= rsget("pushurl")
		FOneItem.fpushimg		= rsget("pushimg")
		FOneItem.fstate			= rsget("state")
		FOneItem.ftestpush		= rsget("testpush")
		FOneItem.fisusing		= rsget("isusing")
		FOneItem.freservedate	= rsget("reservedate")
        FOneItem.fistargetMsg   = rsget("istargetMsg")
        FOneItem.fnoduppDate    = rsget("noduppDate") 
        FOneItem.ftargetKey     = rsget("targetKey") 
        FOneItem.fadmcomment    = rsget("admcomment")
        FOneItem.fbaseIdx       = rsget("baseIdx")
        FOneItem.ftargetState   = rsget("targetState")
        FOneItem.fmayTargetCnt   = rsget("mayTargetCnt")
        
	end if
	rsget.Close
	end Sub

	public sub pushmsg_userinfo()
		dim sqlStr, i, sqlsearch

		if Frectuserid <> "" then
			sqlsearch = sqlsearch & " and l.userid='"&FRectuserid&"'"
		end if
		
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_app_regInfo as l "
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten as p "
		sqlStr = sqlStr & " on l.userid = p.userid "
		sqlStr = sqlStr & " where ((l.appkey=6 and l.appVer>='36') or (l.appkey=5 and l.appVer>='1')) "
		sqlStr = sqlStr & " and l.isusing = 'Y' " & sqlsearch		
        sqlStr = sqlStr & " and ((p.userid is Not NULL)"
        sqlStr = sqlStr & "    or (l.userid in ('10x10green','10x10yellow'))"
        sqlStr = sqlStr & " )" 
        
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close
		
		if FTotalCount < 1 then exit sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " l.regidx , l.appkey , l.deviceid , l.appVer , l.lastact , l.regdate "
		sqlStr = sqlStr & " from db_contents.dbo.tbl_app_regInfo as l "
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten as p "
		sqlStr = sqlStr & " on l.userid = p.userid "
		sqlStr = sqlStr & " where ((l.appkey=6 and l.appVer>='36') or (l.appkey=5 and l.appVer>='1')) "
		sqlStr = sqlStr & " and l.isusing = 'Y' " & sqlsearch	
		sqlStr = sqlStr & " and ((p.userid is Not NULL)"
        sqlStr = sqlStr & "    or (l.userid in ('10x10green','10x10yellow'))"
        sqlStr = sqlStr & " )" 
		sqlStr = sqlStr & " order by l.lastupdate desc"
		
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
				set FItemList(i) = new cpush_item

				FItemList(i).Fregidx		= rsget("regidx")
				FItemList(i).fappkey		= rsget("appkey")
				FItemList(i).fdeviceid		= rsget("deviceid")
				FItemList(i).fappVer		= rsget("appVer")
				FItemList(i).flastact		= rsget("lastact")
							
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub
	
	
	public Function fpushmessage_report()
		dim sqlStr, i, sqlsearch
		
		sqlStr = "exec [db_AppNoti].[dbo].[sp_Ten_Report_MultiMsg] '" & Frectidx & "'"
		rsAppNotiget.Open sqlStr,dbAppNotiget,1
'		response.write rsAppNotiget.RecordCount
'		rsAppNotiget.Close
'		dbget.close()
'		dbAppNotiget.close()
'		Response.End

		if  not rsAppNotiget.EOF  then
			fpushmessage_report = rsAppNotiget.getRows()
		end if
		rsAppNotiget.Close
	end Function
	

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
end Class


'//메인페이지 , 이벤트 공통함수		'/오픈예정 노출함 , 검색페이지용
function Draweventstate2(selectBoxName,selectedId,changeFlag)
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %> id="<%=selectBoxName%>">
		<option value="" <% if selectedId="" then response.write "selected" %>>전체</option>
		<option value="0" <% if selectedId="0" then response.write "selected" %>>작성중</option>
		<option value="1" <% if selectedId="1" then response.write "selected" %>>발송예약</option>
		<option value="7" <% if selectedId="7" then response.write "selected" %>>타겟중</option>
		<option value="9" <% if selectedId="9" then response.write "selected" %>>발송완료</option>

	</select>
<%
end Function

Function Selectappname(v)
	if v = "3" then
		Selectappname = "colorApp(ios)"
	elseif v = "4" then
		Selectappname = "colorApp(android)"
	elseif v = "5" then
		Selectappname = "wishApp(ios)"
	elseif v = "6" then
		Selectappname = "wishApp(android)"
	end If
	Response.write Selectappname
End Function

Function pushmsgstate(v)
	if v = "0" then
		pushmsgstate = "작성중"
	elseif v = "1" then
		pushmsgstate = "<font color='red'>발송예약</font>"
	elseif v = "7" then
		pushmsgstate = "<font color='green'>발송중</font>"
	elseif v = "9" then
		pushmsgstate = "<font color='blue'>발송완료</font>"
	Else
		pushmsgstate = "전체"
	end If
	Response.write pushmsgstate
End Function
%>