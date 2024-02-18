<%
session.codePage = 65001
%>
<%
'##############################################################
'	Description : 오프라인 공통함수
'	History : 2006.12.11 정윤정 생성
'			  2011.02.18 한용민 수정
'##############################################################

'//오프라인매장 구분에 따라서 동적으로 가져 오기	'/한용민 추가
Sub drawSelectBoxOffShopdiv_off(selectBoxName,selectedId,shopdiv,shopall,chplg)
	dim tmp_str,query1
	%>
	<select class="select" name="<%=selectBoxName%>" <%=chplg%>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택하세요</option>

		<% if shopall <> "" then %>
			<option value='all' <%if selectedId="all" then response.write " selected"%>>전체매장공통</option>
		<% end if %>
	<%
		query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user"
		query1 = query1 & " where isusing='Y' "

		if shopdiv <> "" then
			query1 = query1 & " and shopdiv in ("&shopdiv&")"
		end if

   		query1 = query1 & " order by isusing desc, convert(int,shopdiv)+10 asc, userid asc"

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
end sub

Sub drawSelectBoxOffShopdiv_offNotUsinginc(selectBoxName,selectedId,shopdiv,shopall,chplg)
	dim tmp_str,query1
	%>
	<select class="select" name="<%=selectBoxName%>" <%=chplg%>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택하세요</option>

		<% if shopall <> "" then %>
			<option value='all' <%if selectedId="all" then response.write " selected"%>>전체매장공통</option>
		<% end if %>
	<%
		query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user"
		query1 = query1 & " where 1=1 "

		if shopdiv <> "" then
			query1 = query1 & " and shopdiv in ("&shopdiv&")"
		end if

   		query1 = query1 & " order by isusing desc, convert(int,shopdiv)+10 asc, userid asc"

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
end sub

'//오프라인매장 출고처 구분에 따라서 동적으로 가져 오기		'/2013.10.15 한용민 추가
function drawSelectBoxOffuserdiv_off(selectBoxName, selectedId, userdiv, isusing, shopall, chplg)
	dim tmp_str,query1

	query1 = " select"
	query1 = query1 & " c.userid, c.socname"
	query1 = query1 & " from [db_user].[dbo].tbl_user_c c"
	query1 = query1 & " left join [db_shop].[dbo].tbl_shop_user u"
	query1 = query1 & " 	on c.userid=u.userid"
	query1 = query1 & " where 1=1"

	if isusing <> "" then
		query1 = query1 & " and c.isusing = '"&isusing&"'"
	end if
	if userdiv <> "" then
		query1 = query1 & " and c.userdiv in ("&userdiv&")"
	end if

	query1 = query1 & " order by convert(int,isnull(u.shopdiv,99))+10 asc, c.userid asc"

	'response.write query1 & "<Br>"
	rsget.Open query1,dbget,1
	%>
	<select class="select" name="<%=selectBoxName%>" <%=chplg%>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택하세요</option>

		<% if shopall <> "" then %>
			<option value='all' <%if selectedId="all" then response.write " selected"%>>전체출고처공통</option>
		<% end if %>
	<%
	if  not rsget.EOF  then
	rsget.Movefirst

	do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("userid")) then
			tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&" / "&rsget("socname")&"</option>")
		tmp_str = ""
		rsget.MoveNext
	loop
	end if
	rsget.close
	response.write("</select>")
end function

'//매장이름 가져오기	'//2012.08.09 한용민 생성
function getoffshopname(shopid)
dim sql , i

	if shopid = "" then exit function

	sql = "select top 1"
	sql = sql & " userid ,shopname"
	sql = sql & " from db_shop.dbo.tbl_shop_user"
	sql = sql & " where isusing = 'Y' and userid ='"&trim(shopid)&"'"

	'response.write sqlshopinfo &"<br>"
	rsget.Open sql,dbget,1
	if not rsget.EOF  then
		getoffshopname = rsget("shopname")
	end if
	rsget.close
end function

'//매장 대표 화폐 단위 가져오기		'/2012.03.14 한용민 생성
function getoffshopcurrencyUnit(shopid)
dim sql , i

	if shopid = "" then exit function

	sql = "select top 1"
	sql = sql & " userid ,shopname, currencyUnit"
	sql = sql & " from db_shop.dbo.tbl_shop_user"
	sql = sql & " where isusing = 'Y' and userid ='"&trim(shopid)&"'"

	'response.write sqlshopinfo &"<br>"
	rsget.Open sql,dbget,1
	if not rsget.EOF  then
		getoffshopcurrencyUnit = rsget("currencyUnit")
	end if
	rsget.close
end function

function getShopIDbyPosBrand(imakerid)
    Dim strSql, ret
    ret =""

    strSql = " select top 1 d.shopid from db_shop.dbo.tbl_shop_designer d"
    strSql = strSql & " where d.makerid='"&imakerid&"'"
    strSql = strSql & " and d.comm_cd='B023'"

    'response.write strSql & "<br>"
    rsget.CursorLocation = adUseClient
    rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
	IF Not rsget.EOF THEN
		ret = rsget("shopid")
	END IF
	rsget.Close

	getShopIDbyPosBrand = ret
end function

function getDefaultPosBrand(ishopid)
    Dim strSql, ret
    ret =""

    strSql = " select top 1 d.makerid from db_shop.dbo.tbl_shop_designer d"

    IF (ishopid="") then
        strSql = strSql & " where 1=1"
    ELSE
        strSql = strSql & " where d.shopid='"&ishopid&"'"
    END IF

    strSql = strSql & " and d.comm_cd='B023'"
    strSql = strSql & " order by d.makerid"

    'response.write strSql & "<br>"
    rsget.CursorLocation = adUseClient
    rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
	IF Not rsget.EOF THEN
		ret = rsget("makerid")
	END IF
	rsget.Close

	getDefaultPosBrand = ret
end function

function FnDrawOptPosBrand(ishopid, icompname, ibrand)
    Dim strSql, intLoop, arrBrand
    strSql = " select d.makerid,c.socname_kor, d.shopid"
    strSql = strSql & " from db_shop.dbo.tbl_shop_designer d"
    strSql = strSql & " 	Join db_user.dbo.tbl_user_c c"
    strSql = strSql & " 	on d.makerid=c.userid"

    IF (ishopid="") then
        strSql = strSql & " where 1=1"
    ELSE
        strSql = strSql & " where shopid='"&ishopid&"'"
    END IF

    strSql = strSql & " and d.comm_cd='B023'"
    strSql = strSql & " order by d.makerid"

    'response.write strSql & "<br>"
    rsget.CursorLocation = adUseClient
    rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
	IF Not rsget.EOF THEN
		arrBrand = rsget.getRows()
	END IF
	rsget.Close

	IF Not isArray(arrBrand) then
%>
    <input type="hidden" name='<%= icompname %>' value=''>
    <strong>브랜드 ID 미지정</strong>
<%
	ELSE
%>
    <select name='<%= icompname %>'>
<%
		For intLoop =0 To UBOund(arrBrand,2)
%>
	    <option value="<%=arrBrand(0,intLoop)%>" <%IF ibrand = arrBrand(0,intLoop) THEN%>selected<%END IF%>><%=arrBrand(1,intLoop)%> (<%=arrBrand(2,intLoop)%>)</option>
<%
		Next
%>
    </select>
<%
    END IF
end function

'// 오프라인 샵 option list (2006.12.08 정윤정 생성)
'// 2007.03.13 정윤정수정: 화면표시 순서와 별개로 사용여부로 표시 여부 구분
Function fnOptShopName(ByVal selShopID)
	Dim strSql, arrShop, intLoop
	strSql = "SELECT userid, shopname FROM [db_shop].[dbo].[tbl_shop_user] WHERE  vieworder <> 0 and isUsing='Y' ORDER BY vieworder "

    'response.write strSql & "<br>"
	rsget.Open strSql,dbget,1
	IF Not rsget.EOF THEN
		arrShop = rsget.getRows()
	END IF
	rsget.Close

	If isArray(arrShop) THEN
		For intLoop =0 To UBOund(arrShop,2)
%>
	<option value="<%=arrShop(0,intLoop)%>" <%IF selShopID = arrShop(0,intLoop) THEN%>selected<%END IF%>><%=arrShop(1,intLoop)%></option>
<%
		Next
	END If
End Function

'// 오프라인 샵 공통코드 가져오기 (2006.12.08 정윤정)	'//사용안함
'//디비화 시킴 drawoffshop_commoncode 사용
Function fnOptCommonCode(ByVal sKind, byVal selcodeid)
	Dim strSql, arrCode, intLoop
	strSql = "SELECT codeid, codename FROM [db_shop].[dbo].[tbl_offshop_commoncode] WHERE codekind ='"&sKind&"' and useyn ='Y' "

    'response.write strSql & "<br>"
	rsget.Open strSql,dbget,1
	IF Not rsget.EOF THEN
		arrCode = rsget.getRows()
	END IF
	rsget.Close

	If isArray(arrCode) THEN
		For intLoop =0 To UBound(arrCode,2)
%>
	<option value="<%=arrCode(0,intLoop)%>" <%IF selcodeid = arrCode(0,intLoop) THEN%>selected<%END IF%>><%=arrCode(1,intLoop)%></option>
<%
		Next
	END IF
End Function

'// 오프라인 샵 공통코드명 가져오기 (2006.12.11 정윤정)
function fnGetCommonCode(ByVal sKind, ByVal sId)
	Dim strSql
	strSql ="SELECT codename FROM [db_shop].[dbo].[tbl_offshop_commoncode] WHERE codekind ='"&sKind&"' and codeid='"&sId&"' and useyn ='Y' "

    'response.write strSql & "<br>"
	rsget.Open strSql,dbget,1
	IF Not rsget.EOF THEN
		fnGetCommonCode = rsget("codename")
	END IF
	rsget.Close
End Function

'//오프라인 공용코드 (기존디비 멀티로 사용가능하게 수정) 	'/2012.08.02 한용민 생성
function drawoffshop_commoncode(selectBoxName, selectedId, codekind, codegroup, maincodeid, chplg)
   dim tmp_str, sql , sqlsearch

	if selectBoxName = "" or codekind = "" then exit function

	'//코드그룹이 메인인경우
	if codegroup = "MAIN" then
		sqlsearch = sqlsearch & " and codegroup='MAIN'"

		if maincodeid <> "" then
			sqlsearch = sqlsearch & " and maincodeid='"&maincodeid&"'"
		end if

	'//코드그룹이 서브인경우
	elseif codegroup = "SUB" then
		sqlsearch = sqlsearch & " and codegroup='SUB'"

		if maincodeid <> "" then
			sqlsearch = sqlsearch & " and maincodeid='"&maincodeid&"'"
		end if
	end if

	sql = "select"
	sql = sql & " idx ,codekind ,codegroup ,maincodeid ,codeid ,codename ,useyn ,orderno"
	sql = sql & " from db_shop.dbo.tbl_offshop_commoncode"
	sql = sql & " where useyn='Y' and codekind='"&codekind&"' " & sqlsearch
	sql = sql & " order by orderno asc , codeid asc"

	'response.write sql &"<Br>"
	%>
	<select class="select" name="<%=selectBoxName%>" <%= chplg %>>
		<option value='' <%if selectedId="" then response.write " selected"%>>SELECT</option>
	<%
	rsget.Open sql,dbget,1

	if not rsget.EOF then
		rsget.Movefirst

		do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("codeid")) then
			tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("codeid")&"' "&tmp_str&">"& db2html(rsget("codename"))&"</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
	end if
	rsget.close

	response.write("</select>")
end function

'//오프라인 공용코드 (기존디비 멀티로 사용가능하게 수정) 배열버전	'/2012.08.02 한용민 생성
function getoffshop_commoncodearray(codekind, codegroup, maincodeid)
   dim sql , sqlsearch, arrList

	'//코드그룹이 메인인경우
	if codegroup = "MAIN" then
		sqlsearch = sqlsearch & " and codegroup='MAIN'"

		if maincodeid <> "" then
			sqlsearch = sqlsearch & " and maincodeid='"&maincodeid&"'"
		end if

	'//코드그룹이 서브인경우
	elseif codegroup = "SUB" then
		sqlsearch = sqlsearch & " and codegroup='SUB'"

		if maincodeid <> "" then
			sqlsearch = sqlsearch & " and maincodeid='"&maincodeid&"'"
		end if
	end if

	sql = "select"
	sql = sql & " idx ,codekind ,codegroup ,maincodeid ,codeid ,codename ,useyn ,orderno"
	sql = sql & " from db_shop.dbo.tbl_offshop_commoncode"
	sql = sql & " where useyn='Y' and codekind='"&codekind&"' " & sqlsearch
	sql = sql & " order by orderno asc , codeid asc"

	'response.write sql &"<Br>"
	rsget.Open sql,dbget,1
	if not rsget.EOF then
		arrList = rsget.getrows()
	end if
	rsget.close

	getoffshop_commoncodearray = arrList
end function

'//오프라인 공용코드 해당코드가 같은 그룹인지 TRUE FALSE 로 반환 	'/2012.08.02 한용민 생성
function getoffshop_commoncodegroup(codekind, codegroup, maincodeid ,codeid)
   dim tmp_str, sql , sqlsearch
	tmp_str = FALSE

	if codekind = "" or codegroup = "" then exit function

	'//코드그룹이 메인인경우
	if codegroup = "MAIN" then
		sqlsearch = sqlsearch & " and codegroup='MAIN'"

		if maincodeid <> "" then
			sqlsearch = sqlsearch & " and maincodeid='"&maincodeid&"'"
		end if

	'//코드그룹이 서브인경우
	elseif codegroup = "SUB" then
		sqlsearch = sqlsearch & " and codegroup='SUB'"

		if maincodeid <> "" then
			sqlsearch = sqlsearch & " and maincodeid='"&maincodeid&"'"
		end if
		if codeid <> "" then
			sqlsearch = sqlsearch & " and codeid='"&codeid&"'"
		end if
	end if

	sql = "select"
	sql = sql & " idx ,codekind ,codegroup ,maincodeid ,codeid ,codename ,useyn ,orderno"
	sql = sql & " from db_shop.dbo.tbl_offshop_commoncode"
	sql = sql & " where useyn='Y' and codekind='"&codekind&"' " & sqlsearch
	sql = sql & " order by orderno asc , codeid asc"

	'response.write sql &"<Br>"
	rsget.Open sql,dbget,1
	if not rsget.EOF then
		rsget.Movefirst
		do until rsget.EOF
			tmp_str = TRUE
		rsget.MoveNext
		loop
	end if
	rsget.close

	getoffshop_commoncodegroup = tmp_str
end function

'// 오프라인 샵 권한 관리 (2006.12.11 정윤정)
function fnChkAuth(ByVal BctDiv, ByVal BctID, ByVal BctBigo)
	IF ( BctDiv > 100 and BctDiv < 200 ) THEN	'가맹 - 개인아이디
		fnChkAuth = BctBigo
	ELSEIF ( BctDiv > 200 and BctDiv < 300 ) THEN	'직영
		fnChkAuth = BctBigo
	ELSEIF ( BctDiv > 500 and BctDiv < 600 ) THEN	'가맹 - 가맹점아이디
		fnChkAuth = BctID
	ELSE
		fnChkAuth = ""
	END IF
End function

public function GetImageFolerName(byval FItemID)
    GetImageFolerName = GetImageSubFolderByItemid(FItemID)
	''GetImageFolerName = "0" + CStr(FItemID\10000)
end function

'//직영점 가맹점 매장 일경우 매장 정보 가져옴 	'/2011.02.17 한용민 추가
function getoffshopdiv(shopid)
dim sql , i

	if shopid = "" then exit function

	sql = "select top 1"
	sql = sql & " userid ,shopname ,shopdiv"
	sql = sql & " from db_shop.dbo.tbl_shop_user"
	sql = sql & " where isusing = 'Y' and userid ='"&trim(shopid)&"'"

	'response.write sqlshopinfo &"<br>"
	rsget.Open sql,dbget,1

	if not rsget.EOF  then
		getoffshopdiv = rsget("shopdiv")
	end if

	rsget.close

end function

'/어드민 직책 가져옴 '/파트장 :5 , 점장 : 6		'/2011.03.18 한용민 추가
function getjob_sn(empno , userid)
dim sqlStr ,sqlsearch

	if empno = "" and userid = "" then exit function

	if empno <> "" then
		sqlsearch = sqlsearch & " and t.empno = '"&empno&"'"
	end if
	if userid <> "" then
		sqlsearch = sqlsearch & " and p.id = '"&userid&"'"
	end if

	sqlStr = "select top 1 "
	sqlStr = sqlStr & " t.job_sn  , t.username ,p.id"
	sqlStr = sqlStr & " from db_partner.dbo.tbl_user_tenbyten t"
	sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	sqlStr = sqlStr & " 	on t.userid = p.id"
	sqlStr = sqlStr & " where p.isusing = 'Y' " & sqlsearch

	' 퇴사예정자 처리	' 2018.10.16 한용민
	sqlStr = sqlStr & "	and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf

	'response.write sqlStr &"<br>"
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF  then
		getjob_sn = rsget("job_sn")
	end if
	rsget.close
end function

'/어드민 권한 가져옴 	'/2011.03.18 한용민 추가
function getlevel_sn(empno , userid)
dim sqlStr ,sqlsearch

	if empno = "" and userid = "" then exit function

	if empno <> "" then
		sqlsearch = sqlsearch & " and t.empno = '"&empno&"'"
	end if
	if userid <> "" then
		sqlsearch = sqlsearch & " and p.id = '"&userid&"'"
	end if

	sqlStr = "select top 1 "
	sqlStr = sqlStr & " t.username ,p.id"
	'sqlStr = sqlStr & " ,p.level_sn"
	sqlStr = sqlStr & " ,(case when u.userid is null then p.level_sn else 999 end) as level_sn"
	sqlStr = sqlStr & " from db_partner.dbo.tbl_partner p"
	sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten t"
	sqlStr = sqlStr & " 	on p.id=t.userid"

	' 퇴사예정자 처리	' 2018.10.16 한용민
	sqlStr = sqlStr & "		and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf
	sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u"
	sqlStr = sqlStr & "		on p.id = u.userid"
	sqlStr = sqlStr & "		and shopdiv not in ('1')"
	sqlStr = sqlStr & " where p.isusing = 'Y' " & sqlsearch

	'response.write sqlStr &"<br>"
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF  then
		getlevel_sn = rsget("level_sn")
	end if
	rsget.close
end function

'/어드민 부서정보 가져옴	'/2011.03.22 한용민 추가
function getpart_sn(empno , userid)
dim sqlStr ,sqlsearch

	if empno = "" and userid = "" then exit function

	if empno <> "" then
		sqlsearch = sqlsearch & " and t.empno = '"&empno&"'"
	end if
	if userid <> "" then
		sqlsearch = sqlsearch & " and p.id = '"&userid&"'"
	end if

	sqlStr = "select top 1 "
	sqlStr = sqlStr & " t.part_sn , t.username ,p.id"
	sqlStr = sqlStr & " from db_partner.dbo.tbl_user_tenbyten t"
	sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	sqlStr = sqlStr & " 	on t.userid = p.id"
	sqlStr = sqlStr & " 	and p.isusing = 'Y'"
	sqlStr = sqlStr & " where 1=1 " & sqlsearch

	' 퇴사예정자 처리	' 2018.10.16 한용민
	sqlStr = sqlStr & "	and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf

	'response.write sqlStr &"<br>"
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF  then
		getpart_sn = rsget("part_sn")
	end if
	rsget.close
end function

'/매장에 계약된 브랜드만(조건추가)		'/2011.12.08 한용민 생성
function drawBoxDirectIpchulOffShopByMakerchfg(selectBoxName,selectedId,makerid,chfg,comm_cd)
	dim tmp_str,query1

	if makerid = "" then exit function

	query1 = " select u.userid,u.shopname from [db_shop].[dbo].tbl_shop_user u"
	query1 = query1 & "      Join [db_shop].[dbo].tbl_shop_designer d"
	query1 = query1 & "      on u.userid=d.shopid"
	query1 = query1 & " where u.isusing='Y' "
	query1 = query1 & " and d.makerid='" + makerid + "'"

	if comm_cd <> "" then
		query1 = query1 & " and d.comm_cd in ("&comm_cd&")"   ''업체위탁 매장매입만 가능 //가맹점 매입 추가
	end if

	query1 = query1 & " and u.userid<>'streetshop000'"
	query1 = query1 & " and u.userid<>'streetshop800'"
	query1 = query1 & " and u.userid<>'streetshop870'"
	query1 = query1 & " and u.userid<>'streetshop700'"
	query1 = query1 & " order by u.isusing desc, convert(int,u.shopdiv)+10 asc, u.userid asc"

	'response.write query1 &"<Br>"
	rsget.Open query1,dbget,1
%>
	<select class="select" name="<%=selectBoxName%>" <%=chfg%>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
<%
	   if not rsget.EOF  then
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
end function

'/매장에 계약된 브랜드 수	'/2013.02.21 한용민 생성
function getcontractbranditemcount(shopid,makerid)
	dim sqlStr, sqlsearch

	if shopid="" or makerid="" then
		getcontractbranditemcount = 0
		exit function
	end if

	sqlStr = "select"
	sqlStr = sqlStr + " count(*) as cnt"
	sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_designer sd"
	sqlStr = sqlStr + " where sd.shopid='"&shopid&"'"
	sqlStr = sqlStr + " and sd.makerid='"&makerid&"'"
	
'	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
'	sqlStr = sqlStr + " join db_shop.dbo.tbl_shop_designer sd"
'	sqlStr = sqlStr + " 	on s.makerid=sd.makerid"
'	sqlStr = sqlStr + " 	and sd.shopid='"&shopid&"'"
'	sqlStr = sqlStr + " 	and s.makerid='"&makerid&"'"
'	sqlStr = sqlStr + " where s.isusing='Y'"

	'response.write query1 &"<Br>"
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF  then
		getcontractbranditemcount = rsget("cnt")
	else
		getcontractbranditemcount = 0
	end if
   rsget.close
end function

'//상품 매장 계약 정보 가져 오기		'/2012.03.14 한용민 생성
Function getitemshopcontractinfo(shopdiv, shopid, byval makerid)
	Dim sql

	if makerid="" then exit function

	sql = "SELECT"
	sql = sql & " sd.comm_cd, u.shopname, sd.defaultmargin , sd.defaultsuplymargin"
	sql = sql & " from db_shop.dbo.tbl_shop_designer sd"
	sql = sql & " join db_shop.dbo.tbl_shop_user u"
	sql = sql & " 	on sd.shopid=u.userid"

	if shopid<>"" then
		sql = sql & " 	and sd.shopid = '"& shopid &"'"
	end if
	if shopdiv<>"" then
		sql = sql & " 	and u.shopdiv in ("& shopdiv &")"
	end if

	sql = sql & " where sd.makerid='"& makerid &"'"
   	sql = sql & " order by u.isusing desc, convert(int,u.shopdiv)+10 asc, u.userid asc"

	'response.write sql &"<Br>"
	rsget.Open sql,dbget,1
	IF Not rsget.EOF THEN
		getitemshopcontractinfo = rsget.getRows()
	END IF
	rsget.Close
End Function

Sub drawBoxDirectIpchulOffShopByMaker(selectBoxName,selectedId,makerid)
	dim tmp_str,query1

	query1 = " select u.userid,u.shopname from [db_shop].[dbo].tbl_shop_user u"
	query1 = query1 & "      Join [db_shop].[dbo].tbl_shop_designer d"
	query1 = query1 & "      on u.userid=d.shopid"
	query1 = query1 & " where u.isusing='Y' "
	query1 = query1 & " and d.makerid='" + makerid + "'"
	query1 = query1 & " and d.comm_cd in ('B012','B022','B023')"   ''업체위탁 매장매입만 가능 //가맹점 매입 추가
	query1 = query1 & " and u.userid<>'streetshop000'"
	query1 = query1 & " and u.userid<>'streetshop800'"
	query1 = query1 & " and u.userid<>'streetshop870'"
	query1 = query1 & " and u.userid<>'streetshop700'"
	query1 = query1 & " order by u.isusing desc, convert(int,u.shopdiv)+10 asc, u.userid asc"

	'response.write query1 & "<Br>"
	rsget.Open query1,dbget,1
	%>
	<select class="select" name="<%=selectBoxName%>">
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
	%>
	<%
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
end sub

''//업체위탁/매장매입 계약조건 아닌것.
Sub drawSelectBoxShopjumunDesignerNotUpche(selectBoxName,selectedId,shopid,suplyer,comm_cd)
   dim tmp_str,query1

    query1 = " select d.makerid, c.socname_kor ,d.comm_cd from [db_shop].[dbo].tbl_shop_designer d"
	query1 = query1 + " left join [db_user].[dbo].tbl_user_c c on d.makerid=c.userid"
	query1 = query1 + " where d.shopid='" + shopid + "'"

	if suplyer="10x10" then
		query1 = query1 + " and d.comm_cd not in ('B012','B022')"   ''업체위탁 매장매입만 가능
   	else
   	    query1 = query1 + " and d.comm_cd not in ('B012','B022')"   ''업체위탁 매장매입만 가능
   		query1 = query1 + " and d.makerid='" + suplyer + "'"
   	end if

	if comm_cd <> "" then
		query1 = query1 + " and d.comm_cd in ("&comm_cd&")"
	end if

	query1 = query1 + " order by d.makerid"

	'response.write query1 & "<Br>"
	rsget.Open query1,dbget,1
	%>
	<select class='select' name="<%=selectBoxName%>">
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
	<%
	if not rsget.EOF  then
		rsget.Movefirst

		do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("makerid")) then
		   tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("makerid")&"' "&tmp_str&">"&rsget("makerid")&"/"&db2html(rsget("socname_kor"))&"["&GetJungsanGubunName(rsget("comm_cd"))&"]</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
	end if

rsget.close
response.write("</select>")
End Sub

''매장 기주문 수량 업데이트 (수정요망)		'/2012.01.10 이상구 생성
function PreOrderUpdateByBrand_off(masteridx,targetid,shopid)
	dim sqlStr

	if masteridx = "" or targetid = ""  then exit function

	sqlStr = " exec db_summary.[dbo].[sp_Ten_Shop_Stock_PreOrderUpdate] '"&masteridx&"','"&targetid&"','"&shopid&"'"

	'response.write sqlStr &"<Br>"
	dbget.Execute sqlStr
end function

''매장 재고 수량 업데이트		'/2012.01.10 이상구 생성
function recentUpdate_off(masteridx,targetid,shopid)
	dim sqlStr

	if masteridx = "" or targetid = ""  then exit function

	sqlStr = " exec db_summary.[dbo].[sp_Ten_Shop_Stock_RecentUpdate] '"&masteridx&"','"&targetid&"','"&shopid&"'"

	'response.write sqlStr &"<Br>"
	dbget.Execute sqlStr
end function

'// 매출 기준일 처리	'/한용민 생성
function drawmaechuldatefg(selectBoxName,selectedId,changefg)
%>
<select name="<%=selectBoxName%>" <%=changefg%>>
	<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
	<option value='maechul' <%if selectedId="maechul" then response.write " selected"%>>매출기준일</option>
	<option value='jumun' <%if selectedId="jumun" then response.write " selected"%>>주문일</option>
</select>
<%
end function

'/매장 입출고 마스터 리스트		'/2011.08.25 한용민 생성
function drawipchulmaster(selectBoxName,selectedId,shopid ,makerid ,changefg,allgubun)
	dim tmp_str,query1

	query1 = "select top 500 m.idx , m.shopid , m.chargeid"
	query1 = query1 & " from [db_shop].[dbo].tbl_shop_ipchul_master m"
	query1 = query1 & " where m.deleteyn='N'"

	if shopid <> "" then
		query1 = query1 & " and m.shopid = '"&shopid&"'"
	end if
	if makerid <> "" then
		query1 = query1 & " and m.chargeid = '"&makerid&"'"
	end if

	query1 = query1 & " order by m.idx desc"

	rsget.Open query1,dbget,1
%>
	<select name="<%=selectBoxName%>" <%=changefg %> <% if rsget.recordcount = 0 then response.write " disabled" %>>
	<% if allgubun = "Y" then %>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
	<% end if %>
	<%
	if  not rsget.EOF  then
		rsget.Movefirst

		do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("idx")) then
			tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("idx")&"' "&tmp_str&">"&rsget("idx")&" - "&rsget("chargeid")&" ["&rsget("shopid")&"]</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
	else
	%>
		<option value='' <%if selectedId="" then response.write " selected"%>>내역이 없습니다</option>
	<%
	end if
	rsget.close
	response.write("</select>")
end function

'/물류 입출고 마스터 리스트		'/2011.08.25 한용민 추가
function drawcentervsshopipchulmaster(selectBoxName,selectedId,shopid  ,changefg,allgubun)
	dim tmp_str,query1

	query1 = "select top 100 m.idx , m.targetid ,m.targetname , m.baljuid , m.baljuname ,m.baljucode"
	query1 = query1 & " from [db_storage].[dbo].tbl_ordersheet_master m"
	query1 = query1 & " where m.deldt is Null and m.divcode in ('501','502','503') "

	if shopid <> "" then
		query1 = query1 & " and m.baljuid = '"&shopid&"'"
	end if

	query1 = query1 & " order by m.idx desc"

	rsget.Open query1,dbget,1
%>
	<select name="<%=selectBoxName%>" <%=changefg %> <% if rsget.recordcount = 0 then response.write " disabled" %>>
	<% if allgubun = "Y" then %>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
	<% end if %>
	<%
	if  not rsget.EOF  then
		rsget.Movefirst

		do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("idx")) then
			tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("idx")&"' "&tmp_str&">"&rsget("baljucode")&" - "&rsget("baljuname")&"</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
	else
	%>
		<option value='' <%if selectedId="" then response.write " selected"%>>내역이 없습니다</option>
	<%
	end if
	rsget.close
	response.write("</select>")
end function

'// 매출 기준일 처리	'한용민 생성
function drawmaechul_datefg(selectBoxName,selectedId,changefg)
%>
<select name="<%=selectBoxName%>" <%=changefg%>>
	<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
	<option value='maechul' <%if selectedId="maechul" then response.write " selected"%>>매출기준일</option>
	<option value='jumun' <%if selectedId="jumun" then response.write " selected"%>>주문일</option>
</select>
<%
end function

'/요일		'2011.11.17 한용민 생성
function drawweekday_select(selectBoxName,selectedId,changefg)
%>
<select name="<%=selectBoxName%>" <%=changefg%>>
	<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
	<option value='1' <%if selectedId="1" then response.write " selected"%>>일</option>
	<option value='2' <%if selectedId="2" then response.write " selected"%>>월</option>
	<option value='3' <%if selectedId="3" then response.write " selected"%>>화</option>
	<option value='4' <%if selectedId="4" then response.write " selected"%>>수</option>
	<option value='5' <%if selectedId="5" then response.write " selected"%>>목</option>
	<option value='6' <%if selectedId="6" then response.write " selected"%>>금</option>
	<option value='7' <%if selectedId="7" then response.write " selected"%>>토</option>
</select>
<%
end function

'/요일 반환		'2012.10.04 한용민 생성
Function getweekday_select(v)
	if v = "1" then
		getweekday_select = "<font color=""red"">일</font>"
	elseif v = "2" then
		getweekday_select = "월"
	elseif v = "3" then
		getweekday_select = "화"
	elseif v = "4" then
		getweekday_select = "수"
	elseif v = "5" then
		getweekday_select = "목"
	elseif v = "6" then
		getweekday_select = "금"
	elseif v = "7" then
		getweekday_select = "<font color=""blue"">토</font>"
	else
		getweekday_select = v
	end if
End Function

'//날짜(2012-10-25)를 넣으면 요일 출력		'/2012.10.25 한용민 생성
function getweekend(yyyymmdd)
	if yyyymmdd = "" then exit function

	if right(FormatDateTime(yyyymmdd,1),3) = "토요일" then
		getweekend = "<font color='blue'>"&Replace(Right(FormatDateTime(yyyymmdd,1),3),"요일","")&"</font>"
	elseif right(FormatDateTime(yyyymmdd,1),3) = "일요일" then
		getweekend = "<font color='red'>"&Replace(Right(FormatDateTime(yyyymmdd,1),3),"요일","")&"</font>"
	else
		getweekend = Replace(Right(FormatDateTime(yyyymmdd,1),3),"요일","")
	end if
end function

'//날짜(2012-10-25)를 넣으면 주말은 색이 틀림		'/2012.10.25 한용민 생성
function getweekendcolor(yyyymmdd)
	if yyyymmdd = "" then exit function

	if right(FormatDateTime(yyyymmdd,1),3) = "토요일" then
		getweekendcolor = "<font color='blue'>"&yyyymmdd&"</font>"
	elseif right(FormatDateTime(yyyymmdd,1),3) = "일요일" then
		getweekendcolor = "<font color='red'>"&yyyymmdd&"</font>"
	else
		getweekendcolor = yyyymmdd
	end if
end function

'//파트 구성원 가져 오기	'/2011.11.24 한용민 생성
function getpartpeople(ByVal boxname, ByVal selectid, ByVal chscript ,ByVal part_sn)
Dim sqlStr ,sqlsearch, arrList, intLoop

	if part_sn <> "" then
		sqlsearch = sqlsearch & " and t.part_sn in ("&part_sn&")"
	end if

	sqlStr = "select top 500"
	sqlStr = sqlStr & " t.empno  , t.username ,p.id"
	sqlStr = sqlStr & " from db_partner.dbo.tbl_user_tenbyten t"
	sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	sqlStr = sqlStr & " 	on t.userid = p.id"
	sqlStr = sqlStr & " 	and p.isusing = 'Y'"
	sqlStr = sqlStr & " where 1=1 " & sqlsearch

	' 퇴사예정자 처리	' 2018.10.16 한용민
	sqlStr = sqlStr & "	and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf
	sqlStr = sqlStr & " order by t.part_sn asc, t.posit_sn asc ,t.username asc"

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget
	IF not rsget.eof THEN
		arrList = rsget.getRows()
	End IF
	rsget.close

	response.write "<select name='"&boxname&"' "&chscript&">"
%>
	<option value="" <% if selectid = "" then response.write " selected" %>>선택</option>
<%
	If isArray(arrList) THEN

	For intLoop = 0 To UBound(arrList,2)
%>
	<option value="<%=arrList(0,intLoop)%>" <%if arrList(0,intLoop) = selectid then %> selected<%end if%>><%=arrList(1,intLoop)%></option>
<%
	Next

	End IF
	response.write "</select>"
End function

'//전사원 선택창	'/2012.07.17 한용민 생성
function gettenbytenuser(ByVal boxname, ByVal selectid, ByVal chplg ,byval part_sn ,byval scriptplg)
	Dim strSql, arrList, intLoop , tmpuserid ,tmpusername

	strSql = " SELECT top 1"
	strSql = strSql & " userid, username"
	strSql = strSql & " FROM db_partner.dbo.tbl_user_tenbyten t"
	strSql = strSql & " WHERE userid <> ''"
	strSql = strSql & " and userid = '"&selectid&"'"

	'response.write strSql &"<br>"
	rsget.Open strSql,dbget
	IF not rsget.eof THEN
		arrList = rsget.getRows()
	End IF
	rsget.close

	IF isArray(arrList) THEN
		tmpuserid = arrList(0,0)
		tmpusername = arrList(1,0)
	end if
%>
	<input type="text" name="username" value="<%= tmpusername %>" readonly size="10" class="text" <%= chplg %>>
	<input type="button" class="button" value="선택" onClick="pop_tenbytenuser()">
	<input type="hidden" name="<%= boxname %>" value="<%= tmpuserid %>">

	<script language='javascript'>
		function pop_tenbytenuser(){

			<%= scriptplg %>

			var userid = document.getElementsByName("<%= boxname %>")[0].value;
			var username = document.getElementsByName("username")[0].value;

			var pop_tenbytenuser = window.open('/common/offshop/member/PoptenbytenuserList.asp?worker='+userid+'&team=<%= part_sn %>&boxname=<%=boxname%>&username='+username,'pop_tenbytenuser','width=570,height=570,scrollbars=yes');
			pop_tenbytenuser.focus();
		}
	</script>
<%
End function

'//매장 파트 구성원 가져 오기		'/2011.11.24 한용민 생성
function drawpartpeopleshop(ByVal boxname, ByVal selectid, ByVal chscript)
Dim sqlStr ,sqlsearch, arrList, intLoop

	sqlStr = "select top 500"
	sqlStr = sqlStr & " ut.userid  , ut.username ,ut.userid"
	sqlStr = sqlStr + " from db_partner.dbo.tbl_user_tenbyten ut"
	sqlStr = sqlStr + " join db_partner.dbo.tbl_partner_shopuser ps"
	sqlStr = sqlStr + " 	on ps.empno = ut.empno"
	sqlStr = sqlStr + " where ut.isusing=1"

	' 퇴사예정자 처리	' 2018.10.16 한용민
	sqlStr = sqlStr & "	and (ut.statediv ='Y' or (ut.statediv ='N' and datediff(dd,ut.retireday,getdate())<=0))" & vbcrlf
	sqlStr = sqlStr & " order by ut.part_sn asc, ut.posit_sn asc ,ut.username asc"

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget
	IF not rsget.eof THEN
		arrList = rsget.getRows()
	End IF
	rsget.close

	response.write "<select name='"&boxname&"' "&chscript&">"
%>
	<option value="" <% if selectid = "" then response.write " selected" %>>선택</option>
<%
	If isArray(arrList) THEN

	For intLoop = 0 To UBound(arrList,2)
%>
	<option value="<%=arrList(0,intLoop)%>" <%if arrList(0,intLoop) = selectid then %> selected<%end if%>><%=arrList(1,intLoop)%></option>
<%
	Next

	End IF
	response.write "</select>"
End function

'//화폐단위		'/2011.12.13 한용민 생성
function DrawexchangeRate(selectBoxName,selectedId,changeFlag)
	dim tmp_str,query1

	query1 = " select"
	query1 = query1 & " currencyUnit ,exchangeRate ,basedate ,currencyChar"
	query1 = query1 & " from db_shop.dbo.tbl_shop_exchangeRate"

	'response.write query1 &"<Br>"
	rsget.Open query1,dbget,1
	%>
	<select class="select" name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value='' <%if selectedId="" then response.write " selected"%>>CHOICE</option><%

	if not rsget.EOF then
		rsget.Movefirst

		do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("currencyUnit")) then
			tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("currencyUnit")&"' "&tmp_str&">"&rsget("currencyUnit")&"["&rsget("currencyChar")&"]</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
	end if
	rsget.close

	response.write("</select>")
end function

'/매장 배송 구분 셀렉트박스 버전	'/2012-05-18 한용민 생성
Sub drawSelectBoxShopBeasongDiv(selectBoxName,selectedId)
	dim tmp_str,query1
   %>
	<select class="select" name="<%=selectBoxName%>">
		<option value="">선택
		<option value="8" <% if (selectedId="0") then response.write "selected" %> >매장판매
		<option value="2" <% if (selectedId="2") then response.write "selected" %> >업체배송
	</select>
<%
End Sub

'/매장 배송 구분 체크박스 버전		'/2012-05-18 한용민 생성
Sub drawCheckBoxShopBeasongDiv(checkBoxName,checkedId)
	dim tmp_str,query1
   %>

   <!--<input type="checkbox" name="<%'=checkBoxName%>" value="0" <%' if (checkedId = "0") then %>checked<%' end if %>> 매장판매-->
   <input type="checkbox" name="<%=checkBoxName%>" value="2" <% if (checkedId = "2") then %>checked<% end if %>> 업체배송
<%
End Sub

'//디비화 시킴 drawoffshop_commoncode 사용		'//사용안함
function DrawShopDivCombo(selBoxName,selVal)
%>
    <select class='select' name="<%= selBoxName %>" >
	<option value='' <% if selVal="" then response.write " selected" %> >전체</option>
	<option value='1' <% if selVal="1" or selVal="2" then response.write " selected" %> >직영</option>
	<option value='3' <% if selVal="3" or selVal="4" then response.write " selected" %> >가맹</option>
	<option value='5' <% if selVal="5" or selVal="6" then response.write " selected" %> >도매</option>
	<option value='7' <% if selVal="7" or selVal="8" then response.write " selected" %> >해외</option>
	<option value='9' <% if selVal="9" then response.write " selected" %> >기타</option>
	</select>
<%
end Function

'/지정매장과 계약마진 같은 매장 가져오기	'/2011.12.07 한용민 생성
function drawBoxshopipchulcontract(selectBoxName,selectedId,makerid,gubunshopid,chevent)
   dim tmp_str,query1

	if makerid = "" or gubunshopid = "" then exit function

	query1 = " select"
	query1 = query1 & " u.userid,u.shopname"
	query1 = query1 & " from [db_shop].[dbo].tbl_shop_user u"
	query1 = query1 & " Join [db_shop].[dbo].tbl_shop_designer d"
	query1 = query1 & " 	on u.userid=d.shopid"
	query1 = query1 & " left join ("
	query1 = query1 & " 	select top 1 "
	query1 = query1 & " 	shopid , defaultmargin,defaultsuplymargin"
	query1 = query1 & " 	from [db_shop].[dbo].tbl_shop_designer"
	query1 = query1 & " 	where shopid = '"&gubunshopid&"'"
	query1 = query1 & " 	and makerid = '"&makerid&"'"
	query1 = query1 & "	) as t"
	query1 = query1 & "		on d.defaultmargin = t.defaultmargin"
	query1 = query1 & "		and d.defaultsuplymargin = t.defaultsuplymargin"
	query1 = query1 & " where u.isusing='Y'"
	query1 = query1 & " and d.makerid='" + makerid + "'"
	query1 = query1 & " and d.comm_cd in ('B012','B022')"   ''업체위탁 매장매입만 가능
	query1 = query1 & " and u.userid<>'streetshop000'"
	query1 = query1 & " and u.userid<>'streetshop800'"
	query1 = query1 & " and u.userid<>'streetshop870'"
	query1 = query1 & " and u.userid<>'streetshop700'"
	query1 = query1 & " and t.shopid is not null"
	query1 = query1 & " and u.userid <> '"&gubunshopid&"'"	'지정매장은 제외
	query1 = query1 & " and u.shopdiv=1"	'직영매장만
	query1 = query1 & " order by u.isusing desc, convert(int,u.shopdiv)+10 asc, u.userid asc"

	'response.write query1 &"<Br>"
	rsget.Open query1,dbget,1
	%>
	<select class="select" name="<%=selectBoxName%>" <%= chevent %>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%

	if not rsget.EOF then
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
end function

'/매입구분별로 동적으로 가져 오기	'/2011.12.13 한용민 생성
function drawSelectBoxDesignerOffWitakContract(selectBoxName,selectedId,shopid,comm_cd,chevent)
   dim tmp_str,query1

	if shopid = "" then exit function

  	query1 = "select"
  	query1 = query1 + " distinct d.makerid, c.socname_kor"
  	query1 = query1 + " from [db_shop].[dbo].tbl_shop_designer d"
	query1 = query1 + " left join [db_user].[dbo].tbl_user_c c"
	query1 = query1 + " 	on d.makerid=c.userid"
	query1 = query1 + " where d.shopid='" + shopid + "'"

	if comm_cd <> "" then
		query1 = query1 + " and d.comm_cd in ("&comm_cd&") " ''B012','B022','B023'
	end if

	query1 = query1 + " order by d.makerid"

	'response.write query1 & "<Br>"
	rsget.Open query1,dbget,1
	%>
	<select name="<%=selectBoxName%>" <%= chevent %>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%

		if  not rsget.EOF  then
			rsget.Movefirst

			do until rsget.EOF
			if Lcase(selectedId) = Lcase(rsget("makerid")) then
				tmp_str = " selected"
			end if
			response.write("<option value='"&rsget("makerid")&"' "&tmp_str&">"&rsget("makerid")&" ["&db2html(rsget("socname_kor"))&"]</option>")
			tmp_str = ""
			rsget.MoveNext
			loop
		end if
	response.write("</select>")
	rsget.close
End function

'/할인유형구분	: 샾주문디테일 discountKind		'/2012.02.10 한용민 생성
function DrawdiscountKind(selBoxName,selVal,chplg)
%>
    <select class='select' name="<%= selBoxName %>" <%= chplg %>>
		<option value='' <% if selVal="" then response.write " selected" %> >전체</option>
		<option value='1' <% if selVal="1" then response.write " selected" %> >직원할인</option>
		<option value='2' <% if selVal="2" then response.write " selected" %> >샘플할인</option>
		<option value='4' <% if selVal="4" then response.write " selected" %> >패키지할인</option>
		<option value='5' <% if selVal="5" then response.write " selected" %> >고객5%할인</option>
		<option value='6' <% if selVal="6" then response.write " selected" %> >고객할인</option>
		<option value='7' <% if selVal="7" then response.write " selected" %> >tenday10%할인</option>
		<option value='9' <% if selVal="9" then response.write " selected" %> >두타쿠폰북</option>
	</select>
<%
end Function

'// 신규업체 기준처리		'/2012.02.09 한용민 생성
function drawnewupche_datefg(selectBoxName,selectedId,changefg)
%>
<select name="<%=selectBoxName%>" <%=changefg%>>
	<option value='' <%if selectedId="" then response.write " selected"%>>전체</option>
	<option value='ipgo' <%if selectedId="ipgo" then response.write " selected"%>>입고일</option>
	<option value='reg' <%if selectedId="reg" then response.write " selected"%>>등록일</option>
</select>
<%
end function

'//매장 브랜드 계약 정보 가져 오기		'/2012.03.14 한용민 생성
Function getupcheshopcontractinfo(byval shopid, byval makerid)
	Dim sql

	if shopid = "" or makerid = "" then exit Function

	sql = "SELECT top 1"
	sql = sql & " u.shopname ,u.shopdiv ,u.currencyUnit ,u.exchangeRate ,u.exchangeRate"
	sql = sql & " ,s.shopid ,s.makerid ,s.comm_cd ,s.defaultmargin ,s.defaultsuplymargin"
	sql = sql & " ,s.defaultCenterMwDiv"
	sql = sql & " from db_shop.dbo.tbl_shop_designer s"
	sql = sql & " join db_shop.dbo.tbl_shop_user u"
	sql = sql & " 	on s.shopid = u.userid"
	sql = sql & "	and s.makerid = '"&makerid&"'"
	sql = sql & "	and s.shopid = '"&shopid&"'"

	'response.write sql &"<Br>"
	rsget.Open sql,dbget,1
	IF Not rsget.EOF THEN
		getupcheshopcontractinfo = rsget.getRows()
	END IF
	rsget.Close
End Function

'//업체 반품 주소	'/2012.05.02 한용민 생성
function Getpartnerdeliverinfo_off(userid , isusing)
    dim sqlStr ,arrpartnerinfo , sqlsearch

	if userid = "" then exit function

	if isusing <> "" then
		sqlsearch = sqlsearch & " and p.isusing='"&isusing&"'"
	end if

	sqlStr = "select top 1"
	sqlStr = sqlStr + " id , company_name,deliver_name ,deliver_phone ,deliver_hp ,deliver_email"
	sqlStr = sqlStr + " ,return_zipcode ,return_address ,return_address2"
	sqlStr = sqlStr + " from db_partner.dbo.tbl_partner p"
	sqlStr = sqlStr + " where id = '"&userid&"' " & sqlsearch

	'response.write sqlStr &"<br>"
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        arrpartnerinfo    = rsget.getrows()
    end if
    rsget.close

	Getpartnerdeliverinfo_off = arrpartnerinfo
end function

'//텐바이텐과 연동된 외부 매장 가져 오기 select박스		'/2012.05.17 한용민 생성
function drawSelectBoxtenshoplinkothersite(selectBoxName,selectedId,siteseq,shopdiv)
	dim tmp_str,query1 , sqlsearch

	if siteseq = "" then exit function

	if shopdiv <> "" then
		sqlsearch = sqlsearch & " and u.shopdiv in ("&shopdiv&")"
	end if

	query1 = " select"
	query1 = query1 & " u.userid ,u.shopname"
	query1 = query1 & " ,l.othershopid"
	query1 = query1 & " from [db_shop].[dbo].tbl_shop_user u"
	query1 = query1 & " join db_shop.dbo.tbl_shop_othersitelink l"
	query1 = query1 & " 	on u.userid = l.shopid"
	query1 = query1 & " 	and u.isusing='Y'"
	query1 = query1 & " where l.siteseq = "&siteseq&" " & sqlsearch
	query1 = query1 & " order by u.isusing desc, convert(int,u.shopdiv)+10 asc, u.userid asc"

	'response.write query1 &"<Br>"
	%>
		<select class="select" name="<%=selectBoxName%>">
			<option value='' <%if selectedId="" then response.write " selected"%>>선택하세요</option>
	<%
	dbget.Open query1,dbget,1

		if  not dbget.EOF  then
		dbget.Movefirst
			do until dbget.EOF
			if Lcase(selectedId) = Lcase(dbget("userid")) then
				tmp_str = " selected"
			end if

			response.write("<option value='"&dbget("userid")&"' "&tmp_str&">"&dbget("userid")&" / "&dbget("shopname")&" ["&dbget("othershopid")&"]</option>")
			tmp_str = ""
			dbget.MoveNext
		loop
		end if
	dbget.close
	response.write("</select>")
end function

'//텐바이텐과 연동된 외부 매장 가져 오기. 외부 매장 아이디롤 가져올경우		'/2012.05.17 한용민 생성
function gettenshoplinkothersite(byval tenshopid ,byval siteseq)
dim sql

	if tenshopid = "" or siteseq = "" then Exit function

	sql = "select top 1"
	sql = sql & " idx ,siteseq ,shopid ,othershopid ,regdate ,lastupdate ,lastadminuserid"
	sql = sql & " from db_shop.dbo.tbl_shop_othersitelink"
	sql = sql & " where shopid ='"&tenshopid&"'"
	sql = sql & " and siteseq ="&siteseq&""

	'response.write sql &"<Br>"
	rsget.Open sql,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
		gettenshoplinkothersite = rsget("othershopid")
	else
		gettenshoplinkothersite = ""
	end if
	rsget.Close
end function

'//텐바이텐과 연동된 외부 매장 가져 오기. 텐바이텐 매장 아이디롤 가져올경우		'/2012.05.17 한용민 생성
function getothersitelinktenshop(byval othershopid ,byval siteseq)
dim sql

	if othershopid = "" or siteseq = "" then Exit function

	sql = "select top 1"
	sql = sql & " idx ,siteseq ,shopid ,othershopid ,regdate ,lastupdate ,lastadminuserid"
	sql = sql & " from db_shop.dbo.tbl_shop_othersitelink"
	sql = sql & " where othershopid ='"&othershopid&"'"
	sql = sql & " and siteseq ="&siteseq&""

	'response.write sql &"<Br>"
	rsget.Open sql,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
		getothersitelinktenshop = rsget("shopid")
	else
		getothersitelinktenshop = ""
	end if
	rsget.Close
end function

'/업체 어드민 오프라인 상품 등록 권한 여부	'/2012.07.12 한용민 생성
function getupcheitemregyn(makerid)
    dim sqlStr

    if makerid = "" then exit function

    getupcheitemregyn = FALSE

    sqlStr = "select count(*) as CNT"
    sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_designer"
    sqlStr = sqlStr + " where makerid='" + makerid + "'"
    sqlStr = sqlStr + " and itemregyn = 'Y'"

    'response.write sqlStr & "<Br>"
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        getupcheitemregyn = rsget("CNT")>0
    end if
    rsget.close
end function

'//오프라인 수기 주문번호 생성		'//2012.08.08 한용민 생성
function manualordernomake_off(shopid,posid)
	dim tmporderno , tmpidx ,sqlStr

	if shopid = "" or posid = "" then exit function

    sqlStr = "select top 1 idx"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master"
	sqlStr = sqlStr + " order by idx desc"

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
	    tmpidx = rsget("idx") + 1
	else
		tmpidx = 1
	end if
	rsget.close

	tmporderno = right(year(now()),2)		'//년도뒤에두자리
	tmporderno = tmporderno & Format00(2,month(now()))		'//월두자리
	tmporderno = tmporderno & Format00(2,day(now()))		'//일두자리
	tmporderno = tmporderno & right(shopid,3)		'//매장아이디 뒤세자리
	tmporderno = tmporderno & Format00(2,trim(posid))		'//포스아이디
	tmporderno = tmporderno & right(hour(now()),1)
	tmporderno = tmporderno & right(minute(now()),1)
	tmporderno = tmporderno & right(second(now()),1)
	tmporderno = tmporderno & Format00(2,right(tmpidx,2))

	manualordernomake_off = tmporderno
end function

'//물류 innerbox 내역 조회 배열버전		'/2012.08.13 한용민 생성
function getarrinnerbox(siteSeq ,baljudate ,shopid ,innerboxno,innerboxidx ,cartoonboxno)
    dim sqlStr ,sqlsearch

	if baljudate <> "" then
		sqlsearch = sqlsearch & " and convert(varchar(10),baljudate,121) = '"&baljudate&"'"
	end if
	if shopid <> "" then
		sqlsearch = sqlsearch & " and shopid = '"&shopid&"'"
	end if
	if innerboxno <> "" then
		sqlsearch = sqlsearch & " and innerboxno = "&innerboxno&""
	end if
	if innerboxidx <> "" then
		sqlsearch = sqlsearch & " and innerboxidx = "&innerboxidx&""
	end if
	if cartoonboxno <> "" then
		sqlsearch = sqlsearch & " and cartoonboxno = "&cartoonboxno&""
	end if

	sqlStr = "select top 1"
	sqlStr = sqlStr & " '"&siteSeq&"' as siteSeq, convert(varchar(10),baljudate,121) as baljudate"
	sqlStr = sqlStr & " , shopid, innerboxno, innerboxweight, cartoonboxno, cartonboxsongjangdiv"
	sqlStr = sqlStr & " , cartonboxsongjangno ,innerboxidx ,cartoonboxweight"
	sqlStr = sqlStr & " from db_storage.dbo.tbl_cartoonbox_detail"
	sqlStr = sqlStr & " where 1=1 " & sqlsearch

	'response.write sqlStr & "<br>"
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
    	getarrinnerbox = rsget.getrows()
    end if
    rsget.Close
end function

'/텐바이텐 매장 shopseq 가져오기 		'/2012.08.21 한용민 생성
function gettenshopidx(shopid)
   dim sqlstr , sqlsearch

	if shopid = "" then exit function

	if shopid <> "" then
		sqlsearch = sqlsearch & " and u.userid='"&shopid&"'"
	end if

	sqlstr = "select top 1"
	sqlstr = sqlstr & " p.partnerseq"
	sqlstr = sqlstr & " from [db_shop].[dbo].tbl_shop_user u"
	sqlstr = sqlstr & " join db_partner.dbo.tbl_partner p"
	sqlstr = sqlstr & " 	on u.userid = p.id"
	sqlstr = sqlstr & " where u.isusing='Y'"
	sqlstr = sqlstr & " and p.isusing='Y' " & sqlsearch

	'response.write sqlstr &"<Br>"
	rsget.Open sqlstr,dbget,1
	if not rsget.EOF  then
		gettenshopidx = rsget("partnerseq")
	end if
	rsget.close
end function

'//정렬기준		'/2012.08.29 한용민 추가
function drawordertype(selBoxName ,selVal ,chplg ,searchgubun)
%>
    <select name="<%= selBoxName %>" <%= chplg %>>
		<option value="" <% if selVal="" then response.write " selected" %>>정렬선택</option>
		<option value="ea" <% if selVal="ea" then response.write " selected" %>>수량순</option>
		<option value="totalprice" <% if selVal="totalprice" then response.write " selected" %>>매출순</option>
		<option value="gain" <% if selVal="gain" then response.write " selected" %>>수익순</option>

		<% if searchgubun = "I" then %>
			<option value="unitCost" <% if selVal="unitCost" then response.write " selected" %>>객단가순</option>
		<% end if %>
	</select>
<%
end Function

'/출고위탁 있는지 없는지 확인		'/2013.01.05 한용민 생성
function getcwflag(shopid,comm_cd)
   dim sqlstr , sqlsearch

	if shopid = "" or comm_cd = "" then exit function

	if shopid <> "" then
		sqlsearch = sqlsearch & " and d.shopid='"&shopid&"'"
	end if

	if comm_cd <> "" then
		sqlsearch = sqlsearch & " and d.comm_cd='"&comm_cd&"'"
	end if

	sqlstr = "select top 1"
	sqlstr = sqlstr & " d.shopid , d.makerid , d.comm_cd"
	sqlstr = sqlstr & " from db_shop.dbo.tbl_shop_designer d"
	sqlstr = sqlstr & " where 1=1 " & sqlsearch

	'response.write sqlstr &"<Br>"
	rsget.Open sqlstr,dbget,1
	if not rsget.EOF  then
		getcwflag = "1"
	else
		getcwflag = "0"
	end if
	rsget.close
end function

'//어드민표기언어정보		'/2012.09.14 한용민 생성
function getadmindisplaylanguage(shopid)
	dim sql

	if shopid = "" then
		getadmindisplaylanguage = "KOR"
		exit function
	end if

	sql = "select top 1"
	sql = sql & " u.userid as shopid ,u.admindisplang"
	sql = sql & " from db_shop.dbo.tbl_shop_user u"
	sql = sql & " where u.userid = '"&shopid&"'"

	'response.write sql &"<Br>"
	rsget.Open sql,dbget,1
	if not rsget.EOF then
		getadmindisplaylanguage = rsget("admindisplang")
	else
		getadmindisplaylanguage = "KOR"
	end if
	rsget.close
end function

'/넘어온 날짜의 이전주 일요일을 반환함		'/2012.11.02 한용민 생성
Function beforeWeeksunday(tmpdate)

	If DatePart("w",tmpdate) = "1" Then
		beforeWeeksunday = tmpdate
	Else
		beforeWeeksunday = DateAdd("d",((CInt(DatePart("w",tmpdate))-1)*-1),tmpdate)
	End If
End Function

'/사용여부 영문버전		'/2012.11.05 한용민 생성
Function drawSelectBoxisusingYN(selectBoxName,selectedId,chplg)
%>
   <select name="<%=selectBoxName%>" <%= chplg %>>
	   <option value="" <% if selectedId="" then response.write "selected" %>>SELECT</option>
	   <option value="Y" <% if selectedId="Y" then response.write "selected" %>>Y</option>
	   <option value="N" <% if selectedId="N" then response.write "selected" %>>N</option>
   </select>
<%
End Function

'//담당매장 갯수		'/2013.10.28 한용민 생성
function getshopusercount(userid)
dim tmp_str,query1

	if userid = "" then
		getshopusercount = 0
		exit function
	end if

	query1 = "select count(*) as cnt"
	query1 = query1 + " from db_partner.dbo.tbl_user_tenbyten ut"
	query1 = query1 + " join db_partner.dbo.tbl_partner_shopuser ps"
	query1 = query1 + " 	on ps.empno = ut.empno"
	query1 = query1 + " where ut.isusing=1"

	' 퇴사예정자 처리	' 2018.10.16 한용민
	query1 = query1 & "	and (ut.statediv ='Y' or (ut.statediv ='N' and datediff(dd,ut.retireday,getdate())<=0))" & vbcrlf
	query1 = query1 + " and ut.userid = '"&userid&"'"

	'response.write query1 &"<br>"
	rsget.Open query1,dbget,1
	if not rsget.EOF  then
		if rsget("cnt") > 0 then
			getshopusercount = rsget("cnt")
		else
			getshopusercount = 0
		end if
	else
		getshopusercount = 0
	end if
	rsget.close
end function

'//담당매장 가져오기. 배열 버전		'/2013.10.28 한용민 생성
function getpartpeopleshoparray(existsuseridyn)
Dim sqlStr, arrList

	sqlStr = "select top 500"
	sqlStr = sqlStr & " ut.userid  , ut.username ,ut.empno"
	sqlStr = sqlStr & " from db_partner.dbo.tbl_user_tenbyten ut"
	sqlStr = sqlStr & " join db_partner.dbo.tbl_partner_shopuser ps"
	sqlStr = sqlStr & " 	on ps.empno = ut.empno"
	sqlStr = sqlStr & " where ut.isusing=1"

	' 퇴사예정자 처리	' 2018.10.16 한용민
	sqlStr = sqlStr & "	and (ut.statediv ='Y' or (ut.statediv ='N' and datediff(dd,ut.retireday,getdate())<=0))" & vbcrlf

	if existsuseridyn="Y" then
		sqlStr = sqlStr & " and isnull(ut.userid,'')<>''"
	end if

	sqlStr = sqlStr & " order by ut.part_sn asc, ut.posit_sn asc ,ut.username asc"

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget
	IF not rsget.eof THEN
		arrList = rsget.getRows()
	End IF
	rsget.close

	getpartpeopleshoparray = arrList
end function

Sub sbOptJungSanGubun(selValue)
%> 
	<option value="2" <%if selValue ="2" then %>selected<%end if%>>출고위탁</option>  <%'10x10 위탁%>
	<option value="4" <%if selValue ="4" then %>selected<%end if%>>출고매입(10x10매입)</option>
	<option value="5" <%if selValue ="5" then %>selected<%end if%>>출고매입</option>  <% '출고분정산%>
	<option value="6" <%if selValue ="6" then %>selected<%end if%>>업체위탁</option>
	<option value="8" <%if selValue ="8" then %>selected<%end if%>>업체매입</option>
	<option value="9" <%if selValue ="9" then %>selected<%end if%>>가맹점</option>
	<option value="0" <%if selValue ="0" then %>selected<%end if%>>통합</option> 
<%
End Sub

'/2016.09.26 한용민 생성
sub drawcartoonboxtype(selBoxName, selVal, chplg, size_yn, cbm_yn, kg_yn)
%>
    <select name="<%= selBoxName %>" <%= chplg %>>
		<option value="" <% if selVal="" then response.write " selected" %>>SELECT</option>
		<option value="Z1" <% if selVal="Z1" then response.write " selected" %>>Z1 :<% if size_yn="Y" then %> 600x440x270<% end if %><% if cbm_yn="Y" then %> / CBM 0.07128<% end if %><% if kg_yn="Y" then %> / 11.90KG<% end if %></option>
		<option value="Z2" <% if selVal="Z2" then response.write " selected" %>>Z2 :<% if size_yn="Y" then %> 600x440x430<% end if %><% if cbm_yn="Y" then %> / CBM 0.11352<% end if %><% if kg_yn="Y" then %> / 18.96KG<% end if %></option>
		<option value="Z3" <% if selVal="Z3" then response.write " selected" %>>Z3 :<% if size_yn="Y" then %> 600x440x570<% end if %><% if cbm_yn="Y" then %> / CBM 0.15048<% end if %><% if kg_yn="Y" then %> / 25.13KG<% end if %></option>
		<option value="U1" <% if selVal="U1" then response.write " selected" %>>U1 :<% if size_yn="Y" then %> 1000x150x150<% end if %><% if cbm_yn="Y" then %> / CBM 0.02250<% end if %><% if kg_yn="Y" then %> / 3.76KG<% end if %></option>
		<option value="A1" <% if selVal="A1" then response.write " selected" %>>A1 :<% if size_yn="Y" then %> 230x170x80<% end if %><% if cbm_yn="Y" then %> / CBM 0.00313<% end if %><% if kg_yn="Y" then %> / 0.52KG<% end if %></option>
		<option value="A2" <% if selVal="A2" then response.write " selected" %>>A2 :<% if size_yn="Y" then %> 250x180x90<% end if %><% if cbm_yn="Y" then %> / CBM 0.00405<% end if %><% if kg_yn="Y" then %> / 0.68KG<% end if %></option>
		<option value="B1" <% if selVal="B1" then response.write " selected" %>>B1 :<% if size_yn="Y" then %> 280x200x240<% end if %><% if cbm_yn="Y" then %> / CBM 0.01344<% end if %><% if kg_yn="Y" then %> / 2.24KG<% end if %></option>
		<option value="B2" <% if selVal="B2" then response.write " selected" %>>B2 :<% if size_yn="Y" then %> 280x200x240<% end if %><% if cbm_yn="Y" then %> / CBM 0.01344<% end if %><% if kg_yn="Y" then %> / 2.24KG<% end if %></option>
		<option value="B3" <% if selVal="B3" then response.write " selected" %>>B3 :<% if size_yn="Y" then %> 280x220x120<% end if %><% if cbm_yn="Y" then %> / CBM 0.00739<% end if %><% if kg_yn="Y" then %> / 1.23KG<% end if %></option>
		<option value="C1" <% if selVal="C1" then response.write " selected" %>>C1 :<% if size_yn="Y" then %> 340x230x120<% end if %><% if cbm_yn="Y" then %> / CBM 0.00938<% end if %><% if kg_yn="Y" then %> / 1.57KG<% end if %></option>
		<option value="C2" <% if selVal="C2" then response.write " selected" %>>C2 :<% if size_yn="Y" then %> 340x230x240<% end if %><% if cbm_yn="Y" then %> / CBM 0.01877<% end if %><% if kg_yn="Y" then %> / 3.13KG<% end if %></option>
		<option value="D1" <% if selVal="D1" then response.write " selected" %>>D1 :<% if size_yn="Y" then %> 400x280x120<% end if %><% if cbm_yn="Y" then %> / CBM 0.01344<% end if %><% if kg_yn="Y" then %> / 2.24KG<% end if %></option>
		<option value="D2" <% if selVal="D2" then response.write " selected" %>>D2 :<% if size_yn="Y" then %> 400x280x240<% end if %><% if cbm_yn="Y" then %> / CBM 0.02688<% end if %><% if kg_yn="Y" then %> / 4.49KG<% end if %></option>
		<option value="E1" <% if selVal="E1" then response.write " selected" %>>E1 :<% if size_yn="Y" then %> 460x340x120<% end if %><% if cbm_yn="Y" then %> / CBM 0.01877<% end if %><% if kg_yn="Y" then %> / 3.13KG<% end if %></option>
		<option value="E2" <% if selVal="E2" then response.write " selected" %>>E2 :<% if size_yn="Y" then %> 460x340x240<% end if %><% if cbm_yn="Y" then %> / CBM 0.03754<% end if %><% if kg_yn="Y" then %> / 6.27KG<% end if %></option>
		<option value="F1" <% if selVal="F1" then response.write " selected" %>>F1 :<% if size_yn="Y" then %> 570x410x130<% end if %><% if cbm_yn="Y" then %> / CBM 0.03038<% end if %><% if kg_yn="Y" then %> / 5.07KG<% end if %></option>
		<option value="F2" <% if selVal="F2" then response.write " selected" %>>F2 :<% if size_yn="Y" then %> 570x410x250<% end if %><% if cbm_yn="Y" then %> / CBM 0.05843<% end if %><% if kg_yn="Y" then %> / 9.76KG<% end if %></option>
		<option value="F3" <% if selVal="F3" then response.write " selected" %>>F3 :<% if size_yn="Y" then %> 570x410x375<% end if %><% if cbm_yn="Y" then %> / CBM 0.08764<% end if %><% if kg_yn="Y" then %> / 14.64KG<% end if %></option>
	</select>
<%
end sub

'/2016.09.26 한용민 생성
function getcartoonboxtype(cartoonboxtype, arrnum)
	dim tmparr,  tmpval

	if cartoonboxtype="" or arrnum="" then exit function

	if cartoonboxtype="Z1" then
		tmparr = "600x440x270,0.07128,11.90KG"
	elseif cartoonboxtype="Z2" then
		tmparr = "600x440x430,0.11352,18.96KG"
	elseif cartoonboxtype="Z3" then
		tmparr = "600x440x570,0.15048,25.13KG"
	elseif cartoonboxtype="U1" then
		tmparr = "1000x150x150,0.02250,3.76KG"
	elseif cartoonboxtype="A1" then
		tmparr = "230x170x80,0.00313,0.52KG"
	elseif cartoonboxtype="A2" then
		tmparr = "250x180x90,0.00405,0.68KG"
	elseif cartoonboxtype="B1" then
		tmparr = "280x200x240,0.01344,2.24KG"
	elseif cartoonboxtype="B2" then
		tmparr = "280x200x240,0.01344,2.24KG"
	elseif cartoonboxtype="B3" then
		tmparr = "280x220x120,0.00739,1.23KG"
	elseif cartoonboxtype="C1" then
		tmparr = "340x230x120,0.00938,1.57KG"
	elseif cartoonboxtype="C2" then
		tmparr = "40x230x240,0.01877,3.13KG"
	elseif cartoonboxtype="D1" then
		tmparr = "400x280x120,0.01344,2.24KG"
	elseif cartoonboxtype="D2" then
		tmparr = "400x280x240,0.02688,4.49KG"
	elseif cartoonboxtype="E1" then
		tmparr = "460x340x120,0.01877,3.13KG"
	elseif cartoonboxtype="E2" then
		tmparr = "460x340x240,0.03754,6.27KG"
	elseif cartoonboxtype="F1" then
		tmparr = "570x410x130,0.03038,5.07KG"
	elseif cartoonboxtype="F2" then
		tmparr = "570x410x250,0.05843,9.76KG"
	elseif cartoonboxtype="F3" then
		tmparr = "570x410x375,0.08764,14.64KG"
	end if

	tmpval = split(tmparr,",")(arrnum)
	getcartoonboxtype = tmpval
end function

'/해외가격 형식 표기	'/2016.09.26 한용민 생성
function getdisp_price_currencyChar(price, currencyChar)
	dim tmpdisp
    dim bufCurrencyChar : bufCurrencyChar = currencyChar
    if (bufCurrencyChar="WON") THEN bufCurrencyChar="KRW"
    if (bufCurrencyChar="원") THEN bufCurrencyChar="￦"
        
	if price="0" or price="" or isnull(price) then price=0

	if len(bufCurrencyChar)<3 then
		if (bufCurrencyChar="￦") then
			tmpdisp = bufCurrencyChar & " " & FormatNumber(price ,0)	'/한국화폐 소수점 없음
		else
			tmpdisp = bufCurrencyChar & " " & FormatNumber(price ,2)	'/외국화폐 소수점 2자리 까지
		end if
	else
	    if (bufCurrencyChar="KRW") then
	    	tmpdisp = FormatNumber(price ,0) & " " & bufCurrencyChar	'/한국화폐 소수점 없음
	    else
			tmpdisp = FormatNumber(price ,2) & " " & bufCurrencyChar	'/외국화폐 소수점 2자리 까지
	    end if
	end if

	getdisp_price_currencyChar = tmpdisp
end function

'/해외가격 표기	'/2016.09.26 한용민 생성
function getdisp_price(price, currencyChar)
	dim tmpdisp
    dim bufCurrencyChar : bufCurrencyChar = currencyChar
    if (bufCurrencyChar="WON") THEN bufCurrencyChar="KRW"
    if (bufCurrencyChar="원") THEN bufCurrencyChar="￦"
        
	if price="0" or price="" or isnull(price) then price=0

	if len(bufCurrencyChar)<3 then
		if (bufCurrencyChar="￦") then
			tmpdisp = FormatNumber(price ,0)	'/한국화폐 소수점 없음
		else
			tmpdisp = FormatNumber(price ,2)	'/외국화폐 소수점 2자리 까지
		end if
	else
	    if (bufCurrencyChar="KRW") then
	    	tmpdisp = FormatNumber(price ,0)	'/한국화폐 소수점 없음
	    else
			tmpdisp = FormatNumber(price ,2)	'/외국화폐 소수점 2자리 까지
	    end if
	end if

	getdisp_price = tmpdisp
end function

'/ 오프라인 매장 정보	'/2017.10.26 한용민 생성
function getoffshopuser(userid)
   dim sql , sqlsearch, arrList

	if isnull(userid) or userid="" then exit function

	sql = "select top 1 " & vbcrlf
	sql = sql & " u.countrylangcd, u.currencyunit, u.loginsite" & vbcrlf
	sql = sql & " , e.currencyChar, e.exchangeRate, e.multiplerate, e.linkPriceType, e.basedate, e.regdate" & vbcrlf
	sql = sql & " , e.lastupdate, e.reguserid, e.lastuserid" & vbcrlf
	sql = sql & " from db_shop.dbo.tbl_shop_user u" & vbcrlf
	sql = sql & " left join db_item.dbo.tbl_exchangeRate e" & vbcrlf
	sql = sql & " 	on u.countrylangcd = e.countrylangcd" & vbcrlf
	sql = sql & " 	and u.currencyunit = e.currencyunit" & vbcrlf
	sql = sql & " 	and u.loginsite = e.sitename" & vbcrlf
	sql = sql & " where userid = '" & userid & "'" & vbcrlf

	'response.write sql & "<br>"
	rsget.Open sql,dbget,1
	if not rsget.EOF  then
		arrList = rsget.getrows()
	end if
	rsget.Close

	getoffshopuser = arrList
end function
%>
