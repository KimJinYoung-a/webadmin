<%
'##############################################################
'	Description : �������� �����Լ�
'	History : 2006.12.11 ������ ����
'			  2011.02.18 �ѿ�� ����
'##############################################################

'//�������θ��� ���п� ���� �������� ���� ����	'/�ѿ�� �߰�
Sub drawSelectBoxOffShopdiv_off(selectBoxName,selectedId,shopdiv,shopall,chplg)
	dim tmp_str,query1
	%>
	<select class="select" name="<%=selectBoxName%>" <%=chplg%>>
		<option value='' <%if selectedId="" then response.write " selected"%>>�����ϼ���</option>

		<% if shopall <> "" then %>
			<option value='all' <%if selectedId="all" then response.write " selected"%>>��ü�������</option>
		<% end if %>
	<%
		query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user"
		query1 = query1 & " where 1=1 "		'/and isusing='Y'	'/2017.05.10 �ѿ�� ȸ�������� �����Ե� ���̰� �ش޶����.

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

'//�������θ��� ���п� ���� �������� ���� ����	'/������ �߰�
Sub drawSelectBoxOffShopdiv_New(selectBoxName,selectedId,shopdiv,shopall,chplg)
	dim tmp_str,query1
	%>
	<select class="select" name="<%=selectBoxName%>" <%=chplg%>>
		<option value='' <%if selectedId="" then response.write " selected"%>>�����ϼ���</option>

		<% if shopall <> "" then %>
			<option value='all' <%if selectedId="all" then response.write " selected"%>>��ü�������</option>
		<% end if %>
	<%
		query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user"
		query1 = query1 & " where 1=1 "		'/and isusing='Y'	'/2017.05.10 �ѿ�� ȸ�������� �����Ե� ���̰� �ش޶����.

		if shopdiv <> "" then
			query1 = query1 & " and shopdiv in ("&shopdiv&") AND vieworder>=0"
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
		<option value='' <%if selectedId="" then response.write " selected"%>>�����ϼ���</option>

		<% if shopall <> "" then %>
			<option value='all' <%if selectedId="all" then response.write " selected"%>>��ü�������</option>
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

'//�������θ��� ���ó ���п� ���� �������� ���� ����		'/2013.10.15 �ѿ�� �߰�
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
		<option value='' <%if selectedId="" then response.write " selected"%>>�����ϼ���</option>

		<% if shopall <> "" then %>
			<option value='all' <%if selectedId="all" then response.write " selected"%>>��ü���ó����</option>
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

'//�����̸� ��������	'//2012.08.09 �ѿ�� ����
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

'//���� ��ǥ ȭ�� ���� ��������		'/2012.03.14 �ѿ�� ����
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

'/�ֹ�Ȯ�� ����		'/2013.07.04 �ѿ�� ����
function isstateconfirm(statecd, foreign_statecd)
	if statecd="" or foreign_statecd="" then
		isstateconfirm=FALSE
		exit function
	end if

	if statecd = "7" or statecd = "8" or statecd = "9" then
		isstateconfirm = TRUE
	else
		isstateconfirm = FALSE
	end if
end function

'/�ֹ����¿� ���� ���� ��ȯ		'/2013.07.04 �ѿ�� ����
function getstateitemno(statecd, foreign_statecd, jumunitemno, fixitemno)
	if statecd="" or foreign_statecd="" or jumunitemno="" or fixitemno="" then exit function

	if ( isstateconfirm(statecd, foreign_statecd) ) then
		getstateitemno = fixitemno
	else
		getstateitemno = jumunitemno
	end if
end function

'/�ֹ����¿� ���� ���ް� ��ȯ			'/2013.07.04 �ѿ�� ����
function getstatesuplycash(statecd, foreign_statecd, jumunsuplyprice, fixsuplyprice)
	if statecd="" or foreign_statecd="" or jumunsuplyprice="" or fixsuplyprice="" then exit function

	if ( isstateconfirm(statecd, foreign_statecd) ) then
		getstatesuplycash = fixsuplyprice
	else
		getstatesuplycash = jumunsuplyprice
	end if
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
    <strong>�귣�� ID ������</strong>
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

'// �������� �� option list (2006.12.08 ������ ����)
'// 2007.03.13 ����������: ȭ��ǥ�� ������ ������ ��뿩�η� ǥ�� ���� ����
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

'// �������� �� �����ڵ� �������� (2006.12.08 ������)	'//������
'//���ȭ ��Ŵ drawoffshop_commoncode ���
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

'// �������� �� �����ڵ�� �������� (2006.12.11 ������)
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

'//�������� �����ڵ� (������� ��Ƽ�� ��밡���ϰ� ����) 	'/2012.08.02 �ѿ�� ����
function drawoffshop_commoncode(selectBoxName, selectedId, codekind, codegroup, maincodeid, chplg)
   dim tmp_str, sql , sqlsearch

	if selectBoxName = "" or codekind = "" then exit function

	'//�ڵ�׷��� �����ΰ��
	if codegroup = "MAIN" then
		sqlsearch = sqlsearch & " and codegroup='MAIN'"

		if maincodeid <> "" then
			sqlsearch = sqlsearch & " and maincodeid='"&maincodeid&"'"
		end if

	'//�ڵ�׷��� �����ΰ��
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

'//�������� �����ڵ� (������� ��Ƽ�� ��밡���ϰ� ����) �迭����	'/2012.08.02 �ѿ�� ����
function getoffshop_commoncodearray(codekind, codegroup, maincodeid)
   dim sql , sqlsearch, arrList

	'//�ڵ�׷��� �����ΰ��
	if codegroup = "MAIN" then
		sqlsearch = sqlsearch & " and codegroup='MAIN'"

		if maincodeid <> "" then
			sqlsearch = sqlsearch & " and maincodeid='"&maincodeid&"'"
		end if

	'//�ڵ�׷��� �����ΰ��
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

'//�������� �����ڵ� �ش��ڵ尡 ���� �׷����� TRUE FALSE �� ��ȯ 	'/2012.08.02 �ѿ�� ����
function getoffshop_commoncodegroup(codekind, codegroup, maincodeid ,codeid)
   dim tmp_str, sql , sqlsearch
	tmp_str = FALSE

	if codekind = "" or codegroup = "" then exit function

	'//�ڵ�׷��� �����ΰ��
	if codegroup = "MAIN" then
		sqlsearch = sqlsearch & " and codegroup='MAIN'"

		if maincodeid <> "" then
			sqlsearch = sqlsearch & " and maincodeid='"&maincodeid&"'"
		end if

	'//�ڵ�׷��� �����ΰ��
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

'// �������� �� ���� ���� (2006.12.11 ������)
function fnChkAuth(ByVal BctDiv, ByVal BctID, ByVal BctBigo)
	IF ( BctDiv > 100 and BctDiv < 200 ) THEN	'���� - ���ξ��̵�
		fnChkAuth = BctBigo
	ELSEIF ( BctDiv > 200 and BctDiv < 300 ) THEN	'����
		fnChkAuth = BctBigo
	ELSEIF ( BctDiv > 500 and BctDiv < 600 ) THEN	'���� - ���������̵�
		fnChkAuth = BctID
	ELSE
		fnChkAuth = ""
	END IF
End function

public function GetImageFolerName(byval FItemID)
    GetImageFolerName = GetImageSubFolderByItemid(FItemID)
	''GetImageFolerName = "0" + CStr(FItemID\10000)
end function

'//������ ������ ���� �ϰ�� ���� ���� ������ 	'/2011.02.17 �ѿ�� �߰�
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

'/���� ��å ������ '/��Ʈ�� :5 , ���� : 6		'/2011.03.18 �ѿ�� �߰�
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
	sqlStr = sqlStr & " on t.userid = p.id"
	sqlStr = sqlStr & " where p.isusing = 'Y' " & sqlsearch

	' ��翹���� ó��	' 2018.10.16 �ѿ��
	sqlStr = sqlStr & "	and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf

	'response.write sqlStr &"<br>"
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF  then
		getjob_sn = rsget("job_sn")
	end if
	rsget.close
end function

'/���� ���� ������ 	'/2011.03.18 �ѿ�� �߰�
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

	' ��翹���� ó��	' 2018.10.16 �ѿ��
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

'/���� �μ����� ������	'/2011.03.22 �ѿ�� �߰�
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

	' ��翹���� ó��	' 2018.10.16 �ѿ��
	sqlStr = sqlStr & "	and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf

	'response.write sqlStr &"<br>"
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF  then
		getpart_sn = rsget("part_sn")
	end if
	rsget.close
end function

'/���忡 ���� �귣�常(�����߰�)		'/2011.12.08 �ѿ�� ����
function drawBoxDirectIpchulOffShopByMakerchfg(selectBoxName,selectedId,makerid,chfg,comm_cd)
	dim tmp_str,query1

	if makerid = "" then exit function

	query1 = " select u.userid,u.shopname from [db_shop].[dbo].tbl_shop_user u"
	query1 = query1 & "      Join [db_shop].[dbo].tbl_shop_designer d"
	query1 = query1 & "      on u.userid=d.shopid"
	query1 = query1 & " where u.isusing='Y' "
	query1 = query1 & " and d.makerid='" + makerid + "'"

	if comm_cd <> "" then
		query1 = query1 & " and d.comm_cd in ("&comm_cd&")"   ''��ü��Ź ������Ը� ���� //������ ���� �߰�
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
		<option value='' <%if selectedId="" then response.write " selected"%>>����</option>
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

'/���忡 ���� �귣�� ��	'/2013.02.21 �ѿ�� ����
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

'//��ǰ ���� ��� ���� ���� ����		'/2012.03.14 �ѿ�� ����
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
	query1 = query1 & " and d.comm_cd in ('B012','B022','B023')"   ''��ü��Ź ������Ը� ���� //������ ���� �߰�
	query1 = query1 & " and u.userid<>'streetshop000'"
	query1 = query1 & " and u.userid<>'streetshop800'"
	query1 = query1 & " and u.userid<>'streetshop870'"
	query1 = query1 & " and u.userid<>'streetshop700'"
	query1 = query1 & " order by u.isusing desc, convert(int,u.shopdiv)+10 asc, u.userid asc"

	'response.write query1 & "<Br>"
	rsget.Open query1,dbget,1
	%>
	<select class="select" name="<%=selectBoxName%>">
		<option value='' <%if selectedId="" then response.write " selected"%>>����</option><%
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

''//��ü��Ź/������� ������� �ƴѰ�.
Sub drawSelectBoxShopjumunDesignerNotUpche(selectBoxName,selectedId,shopid,suplyer,comm_cd)
   dim tmp_str,query1

    query1 = " select d.makerid, c.socname_kor ,d.comm_cd from [db_shop].[dbo].tbl_shop_designer d"
	query1 = query1 + " left join [db_user].[dbo].tbl_user_c c on d.makerid=c.userid"
	query1 = query1 + " where d.shopid='" + shopid + "'"

	if suplyer="10x10" then
		query1 = query1 + " and d.comm_cd not in ('B012','B022')"   ''��ü��Ź ������Ը� ����
   	else
   	    query1 = query1 + " and d.comm_cd not in ('B012','B022')"   ''��ü��Ź ������Ը� ����
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
		<option value='' <%if selectedId="" then response.write " selected"%>>����</option>
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

function IsShopMakerContractExists(shopid, makerid)
	dim query1
	query1 = " select d.makerid from [db_shop].[dbo].tbl_shop_designer d "
	query1 = query1 + " left join [db_user].[dbo].tbl_user_c c on d.makerid=c.userid "
	query1 = query1 + " where d.shopid='" & shopid & "' "
	query1 = query1 + " and d.makerid = '" & makerid & "' "
	query1 = query1 + " and d.comm_cd not in ('B012','B022') "
	rsget.Open query1,dbget,1
	if not rsget.EOF  then
		IsShopMakerContractExists = True
	else
		IsShopMakerContractExists = False
	end if
	rsget.close
end function

''���� ���ֹ� ���� ������Ʈ (�������)		'/2012.01.10 �̻� ����
function PreOrderUpdateByBrand_off(masteridx,targetid,shopid)
	dim sqlStr

	if masteridx = "" or targetid = ""  then exit function

	sqlStr = " exec db_summary.[dbo].[sp_Ten_Shop_Stock_PreOrderUpdate] '"&masteridx&"','"&targetid&"','"&shopid&"'"

	'response.write sqlStr &"<Br>"
	dbget.Execute sqlStr
end function

''���� ��� ���� ������Ʈ		'/2012.01.10 �̻� ����
function recentUpdate_off(masteridx,targetid,shopid)
	dim sqlStr

	if masteridx = "" or targetid = ""  then exit function

	sqlStr = " exec db_summary.[dbo].[sp_Ten_Shop_Stock_RecentUpdate] '"&masteridx&"','"&targetid&"','"&shopid&"'"

	'response.write sqlStr &"<Br>"
	dbget.Execute sqlStr
end function

'// ���� ������ ó��	'/�ѿ�� ����
function drawmaechuldatefg(selectBoxName,selectedId,changefg)
%>
<select name="<%=selectBoxName%>" <%=changefg%>>
	<option value='' <%if selectedId="" then response.write " selected"%>>����</option>
	<option value='maechul' <%if selectedId="maechul" then response.write " selected"%>>���������</option>
	<option value='jumun' <%if selectedId="jumun" then response.write " selected"%>>�ֹ���</option>
</select>
<%
end function

function IsSameShopContract(shopid1, shopid2, makerid)
	dim sqlStr

	IsSameShopContract = False

	sqlStr = " select top 1 T1.comm_cd "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	( "
	sqlStr = sqlStr + " 		select top 1 d.comm_cd, d.defaultmargin, d.defaultsuplymargin "
	sqlStr = sqlStr + " 		from "
	sqlStr = sqlStr + " 		[db_shop].[dbo].tbl_shop_user u "
	sqlStr = sqlStr + " 		join [db_shop].[dbo].tbl_shop_designer d "
	sqlStr = sqlStr + " 		on "
	sqlStr = sqlStr + " 			u.userid = d.shopid "
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and u.userid = '" & shopid1 & "' "
	sqlStr = sqlStr + " 			and d.makerid = '" & makerid & "' "
	sqlStr = sqlStr + " 			and ((d.comm_cd = 'B031') or (d.comm_cd = 'B031' and d.makerid = 'ithinkso')) "
	sqlStr = sqlStr + " 	) T1 "
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select top 1 d.comm_cd, d.defaultmargin, d.defaultsuplymargin "
	sqlStr = sqlStr + " 		from "
	sqlStr = sqlStr + " 		[db_shop].[dbo].tbl_shop_user u "
	sqlStr = sqlStr + " 		join [db_shop].[dbo].tbl_shop_designer d "
	sqlStr = sqlStr + " 		on "
	sqlStr = sqlStr + " 			u.userid = d.shopid "
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and u.userid = '" & shopid2 & "' "
	sqlStr = sqlStr + " 			and d.makerid = '" & makerid & "' "
	sqlStr = sqlStr + " 			and ((d.comm_cd = 'B031') or (d.comm_cd = 'B013' and d.makerid = 'ithinkso')) "
	sqlStr = sqlStr + " 	) T2 "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and T1.comm_cd = T2.comm_cd "
	sqlStr = sqlStr + " 		and T1.defaultmargin = T2.defaultmargin "
	sqlStr = sqlStr + " 		and T1.defaultsuplymargin = T2.defaultsuplymargin "
	rsget.Open sqlStr, dbget, 1
	''response.write sqlStr & "<br />"

	if not rsget.EOF then
		IsSameShopContract = True
	end if
	rsget.Close

end function

'/���� ����� ������ ����Ʈ		'/2011.08.25 �ѿ�� ����
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
		<option value='' <%if selectedId="" then response.write " selected"%>>����</option>
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
		<option value='' <%if selectedId="" then response.write " selected"%>>������ �����ϴ�</option>
	<%
	end if
	rsget.close
	response.write("</select>")
end function

'/���� ����� ������ ����Ʈ		'/2011.08.25 �ѿ�� �߰�
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
		<option value='' <%if selectedId="" then response.write " selected"%>>����</option>
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
		<option value='' <%if selectedId="" then response.write " selected"%>>������ �����ϴ�</option>
	<%
	end if
	rsget.close
	response.write("</select>")
end function

'// ���� ������ ó��	'�ѿ�� ����
function drawmaechul_datefg(selectBoxName,selectedId,changefg)
%>
<select name="<%=selectBoxName%>" <%=changefg%>>
	<option value='' <%if selectedId="" then response.write " selected"%>>����</option>
	<option value='maechul' <%if selectedId="maechul" then response.write " selected"%>>���������</option>
	<option value='jumun' <%if selectedId="jumun" then response.write " selected"%>>�ֹ���</option>
</select>
<%
end function

'/����		'2011.11.17 �ѿ�� ����
function drawweekday_select(selectBoxName,selectedId,changefg)
%>
<select name="<%=selectBoxName%>" <%=changefg%>>
	<option value='' <%if selectedId="" then response.write " selected"%>>����</option>
	<option value='1' <%if selectedId="1" then response.write " selected"%>>��</option>
	<option value='2' <%if selectedId="2" then response.write " selected"%>>��</option>
	<option value='3' <%if selectedId="3" then response.write " selected"%>>ȭ</option>
	<option value='4' <%if selectedId="4" then response.write " selected"%>>��</option>
	<option value='5' <%if selectedId="5" then response.write " selected"%>>��</option>
	<option value='6' <%if selectedId="6" then response.write " selected"%>>��</option>
	<option value='7' <%if selectedId="7" then response.write " selected"%>>��</option>
</select>
<%
end function

'/���� ��ȯ		'2012.10.04 �ѿ�� ����
Function getweekday_select(v)
	if v = "1" then
		getweekday_select = "<font color=""red"">��</font>"
	elseif v = "2" then
		getweekday_select = "��"
	elseif v = "3" then
		getweekday_select = "ȭ"
	elseif v = "4" then
		getweekday_select = "��"
	elseif v = "5" then
		getweekday_select = "��"
	elseif v = "6" then
		getweekday_select = "��"
	elseif v = "7" then
		getweekday_select = "<font color=""blue"">��</font>"
	else
		getweekday_select = v
	end if
End Function

'//��¥(2012-10-25)�� ������ ���� ���		'/2012.10.25 �ѿ�� ����
function getweekend(yyyymmdd)
	if yyyymmdd = "" then exit function

	if right(FormatDateTime(yyyymmdd,1),3) = "�����" then
		getweekend = "<font color='blue'>"&Replace(Right(FormatDateTime(yyyymmdd,1),3),"����","")&"</font>"
	elseif right(FormatDateTime(yyyymmdd,1),3) = "�Ͽ���" then
		getweekend = "<font color='red'>"&Replace(Right(FormatDateTime(yyyymmdd,1),3),"����","")&"</font>"
	else
		getweekend = Replace(Right(FormatDateTime(yyyymmdd,1),3),"����","")
	end if
end function

'//��¥(2012-10-25)�� ������ �ָ��� ���� Ʋ��		'/2012.10.25 �ѿ�� ����
function getweekendcolor(yyyymmdd)
	if yyyymmdd = "" then exit function

	if right(FormatDateTime(yyyymmdd,1),3) = "�����" then
		getweekendcolor = "<font color='blue'>"&yyyymmdd&"</font>"
	elseif right(FormatDateTime(yyyymmdd,1),3) = "�Ͽ���" then
		getweekendcolor = "<font color='red'>"&yyyymmdd&"</font>"
	else
		getweekendcolor = yyyymmdd
	end if
end function

'//��Ʈ ������ ���� ����	'/2011.11.24 �ѿ�� ����
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

	' ��翹���� ó��	' 2018.10.16 �ѿ��
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
	<option value="" <% if selectid = "" then response.write " selected" %>>����</option>
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

'//����� ����â	'/2012.07.17 �ѿ�� ����
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
	<input type="button" class="button" value="����" onClick="pop_tenbytenuser()">
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

'//���� ��Ʈ ������ ���� ����		'/2011.11.24 �ѿ�� ����
function drawpartpeopleshop(ByVal boxname, ByVal selectid, ByVal chscript)
Dim sqlStr ,sqlsearch, arrList, intLoop

	sqlStr = "select top 500"
	sqlStr = sqlStr & " ut.userid  , ut.username ,ut.userid"
	sqlStr = sqlStr + " from db_partner.dbo.tbl_user_tenbyten ut"
	sqlStr = sqlStr + " join db_partner.dbo.tbl_partner_shopuser ps"
	sqlStr = sqlStr + " 	on ps.empno = ut.empno"
	sqlStr = sqlStr + " where ut.isusing=1"

	' ��翹���� ó��	' 2018.10.16 �ѿ��
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
	<option value="" <% if selectid = "" then response.write " selected" %>>����</option>
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

'//ȭ�����		'/2011.12.13 �ѿ�� ����
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

'/���� ��� ���� ����Ʈ�ڽ� ����	'/2012-05-18 �ѿ�� ����
Sub drawSelectBoxShopBeasongDiv(selectBoxName,selectedId)
	dim tmp_str,query1
   %>
	<select class="select" name="<%=selectBoxName%>">
		<option value="">����
		<option value="0" <% if (selectedId="0") then response.write "selected" %> >�����Ǹ�
		<option value="2" <% if (selectedId="2") then response.write "selected" %> >��ü���
	</select>
<%
End Sub

'/���� ��� ���� üũ�ڽ� ����		'/2012-05-18 �ѿ�� ����
Sub drawCheckBoxShopBeasongDiv(checkBoxName,checkedId)
	dim tmp_str,query1
   %>

   <!--<input type="checkbox" name="<%'=checkBoxName%>" value="0" <%' if (checkedId = "0") then %>checked<%' end if %>> �����Ǹ�-->
   <input type="checkbox" name="<%=checkBoxName%>" value="2" <% if (checkedId = "2") then %>checked<% end if %>> ��ü���
<%
End Sub

'//���ȭ ��Ŵ drawoffshop_commoncode ���		'//������
function DrawShopDivCombo(selBoxName,selVal)
%>
    <select class='select' name="<%= selBoxName %>" >
	<option value='' <% if selVal="" then response.write " selected" %> >��ü</option>
	<option value='1' <% if selVal="1" or selVal="2" then response.write " selected" %> >����</option>
	<option value='3' <% if selVal="3" or selVal="4" then response.write " selected" %> >����</option>
	<option value='5' <% if selVal="5" or selVal="6" then response.write " selected" %> >����</option>
	<option value='7' <% if selVal="7" or selVal="8" then response.write " selected" %> >�ؿ�</option>
	<option value='9' <% if selVal="9" then response.write " selected" %> >��Ÿ</option>
	</select>
<%
end Function

'/��������� ��ึ�� ���� ���� ��������	'/2011.12.07 �ѿ�� ����
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
	query1 = query1 & " and d.comm_cd in ('B012','B022')"   ''��ü��Ź ������Ը� ����
	query1 = query1 & " and u.userid<>'streetshop000'"
	query1 = query1 & " and u.userid<>'streetshop800'"
	query1 = query1 & " and u.userid<>'streetshop870'"
	query1 = query1 & " and u.userid<>'streetshop700'"
	query1 = query1 & " and t.shopid is not null"
	query1 = query1 & " and u.userid <> '"&gubunshopid&"'"	'���������� ����
	query1 = query1 & " and u.shopdiv=1"	'�������常
	query1 = query1 & " order by u.isusing desc, convert(int,u.shopdiv)+10 asc, u.userid asc"

	'response.write query1 &"<Br>"
	rsget.Open query1,dbget,1
	%>
	<select class="select" name="<%=selectBoxName%>" <%= chevent %>>
		<option value='' <%if selectedId="" then response.write " selected"%>>����</option><%

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

'/���Ա��к��� �������� ���� ����	'/2011.12.13 �ѿ�� ����
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
		<option value='' <%if selectedId="" then response.write " selected"%>>����</option><%

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

'/������������	: ������ discountKind		'/2012.02.10 �ѿ�� ����
function DrawdiscountKind(selBoxName,selVal,chplg)
%>
    <select class='select' name="<%= selBoxName %>" <%= chplg %>>
		<option value='' <% if selVal="" then response.write " selected" %> >��ü</option>
		<option value='1' <% if selVal="1" then response.write " selected" %> >��������</option>
		<option value='2' <% if selVal="2" then response.write " selected" %> >��������</option>
		<option value='4' <% if selVal="4" then response.write " selected" %> >��Ű������</option>
		<option value='5' <% if selVal="5" then response.write " selected" %> >��5%����</option>
		<option value='6' <% if selVal="6" then response.write " selected" %> >������</option>
		<option value='7' <% if selVal="7" then response.write " selected" %> >tenday10%����</option>
		<option value='9' <% if selVal="9" then response.write " selected" %> >��Ÿ������</option>
	</select>
<%
end Function

'// �űԾ�ü ����ó��		'/2012.02.09 �ѿ�� ����
function drawnewupche_datefg(selectBoxName,selectedId,changefg)
%>
<select name="<%=selectBoxName%>" <%=changefg%>>
	<option value='' <%if selectedId="" then response.write " selected"%>>��ü</option>
	<option value='ipgo' <%if selectedId="ipgo" then response.write " selected"%>>�԰���</option>
	<option value='reg' <%if selectedId="reg" then response.write " selected"%>>�����</option>
</select>
<%
end function

'//���� �귣�� ��� ���� ���� ����		'/2012.03.14 �ѿ�� ����
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

'//��ü ��ǰ �ּ�	'/2012.05.02 �ѿ�� ����
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

'//�ٹ����ٰ� ������ �ܺ� ���� ���� ���� select�ڽ�		'/2012.05.17 �ѿ�� ����
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
			<option value='' <%if selectedId="" then response.write " selected"%>>�����ϼ���</option>
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

'//�ٹ����ٰ� ������ �ܺ� ���� ���� ����. �ܺ� ���� ���̵�� �����ð��		'/2012.05.17 �ѿ�� ����
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

'//�ٹ����ٰ� ������ �ܺ� ���� ���� ����. �ٹ����� ���� ���̵�� �����ð��		'/2012.05.17 �ѿ�� ����
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

'/��ü ���� �������� ��ǰ ��� ���� ����	'/2012.07.12 �ѿ�� ����
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

'//�������� ���� �ֹ���ȣ ����		'//2012.08.08 �ѿ�� ����
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

	tmporderno = right(year(now()),2)		'//�⵵�ڿ����ڸ�
	tmporderno = tmporderno & Format00(2,month(now()))		'//�����ڸ�
	tmporderno = tmporderno & Format00(2,day(now()))		'//�ϵ��ڸ�
	tmporderno = tmporderno & right(shopid,3)		'//������̵� �ڼ��ڸ�
	tmporderno = tmporderno & Format00(2,trim(posid))		'//�������̵�
	tmporderno = tmporderno & right(hour(now()),1)
	tmporderno = tmporderno & right(minute(now()),1)
	tmporderno = tmporderno & right(second(now()),1)
	tmporderno = tmporderno & Format00(2,right(tmpidx,2))

	manualordernomake_off = tmporderno
end function

'//���� innerbox ���� ��ȸ �迭����		'/2012.08.13 �ѿ�� ����
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

'/�ٹ����� ���� shopseq �������� 		'/2012.08.21 �ѿ�� ����
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

'//���ı���		'/2012.08.29 �ѿ�� �߰�
function drawordertype(selBoxName ,selVal ,chplg ,searchgubun)
%>
    <select name="<%= selBoxName %>" <%= chplg %>>
		<option value="" <% if selVal="" then response.write " selected" %>>���ļ���</option>
		<option value="ea" <% if selVal="ea" then response.write " selected" %>>������</option>
		<option value="totalprice" <% if selVal="totalprice" then response.write " selected" %>>�����</option>
		<option value="gain" <% if selVal="gain" then response.write " selected" %>>���ͼ�</option>

		<% if searchgubun = "I" then %>
			<option value="unitCost" <% if selVal="unitCost" then response.write " selected" %>>���ܰ���</option>
		<% end if %>
	</select>
<%
end Function

'/�����Ź �ִ��� ������ Ȯ��		'/2013.01.05 �ѿ�� ����
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

'//����ǥ��������		'/2012.09.14 �ѿ�� ����
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

'/�Ѿ�� ��¥�� ������ �Ͽ����� ��ȯ��		'/2012.11.02 �ѿ�� ����
Function beforeWeeksunday(tmpdate)

	If DatePart("w",tmpdate) = "1" Then
		beforeWeeksunday = tmpdate
	Else
		beforeWeeksunday = DateAdd("d",((CInt(DatePart("w",tmpdate))-1)*-1),tmpdate)
	End If
End Function

'/��뿩�� ��������		'/2012.11.05 �ѿ�� ����
Function drawSelectBoxisusingYN(selectBoxName,selectedId,chplg)
%>
   <select name="<%=selectBoxName%>" <%= chplg %>>
	   <option value="" <% if selectedId="" then response.write "selected" %>>SELECT</option>
	   <option value="Y" <% if selectedId="Y" then response.write "selected" %>>Y</option>
	   <option value="N" <% if selectedId="N" then response.write "selected" %>>N</option>
   </select>
<%
End Function

'//������ ����		'/2013.10.28 �ѿ�� ����
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

	' ��翹���� ó��	' 2018.10.16 �ѿ��
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

'//������ ��������. �迭 ����		'/2013.10.28 �ѿ�� ����
function getpartpeopleshoparray(existsuseridyn)
Dim sqlStr, arrList

	sqlStr = "select top 500"
	sqlStr = sqlStr & " ut.userid  , ut.username ,ut.empno"
	sqlStr = sqlStr & " from db_partner.dbo.tbl_user_tenbyten ut"
	sqlStr = sqlStr & " join db_partner.dbo.tbl_partner_shopuser ps"
	sqlStr = sqlStr & " 	on ps.empno = ut.empno"
	sqlStr = sqlStr & " 	and ps.firstisusing = 'Y'"
	sqlStr = sqlStr & " where ut.isusing=1"

	' ��翹���� ó��	' 2018.10.16 �ѿ��
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
	<option value="2" <%if selValue ="2" then %>selected<%end if%>>�����Ź</option>  <%'10x10 ��Ź%>
	<option value="4" <%if selValue ="4" then %>selected<%end if%>>������(10x10����)</option>
	<option value="5" <%if selValue ="5" then %>selected<%end if%>>������</option>  <% '��������%>
	<option value="6" <%if selValue ="6" then %>selected<%end if%>>��ü��Ź</option>
	<option value="8" <%if selValue ="8" then %>selected<%end if%>>��ü����</option>
	<option value="9" <%if selValue ="9" then %>selected<%end if%>>������</option>
	<option value="0" <%if selValue ="0" then %>selected<%end if%>>����</option>
<%
End Sub

'/2016.09.26 �ѿ�� ����
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
		<option value="B1" <% if selVal="B1" then response.write " selected" %>>B1 :<% if size_yn="Y" then %> 280x200x120<% end if %><% if cbm_yn="Y" then %> / CBM 0.01344<% end if %><% if kg_yn="Y" then %> / 2.24KG<% end if %></option>
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
		<option value="G1" <% if selVal="G1" then response.write " selected" %>>G1 :<% if size_yn="Y" then %> 680x500x150<% end if %><% if cbm_yn="Y" then %> / CBM 0.05100<% end if %><% if kg_yn="Y" then %> / 00.00KG<% end if %></option>
	</select>
<%
end sub

'/2016.09.26 �ѿ�� ����
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
		tmparr = "280x200x120,0.01344,2.24KG"
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
	elseif cartoonboxtype="G1" then
		tmparr = "680x500x150,0.05100,00.00KG"
	end if

	tmpval = split(tmparr,",")(arrnum)
	getcartoonboxtype = tmpval
end function

'/�ؿܰ��� ���� ǥ��	'/2016.09.26 �ѿ�� ����
function getdisp_price_currencyChar(price, currencyChar)
	dim tmpdisp
    dim bufCurrencyChar : bufCurrencyChar = currencyChar
    if (bufCurrencyChar="WON") THEN bufCurrencyChar="KRW"
    if (bufCurrencyChar="��") THEN bufCurrencyChar="��"

	if price="0" or price="" or isnull(price) then price=0

	if len(bufCurrencyChar)<3 then
		if (bufCurrencyChar="��") then
			tmpdisp = bufCurrencyChar & " " & FormatNumber(price ,0)	'/�ѱ�ȭ�� �Ҽ��� ����
		'2022-08-25 ������ �ϴ� �ּ�ó��
		' elseif (bufCurrencyChar="��") then
		' 	tmpdisp = bufCurrencyChar & " " & FormatNumber(price ,0)	'/�Ϻ�ȭ�� �Ҽ��� ����
		else
			tmpdisp = bufCurrencyChar & " " & FormatNumber(price ,2)	'/�ܱ�ȭ�� �Ҽ��� 2�ڸ� ����
		end if
	else
	    if (bufCurrencyChar="KRW") then
	    	tmpdisp = FormatNumber(price ,0) & " " & bufCurrencyChar	'/�ѱ�ȭ�� �Ҽ��� ����
		'2022-08-25 ������ �ϴ� �ּ�ó��
	    ' elseif (bufCurrencyChar="JPY") then
	    ' 	tmpdisp = FormatNumber(price ,0) & " " & bufCurrencyChar	'/�Ϻ�ȭ�� �Ҽ��� ����
	    else
			tmpdisp = FormatNumber(price ,2) & " " & bufCurrencyChar	'/�ܱ�ȭ�� �Ҽ��� 2�ڸ� ����
	    end if
	end if

	getdisp_price_currencyChar = tmpdisp
end function

'/�ؿܰ��� ǥ��	'/2016.09.26 �ѿ�� ����
function getdisp_price(price, currencyChar)
	dim tmpdisp
    dim bufCurrencyChar : bufCurrencyChar = currencyChar
    if (bufCurrencyChar="WON") THEN bufCurrencyChar="KRW"
    if (bufCurrencyChar="��") THEN bufCurrencyChar="��"

	if price="0" or price="" or isnull(price) then price=0

	if len(bufCurrencyChar)<3 then
		if (bufCurrencyChar="��") then
			tmpdisp = FormatNumber(price ,0)	'/�ѱ�ȭ�� �Ҽ��� ����
		'2022-08-25 ������ �ϴ� �ּ�ó��
		' elseif (bufCurrencyChar="��") then
		' 	tmpdisp = FormatNumber(price ,0)	'/�Ϻ�ȭ�� �Ҽ��� ����
		else
			tmpdisp = FormatNumber(price ,2)	'/�ܱ�ȭ�� �Ҽ��� 2�ڸ� ����
		end if
	else
	    if (bufCurrencyChar="KRW") then
	    	tmpdisp = FormatNumber(price ,0)	'/�ѱ�ȭ�� �Ҽ��� ����
		'2022-08-25 ������ �ϴ� �ּ�ó��
	    ' elseif (bufCurrencyChar="JPY") then
	    ' 	tmpdisp = FormatNumber(price ,0)	'/�Ϻ�ȭ�� �Ҽ��� ����
	    else
			tmpdisp = FormatNumber(price ,2)	'/�ܱ�ȭ�� �Ҽ��� 2�ڸ� ����
	    end if
	end if

	getdisp_price = tmpdisp
end function

'/ �������� ���� ����	'/2017.10.26 �ѿ�� ����
function getoffshopuser(userid)
   dim sql , sqlsearch, arrList

	if isnull(userid) or userid="" then exit function

	sql = "select top 1 " & vbcrlf
	sql = sql & " u.countrylangcd, u.currencyunit, u.loginsite" & vbcrlf
	sql = sql & " , e.currencyChar, e.exchangeRate, e.multiplerate, e.linkPriceType, e.basedate, e.regdate" & vbcrlf
	sql = sql & " , e.lastupdate, e.reguserid, e.lastuserid, u.shopdiv" & vbcrlf
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
