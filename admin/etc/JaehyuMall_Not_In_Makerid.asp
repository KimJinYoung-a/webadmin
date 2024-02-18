<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>
<body onload="javascript:window.resizeTo(900, 770);">
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<%
''20130304 인터페이스변경 - 서동석

	Dim strSql, arrList, vAction, intLoop, vMallGubun, iDelCnt, vPageSize, vCurrPage, vTotalCount, vMakerID, vBrandNameKr, i
	Dim cisextusing : cisextusing="N"
	Dim isValidMakerid : isValidMakerid=False
	Dim arrListBrBrand, arrListLogBrand

	vAction			= Request("action")
	vMallGubun		= NullFillWith(Request("mallgubun"),"")
	vCurrPage		= NullFillWith(Request("cp"),1)
	vMakerID		= Request("makerid")
	vBrandNameKr	= Request("brandnamekr")
	vPageSize = "15"

	if (vMallGubun="lotte") then vMallGubun="lotteCom"   ''' 20130304 추가

	''If vAction = "insert" OR vAction = "delete"  Then
	If vAction = "upsel" Then
		Call Proc()
	ElseIf vAction = "epsel" Then
		Call potalProc()
	End If

	''브랜드 대표 설정 검색
	strSql = "select top 1 isextusing from db_user.dbo.tbl_user_c"
	strSql = strSql & " where userid='"&vMakerID&"'"

	if (vMakerID<>"") then
    	rsget.Open strSql,dbget
    	if Not rsget.Eof then
    	    isValidMakerid = True
    	    cisextusing = rsget("isextusing")
    	end if
    	rsget.close
    end if

	''브랜드별 제휴 사용여부
	strSql = " select top 100 c.userid as MallID, ni.idx, ni.regdate, ni.reguserid"
    strSql = strSql & " from db_user.dbo.tbl_user_c c "
    strSql = strSql & " 	Join db_partner.dbo.tbl_partner_addInfo f "
    strSql = strSql & " 	on c.userid=f.partnerid and c.userid <> 'ezwel' "
    strSql = strSql & " 	and f.pcomType=1 "
    strSql = strSql & " 	and f.pmallSellType=1"
    strSql = strSql & " 	left join db_temp.dbo.tbl_jaehyumall_not_in_makerid ni"
    strSql = strSql & " 	on c.userid=ni.mallGubun and ni.makerid='"&vMakerID&"'"
    strSql = strSql & " where c.isusing='Y' and c.userdiv='50'"
	rsget.Open strSql,dbget
	IF not rsget.EOF THEN
		arrListBrBrand = rsget.getRows()
	END IF
	rsget.close

	'// 로그
	strSql = " select top 50 mallgubun, makerid, useYN, reguserid, regdate from "
    strSql = strSql & " db_log.dbo.tbl_jaehyumall_not_in_makerid_history "
    strSql = strSql & " where "
    strSql = strSql & " 	1 = 1 "

	if (vMallGubun <> "") then
    	strSql = strSql & " 	and mallgubun = '" + CStr(vMallGubun) + "' "
	end if

    strSql = strSql & " 	and makerid = '" + CStr(vMakerID) + "' "
    strSql = strSql & " order by idx desc "
	rsget.Open strSql,dbget
	IF not rsget.EOF THEN
		arrListLogBrand = rsget.getRows()
	END IF
	rsget.close


	iDelCnt =  ((vCurrPage - 1) * vPageSize )
	strSql = "SELECT Count(A.idx) FROM [db_temp].[dbo].[tbl_jaehyumall_not_in_makerid] AS A "
	If vBrandNameKr <> ""  Then
		strSql = strSql & "INNER JOIN [db_user].[dbo].[tbl_user_c] AS C ON A.makerid = C.userid "
	End If
	strSql = strSql & " Where 1=1"
	if (vMallGubun<>"") then
	    strSql = strSql & " AND A.mallgubun = '" & vMallGubun & "'"
    end if

	If vMakerID <> ""  Then
		strSql = strSql & " AND A.makerid = '" & vMakerID & "' "
	End If
	If vBrandNameKr <> ""  Then
		strSql = strSql & " AND C.socname_kor Like '%" & vBrandNameKr & "%' "
	End If

	rsget.Open strSql,dbget
	vTotalCount = rsget(0)
	rsget.close


	strSql = "SELECT Top 15 A.makerid, A.mallgubun,A.regdate,A.idx, A.reguserid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid AS A "

	If vBrandNameKr <> "" Then
		strSql = strSql & "INNER JOIN [db_user].[dbo].[tbl_user_c] AS C ON A.makerid = C.userid "
	End If
	strSql = strSql & " Where 1=1"
	if (vMallGubun<>"") then
	    strSql = strSql & " AND A.mallgubun = '" & vMallGubun & "'"
    end if

	If vMakerID <> ""  Then
		strSql = strSql & " AND A.makerid = '" & vMakerID & "' "
	End If
	If vBrandNameKr <> ""  Then
		strSql = strSql & " AND C.socname_kor Like '%" & vBrandNameKr & "%' "
	End If

	strSql = strSql & "		AND A.idx NOT IN(SELECT TOP "&iDelCnt&" X.idx FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid AS X "
	If vBrandNameKr <> ""  Then
		strSql = strSql & "						INNER JOIN [db_user].[dbo].[tbl_user_c] AS Y ON X.makerid = Y.userid "
	End If
	strSql = strSql & " Where 1=1"
	if (vMallGubun<>"") then
	    strSql = strSql & "						AND X.mallgubun = '" & vMallGubun & "' "
	End If
	If vMakerID <> ""  Then
		strSql = strSql & " 					AND X.makerid = '" & vMakerID & "' "
	End If
	If vBrandNameKr <> ""  Then
		strSql = strSql & " 					AND Y.socname_kor Like '%" & vBrandNameKr & "%' "
	End If

	strSql = strSql & "						ORDER BY X.makerid ASC) "
	strSql = strSql & "ORDER BY A.makerid ASC"
	rsget.Open strSql,dbget

	IF not rsget.EOF THEN
		arrList = rsget.getRows()
	END IF
	rsget.close
%>
<script language="javascript">
function insert_id()
{
	if(frm.in_id.value == ""){
		alert("ID를 입력하세요.");
		frm.in_id.focus();
		return;
	}

	if ((!frm.isall.checked)&&(frm.mallgubun.value.length<1)){
	    alert('등록할 Mall 구분 또는 [모든 Mall에 적용] 체크구분이 필요합니다.');
	    frm.mallgubun.focus();
	    return;
	}

	frm.action.value = "insert";
	frm.submit();
}
function delete_id(){
    var chkExists = false;

    if (document.frm.del_id.length>0){
        for (var i=0;i<document.frm.del_id.length;i++){
            if (document.frm.del_id[i].checked){
                chkExists=true;
                break;
            }
        }
    }else{
        if (document.frm.del_id.checked){
            chkExists=true;
        }
    }

    if (!chkExists){
        alert('선택된 내역이 없습니다.');
        return;
    }

    if (confirm('선택 브랜드에 대해 제휴 판매설정 해제하시겠습니까?')){
	    frm.action.value = "delete";
	    frm.submit();
	}
}

function jsGoPage(iP){
	document.frmpage.cp.value = iP;
	document.frmpage.submit();
}

function chkComp(comp){
    var frm = comp.form;
    for (var i=0;i<frm.elements.length;i++){
        var e=frm.elements[i];
        if (e.name.substring(0,6)=="notin_"){
            e.disabled=(comp.value=="N");
        }
    }

}

function saveUsing(comp){
    if (!confirm('제휴사 브랜드 판매설정을 수정하시겠습니까?')){
        return;
    }

    comp.form.submit();
}

function jsIsusing(ep){
    if (!confirm('포탈 가격비교 판매여부를 수정하시겠습니까?')){
        return;
    }
    ep.form.submit();
}
</script>

<center>
Mall 구분 : <b><%=vMallGubun%></b>
</center>
<br>
<form name="frmsearch" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
		    <td width="90%">Mall 구분 : <% CALL DrawApiMallSelect("mallgubun",vMallGubun) %></td>
		    <td rowspan="4" width="10%"><input type="submit" value="검 색" style="width:50px;height:50px;"></td>
		</tr>
		<tr>
			<td >
			브랜드ID : <input type="text" class="text" name="makerid" value="<%=vMakerID%>" size="20"> <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID(this.form.name,'makerid');" >
			</td>

		</tr>
		<!--
		<tr>
			<td>브랜드명(한글) : <input type="text" class="text" name="brandnamekr" value="<%=vBrandNameKr%>" size="30"></td>
		</tr>
		-->
		</table>
	</td>
</tr>
</table>
</form>
<% if (vMakerID<>"") then %>
<br>
    <table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frmBrdUsing" method="post" action="">
    <input type="hidden" name="action" value="upsel">
    <input type="hidden" name="makerid" value="<%=vMakerID%>">
    <% if (Not isValidMakerid) then %>
    <tr>
        <td align="center" bgcolor="#FFFFFF"><%= vMakerID %>는 올바른 브랜드ID가 아닙니다.</td>
    </tr>
    <% else %>
    <tr align="center" bgcolor="#DDDDDD">
        <td width="200" >몰구분</td>
        <td width="200" >판매설정</td>
        <td width="100" >등록제외자</td>
        <td >등록제외설정일</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" >
        <td >제휴사 전체 사용여부</td>
        <td >
            <input type="radio" name="cisextusing" value="Y" <%=CHKIIF(cisextusing="Y","checked","") %> onClick="chkComp(this)">사용
            <input type="radio" name="cisextusing" value="N" <%=CHKIIF(cisextusing="N","checked","") %> onClick="chkComp(this)">
            <% if cisextusing="N" then %>
            <b>사용안함</b>
            <% else %>
            사용안함
            <% end if %>
        </td>
        <td colspan="2">
        이설정이 [사용안함] 인경우 아래 몰별 설정과 관계없이 판매안함
        </td>
    </tr>
    <tr height="2" bgcolor="#FFFFFF" >
        <td colspan="4"></td>
    </tr>
    <% if isArray(arrListBrBrand) then %>
        <% For intLoop =0 To UBound(arrListBrBrand,2) %>

        	<tr align="center" bgcolor="#FFFFFF" height="30">
        	    <td><%=arrListBrBrand(0,intLoop)%></td>
        		<td>
        		    <% if isNULL(arrListBrBrand(1,intLoop)) then %>
        		    <input type="radio" name="notin_<%=arrListBrBrand(0,intLoop)%>" value="" checked <%=CHKIIF(cisextusing="N","disabled","") %> >등록가능
        		    <input type="radio" name="notin_<%=arrListBrBrand(0,intLoop)%>" value="N" <%=CHKIIF(cisextusing="N","disabled","") %> >등록제외
        		    <% else %>
        		    <input type="radio" name="notin_<%=arrListBrBrand(0,intLoop)%>" value=""  <%=CHKIIF(cisextusing="N","disabled","") %> >등록가능
        		    <input type="radio" name="notin_<%=arrListBrBrand(0,intLoop)%>" value="N" checked <%=CHKIIF(cisextusing="N","disabled","") %> >등록제외
        		    <% end if %>
        		</td>
        		<td><%=arrListBrBrand(3,intLoop)%></td>
        		<td><%=arrListBrBrand(2,intLoop)%></td>
        	</tr>
        <% Next %>
    <% end if %>
    <tr align="center" bgcolor="#FFFFFF">
        <td colspan="4">
         <input type="button" value="제휴몰 브랜드 판매설정 저장" onClick="saveUsing(this)">
        </td>
    </tr>
    <% end if %>
    </form>
    </table>

	<p>

	<% if (isValidMakerid and isArray(arrListLogBrand)) then %>
		<br><br>
		[이전 판매상태]
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="#DDDDDD">
			<td width="200" >몰구분</td>
			<td width="200" >판매상태</td>
			<td width="100" >등록자</td>
			<td >등록일</td>
		</tr>
		<% if isArray(arrListLogBrand) then %>
			<% For intLoop =0 To UBound(arrListLogBrand,2) %>
				<tr align="center" bgcolor="#FFFFFF" height="30">
					<td>
						<% if (arrListLogBrand(0,intLoop) = "") then %>
							제휴몰 전체
						<% else %>
							<%=arrListLogBrand(0,intLoop)%>
						<% end if %>

					</td>
					<td><%=arrListLogBrand(2,intLoop)%></td>
					<td><%=arrListLogBrand(3,intLoop)%></td>
					<td><%=arrListLogBrand(4,intLoop)%></td>
				</tr>
			<% Next %>
		<% end if %>
	<% end if %>
    </table>
<% end if %>
<br>

<% if (vMakerID="") then %>
<form name="frm" action="JaehyuMall_Not_In_Makerid.asp" methd="post" style="margin:0px;">
<input type="hidden" name="action" value="">
<input type="hidden" name="cp" value="<%=vCurrPage%>">

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" class="a">
		<tr>
			<td width="80%">
			<!--
				제외 브랜드ID
				<input type="text" name="in_id" value="" size="10" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ insert_id(); return false;}">
				&nbsp;<input type="checkbox" name="isall" value="o">전체선택(
				)
				<input type="button" value="등록제외 브랜드 설정" onClick="insert_id()">
			-->
			</td>
			<td width="20%" align="right">브랜드수 : <b><%=vTotalCount%></b></td>
		</tr>
		</table>
	</td>
</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
    <td width="30%">몰구분</td>
	<td width="30%">브랜드ID</td>
	<td width="20%">등록일</td>
	<td width="15%">등록자</td>
	<td width="5%">수정</td>
	<!--
	<td width="20%"><input type="button" value="선택 브랜드ID 삭제" onClick="delete_id()"></td>
	-->
</tr>
<%
	IF isArray(arrList) THEN
		For intLoop =0 To UBound(arrList,2)
%>
	<tr align="center" bgcolor="#FFFFFF" height="30">
	    <td><%=arrList(1,intLoop)%></td>
		<td><%=arrList(0,intLoop)%></td>
		<td><%=arrList(2,intLoop)%></td>
		<td><%=arrList(4,intLoop)%></td>
		<td><a href="?mallgubun=<%=vmallgubun%>&makerid=<%=arrList(0,intLoop)%>">[수정]</a></td>
		<!--
		<td><input type="checkbox" name="del_id" value="<%=arrList(3,intLoop)%>"></td>
		-->
	</tr>
<%
		Next
	Else
%>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="5" align="center" class="page_link">[데이터가 없습니다.]</td>
	</tr>
<%
	End If
%>
</table>
</form>

<form name="frmpage" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="cp" value="<%=vCurrPage%>">
<input type="hidden" name="mallgubun" value="<%=vMallGubun%>">
<input type="hidden" name="makerid" value="<%=vMakerID%>">
<input type="hidden" name="brandnamekr" value="<%=vBrandNameKr%>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<%
Dim iStartPage, iEndPage, ix, iTotalPage
iStartPage = (Int((vCurrPage-1)/10)*10) + 1
iTotalPage 	=  int((vTotalCount-1)/vPageSize) +1

If (vCurrPage mod vPageSize) = 0 Then
	iEndPage = vCurrPage
Else
	iEndPage = iStartPage + (10-1)
End If
%>
<tr bgcolor="FFFFFF">
	<td height="30" align="center">
		<% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
		<% else %>[pre]<% end if %>
        <%
			for ix = iStartPage  to iEndPage
				if (ix > iTotalPage) then Exit for
				if Cint(ix) = Cint(vCurrPage) then
		%>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="red">[<%=ix%>]</font></a>
		<%		else %>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
		<%
				end if
			next
		%>
    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
		<% else %>[next]<% end if %>
	</td>
</tr>
</table>
</form>
<% end if %>
<%
function Proc() ''신규.
    Dim strSql

    Dim i_isextusing : i_isextusing = Request("cisextusing")
    Dim vMakerID : vMakerID = Request("makerid")
    strSql = "Update db_user.dbo.tbl_user_c "& VbCRLF
    strSql = strSql & " Set isextusing='"&i_isextusing&"'"& VbCRLF
    strSql = strSql & " where userid='"&vMakerID&"'"& VbCRLF
    dbget.Execute strSql

    dim qItem, mayMallID
    For Each qItem In Request.Form
        if Left(qItem,6)="notin_" then
            ''rw qItem&"=="&Request.Form(qItem)
            mayMallID = Mid(qItem,7,255)

            if (Request.Form(qItem)="N") then ''등록제외설정
                strSql = "IF NOT Exists(select * from [db_temp].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&mayMallID&"' and makerid='"&vMakerID&"')"
                strSql = strSql&" BEGIN"
                strSql = strSql&" Insert into [db_temp].dbo.tbl_jaehyumall_not_in_makerid "
                strSql = strSql&" (makerid,mallgubun,regdate,reguserid)"
                strSql = strSql&" values('"&vMakerID&"','"&mayMallID&"',getdate(),'"&session("ssBctID")&"')"
                strSql = strSql&" END "
                dbget.Execute strSql

                strSql = "IF NOT Exists(select * from [db_outmall].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&mayMallID&"' and makerid='"&vMakerID&"')"
                strSql = strSql&" BEGIN"
                strSql = strSql&" Insert into [db_outmall].dbo.tbl_jaehyumall_not_in_makerid "
                strSql = strSql&" (makerid,mallgubun,regdate,reguserid)"
                strSql = strSql&" values('"&vMakerID&"','"&mayMallID&"',getdate(),'"&session("ssBctID")&"')"
                strSql = strSql&" END "
                dbCTget.Execute strSql
            else                              ''등록가능
                strSql = "IF Exists(select * from [db_temp].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&mayMallID&"' and makerid='"&vMakerID&"')"
                strSql = strSql&" BEGIN"
                strSql = strSql&" delete from [db_temp].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&mayMallID&"' and makerid='"&vMakerID&"'"
                strSql = strSql&" END "
                dbget.Execute strSql

                strSql = "IF Exists(select * from [db_outmall].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&mayMallID&"' and makerid='"&vMakerID&"')"
                strSql = strSql&" BEGIN"
                strSql = strSql&" delete from [db_outmall].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&mayMallID&"' and makerid='"&vMakerID&"'"
                strSql = strSql&" END "
                dbCTget.Execute strSql
            end if
        end if
    Next

    if (i_isextusing="N") then ''N로 저장하면 때려넣음.
        strSql = " Insert into [db_temp].dbo.tbl_jaehyumall_not_in_makerid"
        strSql = strSql&" (makerid,mallgubun,regdate,reguserid)"
        strSql = strSql&" select '"&vMakerID&"', K.mayMallID,getdate(),'"&session("ssBctID")&"'" &VbCRLF
        strSql = strSql&" from (select c.userid as mayMallID from db_user.dbo.tbl_user_c c Join db_partner.dbo.tbl_partner_addInfo f "
        strSql = strSql&"       on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=1 where c.isusing='Y' and c.userdiv='50') K "
        strSql = strSql&" left join [db_temp].dbo.tbl_jaehyumall_not_in_makerid T "
        strSql = strSql&" on K.mayMallID =T.mallgubun and T.makerid='"&vMakerID&"'"
        strSql = strSql&" where T.makerid is NULL"
        dbget.Execute strSql

        strSql = " Insert into [db_outmall].dbo.tbl_jaehyumall_not_in_makerid"
        strSql = strSql&" (makerid,mallgubun,regdate,reguserid)"
        strSql = strSql&" select '"&vMakerID&"', K.mayMallID,getdate(),'"&session("ssBctID")&"'" &VbCRLF
        strSql = strSql&" from (select c.userid as mayMallID from db_AppWish.dbo.tbl_user_c c Join db_AppWish.dbo.tbl_partner_addInfo f "
        strSql = strSql&"       on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=1 where c.isusing='Y' and c.userdiv='50') K "
        strSql = strSql&" left join [db_outmall].dbo.tbl_jaehyumall_not_in_makerid T "
        strSql = strSql&" on K.mayMallID =T.mallgubun and T.makerid='"&vMakerID&"'"
        strSql = strSql&" where T.makerid is NULL"
        dbCTget.Execute strSql
    end if

	strSql = " exec [db_log].[dbo].[usp_Ten_SaveJaehyuMallNotInMakeridChangeInfo] '" + CStr(vMakerID) + "', '" + CStr(session("ssBctID")) + "' "
	dbget.Execute strSql

end function

Function Proc_NotUsing() ''' 더이상 사용안함.
	Dim strSql, vAction, vMakerID, vMallGubun, vResult, vCurrPage, vIsAll, arrList, intLoop
	vAction = Request("action")
	vMallGubun = NullFillWith(Request("mallgubun"),"")
	vCurrPage = NullFillWith(Request("cp"),1)
	vIsAll = NullFillWith(Request("isall"),"")


	If vAction = "insert" Then
		vMakerID = Request("in_id")
		If vIsAll <> "" Then
			strSql = " select c.userid userid " & _
					 " from db_user.dbo.tbl_user_c c Join db_partner.dbo.tbl_partner_addInfo f on c.userid=f.partnerid and f.pcomType=1 where c.isusing='Y' and c.userdiv='50' and f.pmallSellType=1"
			strSql = " select 'interpark' union select 'lotteCom' union select 'lotteimall'"
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				arrList = rsget.getRows()
			END IF
			rsget.close

			IF isArray(arrList) THEN
				For intLoop =0 To UBound(arrList,2)
					vMallGubun = arrList(0,intLoop)
					strSql = "	DECLARE @Temp CHAR(1) " & _
							 "	If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = '" & vMallGubun & "' AND makerid = '" & vMakerID & "') " & _
							 "		BEGIN " & _
							 "			INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_makerid(makerid,mallgubun) VALUES('" & vMakerID & "','" & vMallGubun & "') " & _
							 "		END	"
					dbget.execute strSql

					strSql = "	DECLARE @Temp CHAR(1) " & _
							 "	If NOT EXISTS(SELECT * FROM [db_outmall].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = '" & vMallGubun & "' AND makerid = '" & vMakerID & "') " & _
							 "		BEGIN " & _
							 "			INSERT INTO [db_outmall].dbo.tbl_jaehyumall_not_in_makerid(makerid,mallgubun) VALUES('" & vMakerID & "','" & vMallGubun & "') " & _
							 "		END	"
					dbCTget.execute strSql
				Next
			End If
			vMallGubun = NullFillWith(Request("mallgubun"),"")
		Else
			strSql = "	DECLARE @Temp CHAR(1) " & _
					 "	If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = '" & vMallGubun & "' AND makerid = '" & vMakerID & "') " & _
					 "		BEGIN " & _
					 "			INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_makerid(makerid,mallgubun) VALUES('" & vMakerID & "','" & vMallGubun & "') " & _
					 "		END	"
			dbget.execute strSql, vResult

			strSql = "	DECLARE @Temp CHAR(1) " & _
					 "	If NOT EXISTS(SELECT * FROM [db_outmall].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = '" & vMallGubun & "' AND makerid = '" & vMakerID & "') " & _
					 "		BEGIN " & _
					 "			INSERT INTO [db_outmall].dbo.tbl_jaehyumall_not_in_makerid(makerid,mallgubun) VALUES('" & vMakerID & "','" & vMallGubun & "') " & _
					 "		END	"
			dbget.execute strSql, vResult

			If vResult <> "1" Then
				Response.Write "<script>alert('등록이 되어있는\n브랜드입니다.');location.href='JaehyuMall_Not_In_Makerid.asp?mallgubun=" & vMallGubun & "&cp=" & vCurrPage & "';</script>"
				dbget.close()
				Response.End
			End If
		End If
	ElseIf vAction = "delete" Then
	    dim del_id
		del_id = Replace(Request("del_id")," ","")
		if (Right(del_id,1)=",") then
		    del_id=Left(del_id,Len(del_id)-1)
		end if
		strSql = "DELETE [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE idx in ("&del_id&")"
		dbget.execute strSql

		strSql = "DELETE [db_outmall].dbo.tbl_jaehyumall_not_in_makerid WHERE idx in ("&del_id&")"
		dbCTget.execute strSql
	End IF

	Response.Write "<script>alert('처리되었습니다.');location.href='JaehyuMall_Not_In_Makerid.asp?mallgubun=" & vMallGubun & "&makerid="&vMakerID&"&cp=" & vCurrPage & "';</script>"
	Response.End
End Function

Public Function fnPotalList
	Dim strSql
	strSql = ""
	strSql = strSql & " select E.mallgubun, M.makerid, isnull(M.isusing, 'Y') as isusing, M.regdate, M.lastupdate, M.regid, M.updateid from db_temp.dbo.tbl_Epshop as E left join db_temp.dbo.tbl_Epshop_not_in_makerid as M on E.mallgubun = M.mallgubun and M.makerid = '"&vMakerid&"' "
	rsget.Open strSql,dbget,1
	'response.write strSql
	IF not rsget.EOF THEN
		fnPotalList = rsget.getRows()
	END IF
	rsget.close
End Function

Function potalProc()
    dim qItem, mayMallID
    For Each qItem In Request.Form
        if Left(qItem,10)="epIsusing_" then
'            rw qItem&"=="&Request.Form(qItem)
            mayMallID = Mid(qItem,11,255)
			strSql = "IF NOT Exists(select * from db_temp.dbo.tbl_EpShop_not_in_makerid where mallgubun='"&mayMallID&"' and makerid='"&vMakerID&"')"
			strSql = strSql&" BEGIN"
			strSql = strSql & " INSERT INTO db_temp.dbo.tbl_EpShop_not_in_makerid (makerid, mallgubun, isusing, regdate, regid) VALUES "
			strSql = strSql & " ('"&vMakerID&"', '"&mayMallID&"', '"&Request.Form(qItem)&"' ,getdate(), '"&session("ssBctID")&"') "
            strSql = strSql&" END Else "
			strSql = strSql&" BEGIN"
			strSql = strSql & " UPDATE db_temp.dbo.tbl_EpShop_not_in_makerid SET "
			strSql = strSql & " isusing = '"&Request.Form(qItem)&"'"
			strSql = strSql & " ,lastupdate = getdate()"
			strSql = strSql & " ,updateid = '"&session("ssBctID")&"'"
			strSql = strSql & " WHERE makerid = '"&vMakerID&"' "
			strSql = strSql & " AND mallgubun = '"&mayMallID&"' "
            strSql = strSql&" END "
            dbget.Execute strSql

			strSql = "IF NOT Exists(select * from db_outmall.dbo.tbl_EpShop_not_in_makerid where mallgubun='"&mayMallID&"' and makerid='"&vMakerID&"')"
			strSql = strSql&" BEGIN"
			strSql = strSql & " INSERT INTO db_outmall.dbo.tbl_EpShop_not_in_makerid (makerid, mallgubun, isusing, regdate, regid) VALUES "
			strSql = strSql & " ('"&vMakerID&"', '"&mayMallID&"', '"&Request.Form(qItem)&"' ,getdate(), '"&vMakerID&"') "
            strSql = strSql&" END Else "
			strSql = strSql&" BEGIN"
			strSql = strSql & " UPDATE db_outmall.dbo.tbl_EpShop_not_in_makerid SET "
			strSql = strSql & " isusing = '"&Request.Form(qItem)&"'"
			strSql = strSql & " ,lastupdate = getdate()"
			strSql = strSql & " ,updateid = '"&vMakerID&"'"
			strSql = strSql & " WHERE makerid = '"&vMakerID&"' "
			strSql = strSql & " AND mallgubun = '"&mayMallID&"' "
            strSql = strSql&" END "
            dbCTget.Execute strSql
        end if
    Next
End Function
%>
<br><br>
포탈 가격비교
<%
Dim arrPotalList
arrPotalList = fnPotalList
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<input type="hidden" name="action" value="epsel">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>구분</td>
    <td>판매설정</td>
	<td>등록일</td>
	<td>최종수정일</td>
</tr>
<%
IF isArray(arrPotalList) THEN
	For i =0 To UBound(arrPotalList,2)
%>
<tr align="center" bgcolor="#FFFFFF" >
	<td>
		<%
			Select Case arrPotalList(0,i)
				Case "naverep" response.write "네이버"
				Case "daumep" response.write "다음"
			End Select
		%>
	</td>
	<td>
		<input type="radio" name="epIsusing_<%=arrPotalList(0,i)%>" value="Y" <%=CHKIIF(arrPotalList(2,i)="Y" ,"checked","") %>>판매함
		<input type="radio" name="epIsusing_<%=arrPotalList(0,i)%>" value="N" <%=CHKIIF(arrPotalList(2,i)="N" ,"checked","") %>>판매안함
	</td>
	<td><%= arrPotalList(3,i) %></td>
	<td><%= arrPotalList(4,i) %></td>
</tr>
<%
	Next
End If
%>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="6">
		<input type="button" class="button" value="포탈 가격비교 저장" onClick="jsIsusing(this)">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
