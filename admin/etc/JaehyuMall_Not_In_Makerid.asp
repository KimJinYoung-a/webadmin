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
''20130304 �������̽����� - ������

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

	if (vMallGubun="lotte") then vMallGubun="lotteCom"   ''' 20130304 �߰�

	''If vAction = "insert" OR vAction = "delete"  Then
	If vAction = "upsel" Then
		Call Proc()
	ElseIf vAction = "epsel" Then
		Call potalProc()
	End If

	''�귣�� ��ǥ ���� �˻�
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

	''�귣�庰 ���� ��뿩��
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

	'// �α�
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
		alert("ID�� �Է��ϼ���.");
		frm.in_id.focus();
		return;
	}

	if ((!frm.isall.checked)&&(frm.mallgubun.value.length<1)){
	    alert('����� Mall ���� �Ǵ� [��� Mall�� ����] üũ������ �ʿ��մϴ�.');
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
        alert('���õ� ������ �����ϴ�.');
        return;
    }

    if (confirm('���� �귣�忡 ���� ���� �Ǹż��� �����Ͻðڽ��ϱ�?')){
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
    if (!confirm('���޻� �귣�� �Ǹż����� �����Ͻðڽ��ϱ�?')){
        return;
    }

    comp.form.submit();
}

function jsIsusing(ep){
    if (!confirm('��Ż ���ݺ� �Ǹſ��θ� �����Ͻðڽ��ϱ�?')){
        return;
    }
    ep.form.submit();
}
</script>

<center>
Mall ���� : <b><%=vMallGubun%></b>
</center>
<br>
<form name="frmsearch" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
		    <td width="90%">Mall ���� : <% CALL DrawApiMallSelect("mallgubun",vMallGubun) %></td>
		    <td rowspan="4" width="10%"><input type="submit" value="�� ��" style="width:50px;height:50px;"></td>
		</tr>
		<tr>
			<td >
			�귣��ID : <input type="text" class="text" name="makerid" value="<%=vMakerID%>" size="20"> <input type="button" class="button" value="ID�˻�" onclick="jsSearchBrandID(this.form.name,'makerid');" >
			</td>

		</tr>
		<!--
		<tr>
			<td>�귣���(�ѱ�) : <input type="text" class="text" name="brandnamekr" value="<%=vBrandNameKr%>" size="30"></td>
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
        <td align="center" bgcolor="#FFFFFF"><%= vMakerID %>�� �ùٸ� �귣��ID�� �ƴմϴ�.</td>
    </tr>
    <% else %>
    <tr align="center" bgcolor="#DDDDDD">
        <td width="200" >������</td>
        <td width="200" >�Ǹż���</td>
        <td width="100" >���������</td>
        <td >������ܼ�����</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" >
        <td >���޻� ��ü ��뿩��</td>
        <td >
            <input type="radio" name="cisextusing" value="Y" <%=CHKIIF(cisextusing="Y","checked","") %> onClick="chkComp(this)">���
            <input type="radio" name="cisextusing" value="N" <%=CHKIIF(cisextusing="N","checked","") %> onClick="chkComp(this)">
            <% if cisextusing="N" then %>
            <b>������</b>
            <% else %>
            ������
            <% end if %>
        </td>
        <td colspan="2">
        �̼����� [������] �ΰ�� �Ʒ� ���� ������ ������� �Ǹž���
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
        		    <input type="radio" name="notin_<%=arrListBrBrand(0,intLoop)%>" value="" checked <%=CHKIIF(cisextusing="N","disabled","") %> >��ϰ���
        		    <input type="radio" name="notin_<%=arrListBrBrand(0,intLoop)%>" value="N" <%=CHKIIF(cisextusing="N","disabled","") %> >�������
        		    <% else %>
        		    <input type="radio" name="notin_<%=arrListBrBrand(0,intLoop)%>" value=""  <%=CHKIIF(cisextusing="N","disabled","") %> >��ϰ���
        		    <input type="radio" name="notin_<%=arrListBrBrand(0,intLoop)%>" value="N" checked <%=CHKIIF(cisextusing="N","disabled","") %> >�������
        		    <% end if %>
        		</td>
        		<td><%=arrListBrBrand(3,intLoop)%></td>
        		<td><%=arrListBrBrand(2,intLoop)%></td>
        	</tr>
        <% Next %>
    <% end if %>
    <tr align="center" bgcolor="#FFFFFF">
        <td colspan="4">
         <input type="button" value="���޸� �귣�� �Ǹż��� ����" onClick="saveUsing(this)">
        </td>
    </tr>
    <% end if %>
    </form>
    </table>

	<p>

	<% if (isValidMakerid and isArray(arrListLogBrand)) then %>
		<br><br>
		[���� �ǸŻ���]
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="#DDDDDD">
			<td width="200" >������</td>
			<td width="200" >�ǸŻ���</td>
			<td width="100" >�����</td>
			<td >�����</td>
		</tr>
		<% if isArray(arrListLogBrand) then %>
			<% For intLoop =0 To UBound(arrListLogBrand,2) %>
				<tr align="center" bgcolor="#FFFFFF" height="30">
					<td>
						<% if (arrListLogBrand(0,intLoop) = "") then %>
							���޸� ��ü
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
				���� �귣��ID
				<input type="text" name="in_id" value="" size="10" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ insert_id(); return false;}">
				&nbsp;<input type="checkbox" name="isall" value="o">��ü����(
				)
				<input type="button" value="������� �귣�� ����" onClick="insert_id()">
			-->
			</td>
			<td width="20%" align="right">�귣��� : <b><%=vTotalCount%></b></td>
		</tr>
		</table>
	</td>
</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
    <td width="30%">������</td>
	<td width="30%">�귣��ID</td>
	<td width="20%">�����</td>
	<td width="15%">�����</td>
	<td width="5%">����</td>
	<!--
	<td width="20%"><input type="button" value="���� �귣��ID ����" onClick="delete_id()"></td>
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
		<td><a href="?mallgubun=<%=vmallgubun%>&makerid=<%=arrList(0,intLoop)%>">[����]</a></td>
		<!--
		<td><input type="checkbox" name="del_id" value="<%=arrList(3,intLoop)%>"></td>
		-->
	</tr>
<%
		Next
	Else
%>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="5" align="center" class="page_link">[�����Ͱ� �����ϴ�.]</td>
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
function Proc() ''�ű�.
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

            if (Request.Form(qItem)="N") then ''������ܼ���
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
            else                              ''��ϰ���
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

    if (i_isextusing="N") then ''N�� �����ϸ� ��������.
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

Function Proc_NotUsing() ''' ���̻� ������.
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
				Response.Write "<script>alert('����� �Ǿ��ִ�\n�귣���Դϴ�.');location.href='JaehyuMall_Not_In_Makerid.asp?mallgubun=" & vMallGubun & "&cp=" & vCurrPage & "';</script>"
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

	Response.Write "<script>alert('ó���Ǿ����ϴ�.');location.href='JaehyuMall_Not_In_Makerid.asp?mallgubun=" & vMallGubun & "&makerid="&vMakerID&"&cp=" & vCurrPage & "';</script>"
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
��Ż ���ݺ�
<%
Dim arrPotalList
arrPotalList = fnPotalList
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<input type="hidden" name="action" value="epsel">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>����</td>
    <td>�Ǹż���</td>
	<td>�����</td>
	<td>����������</td>
</tr>
<%
IF isArray(arrPotalList) THEN
	For i =0 To UBound(arrPotalList,2)
%>
<tr align="center" bgcolor="#FFFFFF" >
	<td>
		<%
			Select Case arrPotalList(0,i)
				Case "naverep" response.write "���̹�"
				Case "daumep" response.write "����"
			End Select
		%>
	</td>
	<td>
		<input type="radio" name="epIsusing_<%=arrPotalList(0,i)%>" value="Y" <%=CHKIIF(arrPotalList(2,i)="Y" ,"checked","") %>>�Ǹ���
		<input type="radio" name="epIsusing_<%=arrPotalList(0,i)%>" value="N" <%=CHKIIF(arrPotalList(2,i)="N" ,"checked","") %>>�Ǹž���
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
		<input type="button" class="button" value="��Ż ���ݺ� ����" onClick="jsIsusing(this)">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
