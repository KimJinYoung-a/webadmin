<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%

dim strSql
dim vMallGubun, vMakerID, vAction, i
dim intLoop
Dim cisextusing : cisextusing="N"
Dim isValidMakerid : isValidMakerid = False
Dim arrListBrBrand, arrListLogBrand

vMallGubun	= RequestCheckVar(request("mallgubun"),32)
vMakerID	= session("ssBctID")
vAction		= RequestCheckVar(request("action"),32)

if (vMallGubun="lotte") then vMallGubun="lotteCom"

if (vMakerID = "") then
	response.write "�߸��� �����Դϴ�."
	response.end
end if

'// ============================================================================
If vAction = "upsel" Then
	Call Proc()
ElseIf vAction = "epsel" Then
	Call potalProc()
End If


'// ============================================================================
''�귣�� ��ǥ ���� �˻�
strSql = "select top 1 isextusing from db_user.dbo.tbl_user_c"
strSql = strSql & " where userid='"&vMakerID&"'"

rsget.Open strSql,dbget
if Not rsget.Eof then
	isValidMakerid = True
	cisextusing = rsget("isextusing")
end if
rsget.close

'2014-01-27 ä���� ��û..�Ե����Ŀ� ���� ���̺� ��ϵ� �귣��� ��ü�� �������ϰ� �ش޶��..
Dim onlyLotteMKTmodify
strSql = ""
strSql = strSql & " SELECT COUNT(*) as cnt FROM db_temp.dbo.tbl_Lotte_not_in_makerid_By_KimJinYoung WHERE makerid = '"&vMakerID&"' and isusing = 'Y' "
rsget.Open strSql,dbget
If rsget("cnt") > 0 Then
	onlyLotteMKTmodify = "o"
Else
	onlyLotteMKTmodify = "x"
End If
rsget.Close

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

'2014-01-27 ä���� ��û..�Ե����Ŀ� ���� ���̺� ��ϵ� �귣��� ��ü�� �������ϰ� �ش޶��..�ƾ� ����Ʈ ���x
If onlyLotteMKTmodify = "o" Then
	strSql = strSql & " and c.userid <> 'lotteCom' "
End If

rsget.Open strSql,dbget
IF not rsget.EOF THEN
	arrListBrBrand = rsget.getRows()
END IF
rsget.close

'// �α�
strSql = " select top 5 mallgubun, makerid, useYN, reguserid, regdate from "
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
%>
<script language="javascript">

function chkComp(comp){
    var frm = comp.form;

	if (comp.value == "N") {
		for (var i=0;i<frm.elements.length;i++){
			var e=frm.elements[i];
			if (e.name.substring(0,6)=="notin_"){
				if (e.value == "N") {
					e.checked = true;
				}
			}
		}
	}

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

<form name="frmsearch" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
		    <td width="90%"></td>
		</tr>
		<tr>
			<td >
			�귣��ID : <%= vMakerID %>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<br>
<strong>���޸� ����</strong>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmBrdUsing" method="post" action="">
<input type="hidden" name="action" value="upsel">
<input type="hidden" name="menupos" value="<%=menupos%>">
<% if (Not isValidMakerid) then %>
<tr>
	<td align="center" bgcolor="#FFFFFF"><%= vMakerID %>�� �ùٸ� �귣��ID�� �ƴմϴ�.</td>
</tr>
<% else %>
<tr align="center" bgcolor="#DDDDDD">
	<td width="200" >������</td>
	<td width="200" >�Ǹż���</td>
	<td width="200" >�����</td>
	<td >����������</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td >���޻� �Ǹſ���</td>
	<td >
		<input type="radio" name="cisextusing" value="Y" <%=CHKIIF(cisextusing="Y","checked","") %> onClick="chkComp(this)">�Ǹ���
		<input type="radio" name="cisextusing" value="N" <%=CHKIIF(cisextusing="N","checked","") %> onClick="chkComp(this)">
		<% if cisextusing="N" then %>
		<b>�Ǹž���</b>
		<% else %>
		�Ǹž���
		<% end if %>
	</td>
	<td colspan="2">
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
				<input type="radio" name="notin_<%=arrListBrBrand(0,intLoop)%>" value="" checked <%=CHKIIF(cisextusing="N","disabled","") %> >�Ǹ���
				<input type="radio" name="notin_<%=arrListBrBrand(0,intLoop)%>" value="N" <%=CHKIIF(cisextusing="N","disabled","") %> >�Ǹž���
				<% else %>
				<input type="radio" name="notin_<%=arrListBrBrand(0,intLoop)%>" value=""  <%=CHKIIF(cisextusing="N","disabled","") %> >�Ǹ���
				<input type="radio" name="notin_<%=arrListBrBrand(0,intLoop)%>" value="N" checked <%=CHKIIF(cisextusing="N","disabled","") %> >�Ǹž���
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

<!--
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
-->

<%

function Proc() ''�ű�.
    Dim strSql

    Dim i_isextusing : i_isextusing = Request("cisextusing")
    Dim vMakerID : vMakerID = session("ssBctID")
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
			strSql = strSql & " ('"&vMakerID&"', '"&mayMallID&"', '"&Request.Form(qItem)&"' ,getdate(), '"&vMakerID&"') "
            strSql = strSql&" END Else "
			strSql = strSql&" BEGIN"
			strSql = strSql & " UPDATE db_temp.dbo.tbl_EpShop_not_in_makerid SET "
			strSql = strSql & " isusing = '"&Request.Form(qItem)&"'"
			strSql = strSql & " ,lastupdate = getdate()"
			strSql = strSql & " ,updateid = '"&vMakerID&"'"
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
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
