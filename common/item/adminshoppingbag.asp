<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �¶��� & �������� ���� ��ٱ���
' Hieditor : 2011.08.02 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/stock/shortagestock_cls.asp" -->
<!-- #include virtual="/lib/classes/items/adminshoppingbag/adminshoppingbag_cls.asp" -->
<%
dim IS_HIDE_BUYCASH : IS_HIDE_BUYCASH = False

'// ���԰� �ҽ����� ����, skyer9, 2018-02-14
dim PriceEditEnable : PriceEditEnable = False

dim itemgubunarr , itemidarr , itemoptionarr, itemnoarr, onoffgubun, shopid , i ,acURL ,research
dim obaginsert , userid , isusing ,makerid ,itemid ,itemname ,comm_cd ,cdl ,cdm ,cds ,obag, myorderyn
	itemgubunarr = request("itemgubunarr")
	itemidarr = request("itemidarr")
	itemoptionarr = request("itemoptionarr")
	itemnoarr = request("itemnoarr")
	onoffgubun = requestCheckVar(request("onoffgubun"),10)
	userid = session("ssBctId")
    isusing = requestCheckVar(request("isusing"),1)
    makerid = requestCheckVar(request("makerid"),32)
    itemid = requestCheckVar(request("itemid"),10)
    itemname = requestCheckVar(request("itemname"),64)
    comm_cd = requestCheckVar(request("comm_cd"),32)
    cdl = requestCheckVar(request("cdl"),3)
    cdm = requestCheckVar(request("cdm"),3)
    cds = requestCheckVar(request("cds"),3)
	shopid = requestCheckVar(request("shopid"),32)
    research = requestCheckVar(request("research"),2)
    myorderyn = requestcheckvar(request("myorderyn"),1)

if (research<>"on") and (isusing="") then
    isusing = "Y"
end if
if (research<>"on") and (myorderyn="") then myorderyn="Y"

if C_ADMIN_USER then

'/�����ϰ�� ���� ���常 ��밡��
elseif (C_IS_SHOP) then
	IS_HIDE_BUYCASH = True
	myorderyn = "Y"

	'/���α��� ���� �̸�
	'if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	'end if
else
	myorderyn = "Y"

	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then

		else

		end if
	end if
end if

'/�¿��� ���� ������ �ʱⰪ ��������(OFF)
if onoffgubun = "" then onoffgubun = "OFF"
if onoffgubun = "" then
	response.write "<script>alert('�¶��� & ���������� �����ϴ�'); self.close();</script>"
	dbget.close() : response.end
end if

'//��ٱ��� �߰�
'response.write userid &"/"& onoffgubun&"/"& shopid&"/"& itemgubunarr&"/"& itemidarr&"/"& itemoptionarr&"/"& itemnoarr &"<Br>"
putadminshoppingbag_insert userid, onoffgubun, shopid, itemgubunarr, itemidarr, itemoptionarr, itemnoarr , "self" ,menupos

set obag  = new cadminshoppingbag_list
	obag.FPageSize = 300
	obag.FCurrPage = 1
    obag.frectcdl = cdl
    obag.frectcdm = cdm
    obag.frectcds = cds
    obag.Frectshopid = shopid
    obag.Frectisusing = isusing

    'if onoffgubun = "" and itemgubunarr = "" and itemidarr = "" and itemoptionarr = "" and itemnoarr = "" then
    	obag.Frectmakerid = makerid
    'end if

    obag.Frectitemid = itemid
    obag.Frectitemname = itemname
    obag.Frectcomm_cd = comm_cd
    obag.frectonoffgubun = onoffgubun

	if myorderyn="Y" then
		obag.frectuserid = userid
	end if

	'/�¶��� ��ٱ��� ����Ʈ
    if onoffgubun = "ON" then
    	obag.fadminshoppingbag_on

    '/�������� ��ٱ��� ����Ʈ
    elseif onoffgubun = "OFF" then
        obag.fadminshoppingbag_off

	    'if shopid = "" then
	    '    response.write "<script language='javascript'>"
	    '    response.write "    alert('������ �����ϼž� �ֹ��� ���� �մϴ�');"
	    '    response.write "</script>"
	    'end if
    end if

'//�űԻ�ǰ �߰��� �˾����� �Ѿ ���		'/�����˾����� �׼� �������� ��ä�� �ѱ��
acURL =Server.HTMLEncode("/common/item/adminshoppingbag_process.asp?onoffgubun="&onoffgubun)
%>

<font color="red">�� <%= userid %> ���� <%= onoffgubun %>LINE ��ٱ���</font>

<%
'/�¶��� ��ٱ���
if onoffgubun = "ON" then
%>

<%
'/�������� ��ٱ���
elseif onoffgubun = "OFF" then
%>
    <Br>&nbsp;&nbsp;&nbsp;- �������� �ֹ� : ���걸��(�ٹ�����Ư��/������/���Ư��)
    <br>&nbsp;&nbsp;&nbsp;- ��ü �ֹ� : ���걸��(��üƯ��/��ü����)
    <br>&nbsp;&nbsp;&nbsp;- �ʿ����(7��) = (7���Ǹź� x 1) - (��ȿ��� + ���ֹ���)
	<!-- �˻� ���� -->
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="onoffgubun" value="<%= onoffgubun %>">
	<tr align="center" bgcolor="#FFFFFF" >
	    <td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	    <td align="left">
	        ���� :
	        <% if C_ADMIN_USER then %>
				<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
	        <% elseif (C_IS_SHOP) then %>
	    		<%= shopid %>
	    	<% else %>
				<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
	        <% end if %>

	        ��뿩��:<% drawSelectBoxUsingYN "isusing", isusing %>
	        <!-- #include virtual="/common/module/categoryselectbox.asp"-->
	    </td>
	    <td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
	        <input type="button" class="button_s" value="�˻�" onClick="javascript:reg(frm);">
	    </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
	    <td align="left">
	        �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
	        &nbsp;
	        ��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9" onKeyPress="if (event.keyCode == 13) reg(frm);">
	        &nbsp;
	        ��ǰ�� : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) reg(frm);">
	        ������� : <% drawSelectBoxOFFJungsanCommCDmulti "comm_cd",comm_cd %>
	    </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
	    <td align="left">
	    	<% 'if C_ADMIN_USER then %>
				������ٱ��ϸ�����<input type="checkbox" name="myorderyn" value="Y" <% if myorderyn="Y" then response.write " checked" %>>
			<% 'end if %>
	    </td>
	</tr>
	</form>
	</table>
	<!-- �˻� �� -->
	<br>

	<!-- �׼� ���� -->
	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
	    <td align="left">
	        <input type="button" class="button" value="���ü���" onclick="bageditarr(frmbag)">
	        <input type="button" class="button" value="���û���" onclick="bagdelarr(frmbag)">
	    </td>
	    <td align="right">
			<% if True or (session("ssBctCname") = "�̻�") then %>
			<input type="button" value="����ǰ�߰�" onclick="jsAddNewItemOFF(frm, '<%= shopid %>', '<%= acURL %>');" class="button">
			<% else %>
	    	<input type="button" value="����ǰ�߰�" onclick="addnewItem('<%=onoffgubun%>',frm,'<%=shopid%>','<%=acURL%>');" class="button">
			<% end if %>
	    	<%' if shopid <> "" then %>
		        <% if obag.FresultCount>0 then %>
		            <input type="button" class="button" value="�����ֹ��ۼ�(�ٹ����ٹ���)" onclick="AddArr(frmArrupdate,'<%=C_IS_SHOP%>')">
		        <% end if %>
		        <% if obag.FresultCount>0 then %>
		        	<%' if makerid <> "" or comm_cd = "B012" or comm_cd = "B022" then %>
		            	<input type="button" class="button" value="�����ֹ��ۼ�(��ü)" onclick="AddArr_upche(frmArrupdate,'<%=C_IS_SHOP%>')">
		            <%' end if %>
		        <% end if %>
		    <%' end if %>
	    </td>
	</tr>
	</table>
	<!-- �׼� �� -->

	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
	    <td colspan="20">
	        �˻���� : <b><%= obag.FTotalcount %></b> ���ִ� 300�Ǳ��� ���� �˴ϴ�.
	    </td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td>����</td>
	    <td>
	    	����ó
	    </td>
	    <td>�귣��</td>
	    <td>��ǰ�ڵ�</td>
	    <td>�̹���</td>
	    <td>��ǰ��<br>[�ɼǸ�]</td>
	    <td>�ǸŰ�</td>

	    <% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
	    	<td>��ü<br>���԰�</td>
	    <% end if %>

	    <td>����<br>���ް�</td>
    	<td>��ȿ<br>���</td>
	    <td>
	    	�Ǹż���<br>(7��)
	    </td>
	    <td>
	        �ʿ����<br>(7��)
	    </td>
	    <td>����</td>
	    <td>�����</td>
	    <td>���</td>
	</tr>
	<% if obag.FresultCount > 0 then %>
	<% for i=0 to obag.FresultCount -1 %>
	<form method="get" action="" name="frmBuyPrc<%=i%>">

	<% if obag.FItemlist(i).Fisusing="N" then %>
		<tr bgcolor="#EEEEEE" align="center">
	<% else %>
		<tr bgcolor="#FFFFFF" align="center">
	<% end if %>
	<input type="hidden" name="onlinebuycash" value="<%= obag.FItemlist(i).fonlinebuycash %>">
	<input type="hidden" name="onlinemwdiv" value="<%= obag.FItemlist(i).fonlinemwdiv %>">
	<input type="hidden" name="bagidx" value="<%= obag.FItemlist(i).fbagidx %>">
	<input type="hidden" name="itemgubun" value="<%= obag.FItemlist(i).fitemgubun %>">
	<input type="hidden" name="itemid" value="<%= obag.FItemlist(i).fitemid %>">
	<input type="hidden" name="itemoption" value="<%= obag.FItemlist(i).fitemoption %>">
	<input type="hidden" name="shopitemprice" value="<%= obag.FItemlist(i).fshopitemprice %>">
	<input type="hidden" name="itemname" value="<%= obag.FItemlist(i).fshopitemname %>">
	<input type="hidden" name="itemoptionname" value="<%= obag.FItemlist(i).fshopitemoptionname %>">
	<input type="hidden" name="makerid" value="<%= obag.FItemlist(i).fmakerid %>">
	<input type="hidden" name="comm_cd" value="<%= obag.FItemlist(i).fcomm_cd %>">
	<% if IS_HIDE_BUYCASH = True then %>
	<input type="hidden" name="shopsuplycash" value="-1">
	<% else %>
	<input type="hidden" name="shopsuplycash" value="<%= obag.FItemlist(i).fshopsuplycash %>">
	<% end if %>
	<input type="hidden" name="shopbuyprice" value="<%= obag.FItemlist(i).fshopbuyprice %>">
	<input type="hidden" name="shopid" value="<%= obag.FItemlist(i).fshopid %>">
	    <td width=20>
	        <input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
	    </td>
	    <td>
	    	<%= obag.FItemlist(i).fshopname %>
	        <Br><%= obag.FItemlist(i).fshopid %>
	    </td>
	    <td width=100>
	        <%= GetdeliverGubunName(obag.FItemlist(i).fcomm_cd) %><br>(<%= obag.FItemlist(i).fcomm_name %>)
	    </td>
	    <td>
	        <a href="javascript:searchmakerid('<%= obag.FItemlist(i).fmakerid %>',frm);" onfocus="this.blur()"><%= obag.FItemlist(i).fmakerid %></a>
	    </td>
	    <td width=80>
	        <%= obag.FItemlist(i).Fitemgubun %><%=  CHKIIF(obag.FItemlist(i).Fitemid>=1000000,Format00(8,obag.FItemlist(i).Fitemid),Format00(6,obag.FItemlist(i).Fitemid)) %><%= obag.FItemlist(i).Fitemoption %>
	        <% if obag.FItemlist(i).Fitemgubun="10" then %>
	        	<Br><a href="<%=wwwUrl%>/shopping/category_prd.asp?itemid=<%=obag.FItemlist(i).Fitemid%>" target="_blink" onfocus="this.blur()">[��]</a>
	        <% end if %>
	    </td>
	    <td width=50>
	        <img src="<%= obag.FItemlist(i).GetImageSmall %>" width=50 height=50 border=0>
	    </td>
	    <td align="left">
	        <%= obag.FItemlist(i).fshopitemname %><Br>
	        <% if obag.FItemlist(i).fshopitemoptionname <> "" then %>
	            [<%=obag.FItemlist(i).fshopitemoptionname%>]
	        <% end if %>
	    </td>
	    <td align="right" width=60>
	        <%= FormatNumber(obag.FItemlist(i).fshopitemprice,0) %>
	    </td>

	    <% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		    <td align="right" width=80>
		        <%= FormatNumber(obag.FItemlist(i).fshopsuplycash,0) %>

		        <% if obag.FItemlist(i).fcentermwdiv="M" then %>
		        	<p>ON:<%= FormatNumber(obag.FItemlist(i).fonlinebuycash,0) %>
		        <% end if %>
		    </td>
		<% end if %>

	    <td align="right" width=60>
	        <%= FormatNumber(obag.FItemlist(i).fshopbuyprice,0) %>
	    </td>
	    <td width=60>
	        <%= obag.FItemlist(i).getAvailStock %>     <!--��ȿ���-->
	    </td>
	    <td width=60>
	        <%= obag.FItemlist(i).fsell7days %> (7��)      <!--�Ǹż���-->
	    </td>
	    <td width=60>
	        <!-- ���ʿ���� -->
	        <% if obag.FItemlist(i).frequire7daystock > 0 then %>
	            <a href="javascript:inputiteno('<%= obag.FItemlist(i).frequire7daystock %>',frmBuyPrc<%= i %>);" onfocus="this.blur()"><p>
	            <font color="red"><%= obag.FItemlist(i).frequire7daystock %> (7��)</font>
	            </a>
	        <% else %>
	           0 (7��)
	        <% end if %>
	    </td>
	    <td width=60>
	        <input type="text" class="text" name="itemno" value="<%= obag.FItemlist(i).fitemno %>" size="4" maxlength="4" onKeyDown="CheckThis(frmBuyPrc<%= i %>);">
	    </td>
	    <td width=90>
	        <%= obag.FItemlist(i).fuserid %>
	    </td>
	    <td>
	        <% if obag.FItemList(i).Fpreorderno>0 then %>
	        	���ֹ�:
	            <% if obag.FItemList(i).Fpreorderno<>obag.FItemList(i).Fpreordernofix then response.write CStr(obag.FItemList(i).Fpreorderno) + " -> " %>
	        	<%= obag.FItemList(i).Fpreordernofix %><br>
	        <% end if %>
	    </td>
	</tr>
	</form>
	<% next %>

	<% else %>

	<tr bgcolor="#FFFFFF">
	    <td colspan="20" align="center">[��ٱ��Ͽ� ��ǰ�� �����ϴ�.]</td>
	</tr>
	<% end if %>
	<form name="frmArrupdate" method="post" action="">
	    <input type="hidden" name="mode" value="arrins">
	    <input type="hidden" name="itemgubunarr2" value="">
	    <input type="hidden" name="itemidadd2" value="">
	    <input type="hidden" name="itemoptionarr2" value="">
	    <input type="hidden" name="sellcasharr2" value="">
	    <input type="hidden" name="buycasharr2" value="">
	    <input type="hidden" name="suplycasharr2" value="">
	    <input type="hidden" name="itemnoarr2" value="">
	    <input type="hidden" name="itemnamearr2" value="">
	    <input type="hidden" name="itemoptionnamearr2" value="">
	    <input type="hidden" name="designerarr2" value="">
	    <input type="hidden" name="shopid" value="<%=shopid%>">
	    <input type="hidden" name="suplyer" value="10x10">
	    <input type="hidden" name="idx" value="0">
	    <input type="hidden" name="chargeid" value="<%=makerid%>">
	    <input type="hidden" name="shopbuypricearr2" value="">
	    <input type="hidden" name="isreq" value="Y">
	    <input type="hidden" name="bagidxarr">
	    <input type="hidden" name="cwflag">
	</form>
	<form name="frmbag" method="post" action="">
		<input type="hidden" name="mode">
		<input type="hidden" name="bagidxarr">
	    <input type="hidden" name="onoffgubun">
	    <input type="hidden" name="itemgubunarr">
	    <input type="hidden" name="itemidarr">
	    <input type="hidden" name="itemoptionarr">
	    <input type="hidden" name="itemnoarr">
	    <input type="hidden" name="makerid">
	    <input type="hidden" name="shopid" >
	</form>
	</table>

	<!-- �׼� ���� -->
	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
	    <td align="left">
	        <input type="button" class="button" value="���ü���" onclick="bageditarr(frmbag)">
	        <input type="button" class="button" value="���û���" onclick="bagdelarr(frmbag)">
	    </td>
	    <td align="right">
	    	<input type="button" value="����ǰ�߰�" onclick="addnewItem('<%=onoffgubun%>',frm,'<%=shopid%>','<%=acURL%>');" class="button">
	    	<%' if shopid <> "" then %>
		        <% if obag.FresultCount>0 then %>
		            <input type="button" class="button" value="�����ֹ��ۼ�(�ٹ����ٹ���)" onclick="AddArr(frmArrupdate,'<%=C_IS_SHOP%>')">
		        <% end if %>
		        <% if obag.FresultCount>0 then %>
		        	<%' if makerid <> "" or comm_cd = "B012" or comm_cd = "B022" then %>
		            	<input type="button" class="button" value="�����ֹ��ۼ�(��ü)" onclick="AddArr_upche(frmArrupdate,'<%=C_IS_SHOP%>')">
		            <%' end if %>
		        <% end if %>
		    <%' end if %>
	    </td>
	</tr>
	</table>
	<!-- �׼� �� -->
<% end if %>
<iframe id="view" name="view" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<%
set obag = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
