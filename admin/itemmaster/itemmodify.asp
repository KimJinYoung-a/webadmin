<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ����
' History : 2009.04.17 ���ʻ����ڸ�
'			2016.07.06 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
CONST CBASIC_IMG_MAXSIZE = 560   'KB
CONST CMAIN_IMG_MAXSIZE = 640   'KB

dim itemid, oitem, oitemvideo
dim makerid
dim chkMWAuth 'mw ���氡���� �������� üũ 
dim rentalItemFlag

itemid = request("itemid")
makerid = request("makerid")
menupos = request("menupos")
if (itemid = "") then
    response.write "<script>alert('�߸��� �����Դϴ�.'); history.back();</script>"
    dbget.close()	:	response.End
end if


'==============================================================================
set oitem = new CItem

oitem.FRectItemID = itemid
oitem.GetOneItem

Set oitemvideo = New CItem
oitemvideo.FRectItemId = itemid
oitemvideo.FRectItemVideoGubun = "video1"
oitemvideo.GetItemContentsVideo

dim oitemAddImage
set oitemAddImage = new CItemAddImage
oitemAddImage.FRectItemID = itemid

if oitem.FResultCount>0 then
    ''��ǰ �߰��̹��� ����.
    oitemAddImage.GetOneItemAddImageList
end if

''������ǰ ��� ����
dim strItemRelation
strItemRelation = GetItemRelationStr(itemid)

'==============================================================================
''��ü �⺻��� ���� 
dim defaultmargin, defaultmaeipdiv, defaultFreeBeasongLimit, defaultDeliverPay, defaultDeliveryType
dim jungsangubun, companyno
dim sqlStr
sqlStr = "select c.defaultmargine, c.maeipdiv as defaultmaeipdiv, "
sqlStr = sqlStr + " IsNULL(c.defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit,"
sqlStr = sqlStr + " IsNULL(c.defaultDeliverPay,0) as defaultDeliverPay,"
sqlStr = sqlStr + " IsNULL(c.defaultDeliveryType,'') as defaultDeliveryType"
sqlStr = sqlStr + "  , p.jungsan_gubun, p.company_no "
sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c as c "
sqlStr = sqlStr + "  inner join db_partner.dbo.tbl_partner as p on c.userid = p.id " 
sqlStr = sqlStr + " where c.userid='" & oitem.FOneItem.Fmakerid & "'" 
rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        defaultmargin           = rsget("defaultmargine")
        defaultmaeipdiv         = rsget("defaultmaeipdiv")
        defaultFreeBeasongLimit = rsget("defaultFreeBeasongLimit")
        defaultDeliverPay       = rsget("defaultDeliverPay")
        defaultDeliveryType     = rsget("defaultDeliveryType")
        jungsangubun						= rsget("jungsan_gubun")
        companyno							= rsget("company_no")
    end if
rsget.close

'==============================================================================
'���ϸ���
dim sailmargine, orgmargine, margine

''����
if oitem.FOneItem.Fsailprice<>0 then
	sailmargine = 100-CCur(oitem.FOneItem.Fsailsuplycash/oitem.FOneItem.Fsailprice*100*100*100*100)/100/100/100
else
	sailmargine = 0
end if

if oitem.FOneItem.Forgprice<>0 then
	orgmargine = 100-CCur(oitem.FOneItem.Forgsuplycash/oitem.FOneItem.Forgprice*100*100*100*100)/100/100/100
else
	orgmargine = 0
end if

if oitem.FOneItem.Fsellcash<>0 then
	margine = 100-CCur(oitem.FOneItem.Fbuycash/oitem.FOneItem.Fsellcash*100*100*100*100)/100/100/100     ''''*100*100 / 100/100 �߰�
else
	margine = 0
end if


'mw ���氡�� �������� üũ
chkMWAuth = False
IF (Not oitem.FOneItem.FisCurrStockExists)  or C_ADMIN_AUTH  THEN chkMWAuth = True 

'// ��Ż ��ǰ�� �ϴ� �׽�Ʈ�� Ư�� ������ ������
If C_ADMIN_AUTH Then
	rentalItemFlag = true
Else
	rentalItemFlag = true
End If
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
<!-- #include file="./itemmodify_javascript.asp"-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>��ǰ����</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br><b>��ϵ� ��ǰ�� �����մϴ�.</b>
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>

<form name="itemreg" method="post" action="<%= ItemUploadUrl %>/linkweb/items/itemmodifyWithImage_process.asp" onsubmit="return false;" enctype="multipart/form-data" style="margin:0;">
<input type="hidden" name="itemid" value="<%= oitem.FOneItem.Fitemid %>">
<input type="hidden" name="designerid" value="<%= oitem.FOneItem.Fmakerid %>">
<input type="hidden" name="orgprice" value="<%= oitem.FOneItem.Forgprice %>">
<input type="hidden" name="orgsuplycash" value="<%= oitem.FOneItem.Forgsuplycash %>">
<!-- ��ü �⺻ ��� ���� -->
<input type="hidden" name="defaultmargin" value="<%= defaultmargin %>">
<input type="hidden" name="defaultmaeipdiv" value="<%= defaultmaeipdiv %>">
<input type="hidden" name="defaultFreeBeasongLimit" value="<%= defaultFreeBeasongLimit %>">
<input type="hidden" name="defaultDeliverPay" value="<%= defaultDeliverPay %>">
<input type="hidden" name="defaultDeliveryType" value="<%= defaultDeliveryType %>">
<input type="hidden" name="jungsangubun" value="<%=jungsangubun%>">
<input type="hidden" name="companyno" value="<%=companyno%>">
<input type="hidden" name="sellreservedate" value="<%=oitem.FOneItem.Fsellreservedate%>"><!--���¿�����-->
<input type="hidden" name="chkModSR" value="N"><!--���¿��� ��ҿ���-->
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>

</table>
<!-- ǥ ��ܹ� ��-->

<!-- 1.�Ϲ����� --> 
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>1.�Ϲ�����</strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�귣��ID :</td>
	<!--td bgcolor="#FFFFFF" colspan="3"><% 'SelectBoxDesignerItem oitem.FOneItem.Fmakerid %> (����ü�� ǥ�õ˴ϴ�)</td-->
	<td bgcolor="#FFFFFF" colspan="3"><% NewDrawSelectBoxDesignerChangeMargin "designer", oitem.FOneItem.Fmakerid, "marginData", "TnDesignerNMargineAppl" %></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemname" maxlength="64" size="50" class="text" id="[on,off,off,off][��ǰ��]" value="<%= Replace(oitem.FOneItem.Fitemname,"""","&quot;") %>">&nbsp;
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ǰ�� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemnameEng" maxlength="64" size="60" class="text_ro" readonly id="[off,off,off,off][������ǰ��]" value="<%= Replace(oitem.FOneItem.FitemnameEng,"""","&quot;") %>">&nbsp;
		<input type="button" value="�ٱ��� ���� <%=chkIIF(oitem.FOneItem.FitemnameEng="" or isnull(oitem.FOneItem.FitemnameEng),"���","����")%>" class="button" onclick="popMultiLangEdit(<%= oitem.FOneItem.Fitemid %>)" />
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰī�� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="designercomment" size="60" maxlength="128" class="text" id="[off,off,off,off][��ǰī��]" value="<%= Replace(oitem.FOneItem.Fdesignercomment,"""","&quot;") %>">
	</td>
</tr>
</table>

<!-- 2.���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left" >
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>2.����</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="���/���� ���� ���� ī�װ�" style="cursor:help;">���� ī�װ� :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<table class=a>
		<tr>
			<td><%=getCategoryInfo(oitem.FOneItem.Fitemid)%></td>
			<td valign="bottom"><input type="button" value="�߰�" class="button" onClick="popCateSelect('<%=oitem.FOneItem.Fitemid%>')"></td>
		</tr>
		</table>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="����Ʈ�� ������ ī�װ�" style="cursor:help;">���� ī�װ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td><%=getDispCategory(oitem.FOneItem.Fitemid)%></td>
			<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
		</tr>
		</table>
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
	<td bgcolor="#FFFFFF" >
		<label><input type="radio" name="itemdiv" value="01" <%=chkIIF(oitem.FOneItem.Fitemdiv="01","checked","")%> onClick="this.form.requireMakeDay.value=0;document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">�Ϲݻ�ǰ</label>
		<br>
		<label><input type="radio" name="itemdiv" value="<%= oitem.FOneItem.Fitemdiv %>" <%=chkIIF(oitem.FOneItem.Fitemdiv="06" or oitem.FOneItem.Fitemdiv="16","checked","")%> onClick="document.getElementById('lyRequre').style.display='block';checkItemDiv(this);">�ֹ� ���ۻ�ǰ</label>
		<input type="checkbox" name="reqMsg" value="10" <%=chkIIF(oitem.FOneItem.Fitemdiv="06","checked","")%> <%=chkIIF(oitem.FOneItem.Fitemdiv="06" or oitem.FOneItem.Fitemdiv="16","","disabled")%> onClick="checkItemDiv(this);">�ֹ����� ���� �ʿ�<font color=red>(�ֹ��� �̴ϼȵ� ���۹����� �ʿ��Ѱ�� üũ)</font>
		<br>
		<label><input type="radio" name="itemdiv" value="08" <%=chkIIF(oitem.FOneItem.Fitemdiv="08","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">Ƽ�ϻ�ǰ</label>
		<label><input type="radio" name="itemdiv" value="09" <%=chkIIF(oitem.FOneItem.Fitemdiv="09","checked","")%> >Present��ǰ</label>
		<label><input type="radio" name="itemdiv" value="11" <%=chkIIF(oitem.FOneItem.Fitemdiv="11","checked","")%> >��ǰ�ǻ�ǰ</label>
		<label><input type="radio" name="itemdiv" value="18" <%=chkIIF(oitem.FOneItem.Fitemdiv="18","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">�����ǰ</label>

		<% if oitem.FOneItem.Fitemdiv ="07" then %> <!-- 2014������ �ܵ����� ��ǰ > reserveItemTp=1 / ����� ��������(ȸ���� ���� ����) -->
			<label><input type="radio" name="itemdiv" value="07" <%=chkIIF(oitem.FOneItem.Fitemdiv="07","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">�������ѻ�ǰ</label>
		<% end if %>
		<% if oitem.FOneItem.Fitemdiv ="82" then %>
			<label><input type="radio" name="itemdiv" value="82" <%=chkIIF(oitem.FOneItem.Fitemdiv="82","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">���ϸ����� ��ǰ</label>
		<% end if %>

		<label><input type="radio" name="itemdiv" value="75" <%=chkIIF(oitem.FOneItem.Fitemdiv="75","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">���ⱸ����ǰ</label>

		<% If rentalItemFlag Then %>
			<label><input type="radio" name="itemdiv" value="30" <%=chkIIF(oitem.FOneItem.Fitemdiv="30","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">��Ż��ǰ</label>
		<% End If %>
		<label><input type="radio" name="itemdiv" value="23" <%=chkIIF(oitem.FOneItem.Fitemdiv="23","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">B2B��ǰ</label>
		<label><input type="radio" name="itemdiv" value="17" <%=chkIIF(oitem.FOneItem.Fitemdiv="17","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">�����������ǰ</label>
	</td>
	<td bgcolor="#FFFFFF">
	    <div id="lyRequre" style="<%=chkIIF((oitem.FOneItem.Fitemdiv ="06") or (oitem.FOneItem.Fitemdiv ="16"),"","display:none;")%>padding-left:22px;">
		�������ۼҿ��� <input type="text" name="requireMakeDay" value="<%=oitem.FOneItem.FrequireMakeDay%>" size="2" class="text" id="[off,on,off,off][�������ۼҿ���]">��
		<font color="red">(��ǰ�߼��� ��ǰ���� �Ⱓ)</font>
		</div>
	</td>
</tr>
<% if (oitem.FOneItem.IsReserveOnlyItem) then %>
<!-- ������ �ý����� only 2012/03/26 �߰�-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ܵ�(����)���� :</td>
	<td bgcolor="#FFFFFF" colspan="2">
	    <label><input type="radio" name="reserveItemTp" value="0" <%=chkIIF(oitem.FOneItem.FreserveItemTp="0" And oitem.FOneItem.Fitemdiv <>"30","checked","")%>>�Ϲ�</label>
		<label><input type="radio" name="reserveItemTp" value="1" <%=chkIIF(oitem.FOneItem.FreserveItemTp="1" or oitem.FOneItem.Fitemdiv="30","checked","")%>>�ܵ�(����)���Ż�ǰ</label>
	</td>
</tr>
<% end if %>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ٹ����� �������� :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<label><input type="radio" name="tenOnlyYn" value="Y" <%=chkIIF(oitem.FOneItem.FtenOnlyYn="Y","checked","")%>>������ǰ</label>
		<label><input type="radio" name="tenOnlyYn" value="N" <%=chkIIF(oitem.FOneItem.FtenOnlyYn="N","checked","")%>>�Ϲݻ�ǰ</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ ���� ��ǰ :</td>
	<td width="85%" bgcolor="#FFFFFF" colspan="2">
	<input type="hidden" name="availPayType" value="<%= oitem.FOneItem.FavailPayType %>">
	<% if (oitem.FOneItem.FavailPayType = "9") then %>
		������
	<% elseif (oitem.FOneItem.FavailPayType = "8") then %>
		����Ʈ������
	<% elseif (oitem.FOneItem.FavailPayType = "0") then %>
		�Ϲ�
	<% else %>
		<%= oitem.FOneItem.FavailPayType %>
	<% end if %>
	</td>
</tr>
</table>

<!-- 3.�������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>3.��������</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">���ݼ��� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
		<tr align="center">
			<td height="25" width="90" bgcolor="#DDDDFF">����</td>
			<td width="100" bgcolor="#DDDDFF">�Һ��ڰ�</td>
			<td width="100" bgcolor="#DDDDFF">���ް�</td>
			<td width="100" bgcolor="#DDDDFF">����</td>
			<td bgcolor="#DDDDFF">&nbsp;</td>
		</tr>
		<tr>
			<td height="25" bgcolor="#FFFFFF"><label><input type="radio" name="sailyn" onClick="TnCheckSailYN(itemreg)" value="N" <% if oitem.FOneItem.Fsailyn = "N" then response.write "checked" %>> ���󰡰�</label></td>
			<td bgcolor="#FFFFFF" align="center">
			<% if oitem.FOneItem.Fsailyn = "N" then %>
				<input type="text" name="sellcash" maxlength="16" size="8" class="text" id="[on,on,off,off][�Һ��ڰ�]" value="<%= oitem.FOneItem.Fsellcash %>" onkeyup="CalcuAuto(itemreg);">��
			<% else %>
				<input type="text" name="sellcash" maxlength="16" size="8" class="text" id="[on,on,off,off][�Һ��ڰ�]" value="<%= oitem.FOneItem.Forgprice %>" onkeyup="CalcuAuto(itemreg);">��
			<% end if %>
			</td>
			<td bgcolor="#FFFFFF" align="center">
			<% if oitem.FOneItem.Fsailyn = "N" then %>
				<input type="text" name="buycash" maxlength="16" size="8" class="text" id="[on,on,off,off][���ް�]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Fbuycash %>">��
			<% else %>
				<input type="text" name="buycash" maxlength="16" size="8" class="text" id="[on,on,off,off][���ް�]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Forgsuplycash %>">��
			<% end if %>
			</td>
			<% if oitem.FOneItem.Fsailyn = "N" then %>
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="margin" maxlength="32" size="5" class="text" id="[on,off,off,off][����]" value="<%= margine %>">%
			</td>
			<% else %>
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="margin" maxlength="32" size="5" class="text" id="[on,off,off,off][����]" value="<%= orgmargine %>">%
			</td>
			<% end if %>
			<td bgcolor="#FFFFFF" align="left">
				<input type="button" value="���ް� �ڵ����" class="button" onclick="CalcuAuto(itemreg);">
			</td>
		</tr>
		<tr>
			<td height="25" bgcolor="#FFFFFF"><label><input type="radio" name="sailyn" onClick="TnCheckSailYN(itemreg)" value="Y" <% if oitem.FOneItem.Fsailyn = "Y" then response.write "checked" %>> ���ΰ���</label></td>
			<input type="hidden" name="sailpricevat">
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="sailprice" maxlength="16" size="8" class="text" id="[on,on,off,off][���μҺ��ڰ�]" value="<%= oitem.FOneItem.Fsailprice %>"  onkeyup="CalcuAuto(itemreg);">��
			</td>
			<input type="hidden" name="sailsuplycashvat">
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="sailsuplycash" maxlength="16" size="8" class="text" id="[on,on,off,off][���ΰ��ް�]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Fsailsuplycash %>">��
			</td>
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="sailmargin" maxlength="32" size="5" class="text" id="[on,off,off,off][���θ���]" value="<%= sailmargine %>">%
			</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="button" value="���ް� �ڵ����" class="button" onclick="CalcuAuto(itemreg);">
				<%
					dim itemSalePer : itemSalePer=0
					if oitem.FOneItem.Fsailyn="Y" then
						itemSalePer = oitem.FOneItem.Forgprice - oitem.FOneItem.Fsailprice
						itemSalePer = itemSalePer/oitem.FOneItem.Forgprice*100
					end if
				%>
				<span id="lyrPct"><% if itemSalePer>0 then %>������: <font color="#EE0000"><strong><%=formatNumber(itemSalePer,1)%>%</strong></font><% end if %></span>
			</td>
		</tr>
		</table>
		<br>
		- ���ް��� <b>�ΰ��� ���԰�</b>�Դϴ�.<br>
		- �Һ��ڰ�(���ΰ�)�� ����(���θ���)�� �Է��ϰ� [���ް��ڵ����] ��ư�� ������ ���ް��� ���ϸ����� �ڵ����˴ϴ�.
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">���ϸ��� :</td>
	<td width="35%" bgcolor="#FFFFFF"><input type="text" name="mileage" maxlength="32" size="10" class="text" id="[on,on,off,off][���ϸ���]" value="<%= oitem.FOneItem.Fmileage %>">point</td>
	<td width="15%" bgcolor="#DDDDFF">����, �鼼 ���� :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="vatinclude" value="Y" <% if oitem.FOneItem.Fvatinclude = "Y" then response.write "checked" %>>����</label>
		<label><input type="radio" name="vatinclude" value="N" <% if oitem.FOneItem.Fvatinclude = "N" then response.write "checked" %>>�鼼</label>
	</td>
</tr>
</table>

<!-- 4.�������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>4.��������</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�ڵ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<%= oitem.FOneItem.Fitemid %>
		&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="�̸�����" class="button" onclick="window.open('http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FOneItem.Fitemid %>');">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="��ǰ �� �Ӽ�" style="cursor:help;">��ǰ�Ӽ� :</td>
	<td id="lyrItemAttribAdd" bgcolor="#FFFFFF" colspan="3"></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ü��ǰ�ڵ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="upchemanagecode" class="text" id="[off,off,off,off][��ü��ǰ�ڵ�]" value="<%= oitem.FOneItem.Fupchemanagecode %>" size="20" maxlength="32">
		(��ü���� �����ϴ� �ڵ� �ִ� 32�� - ����/���ڸ� ����)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">ISBN :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		ISBN 13 <input type="text" name="isbn13" class="text" value="<%= oitem.FOneItem.Fisbn13 %>" size="13" maxlength="13">
		/ �ΰ���ȣ <input type="text" name="isbn_sub" class="text" value="<%= oitem.FOneItem.FisbnSub %>" size="5" maxlength="5"><br />
		ISBN 10 <input type="text" name="isbn10" class="text" value="<%= oitem.FOneItem.Fisbn10 %>" size="10" maxlength="10"> (Optional)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ǰ��� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	    <input type="text" name="relateItems" value="<%=strItemRelation%>" size="40" class="text" id="[off,off,off,off][������ǰ]">
	    (������ǰ�� �ִ� 6������ ��ϰ���, ��ǰ��ȣ�� �޸�(,)�� �����Ͽ� �Է�)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�Ǹſ��� :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="sellyn" value="Y" <% if oitem.FOneItem.Fsellyn = "Y" then response.write "checked" %>>�Ǹ���</label>&nbsp;&nbsp;
		<label><input type="radio" name="sellyn" value="S" <% if oitem.FOneItem.Fsellyn = "S" then response.write "checked" %>>�Ͻ�ǰ��</label>&nbsp;&nbsp;
		<label><input type="radio" name="sellyn" value="N" <% if oitem.FOneItem.Fsellyn = "N" then response.write "checked" %>>�Ǹž���</label> 
	<%IF (oitem.FOneItem.Fsellreservedate)<> "" THEN %><font color="blue">[���¿���: <%=oitem.FOneItem.Fsellreservedate%>]</font><%END IF%>
	</td>
	<td width="15%" bgcolor="#DDDDFF">��뿩�� :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="isusing" value="Y" onclick="TnChkIsUsing(this.form)" <%=chkIIF(oitem.FOneItem.Fisusing="Y","checked","")%>>�����</label>&nbsp;&nbsp;
		<label><input type="radio" name="isusing" value="N" onclick="TnChkIsUsing(this.form)" <%=chkIIF(oitem.FOneItem.Fisusing="N","checked","")%>>������</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ����� :</td>
	<td bgcolor="#FFFFFF" colspan="3"><%= oitem.FOneItem.FRegDate %></td>
</tr>
</table>

<!-- 5.�⺻���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>5.�⺻����</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="makername" maxlength="32" size="25" class="text" id="[on,off,off,off][������]" value="<%= oitem.FOneItem.Fmakername %>">&nbsp;(������ü��)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		 <p> 
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="0" <%if isNull(oitem.FOneItem.Fsourcekind) or oitem.FOneItem.Fsourcekind="0" then%>checked<%end if%> onClick="jsSetArea(this.value);"> ��ǰ ��</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="1" <%if oitem.FOneItem.Fsourcekind="1" then%>checked<%end if%> onClick="jsSetArea(this.value);"> ����깰</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="2" <%if oitem.FOneItem.Fsourcekind="2" then%>checked<%end if%> onClick="jsSetArea(this.value);"> ���깰</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="3" <%if oitem.FOneItem.Fsourcekind="3" then%>checked<%end if%> onClick="jsSetArea(this.value);"> ��깰</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="4" <%if oitem.FOneItem.Fsourcekind="4" then%>checked<%end if%> onClick="jsSetArea(this.value);"> ����갡��ǰ</span>
	  </p>
	  <p><input type="text" name="sourcearea" maxlength="64" size="64" class="text" id="[on,off,off,off][������]"  value="<%= oitem.FOneItem.Fsourcearea %>"/></p>
	  <div id="dvArea0" style="display:<%if isNull(oitem.FOneItem.Fsourcekind) or oitem.FOneItem.Fsourcekind="0" then%>block<%else%>none<%end if%>;">
	  <p><strong>ex: �ѱ�, �߱�, �߱�OEM, �Ϻ� �� </strong></BR>
	   - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	  <div id="dvArea1" style="display:<%if oitem.FOneItem.Fsourcekind ="1" then%>block<%else%>none<%end if%>;">
	  <p><strong>������ :</strong> ����, ������ �Ǵ� �á�����, �á�����(���ѹα�, �ѱ�X)  <span style="margin-right:10px;">ex. ��(����)</span></BR>
	   <strong>���Ի� :</strong> ������� ���Ա����� <span style="margin-right:10px;">ex. ����(�߱���)</span></BR>
	   - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	  <div id="dvArea2" style="display:<%if oitem.FOneItem.Fsourcekind ="2" then%>block<%else%>none<%end if%>;">
	  <p><strong>������ :</strong> ����,������ �Ǵ� �����ػ�(��� ���깰�� �á����� ����)   <span style="margin-right:10px;">ex. ��ġ(����), ��¡��(�����ػ�)</span> </BR>
	  	<strong>����� :</strong> ����� �Ǵ� �����(�ؿ���)   <span style="margin-right:10px;">ex. ��ġ[�����(�뼭��)]</span> </BR>
	    <strong>���Ի� :</strong> ������� ���Ա����� <span style="margin-right:10px;">ex. ���(�߱���)</span></BR>
	   - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	  <div id="dvArea3" style="display:<%if oitem.FOneItem.Fsourcekind ="3" then%>block<%else%>none<%end if%>;">
	  <p>�Ұ���� ��� ������ ����(�ѿ�/����/���ұ���) �� ������   <span style="margin-right:10px;">ex. ����(Ⱦ���� �ѿ�), ����(ȣ�ֻ�)</span></BR>
	  - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	  <div id="dvArea4" style="display:<%if oitem.FOneItem.Fsourcekind ="4" then%>block<%else%>none<%end if%>;">
	  <p><strong>98%�̻� ���ᰡ �ִ� ���:</strong>  �Ѱ��� ���Ḹ ǥ�� ����    <span style="margin-right:10px;">ex. ����(�̱���)</span> </BR>
	  	<strong>���� ���Ḧ ����� ���:</strong> ȥ�պ����� ���� ������ 2�� ����   <span style="margin-right:10px;">ex. ������[�а���(�̱���),���尡��(������)]</span></BR>
	  - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div> 
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemWeight" maxlength="12" size="8" class="text" id="[on,off,off,off][��ǰ����]" style="text-align:right" value="<%= oitem.FOneItem.FitemWeight %>">g &nbsp;(�׷������� �Է�, ex:1.5kg�� 1500) / �ؿܹ�۽� ��ۺ� ������ ���� ���̹Ƿ� ��Ȯ�� �Է�.
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�˻�Ű���� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="keywords" maxlength="250" size="50" class="text" id="[on,off,off,off][�˻�Ű����]" value="<%= oitem.FOneItem.Fkeywords %>">&nbsp;(�޸��α��� ex: Ŀ��,Ƽ����,����)
	</td>
</tr>
</table>

<!-- 5-1.ǰ������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-left:5px;"><strong>- ǰ������� </strong> &nbsp;<font color=gray>��ǰ����������� ���� ���� ������ ���� �Ʒ� ������ ��Ȯ�� �Է����ֽñ� �ٶ��ϴ�.</font></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">ǰ���� :</td>
	<td bgcolor="#FFFFFF">
		<% DrawInfoDiv "infoDiv", oitem.FOneItem.FinfoDiv, " onchange='chgInfoDiv(this.value);'" %>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:<%=chkIIF(isNull(oitem.FOneItem.FinfoDiv) or oitem.FOneItem.FinfoDiv="","none","")%>;">
	<td height="30" width="15%" bgcolor="#F8DDFF">ǰ�񳻿� :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList">
	<%
		if Not(isNull(oitem.FOneItem.FinfoDiv) or oitem.FOneItem.FinfoDiv="") then
			Server.Execute("act_itemInfoDivForm.asp")
		end if
	%>
	</td>
</tr>
<tr align="left">
	<td height="25" colspan="2" bgcolor="#FDFDFD"><font color="darkred">��ǰ���������� ������ ���� �Ǿ��ִ��� ��Ȯ�� �Է¹ٶ��ϴ�. ����Ȯ�ϰų� �߸��� ���� �Է½�, �׿� ���� å���� ���� ���� �ֽ��ϴ�.</font></td>
</tr>
<tr align="left" id="lyItemSrc" style="display:<%=chkIIF(oitem.FOneItem.FinfoDiv="35","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemsource" maxlength="64" size="50" class="text" value="<%= oitem.FOneItem.Fitemsource %>">&nbsp;(ex:�ö�ƽ,����,��,...)
	</td>
</tr>
<tr align="left" id="lyItemSize" style="display:<%=chkIIF(oitem.FOneItem.FinfoDiv="35","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ������ :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemsize" maxlength="64" size="50" class="text" value="<%= oitem.FOneItem.Fitemsize %>">&nbsp;(ex:7.5x15(cm))
	</td>
</tr>
</table>
<!-- 5-2.������������ -->
<%
dim arrAuth, r, real_safetydiv, real_safetynum, safetyDivList
arrAuth = oitem.FAuthInfo
if isArray(arrAuth) THEN
	For r =0 To UBound(arrAuth,2)
		real_safetydiv = real_safetydiv & arrAuth(0,r)
		if r <> UBound(arrAuth,2) then real_safetydiv = real_safetydiv & "," end if
		
		real_safetynum = real_safetynum & arrAuth(1,r)
		if r <> UBound(arrAuth,2) then real_safetynum = real_safetynum & "," end if
		
		safetyDivList = safetyDivList & "<p class='tPad05' id='l"&arrAuth(0,r)&"'>"
		safetyDivList = safetyDivList & "- "&fnSafetyDivCodeName(arrAuth(0,r))&"("&CHKIIF(arrAuth(1,r)="x","������ȣ ����",arrAuth(1,r))&")"
		safetyDivList = safetyDivList & " <input type='button' value='����' class='btn3 btnIntb' onClick='jsSafetyDivListDel("&arrAuth(0,r)&");'>"
		safetyDivList = safetyDivList & "</p>"
	Next
end if
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-left:5px;"><strong>- ������������</strong></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">
		����������� :
		<input type="button" value="�������� �ʼ� ǰ�� Ȯ��" onclick="jsSafetyPopup();" class="button" />
	</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
		<tr align="left" height="30">
			<td bgcolor="#FFFFFF">
				<label><input type="radio" name="safetyYn" value="Y" <%=chkIIF(oitem.FOneItem.FsafetyYn="Y","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> ���</label>
				<label><input type="radio" name="safetyYn" value="N" <%=chkIIF(oitem.FOneItem.FsafetyYn="N","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> ���ƴ�</label>
				<label><input type="radio" name="safetyYn" value="I" <%=chkIIF(oitem.FOneItem.FsafetyYn="I","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> ��ǰ���� ǥ��</label>
				<label><input type="radio" name="safetyYn" value="S" <%=chkIIF(oitem.FOneItem.FsafetyYn="S","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> ���������ؼ�</label>
				<input type="hidden" name="auth_go_catecode" id="auth_go_catecode" value="">
				<input type="hidden" name="real_safetydiv" id="real_safetydiv" value="<%=real_safetydiv%>">
				<input type="hidden" name="real_safetynum" id="real_safetynum" value="<%=real_safetynum%>">
				<input type="hidden" name="real_safetyidx" id="real_safetyidx" value="">
				<input type="hidden" name="real_safetynum_delete" id="real_safetynum_delete" value="">
				<input type="hidden" name="real_safetydiv_delete" id="real_safetydiv_delete" value="">
			</td>
		</tr>
		<tr align="left">
			<td bgcolor="#FFFFFF">
				<% drawSelectBoxSafetyDivCode "safetyDiv", "", oitem.FOneItem.FsafetyYn, "" %>
				������ȣ <input type="text" name="safetyNum" id="[off,off,off,off][�������� ������ȣ]" <%=chkIIF(oitem.FOneItem.FsafetyYn<>"Y","disabled","")%> size="35" maxlength="25" value="" /><%'=oitem.FOneItem.FsafetyNum%>
				<input type="button" id="safetybtn" value="��   ��" onclick="jsSafetyAuth();" class="button">
				<input type="hidden" name="issafetyauth" id="issafetyauth" value="">
			</td>
		</tr>
		<tr align="left">
			<td bgcolor="#FFFFFF">
				<div id="safetyDivList">
					<%=safetyDivList%>
				</div>
				<div id="safetyYnI" style="display:none;">
					<font color="blue">��ǰ ���� ǥ��(ǥ���� ��ǰ�ΰ�� ��ǰ �� �������� ������ȣ�� �𵨸�, KC ��ũ�� �� ǥ�����ּ���.)</font>
				</div>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="left">
	<td bgcolor="#FFFFFF" colspan=2>
		* ���������� �Է� �� �ϰų�, �߸��� ���������� �Է��� ��� �߰� <strong><font color='red'>��� �Ǹ����� �Ǵ� ����</font></strong> �˴ϴ�.<br>
		* <strong><font color='red'>���������ؼ�</font></strong> ����ϰ�� ������ȣ�� ������, KC��ũ�� ǥ������ �ʾƾ� �˴ϴ�.<br>
		* �Է��� ���������� ��ǰ�����������Ϳ��� ������ ������ �������� ��ȸ�Ǹ�, <strong><font color='red'>�������� ���� ������ ����� �Ұ�</font></strong>���մϴ�.<br>
		* �������� ���������� �Է��������� �ұ��ϰ� ����� �ȵɰ�쿡 "��ǰ���� ǥ��"�� ������ �����ϸ�, ��ǰ �� �������� �𵨸�� ǥ���� ��ǰ�ΰ�� ������ȣ,KC��ũ�� ǥ���ؾ� �մϴ�.<br>
		* ������������ ���� ���Ǵ� Ȩ������(<u><a href="http://safetykorea.kr" target="_blank">http://safetykorea.kr</a></u>)�� Ȯ���� �ֽñ� �ٶ��ϴ�.
	</td>
</tr>
</table>

<!-- 6.������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>6.�������</strong>
    </td>
    <td align="right">
    	<input type="button" class="button" value="����������� ����" onclick="TnAutoChkDeliver()">
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����Ư������ :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	    <% IF chkMWAuth THEN %>
		<label><input type="radio" name="mwdiv" value="M" onclick="TnCheckUpcheYN(this.form);" <% if oitem.FOneItem.Fmwdiv = "M" then response.write "checked" %>>����</label>
		<label><input type="radio" name="mwdiv" value="W" onclick="TnCheckUpcheYN(this.form);" <% if oitem.FOneItem.Fmwdiv = "W" then response.write "checked" %>>Ư��</label>
		<label><input type="radio" name="mwdiv" value="U" onclick="TnCheckUpcheYN(this.form);" <% if oitem.FOneItem.Fmwdiv = "U" then response.write "checked" %>>��ü���</label>
		&nbsp;&nbsp; - ����Ư�����п� ���� ��۱����� �޶����ϴ�. ��۱����� Ȯ�����ּ���.
		<%ELSE%> 
		<%= fnColor(oitem.FOneItem.Fmwdiv,"mw") %>
		<input type="hidden" name="mwdiv" value="<%=oitem.FOneItem.Fmwdiv%>">
		<%END IF%>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��۱��� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverytype" value="1" onclick="TnCheckUpcheDeliverYN(this.form);" <% if oitem.FOneItem.Fdeliverytype = "1" then response.write "checked" %>>�ٹ����ٹ��</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);" <% if oitem.FOneItem.Fdeliverytype = "2" then response.write "checked" %>>��ü(����)���</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="4" onclick="TnCheckUpcheDeliverYN(this.form);" <% if oitem.FOneItem.Fdeliverytype = "4" then response.write "checked" %>>�ٹ����ٹ�����</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);" <% if oitem.FOneItem.Fdeliverytype = "9" then response.write "checked" %>>��ü���ǹ��(���� ��ۺ�ΰ�)</label>
		<label><input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);" <% if oitem.FOneItem.Fdeliverytype = "7" then response.write "checked" %>>��ü���ҹ��</label>
		<% if oitem.FOneItem.Fdeliverytype = "6" then %>
		<label><input type="radio" name="deliverytype" value="6" onclick="TnCheckUpcheDeliverYN(this.form);" checked><font color="darkred">�������</font></label>
		<% end if %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��۹�� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverfixday" value="" <%=chkIIF(Trim(oitem.FOneItem.Fdeliverfixday)="" or IsNull(oitem.FOneItem.Fdeliverfixday),"checked","")%> onclick="TnCheckFixday(this.form)">�ù�(�Ϲ�)</label>&nbsp;
		<label><input type="radio" name="deliverfixday" value="X" <%=chkIIF(oitem.FOneItem.Fdeliverfixday="X","checked","")%> onclick="TnCheckFixday(this.form)">ȭ��</label>&nbsp;
		<label><input type="radio" name="deliverfixday" value="C" <%=chkIIF(oitem.FOneItem.Fdeliverfixday="C","checked","")%> onclick="TnCheckFixday(this.form)">�ö��������</label>
		<label><input type="radio" name="deliverfixday" value="G" <%=chkIIF(oitem.FOneItem.Fdeliverfixday="G","checked","")%> onclick="TnCheckFixday(this.form)">�ؿ�����</label>
		<span id="lyrFreightRng" style="display:<%=chkIIF(oitem.FOneItem.Fdeliverfixday="X","","none")%>;">
			<br />&nbsp;
			��ǰ/��ȯ �� ȭ����� ���(��) :
			�ּ� <input type="text" name="freight_min" class="text" size="6" value="<%=oitem.FOneItem.Ffreight_min%>" style="text-align:right;">�� ~
			�ִ� <input type="text" name="freight_max" class="text" size="6" value="<%=oitem.FOneItem.Ffreight_max%>" style="text-align:right;">��
		</span>
		<br>&nbsp;<font color="red">(�ö�� ��ǰ�� ��츸 �����ǹ��, ������, �ö�������� �ɼ��� ��밡���մϴ�.)</font>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverarea" value="" <%=chkIIF(Trim(oitem.FOneItem.Fdeliverarea)="" or IsNull(oitem.FOneItem.Fdeliverarea),"checked","")%>>�������</label>&nbsp;
		<label><input type="radio" name="deliverarea" value="C" <%=chkIIF(oitem.FOneItem.Fdeliverarea="C","checked","")%> >�����ǹ��</label>&nbsp;
		<label><input type="radio" name="deliverarea" value="S" <%=chkIIF(oitem.FOneItem.Fdeliverarea="S","checked","")%> >������</label>
		<label><input type="checkbox" name="deliverOverseas" value="Y" <% if oitem.FOneItem.FdeliverOverseas="Y" then response.write "checked" %> title="�ؿܹ���� ��ǰ���԰� �Է��� �ž� �Ϸ�˴ϴ�.">�ؿܹ��</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">���尡�ɿ��� :</td>
	<td width="85%" colspan="3" bgcolor="#FFFFFF">
		<%= oitem.FOneItem.Fpojangok %> <!-- �б����� ���� ���� ������ �ٸ������� popup ����. -->
	</td>
</tr>

<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">���԰����� :</td>
	<td width="85%" colspan="3" bgcolor="#FFFFFF">
		<input type="text" name="reipgodate" class="text" id="[off,off,off,off][���԰�����]" size="10" value="<%= oitem.FOneItem.FreipgoDate %>" maxlength="10">
		<a href="javascript:calendarOpen(itemreg.reipgodate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		<a href="javascript:ClearVal(itemreg.reipgodate);"><img src="/images/icon_delete2.gif" width="20" border="0" align="absmiddle"></a>
	</td>
</tr>
</table>

<!-- 7.�ɼ����� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>7.�ɼ�����</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ɼǱ��� :</td>
	<input type="hidden" name="optioncnt" value="<%= oitem.FOneItem.Foptioncnt %>">
	<td width="85%" colspan="3" bgcolor="#FFFFFF">
	<% if oitem.FOneItem.Foptioncnt < 1 then %>
		�ɼǻ�����
	<% else %>
		�ɼǻ����
	<% end if %>
	</td>
</tr>
<tr align="left">
	<td width="15%" bgcolor="#DDDDFF">�ɼǼ��� :</td>
	<td width="85%" bgcolor="#FFFFFF" colspan="3">
		- �ɼ������� �ɼ�â���� ���������մϴ�.<br>
		- �ɼ��� �߰��� ���������� ������ �Ұ����մϴ�. ��Ȯ�� �Է��ϼ���.<br>
		- ���������� �ɼ��� ���� ���, �ɼ�â���� ������ �����մϴ�.<br>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" rowspan="2">������ :</td>
	<td width="85%" colspan="3" bgcolor="#FFFFFF">
	  - ��ǰ �������� [��ǰ �÷� ����]���� �Ͻ� �� �ֽ��ϴ�.
	</td>
</tr>
</table>

<!-- 8.�������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>8.��������</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td width="15%" bgcolor="#DDDDFF">�����Ǹű��� :</td>
	<td width="85%" colspan="3" bgcolor="#FFFFFF">
		<label><input type="radio" name="limityn" value="N" onClick="TnCheckLimitYN(itemreg)" <% if oitem.FOneItem.Flimityn = "N" then response.write "checked" %>>�������Ǹ�</label>&nbsp;&nbsp;
		<label><input type="radio" name="limityn" value="Y" onClick="TnCheckLimitYN(itemreg)" <% if oitem.FOneItem.Flimityn = "Y" then response.write "checked" %>>�����Ǹ�</label>
	  <div id="dvDisp" style="display:none;" >
			&nbsp;-> �������⿩��: 
			<input type="radio" name="limitdispyn" value="Y" <%IF oitem.FOneItem.Flimitdispyn="Y"  THEN%>checked<%END IF%>>���� 
			<input type="radio" name="limitdispyn" value="N" <%IF oitem.FOneItem.Flimitdispyn="N" or oitem.FOneItem.Flimitdispyn ="" THEN%>checked<%END IF%>>�����
		</div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�������� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="limitno" maxlength="32" size="8" readonly style="background-color:#E6E6E6;" class="text" id="[off,on,off,off][��������]" value="<%= oitem.FOneItem.Flimitno %>">
		-
		<input type="text" name="limitsold" maxlength="32" size="8" readonly style="background-color:#E6E6E6;" class="text" id="[off,on,off,off][�����Ǹ�]" value="<%= oitem.FOneItem.Flimitsold %>">
		=
		<input type="text" name="limitstock" maxlength="32" size="8" readonly style="background-color:#E6E6E6;" class="text" id="[off,on,off,off][�������]" value="<%= (oitem.FOneItem.Flimitno - oitem.FOneItem.Flimitsold) %>">(��)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ּ�/�ִ� �Ǹż� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		�ּ�
		<input type="text" name="orderMinNum" maxlength="5" size="5" class="text" id="[off,on,off,off][�ּ��Ǹż�]" value="<%= oitem.FOneItem.ForderMinNum %>">
		/ �ִ�
		<input type="text" name="orderMaxNum" maxlength="5" size="5" class="text" id="[off,on,off,off][�ִ��Ǹż�]" value="<%= oitem.FOneItem.ForderMaxNum %>">
		(�� �ֹ��� �Ǹ� ���� ��)
	</td>
</tr>
</table>

<!-- 9.��ǰ���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>9.��ǰ����</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ ���� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="usinghtml" value="N" <%=chkIIF(oitem.FOneItem.Fusinghtml="N","checked","") %>>�Ϲ�TEXT</label>
		<label><input type="radio" name="usinghtml" value="H" <%=chkIIF(oitem.FOneItem.Fusinghtml="H","checked","") %>>TEXT+HTML</label>
		<label><input type="radio" name="usinghtml" value="Y" <%=chkIIF(oitem.FOneItem.Fusinghtml="Y","checked","") %>>HTML���</label>
		<br>
		<textarea name="itemcontent" rows="18" class="textarea" style="width:100%" id="[on,off,off,off][��ǰ����]"><%= oitem.FOneItem.Fitemcontent %></textarea>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ ������ :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="itemvideo" rows="5" class="textarea" cols="80" id="[off,off,off,off][�����۵�����]"><%=oitemvideo.FOneItem.FvideoFullUrl%></textarea>
	    <br>�� Youtube, Vimeo ������ ����(Youtube : �ҽ��ڵ尪 �Է�, Vimeo : �Ӻ����� �Է�)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ֹ��� ���ǻ��� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="ordercomment" rows="5" cols="90" class="textarea" id="[off,off,off,off][���ǻ���]"><%= oitem.FOneItem.Fordercomment %></textarea><br>
		<font color="red">Ư���� ��۱Ⱓ�̳� �ֹ��� Ȯ���ؾ߸� �ϴ� ����</font>�� �Է��Ͻø� ���Ҹ��̳� ȯ���� ���ϼ� �ֽ��ϴ�.
	</td>
</tr>
</table>

<!-- 10.�̹������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-bottom:5px;">
      <img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>10.�̹�������</strong>
		<br>- �ٹ����ٿ��� �̹����� ����� ��쿡�� �ʼ��׸��� �⺻�̹����� �Է��Ͻñ� �ٶ��ϴ�.
		<br>- �̹����� <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> ���� �ø��� �� �ֽ��ϴ�.
		<br>&nbsp;&nbsp;(�̹�������� <font color=red>���μ������� ������</font>�� �԰ݿ� ���� �ʰ� ������ּ���. �԰��ʰ��� ����� ���� �ʽ��ϴ�.)
		<br>- <font color=red>����޿��� Save For Web����, Optimizeüũ, ������ 80%����</font>�� ����� �� �÷��ֽñ� �ٶ��ϴ�.
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�⺻�̹��� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Fbasicimage <> "") then %>
		<div id="divimgbasic" style="display:block;">
		<img src="<%= oitem.FOneItem.Fbasicimage %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgbasic" style="display:none;"></div>
	  <% end if %>
	  <input type="file" name="imgbasic" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
	  <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgbasic,40, 1000, 1000)"> (<font color=red>�ʼ�</font>,1000X1000,<b><font color="red">jpg</font></b>)
	  <input type="hidden" name="basic">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�������̹���(�ڵ�����) :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	<% if (oitem.FOneItem.Ficon1image <> "") then %>
		<img src="<%= oitem.FOneItem.Ficon1image %>" width="200" height="200">
	<% end if %>
	<% if (oitem.FOneItem.Ficon2image <> "") then %>
		<img src="<%= oitem.FOneItem.Ficon2image %>" >
	<% end if %>
	<% if (oitem.FOneItem.Flistimage120 <> "") then %>
		<img src="<%= oitem.FOneItem.Flistimage120 %>" width="120" height="120">
	<% end if %>
	<% if (oitem.FOneItem.Flistimage <> "") then %>
		<img src="<%= oitem.FOneItem.Flistimage %>" width="100" height="100">
	<% end if %>
	<% if (oitem.FOneItem.Fsmallimage <> "") then %>
		<img src="<%= oitem.FOneItem.Fsmallimage %>" width="50" height="50">
	<% end if %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����(����)�̹��� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Fmaskimage <> "") then %>
		<div id="divimgmask" style="display:block;">
		<img src="<%= oitem.FOneItem.Fmaskimage %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgmask" style="display:none;"></div>
	  <% end if %>
	  <input type="file" name="imgmask" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
	  <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgmask,40, 1000, 1000)"> (<font color=red>�ʼ�</font>,1000X1000,<b><font color="red">jpg</font></b>)
	  <input type="hidden" name="mask">
	</td>
</tr>

<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ٹ����ٱ⺻�̹��� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Ftentenimage <> "") then %>
		<div id="divimgtenten" style="display:block;">
		<img src="<%= oitem.FOneItem.Ftentenimage %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgtenten" style="display:none;"></div>
	  <% end if %>
	  <input type="file" name="imgtenten" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
	  <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgtenten,40, 1000, 1000)"> (����,1000X1000,<b><font color="red">jpg</font></b>)
	  <input type="hidden" name="tenten">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ٹ����ٱ⺻������̹���(�ڵ�����) :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	<% if (oitem.FOneItem.Ftentenimage1000 <> "") then %>
		<img src="<%= oitem.FOneItem.Ftentenimage1000 %>" width="400" height="400" title="1000*1000�̹���">
	<% end if %>
	<% if (oitem.FOneItem.Ftentenimage600 <> "") then %>
		<img src="<%= oitem.FOneItem.Ftentenimage600 %>" width="300" height="300" title="600*600�̹���">
	<% end if %>
	<% if (oitem.FOneItem.Ftentenimage400 <> "") then %>
		<img src="<%= oitem.FOneItem.Ftentenimage400 %>" width="200" height="200" title="400*400�̹���">
	<% end if %>
	<% if (oitem.FOneItem.Ftentenimage200 <> "") then %>
		<img src="<%= oitem.FOneItem.Ftentenimage200 %>" width="150" height="150" title="200*200�̹���">
	<% end if %>
	<% if (oitem.FOneItem.Ftentenimage50 <> "") then %>
		<img src="<%= oitem.FOneItem.Ftentenimage50 %>" width="50" height="50" title="50*50�̹���">
	<% end if %>
	</td>
</tr>

<tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���1 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitemAddImage.GetImageAddByIdx(0,1) <> "") then %>
		<div id="divimgadd1" style="display:block;">
		<img src="<%=oitemAddImage.GetImageAddByIdx(0,1) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd1" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd1" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd1,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add1">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���2 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitemAddImage.GetImageAddByIdx(0,2) <> "") then %>
		<div id="divimgadd2" style="display:block;">
		<img src="<%=oitemAddImage.GetImageAddByIdx(0,2) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd2" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd2" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd2,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add2">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���3 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitemAddImage.GetImageAddByIdx(0,3) <> "") then %>
		<div id="divimgadd3" style="display:block;">
		<img src="<%=oitemAddImage.GetImageAddByIdx(0,3) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd3" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd3" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd3,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add3">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���4 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitemAddImage.GetImageAddByIdx(0,4) <> "") then %>
		<div id="divimgadd4" style="display:block;">
		<img src="<%=oitemAddImage.GetImageAddByIdx(0,4) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd4" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd4" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd4,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add4">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���5 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitemAddImage.GetImageAddByIdx(0,5) <> "") then %>
		<div id="divimgadd5" style="display:block;">
		<img src="<%=oitemAddImage.GetImageAddByIdx(0,5) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd5" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd5" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd5,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add5">
	</td>
</tr>
</table>
<%
	Dim cImg, k, vArr, j
	set cImg = new CItemAddImage
	cImg.FRectItemID = itemid
	vArr = cImg.GetAddImageListIMGTYPE1
%>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="imgIn">
	<% If isArray(vArr) Then
			If vArr(3,UBound(vArr,2)) > 0 Then
			For k = 1 To vArr(3,UBound(vArr,2))
	%>
			  <tr align="left">
			  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #<%= (k) %> :</td>
			  	<td bgcolor="#FFFFFF">
		  		<%
		  		If cImg.IsImgExist(vArr,k) Then
		    		For j = 0 To UBound(vArr,2)
		    			If CStr(vArr(3,j)) = CStr(k) AND (vArr(4,j) <> "" and isNull(vArr(4,j)) = False) Then
							Response.Write "<div id=""divaddimgname"&(k)&""" style=""display:block;""><img src=""" & webImgUrl & "/item/contentsimage/" & GetImageSubFolderByItemid(vArr(1,j)) & "/" & vArr(4,j) & """ height=""250""></div>"
							Exit For
		    			End If
		    		Next
				Else
					Response.Write "<div id=""divaddimgname"&(k)&""" style=""display:none;""></div>"
				End If
				%>
			      <input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40, <%= (k-1) %>);" class="text" size="40">
			      <input type="button" value="#<%= (k) %> �̹��������" class="button" onClick="ClearImage2(this.form.addimgname<%=CHKIIF(vArr(3,UBound(vArr,2))=1,"","["&(k-1)&"]")%>,40, 1000, 1000, <%= (k-1) %>)"> (����,800X1600, Max 800KB,jpg,gif)
			      <input type="hidden" name="addimggubun" value="<%= (k) %>">
			      <input type="hidden" name="addimgdel" value="">
			  	</td>
			  </tr>
	<%
			Next
			End IF
		Else
	%>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">PC��ǰ�����̹��� #1 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname1" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40,0);" class="text" size="40">
				<input type="button" value="#1 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[0],40, 800, 1600, 0)"> (����,800X1600, Max 800KB,jpg,gif)
				<input type="hidden" name="addimggubun" value="1">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">PC��ǰ�����̹��� #2 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname2" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40,1);" class="text" size="40">
				<input type="button" value="#2 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[1],40, 800, 1600, 1)"> (����,800X1600, Max 800KB,jpg,gif)
				<input type="hidden" name="addimggubun" value="2">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">PC��ǰ�����̹��� #3 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname3" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40,2);" class="text" size="40">
				<input type="button" value="#3 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[2],40, 800, 1600, 2)"> (����,800X1600, Max 800KB,jpg,gif)
				<input type="hidden" name="addimggubun" value="3">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
	<%
	   End IF %>
</table>
<%	set cImg = nothing %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
  	<td bgcolor="#FFFFFF" height="30">
      <input type="button" value="PC��ǰ�����̹����߰�" class="button" onClick="InsertImageUp()">
      <font color="red">* ���ε尡 �� �̹����� ����� �ȳ����� ���ΰ�ħ(CTRL + F5(��Ʈ�� F5 ��ư))�� ���ּ���.</font>
  	</td>
  </tr>
</table>

<%
	Dim cmImg, mk, vmArr, mj
	set cmImg = new CItemAddImage
	cmImg.FRectItemID = itemid
	vmArr = cmImg.GetAddImageListIMGTYPE2
%>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="MobileimgIn">
	<% If isArray(vmArr) Then
			If vmArr(3,UBound(vmArr,2)) > 0 Then
			For mk = 1 To vmArr(3,UBound(vmArr,2))
	%>

			  <tr align="left">
				<td height="30" width="15%" bgcolor="#DDDDFF">����ϻ�ǰ���̹��� #<%= (mk) %> :</td>
				<td bgcolor="#FFFFFF">
				<%
				If cmImg.IsImgExist(vmArr,mk) Then
					For mj = 0 To UBound(vmArr,2)
						If CStr(vmArr(3,mj)) = CStr(mk) AND (vmArr(4,mj) <> "" and isNull(vmArr(4,mj)) = False) Then
							Response.Write "<div id=""divaddmobileimgname"&(mk)&""" style=""display:block;""><img src=""" & webImgUrl & "/item/contentsimage/" & GetImageSubFolderByItemid(vmArr(1,mj)) & "/" & vmArr(4,mj) & """ height=""250""></div>"
							Exit For
						End If
					Next
				Else
					Response.Write "<div id=""divaddmobileimgname"&(mk)&""" style=""display:none;""></div>"
				End If
				%>
				  <input type="file" name="addmobileimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40, <%= (mk-1) %>);" class="text" size="40">
				  <input type="button" value="#<%= (mk) %> �̹��������" class="button" onClick="ClearImage3(this.form.addmoblieimgname<%=CHKIIF(vmArr(3,UBound(vmArr,2))=1,"","["&(mk-1)&"]")%>,40, 640, 1200, <%= (mk-1) %>)"> (����,400X800, Max 400KB,jpg,gif)
				  <input type="hidden" name="addmobileimggubun" value="<%= (mk) %>">
				  <input type="hidden" name="addmobileimgdel" value="">
				</td>
			  </tr>
	<%
			Next
			End IF
		Else
	%>
		  <tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">����ϻ�ǰ���̹��� #1 :</td>
			<td bgcolor="#FFFFFF">
			  <input type="file" name="addmobileimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40);" class="text" size="40">
			  <input type="button" value="#1 �̹��������" class="button" onClick="ClearImage2(this.form.addmobileimgname[0],40, 640, 1200)"> (����,640X1200, Max 400KB,jpg,gif)
				<input type="hidden" name="addmobileimggubun" value="1">
				<input type="hidden" name="addmobileimgdel" value="">
			</td>
		  </tr>
		  <tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">����ϻ�ǰ���̹��� #2 :</td>
			<td bgcolor="#FFFFFF">
			  <input type="file" name="addmobileimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40);" class="text" size="40">
			  <input type="button" value="#2 �̹��������" class="button" onClick="ClearImage2(this.form.addmobileimgname[1],40, 640, 1200)"> (����,640X1200, Max 400KB,jpg,gif)
				<input type="hidden" name="addmobileimggubun" value="2">
				<input type="hidden" name="addmobileimgdel" value="">
			</td>
		  </tr>
		  <tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">����ϻ�ǰ���̹��� #3 :</td>
			<td bgcolor="#FFFFFF">
			  <input type="file" name="addmobileimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40);" class="text" size="40">
			  <input type="button" value="#3 �̹��������" class="button" onClick="ClearImage2(this.form.addmobileimgname[2],40, 640, 1200)"> (����,640X1200, Max 400KB,jpg,gif)
				<input type="hidden" name="addmobileimggubun" value="3">
				<input type="hidden" name="addmobileimgdel" value="">
			</td>
		  </tr>
	<%
	   End IF %>
</table>
<%	set cmImg = nothing %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>�� ����� ��ǰ�� �̹����� ������ �� �������� ��ü �˴ϴ�. html�� ������� ���� �����̿��� �������� ���ε� ���ֽñ� �ٶ��ϴ�.<br>�� ����� ��ǰ�󼼿��� �̹����� �߶� �÷��ֽñ� �ٶ��ϴ�.</strong></font>
 	</td>
 </tr>
  <tr align="left">
  	<td bgcolor="#FFFFFF">
      <input type="button" value="����ϻ�ǰ���̹����߰�" class="button" onClick="InsertMobileImageUp()">
      <font color="red">* ���ε尡 �� �̹����� ����� �ȳ����� ���ΰ�ħ(CTRL + F5(��Ʈ�� F5 ��ư))�� ���ּ���.</font>
  	</td>
  </tr>
</table>


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>�� ������ ��ǰ�����̹����� ������� �ʰ� ��ǰ�����̹����� ����մϴ�. ������ ��ϵ� ��ǰ�����̹����� ����� �ϵ� �߰� ������ �����ʰ� ������ �˴ϴ�.</strong></font>
 	</td>
 </tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #1 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Fmainimage <> "") then %>
		<div id="divimgmain" style="display:block;">
		<img src="<%=oitem.FOneItem.Fmainimage %>" width="400">
		</div>
	  <% else %>
	  <div id="divimgmain" style="display:none;"></div>
	  <% end if %>
		<input type="button" value="�̹��������" class="button" onClick="oldClearImage('main', 40, 800, 1600)"> (����,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="main">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #2:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Fmainimage2 <> "") then %>
		<div id="divimgmain2" style="display:block;">
		<img src="<%=oitem.FOneItem.Fmainimage2 %>" width="400">
		</div>
	  <% else %>
	  <div id="divimgmain2" style="display:none;"></div>
	  <% end if %>
		<input type="button" value="�̹��������" class="button" onClick="oldClearImage('main2', 40, 800, 1600)"> (����,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="main2">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #3:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Fmainimage3 <> "") then %>
		<div id="divimgmain3" style="display:block;">
		<img src="<%=oitem.FOneItem.Fmainimage3 %>" width="400">
		</div>
	  <% else %>
	  <div id="divimgmain3" style="display:none;"></div>
	  <% end if %>
		<input type="button" value="�̹��������" class="button" onClick="oldClearImage('main3', 40, 800, 1600)"> (����,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="main3">
	</td>
</tr>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <!--<input type="button" value="�����ϱ�" class="button" onClick="SubmitSave()">//-->
          <input type="button" value="����ϱ�" class="button" onClick="self.close()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->
</form>

<% if application("Svr_Info")	= "Dev" then %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="600" height="600"></iframe>
<% else %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>

<p>
<script type="text/javascript">

// ����Ư������ �� ��۱��м���
TnCheckUpcheYN(itemreg);
for (var i = 0; i < itemreg.elements.length; i++) {
    if (itemreg.elements[i].name == "deliverytype") {
        if (itemreg.elements[i].value == "<%= oitem.FOneItem.Fdeliverytype %>") {
            itemreg.elements[i].checked = true;
        }
    }
}

// ����
TnSilentCheckLimitYN(itemreg);
// ����
CheckSailEnDisabled(itemreg);

itemreg.designer.readOnly = true;

	// ��������üũ. ���ȹ�
	jsSafetyCheck('<%= oitem.FOneItem.FsafetyYn %>','');
</script>

<%
set oitem = Nothing
set oitemAddImage = Nothing
Set oitemvideo = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->