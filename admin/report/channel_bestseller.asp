<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ä�� ����Ʈ����
' History : ������ ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/category_reportcls.asp"-->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdate,searchnextdate
dim orderserial,itemid,oreport
dim topn,cdl,cdm,page
dim ckpointsearch,ckipkumdiv4
dim ix,iy,cknodate
dim order_desum
dim rectdispy, rectselly, ordertype, rdsite
dim oldlist, sitename
dim sellchnl, userlevel
dim vPurchasetype, inc3pl
Dim dispCate, DlvType
Dim optExists, research

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
cdl = request("cdl")
cdm = request("cdm")
orderserial = request("orderserial")
itemid = request("itemid")
topn = request("topn")
ckpointsearch = request("ckpointsearch")
cknodate = request("cknodate")
order_desum = request("order_desum")
rectdispy = request("rectdispy")
rectselly = request("rectselly")
ordertype = request("ordertype")
if ordertype="" then ordertype="ea"
oldlist = request("oldlist")
rdsite = request("rdsite")
sitename = request("sitename")
sellchnl  = NullFillWith(request("sellchnl"),"")
dispCate = requestCheckvar(request("disp"),16)

vPurchasetype = request("purchasetype")
inc3pl = request("inc3pl")
optExists = request("optExists")
research = request("research")
userlevel = request("userlevel")
DlvType = request("dlvtype")

''�⺻���� �ɼǺ��� ����
If (research = "") Then
	optExists = ""
End If

if sitename<>"" then
	if rdsite<>"on" then
		rdsite = sitename
	else
		sitename = ""
	end if
end if

if (yyyy1="") then
	nowdate = Left(CStr(now()),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
else
	nowdate = Left(CStr(DateSerial(yyyy1 , mm1 , dd1)),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

topn = request("topn")
if (topn="") then topn=100

set oreport = new CCategoryReport

if cknodate="" then
	oreport.FRectFromDate = yyyy1 + "-" + mm1 + "-" + dd1
	oreport.FRectToDate = searchnextdate
end if

oreport.FRectCD1 = cdl
oreport.FRectCD2 = cdm
oreport.FPageSize = topn
oreport.FCurrPage = page
oreport.FRectDispY = rectdispy
oreport.FRectSellY = rectselly
oreport.FRectRdsite = rdsite
oreport.FRectOrdertype = ordertype
oreport.FRectOldJumun = oldlist
oreport.FRectSellChannelDiv = sellchnl
oreport.FRectPurchasetype = vPurchasetype ''2014/01/27
oreport.FRectInc3pl = inc3pl  ''2014/01/27
oreport.FRectDispCate = dispCate
oreport.FRectOptExists = optExists
oreport.FRectUserLevel = userlevel
oreport.FRectDlvType = DlvType
oreport.ONSearchCategoryBestseller

'// ����Ʈ���зΰ˻�����
Sub Drawsitename(selectboxname, sitename)		'�˻��ϰ����ϴ� ���� ����Ʈ �ڽ����ӿ� �ְ�, ��� �ִ� ���� �˻�._selectboxname�� sub���������� ����
	dim userquery, tem_str

	response.write "<select name='" & selectboxname & "' class='select'>"		'�˻��ϰ����ϴ� ���� ����Ʈ �������� �ϰ�
	response.write "<option value=''"							'�ɼ��� ���� ������
		if sitename ="" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">��ü</option>"								'�����̶� �ܾ ��������.
	response.write "<option value='10x10' "
		if sitename ="10x10" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">10x10</option>"

	'����� �˻� �ɼ� ���� DB���� ��������
	userquery = " select id from [db_partner].[dbo].tbl_partner"
	userquery = userquery + " where 1=1"
	userquery = userquery + " and id <> '' and id is not null"
	userquery = userquery + " and userdiv= '999' and isusing='Y' "
	userquery = userquery + " group by id"

	rsget.Open userquery, dbget, 1

	if not rsget.EOF then
		do until rsget.EOF
			if Lcase(sitename) = Lcase(rsget("id")) then 	'�˻��� �̸��� db�� ����� �̸��� ���ؼ� �´ٸ�, //
				tem_str = " selected"								'// �˻���� ����
			end if

			response.write "<option value='" & rsget("id") & "' " & tem_str & ">" & rsget("id") & "</option>"
			tem_str = ""				'rsget�� id�� �����ϰ� �˻��� ������ ����
			rsget.movenext
		loop
	end if
	rsget.close
	response.write "</select>"
End Sub
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

function ViewOrderDetail(itemid){
    var popwin = window.open("http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" + itemid,"category_prd");
    popwin.focus();
}

function ViewUserInfo(frm){
	//var popwin;
    //popwin = window.open('','userinfo');
    frm.target = 'userinfo';
    frm.action="viewuserinfo.asp"
	frm.submit();

}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function MonthDiff(d1, d2) {
	d1 = d1.split("-");
	d2 = d2.split("-");

	d1 = new Date(d1[0], d1[1] - 1, d1[2]);
	d2 = new Date(d2[0], d2[1] - 1, d2[2]);

	var d1Y = d1.getFullYear();
	var d2Y = d2.getFullYear();
	var d1M = d1.getMonth();
	var d2M = d2.getMonth();

	return (d2M+12*d2Y)-(d1M+12*d1Y);
}

function ReSearch(ifrm){
	var v = ifrm.topn.value;
	if (!IsDigit(v)){
		alert('���ڸ� �����մϴ�.');
		ifrm.topn.focus();
		return;
	}

	if (v>1000){
		alert('õ�� ���ϸ� �˻������մϴ�.');
		ifrm.topn.focus();
		return;
	}

	if ((CheckDateValid(ifrm.yyyy1.value, ifrm.mm1.value, ifrm.dd1.value) == true) && (CheckDateValid(ifrm.yyyy2.value, ifrm.mm2.value, ifrm.dd2.value) == true)) {
		if (MonthDiff(ifrm.yyyy1.value + "-" + ifrm.mm1.value + "-" + ifrm.dd1.value, ifrm.yyyy2.value + "-" + ifrm.mm2.value + "-" + ifrm.dd2.value) >= 3) {
			alert("�ִ� 3���������� �˻��� �����մϴ�.");
			return;
		}

		ifrm.submit();
	}

	//ifrm.submit();
}
</script>

	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="70" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
		<td align="left">
			<table class="a" border="0" cellpadding="3">
			<tr>
				<td class="a" >
				�Ⱓ:
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6������������
			</td>
		</tr>
		<tr>
			<td>
				����ī�װ�:
				<% SelectBoxCategoryLarge cdl %>&nbsp;
				<% if cdl="110" then DrawSelectBoxCategoryMid "cdm",cdl,cdm %>
				&nbsp;&nbsp;����ī�װ�: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->


			</td>
		</tr>
		<tr>
			<td>
			  ����Ʈ: <% Drawsitename "sitename",sitename %>
	      &nbsp;&nbsp;ä�α���:
	        <% drawSellChannelComboBoxGroup "sellchnl",sellchnl %>  
			&nbsp;&nbsp;��������: 
			<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
			&nbsp;&nbsp;<b>����ó:</b>
		    <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
		    &nbsp;&nbsp;ȸ�����:
		    <% Call DrawselectboxUserLevel("userlevel", userlevel, "") %>
			&nbsp;&nbsp;��۱���:
			<select name="dlvtype" class="select">
				<option value="">��ü</option>
				<option value="N" <%=chkIIF(DlvType="N","selected","")%>>�ٹ����� ���</option>
				<option value="Y" <%=chkIIF(DlvType="Y","selected","")%>>��ü ���</option>
			</select>
			</td>
		</tr>
		<tr>
			<td>
				<input type="checkbox" name="rectselly" <% if rectselly="on" then response.write "checked" %> >�Ǹ��ϴ¾����۸�
				<input type="checkbox" name="rectdispy" <% if rectdispy="on" then response.write "checked" %> >�����ϴ¾����۸�
				<input type="checkbox" name="rdsite" <% if rdsite="on" then response.write "checked" %> >������ǸŸ�
				&nbsp;&nbsp;����:
				<input type="radio" name="ordertype" value="ea" <% if ordertype="ea" then response.write "checked" %>>������
				<input type="radio" name="ordertype" value="totalprice" <% if ordertype="totalprice" then response.write "checked" %>>�����
				<input type="radio" name="ordertype" value="gain" <% if ordertype="gain" then response.write "checked" %>>���ͼ�
				<input type="radio" name="ordertype" value="unitCost" <% if ordertype="unitCost" then response.write "checked" %>>���ܰ���
				   &nbsp;&nbsp; �˻����� :
				<input type="text" name="topn" value="<%= topn %>" size="7" maxlength="6" >
				&nbsp;<input type="checkbox" name="optExists" <%= ChkIIF(optExists="on","checked","") %> >�ɼǺ��� ����
			</td>
		</tr>

	</table>
	</td>
	<td class="a" align="center"  bgcolor="<%= adminColor("gray") %>">
			<a href="javascript:ReSearch(frm);"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</table>
	</form>
	<br>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#E6E6E6">
	<td colspan="12" height="25" align="right">�˻���� : �� <font color="red"><% = oreport.FResultCount %></font>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr bgcolor="#E6E6E6">
	<td width="30" align="center">����</td>
	<td width="50" align="center">�̹���</td>
	<td width="50" align="center">��ǰ��ȣ</td>
	<td  align="center">��ǰ</td>
<% If optExists = "" Then %>
	<td width="100" align="center">�귣��ID</td>
<% Else %>
	<td width="50">�ܰ�</td>
	<td width="100" align="center">�귣��ID</td>
	<td width="80" align="center">�ɼ�</td>
<% End If %>
	<td width="65" align="center">�ǸŰ���</td>
	<td width="65" align="center">(����)�ǸŰ�</td>
	<td width="100" align="center">�ǸŰ���</td>
	<td width="100" align="center">���԰���</td>
	<td width="100" align="center">����</td>
	<td width="70" align="center">������</td>
</tr>
<% if oreport.FResultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="12" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% else %>
	<% for ix=0 to oreport.FResultCount -1 %>
<%
Dim totalsumprice, totalbuyprice, totalitemno
totalitemno   =  totalitemno + oreport.FItemList(ix).FItemNo
totalsumprice =  totalsumprice + oreport.FItemList(ix).Fselltotal
totalbuyprice =  totalbuyprice + oreport.FItemList(ix).Fbuytotal

%>
	<tr class="a" bgcolor="#FFFFFF" height="50">
		<td align="center"><%=ix+1%></td>
		<td><img src="<%= oreport.FItemList(ix).FImageSmall %>" width=50></td>
		<td align="center" height="25"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oreport.FItemList(ix).FItemID %>" class="zzz" target="_blank"><%= oreport.FItemList(ix).FItemID  %></a></td>
		<td align="center"><%= oreport.FItemList(ix).FItemName %></td>
	<% If optExists = "" Then %>
		<td align="center"><%= oreport.FItemList(ix).FMakerid %></td>
	<% Else %>
		<td align="center"><%= Chkiif(optExists="on", FormatNumber(oreport.FItemList(ix).FItemCost,0), "") %></td>
		<td align="center"><%= oreport.FItemList(ix).FMakerid %></td>
		<% if (oreport.FItemList(ix).FItemOptionStr="") then %>
			<td align="center">&nbsp;</td>
		<% else %>
			<td align="center"><%= oreport.FItemList(ix).FItemOptionStr %></td>
		<% end if %>
	<% End If %>
		<td align="center"><%= oreport.FItemList(ix).FItemNo %></td>
		<td align="right"><%= FormatNumber(oreport.FItemList(ix).Forgprice,0) %>
		    <%
			'���ΰ�
			if oreport.FItemList(ix).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>("&CLng((oreport.FItemList(ix).Forgprice-oreport.FItemList(ix).Fsailprice)/oreport.FItemList(ix).Forgprice*100) & "%��)" & FormatNumber(oreport.FItemList(ix).Fsailprice,0) & "</font>"
			end if
			'������
			if oreport.FItemList(ix).FitemCouponYn="Y" then
				Select Case oreport.FItemList(ix).FitemCouponType
					Case "1"
						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oreport.FItemList(ix).GetCouponAssignPrice(),0) & "</font>"
					Case "2"
						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oreport.FItemList(ix).GetCouponAssignPrice(),0) & "</font>"
				end Select

			end if%></td>
		<td align="right"><%= FormatNumber(oreport.FItemList(ix).Fselltotal,0) %></td>
		<td align="right"><%= FormatNumber(oreport.FItemList(ix).Fbuytotal,0) %></td>
		<td align="right"><%= FormatNumber(oreport.FItemList(ix).Fselltotal-oreport.FItemList(ix).Fbuytotal,0) %></td>
	    <td align="center">
	        <% if oreport.FItemList(ix).Fselltotal<>0 then %>
	        <%= 100-CLng(oreport.FItemList(ix).Fbuytotal/oreport.FItemList(ix).Fselltotal*100*100)/100 %> %
	        <% end if %>
	    </td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
	    <td colspan="2" align="center">Total</td>
	    <td colspan="<%=CHKIIF(optExists ="", "3", "5") %>"></td>
	    <td align="center"><%= FormatNumber(totalitemno,0) %></td>
	    <td>&nbsp;</td>
	    <td align="right"><%= FormatNumber(totalsumprice,0) %></td>
	    <td align="right"><%= FormatNumber(totalbuyprice,0) %></td>
	    <td align="right"><%= FormatNumber(totalsumprice-totalbuyprice,0) %></td>
	    <td align="center">
	        <% if totalsumprice<>0 then %>
	        <%= 100-CLng(totalbuyprice/totalsumprice*100*100)/100 %> %
	        <% end if %>
	    </td>
	</tr>
<% end if %>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
