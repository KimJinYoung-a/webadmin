<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  OFFSHOP ����
' History : 2009.04.07 ������ ����
'			2010.08.04 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbhelper.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/order/clsems_serviceArea.asp" -->

<%
''������ ����� ���� ����.. partner , partner_group

dim ochargeuser, ogroup ,shopid, mode ,userpass, shopname ,shopphone, shopzipcode, shopaddr1, shopaddr2
dim manname, manhp, manphone, manemail ,shopdiv,isusing,stockbasedate,shopsocno, shopceoname, vieworder
dim currencyUnit , multipleRate , i , pyeong
dim shopCountryCode, decimalPointLen, decimalPointCut, exchangeRate
	shopid = RequestCheckVar(request("shopid"),32)
	mode   = RequestCheckVar(request("mode"),32)
	userpass    = request("userpass")
	shopname    = html2db(request("shopname"))
	shopphone   = request("shopphone")
	shopzipcode = request("shopzipcode")
	shopaddr1   = html2db(request("shopaddr1"))
	shopaddr2   = html2db(request("shopaddr2"))
	manname     = html2db(request("manname"))
	manhp       = request("manhp")
	manphone    = request("manphone")
	manemail    = html2db(request("manemail"))
	shopdiv     = request("shopdiv")
	isusing     = request("isusing")
	stockbasedate = request("stockbasedate")
	shopsocno   = request("shopsocno")
	shopceoname = html2db(request("shopceoname"))
	vieworder	= request("vieworder")
	currencyUnit = request("currencyUnit")
	multipleRate = request("multipleRate")
	pyeong    = request("pyeong")
	shopCountryCode = request("shopCountryCode")
    decimalPointLen = request("decimalPointLen")
    decimalPointCut = request("decimalPointCut")
    exchangeRate    = request("exchangeRate")

dim ismobileusing, mobileshopname, mobileworkhour, mobileclosedate, mobiletel, mobileaddr, mobilebysubway, mobilebybus, mobilelatitude, mobilelongitude
	ismobileusing    	= request("ismobileusing")
	mobileshopname    	= html2db(request("mobileshopname"))
	mobileworkhour    	= html2db(request("mobileworkhour"))
	mobileclosedate    	= html2db(request("mobileclosedate"))
	mobiletel    		= html2db(request("mobiletel"))
	mobileaddr    		= html2db(request("mobileaddr"))
	mobilebysubway    	= html2db(request("mobilebysubway"))
	mobilebybus    		= html2db(request("mobilebybus"))
	mobilelatitude    	= request("mobilelatitude")
	mobilelongitude    	= request("mobilelongitude")

dim sqlStr
if (mode="edit") then
	sqlStr = "update [db_shop].[dbo].tbl_shop_user" + VbCrlf
	sqlStr = sqlStr + " set userpass='" + userpass + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopname='" + shopname + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopphone='" + shopphone + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopCountryCode='" + shopCountryCode + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopzipcode='" + shopzipcode + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopaddr1='" + shopaddr1 + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopaddr2='" + shopaddr2 + "'" + VbCrlf
	sqlStr = sqlStr + " ,manname='" + manname + "'" + VbCrlf
	sqlStr = sqlStr + " ,manhp='" + manhp + "'" + VbCrlf
	sqlStr = sqlStr + " ,manphone='" + manphone + "'" + VbCrlf
	sqlStr = sqlStr + " ,manemail='" + manemail + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopdiv='" + shopdiv + "'" + VbCrlf
	sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VbCrlf
	sqlStr = sqlStr + " ,stockbasedate='" + stockbasedate + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopsocno='" + shopsocno + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopceoname='" + shopceoname + "'" + VbCrlf
	sqlStr = sqlStr + " ,vieworder='" + vieworder + "'" + VbCrlf
	sqlStr = sqlStr + " ,currencyUnit='" + currencyUnit + "'" + VbCrlf
	sqlStr = sqlStr + " ,exchangeRate=" + exchangeRate + "" + VbCrlf
	sqlStr = sqlStr + " ,multipleRate='" + multipleRate + "'" + VbCrlf
	sqlStr = sqlStr + " ,decimalPointLen=" + decimalPointLen + "" + VbCrlf
	sqlStr = sqlStr + " ,decimalPointCut=" + decimalPointCut + "" + VbCrlf
	sqlStr = sqlStr + " ,pyeong=" + pyeong + "" + VbCrlf

	sqlStr = sqlStr + " ,ismobileusing='" + CStr(ismobileusing) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobileshopname='" + CStr(mobileshopname) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobileworkhour='" + CStr(mobileworkhour) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobileclosedate='" + CStr(mobileclosedate) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobiletel='" + CStr(mobiletel) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobileaddr='" + CStr(mobileaddr) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobilebysubway='" + CStr(mobilebysubway) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobilebybus='" + CStr(mobilebybus) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobilelatitude='" + CStr(mobilelatitude) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobilelongitude='" + CStr(mobilelongitude) + "' " + VbCrlf
	sqlStr = sqlStr + " where userid='" + shopid + "'" + VbCrlf

	rsget.Open sqlStr,dbget,1
	response.write "<script>alert('OK');opener.location.reload();self.close();</script>"
end if

set ochargeuser = new COffShopChargeUser
	ochargeuser.FRectShopID = shopid
	ochargeuser.GetOffShopList

Dim IsForeignShop : IsForeignShop=ochargeuser.FItemList(0).IsForeignShop

set ogroup = new CPartnerGroup
	ogroup.FRectGroupid = ochargeuser.FItemList(0).Fgroupid
	ogroup.GetOneGroupInfo


Dim oems : SET oems = New CEms
    oems.FRectCurrPage = 1
    oems.FRectPageSize = 200
    oems.FRectisUsing  = "Y"
    oems.GetServiceAreaList

%>

<script language='javascript'>

function CopyZip(flag,post1,post2,add,dong){
	frmedit.shopzipcode.value= post1 + "-" + post2;
	frmedit.shopaddr1.value= add;
	frmedit.shopaddr2.value= dong;
}

function popZip(flag){
	var popwin = window.open("/lib/searchzip3.asp?target=" + flag,"searchzip3","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function editShopInfo(frm){
	if (frm.userpass.value.length<4){
		alert('�н������ 4�� �̻��Դϴ�.');
		return;
	}
<% if (Not IsForeignShop) then %>
    if (frm.groupid.value.length<1){
		alert('����������� ����ϼ���.');
		return;
	}
<% end if %>

	if (frm.shopname.value.length<1){
		alert('�� �̸��� �Է��ϼ���.');
		return;
	}
<% if (Not IsForeignShop) then %>
    if (frm.shopzipcode.value.length!=7){
        alert('�����ȣ�� �Է��ϼ���.');
		return;
    }
 <% end if %>

 	if (frm.ismobileusing[0].checked == true) {
 		// ����� ǥ������

	    if (frm.mobileshopname.value.length<1){
			alert('����ϼ����� �Է��ϼ���.');
			frm.mobileshopname.focus();
			return;
		}

	    if (frm.mobileworkhour.value.length<1){
			alert('�����ð��� �Է��ϼ���.');
			frm.mobileworkhour.focus();
			return;
		}

	    if (frm.mobileclosedate.value.length<1){
			alert('�������� �Է��ϼ���.');
			frm.mobileclosedate.focus();
			return;
		}

	    if (frm.mobiletel.value.length<1){
			alert('��ǥ��ȭ�� �Է��ϼ���.');
			frm.mobiletel.focus();
			return;
		}

	    if (frm.mobileaddr.value.length<1){
			alert('������ּҸ� �Է��ϼ���.');
			frm.mobileaddr.focus();
			return;
		}
 	}

    if (frm.mobilelatitude.value.length<1){
		frm.mobilelatitude.value.length = 0.0;
	} else {
		if (frm.mobilelatitude.value.length*0 != 0) {
			alert('������ ���ڸ� �Է°����մϴ�.');
			frm.mobilelatitude.focus();
			return;
		}
	}

    if (frm.mobilelongitude.value.length<1){
		frm.mobilelongitude.value.length = 0.0;
	} else {
		if (frm.mobilelongitude.value.length*0 != 0) {
			alert('������ ���ڸ� �Է°����մϴ�.');
			frm.mobilelongitude.focus();
			return;
		}
	}

	var ret = confirm('�����Ͻðڽ��ϱ�?');
	if (ret){
		frm.submit();
	}
}

function emsBoxChange(obj) {
	var shopCountryCode = obj.value;

	if (shopCountryCode == "") {
		return;
	}

	if (shopCountryCode == "KR") {
		frmedit.btnsearchzipcode.disabled = false;
		return;
	} else {
		frmedit.btnsearchzipcode.disabled = true;
		return;
	}
}

function clearZipcode() {
	frmedit.shopzipcode.value = "";
	frmedit.shopaddr1.value = "";
}

function popUploadShopimage(frm) {
	var mode, imagekind, pk;

	if (frm.mobileshopimage.value == "") {
		mode = "addimage";
	} else {
		mode = "editimage";
	}

	imagekind = "mobileshopimage";
	pk = frm.shopid.value;


	var popwin = window.open("/common/pop_upload_image.asp?mode=" + mode + "&imagekind=" + imagekind + "&pk=" + pk + "&50X50=Y","popUploadShopimage","width=390 height=120 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popUploadShopmap(frm) {
	var mode, imagekind, pk;

	if (frm.mobilemapimage.value == "") {
		mode = "addimage";
	} else {
		mode = "editimage";
	}

	imagekind = "mobilemapimage";
	pk = frm.shopid.value;


	var popwin = window.open("/common/pop_upload_image.asp?mode=" + mode + "&imagekind=" + imagekind + "&pk=" + pk,"popUploadShopmap","width=390 height=120 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmedit" method="post" action="">
	<input type="hidden" name="shopid" value="<%= shopid %>">
	<input type="hidden" name="mode" value="edit">

	<% if ochargeuser.FresultCount >0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="4">
			<img src="/images/icon_star.gif" border="0" align="absbottom">
			<b>OFFSHOP ����</b>
		</td>
	</tr>
	<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="4"><b>1.���������</b></td>
	</tr>
	<tr height="25">
		<td width="120" bgcolor="<%= adminColor("tabletop") %>">��ü�ڵ�</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text_ro" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>" size="7" maxlength="5" readonly>
			<% if ogroup.FOneItem.FGroupId<>"" then %>
			<input type="button" class="button" value="�������������" onclick="PopUpcheInfoEdit('<%= ogroup.FOneItem.Fgroupid %>')">
			<% end if %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>" width="120">ȸ���(��ȣ)</td>
		<td bgcolor="#FFFFFF" width="200"><%= ogroup.FOneItem.FCompany_name %></td>
		<td bgcolor="<%= adminColor("tabletop") %>" width="120">��ǥ��</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fceoname %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����ڹ�ȣ</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_no %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_gubun %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����������</td>
		<td colspan="3" bgcolor="#FFFFFF">
			[<%= ogroup.FOneItem.Fcompany_zipcode %>]&nbsp;
			<%= ogroup.FOneItem.Fcompany_address %>&nbsp;
			<%= ogroup.FOneItem.Fcompany_address2 %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_uptae %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_upjong %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<%= ogroup.FOneItem.Fjungsan_name %>&nbsp;
			<%= ogroup.FOneItem.Fjungsan_email %>&nbsp;
			<%= ogroup.FOneItem.Fjungsan_hp %>
		</td>
	</tr>

	<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="4"><b>2.Shop����</b></td>
	</tr>
	<tr height="25">
		<td width="90" bgcolor="<%= adminColor("tabletop") %>">ShopID</td>
		<td bgcolor="#FFFFFF" width="200"><%= ochargeuser.FItemList(0).Fuserid %></td>
		<td width="90" bgcolor="<%= adminColor("tabletop") %>">Password</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="userpass" value="<%= ochargeuser.FItemList(0).Fuserpass %>" size="16" maxlength="16"></td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">Shop��</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="shopname" value="<%= ochargeuser.FItemList(0).Fshopname %>" size="20" maxlength="64"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">Shop����</td>
		<td bgcolor="#FFFFFF">
		    <% Call drawSelectBoxShopDiv("shopdiv",ochargeuser.FItemList(0).Fshopdiv) %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">Shop��ȭ��ȣ</td>
		<td colspan="3" bgcolor="#FFFFFF"><input type="text" class="text" name="shopphone" value="<%= ochargeuser.FItemList(0).Fshopphone %>" size="16" maxlength="16"></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">Shop�ּ�</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text_ro" name="shopzipcode" value="<%= ochargeuser.FItemList(0).Fshopzipcode %>" size="7" maxlength="7" <%= CHKIIF(IsForeignShop,"","ReadOnly") %> >
			<input type="button" class="button" name="btnsearchzipcode" value="�����ȣ�˻�" onclick="javascript:popZip('s');">
			<input type="button" class="button" value="����" onclick="javascript:clearZipcode();">
			 ( -�뽬 ���� �Է� )<br>
			<input type="text" class="text_ro" name="shopaddr1" value="<%= ochargeuser.FItemList(0).Fshopaddr1 %>" size="60" maxlength="64"><br>
			<input type="text" class="text" name="shopaddr2" value="<%= ochargeuser.FItemList(0).Fshopaddr2 %>" size="60" maxlength="64"><br>
				<select name="shopCountryCode" class="select" style="width:200px;height:20px;" onChange="emsBoxChange(this);">
				<option value="">��������</option>
				<option value="KR" <% if (ochargeuser.FItemList(0).FshopCountryCode = "KR") then %>selected<% end if %>>���ѹα�</option>
				<% for i=0 to oems.FREsultCount-1 %>
				<option value="<%= oems.FItemList(i).FcountryCode %>" <% if (ochargeuser.FItemList(0).FshopCountryCode = oems.FItemList(i).FcountryCode) then %>selected<% end if %>><%= oems.FItemList(i).FcountryNameKr %>(<%= oems.FItemList(i).FcountryNameEn %>)</option>
				<% next %>
				</select>
			<!--
			<input type="text" class="text_ro" name="" value="KR" size="2" maxlength="4" readonly>
			<input type="text" class="text_ro" name="" value="���ѹα�" size="16" maxlength="16" readonly>EMS���� ����
			<input type="button" class="button" value="�����ڵ�˻�" onclick=""><br>
			-->
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�Ŵ���</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manname" value="<%= ochargeuser.FItemList(0).Fmanname %>" size="16" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ŵ���Phone</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manphone" value="<%= ochargeuser.FItemList(0).Fmanphone %>" size="16" maxlength="16"></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�Ŵ���Email</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manemail" value="<%= ochargeuser.FItemList(0).Fmanemail %>" size="25" maxlength="128"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ŵ���HP</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manhp" value="<%= ochargeuser.FItemList(0).Fmanhp %>" size="16" maxlength="16"></td>
	</tr>

	<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="4"><b>3.��Ÿ����</b></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">���ͺ�</td>
		<td colspan="3" bgcolor="#FFFFFF">
			�⺻������:
			<input type="text" class="text" name="defaultmargine" value="" size="3" maxlength="3">%
			&nbsp;
			���ͺ񱸺�:
			<select class="select" name="chargegubun">
				<option value="A">����</option>
				<option value="M">����</option>
				<option value="Y">�ⳳ</option>
			</select>
			&nbsp;
			���ͺ�:
			<input type="text" class="text" name="franchizecharge" value="" size="7" maxlength="7">��
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">ȭ�����</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<%
			'//������ ������� �� �� �ھ� ����
			'if isnull(ochargeuser.FItemList(0).fcurrencyUnit) then
			%>
		 		<%' DrawexchangeRate "currencyUnit","WON","" %>
		 	<%' else %>
				<% DrawexchangeRate "currencyUnit",ochargeuser.FItemList(0).fcurrencyUnit,"" %>
				&nbsp;ȯ�� <input type="text" class="text" name="exchangeRate" value="<%= ochargeuser.FItemList(0).FexchangeRate %>" size=12 maxlength=12>
			<%' end if %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">ȭ��Ҽ���</td>
		<td colspan="3" bgcolor="#FFFFFF">
		    ǥ�� <input type="text" class="text" name="decimalPointLen" value="<%= ochargeuser.FItemList(0).FdecimalPointLen %>" size=2 maxlength=2> �ڸ�
		    ���� <input type="text" class="text" name="decimalPointCut" value="<%= ochargeuser.FItemList(0).FdecimalPointCut %>" size=2 maxlength=2> �ڸ�
		</td>
	</tr>

	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td bgcolor="<%= adminColor("tabletop") %>">�ؿܸ������</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<%
			'//���� ����� ������� 1.0 �� �ھ� ����
			'if isnull(ochargeuser.FItemList(0).fcurrencyUnit) then
		 		'response.write "<input type='text' name='multipleRate' value='1.0' size=10 maxlength=10>"
		 	'else
				response.write "<input type='text' name='multipleRate' value='"& ochargeuser.FItemList(0).fmultipleRate &"' size=10 maxlength=10>"
			'end if
			%>

			ex) �ǸŰ� x �������(1.0) = �����ǸŰ�
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">���������</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text" name="pyeong" value="<%= ochargeuser.FItemList(0).fpyeong %>" size=5 maxlength=5>
		</td>
	</tr>
	<!--
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<table border=0 cellspacing=0 cellpadding=0 class=a>
			<tr>
				<td width=80>���걸��:</td>
				<td>
					<select class="select" name="jungsangubun">
					<option value="A">����</option>
					<option value="M">����</option>
					</select>
				</td>
			</tr>
			<tr>
				<td width=80>����:</td>
				<td>
					<input type="text" class="text" name="bankname" value="" size="7" maxlength="7">(����,��ȭ,����..)
				</td>
			</tr>
			<tr>
				<td width=80>����:</td>
				<td>
					<input type="text" class="text" name="bankacct" value="" size="16" maxlength="32">(-�뽬����)
				</td>
			</tr>
			<tr>
				<td width=80>������:</td>
				<td>
					<input type="text" class="text" name="acctname" value="" size="16" maxlength="32">
				</td>
			</tr>
			</table>
		</td>
	</tr>
	-->
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text" name="stockbasedate" value="<%= ochargeuser.FItemList(0).Fstockbasedate %>" >
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��뱸��</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" <% if ochargeuser.FItemList(0).FIsUsing="Y" then response.write "checked" %> >�����
		<input type="radio" name="isusing" value="N" <% if ochargeuser.FItemList(0).FIsUsing="N" then response.write "checked" %> >������
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">ȭ��ǥ�ü���</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text" name="vieworder" value="<%= ochargeuser.FItemList(0).Fvieworder%>" size="2">	(0 �ϰ�� ȭ��ǥ�þ���.)
		</td>
	</tr>

	<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="4"><b>5.�����ǥ������</b></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�����ǥ�ÿ���</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="radio" name="ismobileusing" value="Y" <% if ochargeuser.FItemList(0).Fismobileusing="Y" then response.write "checked" %> >ǥ����
			<input type="radio" name="ismobileusing" value="N" <% if ochargeuser.FItemList(0).Fismobileusing<>"Y" then response.write "checked" %> >ǥ�þ���
		</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����ϼ���</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text" name="mobileshopname" value="<%= ochargeuser.FItemList(0).Fmobileshopname %>" size=32 maxlength=32>
		</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">���̹���(400X400)</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<% if (ochargeuser.FItemList(0).Fmobileshopimage <> "") then %>
				<img src="<%= ochargeuser.FItemList(0).GetMobileShopImage50X50 %>"><br>
				<img src="<%= ochargeuser.FItemList(0).GetMobileShopImage %>"><br>
				<input type="button" class="button" value="�����ϱ�" onclick="popUploadShopimage(frmedit)">
			<% else %>
				<input type="button" class="button" value="����ϱ�" onclick="popUploadShopimage(frmedit)">
			<% end if %>
			<input type="hidden" name="mobileshopimage" value="<%= ochargeuser.FItemList(0).Fmobileshopimage %>">
		</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�����ð�</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text" name="mobileworkhour" value="<%= ochargeuser.FItemList(0).Fmobileworkhour %>" size=50 maxlength=100>
		</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text" name="mobileclosedate" value="<%= ochargeuser.FItemList(0).Fmobileclosedate %>" size=50 maxlength=100>
		</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǥ��ȭ</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text" name="mobiletel" value="<%= ochargeuser.FItemList(0).Fmobiletel %>" size=16 maxlength=16>
		</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">������ּ�</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text" name="mobileaddr" value="<%= ochargeuser.FItemList(0).Fmobileaddr %>" size=60 maxlength=60>
		</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�൵(400X400)</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<% if (ochargeuser.FItemList(0).Fmobilemapimage <> "") then %>
				<img src="<%= ochargeuser.FItemList(0).GetMobileMapImage %>"><br>
				<input type="button" class="button" value="�����ϱ�" onclick="popUploadShopmap(frmedit)">
			<% else %>
				<input type="button" class="button" value="����ϱ�" onclick="popUploadShopmap(frmedit)">
			<% end if %>
			<input type="hidden" name="mobilemapimage" value="<%= ochargeuser.FItemList(0).Fmobilemapimage %>">
		</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">���߱�������ö</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<textarea class="textarea" cols="60" rows="4" name="mobilebysubway"><%= ochargeuser.FItemList(0).Fmobilebysubway %></textarea>
		</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">���߱������</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<textarea class="textarea" cols="60" rows="4" name="mobilebybus"><%= ochargeuser.FItemList(0).Fmobilebybus %></textarea>
		</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text" name="mobilelatitude" value="<%= ochargeuser.FItemList(0).Fmobilelatitude %>" size=16 maxlength=16>
		</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�浵</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text" name="mobilelongitude" value="<%= ochargeuser.FItemList(0).Fmobilelongitude %>" size=16 maxlength=16>
		</td>
	</tr>

	<tr height="25" bgcolor="#FFFFFF">
		<td colspan="4" align="center"><input type="button" class="button" value="��������" onclick="editShopInfo(frmedit)"></td>
	</tr>
	<% end if %>

	</form>
</table>


	<p>
	<!--
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if ochargeuser.FresultCount >0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="4">
			<img src="/images/icon_star.gif" border="0" align="absbottom">
			<b>���������</b>
		</td>
	</tr>
	<tr height="25">
		<td width="80">ShopID</td>
		<td bgcolor="#FFFFFF"><%= ochargeuser.FItemList(0).Fuserid %></td>
		<td width="80">��ü�ڵ�</td>
		<td bgcolor="#FFFFFF">G1000&nbsp;<input type="button" class="button" value="��ü��������"></td>
	</tr>
	<tr height="25">
		<td width="100">ȸ���(��ȣ)</td>
		<td bgcolor="#FFFFFF">(��)�ٹ�����</td>
		<td width="80">��ǥ��</td>
		<td bgcolor="#FFFFFF">��â��</td>
	</tr>
	<tr height="25">
		<td width="100">����ڹ�ȣ</td>
		<td bgcolor="#FFFFFF">211-87-00620</td>
		<td width="80">��������</td>
		<td bgcolor="#FFFFFF">����</td>
	</tr>
<% end if %>
</table>
-->

<%
set ochargeuser = Nothing
set ogroup = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->