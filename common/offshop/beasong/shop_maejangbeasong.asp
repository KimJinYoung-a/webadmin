<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : �������� ���
' Hieditor : 2011.02.28 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<% '<!-- #include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->
<%

dim showshopselect, loginidshopormaker

showshopselect = false
loginidshopormaker = ""

if C_ADMIN_USER then
	showshopselect = true
	loginidshopormaker = request("shopid")

elseif (C_IS_SHOP) then
	'����/������
	loginidshopormaker = C_STREETSHOPID
else
	if (C_IS_Maker_Upche) then
		loginidshopormaker = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
			loginidshopormaker = "--"		'ǥ�þ��Ѵ�. ����.
		else
			showshopselect = true
			loginidshopormaker = request("shopid")
		end if
	end if
end if

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2 , nowdate,searchnextdate,BeasongCom
dim dateback, SearchGubun ,SearchType, SearchValue ,ojumun ,ix,iy
	yyyy1   = requestCheckVar(request("yyyy1"),4)
	mm1     = requestCheckVar(request("mm1"),2)
	dd1     = requestCheckVar(request("dd1"),2)
	yyyy2   = requestCheckVar(request("yyyy2"),4)
	mm2     = requestCheckVar(request("mm2"),2)
	dd2     = requestCheckVar(request("dd2"),2)
	SearchType  = requestCheckVar(request("SearchType"),16)
	SearchValue = requestCheckVar(request("SearchValue"),16)
	SearchGubun = requestCheckVar(request("SearchGubun"),16)

	if SearchGubun = "" then SearchGubun = "0"

	nowdate = Left(CStr(now()),10)

	if (yyyy1="") then
		yyyy1 = Left(nowdate,4)
		mm1   = Mid(nowdate,6,2)
		dd1   = Mid(nowdate,9,2)
		yyyy2 = yyyy1
		mm2   = mm1
		dd2   = dd1
	end if

	searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

set ojumun = new cupchebeasong_list
	ojumun.FRectSearchType  = SearchType
	ojumun.FRectSearchValue = SearchValue
	ojumun.frectshopid = loginidshopormaker

	'/����� �����ΰ��
	if (SearchGubun = "1") then
		ojumun.FRectRegStart = DateSerial(yyyy1 , mm1 , dd1)
		ojumun.FRectRegEnd   = searchnextdate
	end if

	ojumun.fshop_maejangbaesong()

'/�ù�� �ϰ�����
Sub drawSelectBoxDeliverCompanyAssign(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>" onChange="AssignDeliverSelect(this);">
     <option value='' <%if selectedId="" then response.write " selected"%>>�ù�缱��</option><%
   query1 = " select top 100 divcd,divname from [db_order].[dbo].tbl_songjang_div where isUsing='Y' "
   query1 = query1 + " order by divcd"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("divcd")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("divcd")&"' "&tmp_str&">" & "" & replace(db2html(rsget("divname")),"'","") &  "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

'/�⺻�ù��.
dim idefaultSongjangDiv
	idefaultSongjangDiv = CStr(fnGetUpcheDefaultSongjangDiv(session("ssBctID")))
%>

<script language='javascript'>

function AssignDeliverSelect(comp){
    var frm = comp.form;
	var selecidx = comp.selectedIndex;
	var frm;

    if (frm.detailidx.length>1){
    	for (var i=0;i<frm.songjangdiv.length;i++){
    	    frm.songjangdiv[i][selecidx].selected=true;
    	}
    }else{
        frm.songjangdiv[selecidx].selected=true;
    }
}

function ShowOrderInfo(masteridx){
	var ShowOrderInfo = window.open('/common/offshop/beasong/upche_viewordermaster.asp?masteridx='+masteridx,'ShowOrderInfo','width=800,height=768,scrollbars=yes,resizable=yes');
	ShowOrderInfo.focus();
}

function CheckThis(comp,i){
    var frm = comp.form;

	if (comp.value.length>5){
	    if (frm.songjangno.length>1){
	        frm.detailidx[i].checked=true;
	        AnCheckClick(frm.detailidx[i]);
        }else{
            frm.detailidx.checked=true;
            AnCheckClick(frm.detailidx);
        }
	}
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function switchCheckBox(comp){
    var frm = comp.form;

	if(frm.detailidx.length>1){
		for(i=0;i<frm.detailidx.length;i++){
		    if (!frm.detailidx[i].disabled){
    			frm.detailidx[i].checked = comp.checked;
    			AnCheckClick(frm.detailidx[i]);
			}
		}
	}else{
	    if (!frm.detailidx.disabled){
    		frm.detailidx.checked = comp.checked;
    		AnCheckClick(frm.detailidx);
    	}
	}
}

function CheckNFinish(frm){
	var pass = false;
    var ordernoArr = "";
    var songjangnoArr  = "";
    var songjangdivArr = "";
    var detailidxArr   = "";

<% if (showshopselect = true) and (loginidshopormaker = "") then %>
	alert("���� ���� �����ϰ� �˻��ϼ���.");
	return;
<% end if %>

    if (!frm.detailidx){
        alert("���� �ֹ��� �����ϴ�.");
		return;
    }

    if(frm.detailidx.length>1){
    	for (var i=0;i<frm.detailidx.length;i++){
    	    pass = (pass||frm.detailidx[i].checked);
    	}
    }else{
        pass = frm.detailidx.checked;
    }

	if (!pass) {
		alert("���� �ֹ��� �����ϴ�.");
		return;
	}

    if(frm.detailidx.length>1){
    	for (var i=0;i<frm.detailidx.length;i++){
    		if (frm.detailidx[i].checked){
    			if (frm.songjangdiv[i].value.length<1){
    				alert("�ù�縦 �����Ͻñ� �ٶ��ϴ�.");
    				frm.songjangdiv[i].focus();
    				return;
    			}else if (trim(frm.songjangno[i].value).length<1){
    				alert("�����ȣ�� �Է��Ͻñ� �ٶ��ϴ�.");
    				frm.songjangno[i].focus();
    				return;
    			}

    			ordernoArr = ordernoArr + frm.orderno[i].value + ",";
				songjangnoArr  = songjangnoArr   + frm.songjangno[i].value + ",";
				songjangdivArr = songjangdivArr + frm.songjangdiv[i].value + ",";
				detailidxArr   = detailidxArr + frm.detailidx[i].value + ",";
    		}
    	}
    }else{
        if (frm.detailidx.checked){
			if (frm.songjangdiv.value.length<1){
				alert("�ù�縦 �����Ͻñ� �ٶ��ϴ�.");
				return;
			}else if (trim(frm.songjangno.value).length<1){
				alert("�����ȣ�� �Է��Ͻñ� �ٶ��ϴ�.");
				frm.songjangno.focus();
				return;
			}
		}
		ordernoArr = ordernoArr + frm.orderno.value + ",";
		songjangnoArr  = songjangnoArr   + frm.songjangno.value + ",";
		songjangdivArr = songjangdivArr + frm.songjangdiv.value + ",";
		detailidxArr   = detailidxArr + frm.detailidx.value + ",";
    }

	if (confirm("���� �ֹ� �����͸� ��� �Ϸ� ó�� �Ͻðڽ��ϱ�?")){
	    frm.ordernoArr.value = ordernoArr;
	    frm.songjangnoArr.value  = songjangnoArr;
        frm.songjangdivArr.value = songjangdivArr;
        frm.detailidxArr.value   = detailidxArr;

        frm.mode.value='SongjangInput';
        frm.action='/common/offshop/beasong/shopbeasong_Process.asp';
		frm.submit();
	}
}

function trim(theString){
   var resultString = theString;

   if (theString.indexOf(" ") == 0) {
        resultString = theString.substring(1, theString.length);
   }

   if (resultString.lastIndexOf(" ") == resultString.length) {
        resultString = resultString.substring(1,theString.length-1);
   }

   return resultString
}

function EnDisabledDateBox(){
	var bool = (frm.SearchGubun.value=="0");
	document.frm.yyyy1.disabled = bool;
	document.frm.yyyy2.disabled = bool;
	document.frm.mm1.disabled = bool;
	document.frm.mm2.disabled = bool;
	document.frm.dd1.disabled = bool;
	document.frm.dd2.disabled = bool;
}

function chksubmit(){
    var frm = document.frm;

    if ((frm.searchType.value.length>0)&&(frm.searchValue.value.length<1)){
        alert('�˻� ������ �Է��ϼ���.');
        frm.searchValue.focus();
        return;
    }

    if ((frm.searchType.value=="orderno")||(frm.searchType.value=="itemid")){
        if (!IsDigit(frm.searchValue.value)){
            alert('�˻� ������ ���ڸ� �����մϴ�.');
            frm.searchValue.focus();
            return;
        }
    }
    frm.submit();
}

function popMisendInput(detailidx){
    var popwin = window.open('/common/offshop/beasong/upche_popMisendInput.asp?detailidx=' + detailidx,'popMisendInput','width=600,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" onsubmit="chksubmit(); return false">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left" bgcolor="#FFFFFF">
		ShopID :
		<% if (showshopselect = true) then %>
			<% 'drawSelectBoxOffShop "shopid",loginidshopormaker %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",loginidshopormaker, "21") %>
		<% else %>
			<%= loginidshopormaker %>
		<% end if %>
		&nbsp;
		<select class="select" name="searchType" >
			<option value="">�˻�����</option>
			<option value="orderno" <%= ChkIIF(searchType="orderno","selected","") %> >�ֹ���ȣ</option>
			<option value="itemid" <%= ChkIIF(searchType="itemid","selected","") %> >��ǰ�ڵ�</option>
			<option value="buyname" <%= ChkIIF(searchType="buyname","selected","") %> >������</option>
			<option value="reqname" <%= ChkIIF(searchType="reqname","selected","") %> >������</option>
		</select>
		<input type="text" class="text" name="searchValue" value="<%= searchValue %>" size="13" maxlength="11">
		&nbsp;
		�����:
		<select class="select" name="SearchGubun" OnChange="EnDisabledDateBox()">
			<option value="0" <% if SearchGubun="0" then response.write "selected" %> >����� ��ü
			<option value="1" <% if SearchGubun="1" then response.write "selected" %> >��� �Ϸ���
		</select>

		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		(�����)
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:chksubmit();">
	</td>
</tr>
</form>
</table>

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
    	<input type="button" class="button" value="�����ֹ� ���ó��" onclick="CheckNFinish(frmbalju)">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmbalju" method="post" action="">
<input type="hidden" name="mode">
<input type="hidden" name="ordernoArr" value="">
<input type="hidden" name="songjangnoArr" value="">
<input type="hidden" name="songjangdivArr" value="">
<input type="hidden" name="detailidxArr" value="">
<input type="hidden" name="isall" value="">
<% if (showshopselect = true) then %>
	<% '�����϶��� �����̵� �ѱ��. %>
	<input type="hidden" name="shopid" value="<%= loginidshopormaker %>">
<% end if %>
<tr bgcolor="FFFFFF">
	<td height="25" colspan="15">
		�˻���� : <b><% = ojumun.FresultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="switchCheckBox(this);"></td>
	<td>IDX</td>
	<td>�����</td>
	<td>�ֹ���ȣ</td>
	<td>�ֹ���</td>
	<td>������</td>
	<td>��ǰ�ڵ�</td>
	<td>��ǰ��<font color="blue">&nbsp;[�ɼ�]</font></td>
	<td>����</td>
	<td>�ֹ��뺸��</td>
	<td>�����<br><font color="#AAAAAA">�������</font></td>
	<td>�����</td>
	<td><% drawSelectBoxDeliverCompanyAssign "defaultsongjangdiv","" %></td>
	<td align="center">������ȣ</td>
</tr>
<% if ojumun.FresultCount > 0 then %>
<% for ix=0 to ojumun.FresultCount-1 %>
<input type="hidden" name="orderno" value="<%= ojumun.FItemList(ix).Forderno %>">
<tr align="center" class="a" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="detailidx" value="<%= ojumun.FItemList(ix).Fdetailidx %>" onClick="AnCheckClick(this);" <%= CHKIIF(ojumun.FItemList(ix).FMisendReason="05","disabled","") %>></td>
	<td><%= ojumun.FItemList(ix).fdetailidx %></td>
	<td><%= ojumun.FItemList(ix).fshopname %></td>
	<td><a href="javascript:ShowOrderInfo('<%= ojumun.FItemList(ix).Fmasteridx %>')"><%= ojumun.FItemList(ix).Forderno %></a></td>
	<td><%= ojumun.FItemList(ix).FBuyname %></td>
	<td><%= ojumun.FItemList(ix).FReqname %></td>
	<td><%= ojumun.fitemlist(ix).fitemgubun %>-<%= CHKIIF(ojumun.fitemlist(ix).FitemID>=1000000,Format00(8,ojumun.fitemlist(ix).FitemID),Format00(6,ojumun.fitemlist(ix).FitemID)) %>-<%= ojumun.fitemlist(ix).fitemoption %></td>
	<td align="left">
		<%= ojumun.FItemList(ix).FItemname %>
		<% if (ojumun.FItemList(ix).FItemoption<>"") then %>
		<font color="blue">[<%= ojumun.FItemList(ix).FItemoption %>]</font>
		<% end if %>
	</td>
	<td><%= ojumun.FItemList(ix).FItemno %></td>
	<td><acronym title="<%= ojumun.FItemList(ix).Fbaljudate %>"><%= left(ojumun.FItemList(ix).Fbaljudate,10) %></acronym></td>
	<td><acronym title="<%= ojumun.FItemList(ix).Fbeasongdate %>"><%= left(ojumun.FItemList(ix).Fbeasongdate,10) %></acronym></td>
	<td>
		D+
		<% if IsNULL(ojumun.FItemList(ix).Fbaljudate) then %>
		    0
		<% elseif IsNULL(ojumun.FItemList(ix).Fbeasongdate) then %>
		    <%= datediff("d",(left(ojumun.FItemList(ix).Fbaljudate,10)) , (left(now,10)) ) %>
		<% else %>
			<% if datediff("d",(left(ojumun.FItemList(ix).Fbaljudate,10)) , (left(ojumun.FItemList(ix).Fbeasongdate,10))) < 0 then %>
			0
			<% else %>
			<%= datediff("d",(left(ojumun.FItemList(ix).Fbaljudate,10)) , (left(ojumun.FItemList(ix).Fbeasongdate,10)) ) %>
			<% end if %>
		<% end if %>
	</td>
	<td>
	    <% if (IsNULL(ojumun.FItemList(ix).FSongjangdiv) or (ojumun.FItemList(ix).FSongjangdiv=0)) then  %>
	        <% drawSelectBoxDeliverCompany "songjangdiv",idefaultSongjangDiv %>
	    <% else %>
	        <% drawSelectBoxDeliverCompany "songjangdiv",ojumun.FItemList(ix).FSongjangdiv %>
	    <% end if %>
	</td>
	<td><input type="text" class="text" name="songjangno" size="16" value="<%= ojumun.FItemList(ix).FSongjangno %>" onKeyup="CheckThis(this,'<%= ix %>');" maxlength=16 <%= CHKIIF(ojumun.FItemList(ix).FMisendReason="05","readonly style='background:#EEEEEE'","") %>></td>
</tr>
<% next %>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="14" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
</form>
</table>

<form name="frmshow" method="post">
	<input type="hidden" name="orderno" value="">
</form>

<form name="frmArrInput" method="post">
	<input type="hidden" name="detailidxArr" value="">
	<input type="hidden" name="iSall" value="">
	<input type="hidden" name="mode">
</form>

<script language='javascript'>
    document.onload = EnDisabledDateBox();
</script>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->