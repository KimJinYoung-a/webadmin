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
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2 , nowdate,searchnextdate,BeasongCom
dim dateback, SearchGubun ,SearchType, SearchValue ,ojumun ,i,iy
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
	ojumun.FRectDesignerID = session("ssBctID")

	'/��������
	ojumun.FRectMisendReason = "AA"

	'/����� �����ΰ��
	if (SearchGubun = "1") then
		ojumun.FRectRegStart = DateSerial(yyyy1 , mm1 , dd1)
		ojumun.FRectRegEnd   = searchnextdate
	end if

	ojumun.fDesignerDateBaljuinputlist()

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

function BatchSongjangInput(frm){
    var detailidxArr = '';
    if (!frm.detailidx){
        alert("�ϰ� ����� �ֹ��� �����ϼ���.");
		return;
    }

    if(frm.detailidx.length>1){
    	for (var i=0;i<frm.detailidx.length;i++){

    	    if (frm.detailidx[i].checked){
    	        detailidxArr = detailidxArr + frm.detailidx[i].value + ',';
    	    }
    	}
    }else{
        if (frm.detailidx.checked){
            detailidxArr = frm.detailidx.value;
        }
    }

	if (detailidxArr.length<1) {
		alert("�ϰ� ����� �ֹ��� �����ϼ���.");
		return;
	}

    var popwin = window.open('','BatchSongjangInput','width=600,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();

    document.frmArrInput.detailidxArr.value=detailidxArr;
    document.frmArrInput.iSall.value ="";
	document.frmArrInput.target = "BatchSongjangInput";
    document.frmArrInput.mode.value='SongjangInput';
	document.frmArrInput.action='upche_BatchSongjangInput.asp';
	document.frmArrInput.submit();
}

function CheckNFinish(frm){
	var pass = false;
    var ordernoArr = "";
    var songjangnoArr  = "";
    var songjangdivArr = "";
    var detailidxArr   = "";

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
        frm.action='/common/offshop/beasong/upche_beasong_Process.asp';
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
	    <input type="button" class="button" value="���� File �ϰ����" onclick="BatchSongjangInput(frmbalju)">
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

<tr bgcolor="FFFFFF">
	<td height="25" colspan="15">
		�˻���� : <b><% = ojumun.FresultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="switchCheckBox(this);"></td>
	<td>�ϷĹ�ȣ</td>
	<td>�ֹ���ȣ</td>
	<td>������</td>
	<td>��ǰ�ڵ�</td>
	<td>��ǰ��<font color="blue">&nbsp;[�ɼ�]</font></td>
	<td>���ް�</td>
	<td>�ǸŰ�</td>
	<td>����</td>
	<td>�ֹ��뺸��</td>
	<td>�����<br><font color="#AAAAAA">�������</font></td>
	<td>�����</td>
	<td><% drawSelectBoxDeliverCompanyAssign "defaultsongjangdiv","" %></td>
	<td align="center">������ȣ</td>
	<!--<td align="center">��������<br>����</td>-->
</tr>
<% if ojumun.FresultCount > 0 then %>
<% for i=0 to ojumun.FresultCount-1 %>
<input type="hidden" name="orderno" value="<%= ojumun.FItemList(i).Forderno %>">
<tr align="center" class="a" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="detailidx" value="<%= ojumun.FItemList(i).Fdetailidx %>" onClick="AnCheckClick(this);" <%= CHKIIF(ojumun.FItemList(i).FMisendReason="05","disabled","") %>></td>
	<td><%= ojumun.FItemList(i).fdetailidx %></td>
	<td><a href="javascript:ShowOrderInfo('<%= ojumun.FItemList(i).Fmasteridx %>')"><%= ojumun.FItemList(i).Forderno %></a></td>
	<td><%= ojumun.FItemList(i).FReqname %></td>
	<td><%= ojumun.fitemlist(i).fitemgubun %>-<%= FormatCode(ojumun.fitemlist(i).FitemID) %>-<%= ojumun.fitemlist(i).fitemoption %></td>
	<td align="left">
		<%= ojumun.FItemList(i).FItemname %>
		<% if (ojumun.FItemList(i).FItemoption<>"") then %>
		<font color="blue">[<%= ojumun.FItemList(i).FItemoption %>]</font>
		<% end if %>
	</td>
	<td><%= FormatNumber(ojumun.fitemlist(i).fsuplyprice,0) %></td>
	<td><%= FormatNumber(ojumun.fitemlist(i).fsellprice,0) %></td>
	<td><%= ojumun.FItemList(i).FItemno %></td>
	<td><acronym title="<%= ojumun.FItemList(i).Fbaljudate %>"><%= left(ojumun.FItemList(i).Fbaljudate,10) %></acronym></td>
	<td><acronym title="<%= ojumun.FItemList(i).Fbeasongdate %>"><%= left(ojumun.FItemList(i).Fbeasongdate,10) %></acronym></td>
	<td>
		D+
		<% if IsNULL(ojumun.FItemList(i).Fbaljudate) then %>
		    0
		<% elseif IsNULL(ojumun.FItemList(i).Fbeasongdate) then %>
		    <%= datediff("d",(left(ojumun.FItemList(i).Fbaljudate,10)) , (left(now,10)) ) %>
		<% else %>
			<% if datediff("d",(left(ojumun.FItemList(i).Fbaljudate,10)) , (left(ojumun.FItemList(i).Fbeasongdate,10))) < 0 then %>
			0
			<% else %>
			<%= datediff("d",(left(ojumun.FItemList(i).Fbaljudate,10)) , (left(ojumun.FItemList(i).Fbeasongdate,10)) ) %>
			<% end if %>
		<% end if %>
	</td>
	<td>
	    <% if (IsNULL(ojumun.FItemList(i).FSongjangdiv) or (ojumun.FItemList(i).FSongjangdiv=0)) then  %>
	        <% drawSelectBoxDeliverCompany "songjangdiv",idefaultSongjangDiv %>
	    <% else %>
	        <% drawSelectBoxDeliverCompany "songjangdiv",ojumun.FItemList(i).FSongjangdiv %>
	    <% end if %>
	</td>
	<td><input type="text" class="text" name="songjangno" size="16" value="<%= ojumun.FItemList(i).FSongjangno %>" onKeyup="CheckThis(this,'<%= i %>');" maxlength=16 <%= CHKIIF(ojumun.FItemList(i).FMisendReason="05","readonly style='background:#EEEEEE'","") %>></td>
	<!--<td>
	    <%' if (ojumun.FItemList(i).isMisendAlreadyInputed) then %>
        	<a href="javascript:popMisendInput('<%= ojumun.FItemList(i).Fdetailidx %>');"><%'= ojumun.FItemList(i).getMisendText %></a>
        <%' end if %>
	</td>-->
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
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->