<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->

<%

dim itemid, oitem
dim makerid

itemid = requestCheckVar(request("itemid"),20)
makerid = requestCheckVar(request("makerid"),50)
menupos = requestCheckVar(request("menupos"),10)
if (itemid = "") then
    response.write "<script>alert('�߸��� �����Դϴ�.'); history.back();</script>"
    dbget.close()	:	response.End
end if


'==============================================================================
set oitem = new CItem

oitem.FRectItemID = itemid
oitem.FRectMakerId = session("ssBctID")
oitem.GetOneItem

if (oitem.FResultCount < 1) then
    response.write "<script>alert('�߸��� �����Դϴ�..');</script>"
    dbget.close()	:	response.End
end if


'==============================================================================
'���ϸ���
dim sailmargine
'���ݰ��
if oitem.FOneItem.Fsailyn="Y" then
	 if oitem.FOneItem.Fvatinclude = "Y" then
			on error resume next
			sailmargine = fix((CLng(oitem.FOneItem.Fsailprice)-Clng(oitem.FOneItem.Fsailsuplycash))/CLng(oitem.FOneItem.Fsailprice)*100*100)/100
			if Err then
				sailmargine = 0
			end if
	 else
			on error resume next
			sailmargine = fix((CLng(oitem.FOneItem.Fsailprice)-Clng(oitem.FOneItem.Fsailsuplycash)-CLng(oitem.FOneItem.Fbuyvat))/CLng(oitem.FOneItem.Fsailprice)*100*100)/100
			if Err then
				sailmargine = 0
			end if
	 end if
else
    sailmargine = 0
end if


'==============================================================================
Sub SelectBoxDesignerItem(selectedId)
   dim query1,tmp_str
   %><select name="designer" onchange="TnDesignerNMargineAppl(this.value);">
     <option value='' <%if selectedId="" then response.write " selected"%>>-- ��ü���� --</option><%
   query1 = " select userid,socname_kor,defaultmargine from [db_user].[dbo].tbl_user_c"
'   query1 = query1 + " where isusing='Y'"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")& "," & rsget("defaultmargine") & "' "&tmp_str&">" & rsget("userid") & "  [" & replace(db2html(rsget("socname_kor")),"'","") & "]" & "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub


%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script>
// ============================================================================
// ī�װ����(������;2010-09-13 ������-MD��û�� ���� ����)
/*
function editCategory(cdl,cdm,cds){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cds=" + cds ;

	popwin = window.open('/common/module/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function setCategory(cd1,cd2,cd3,cd1_name,cd2_name,cd3_name){
	var frm = document.itemreg;
	frm.cd1.value = cd1;
	frm.cd2.value = cd2;
	frm.cd3.value = cd3;
	frm.cd1_name.value = cd1_name;
	frm.cd2_name.value = cd2_name;
	frm.cd3_name.value = cd3_name;
}
*/

function popMultiLangEdit(iid) {
	window.open("/common/item/pop_MultiLangItemCont.asp?itemid="+iid+"&lang=EN", "multiLang_win", "width=1280, height=960, scrollbars=yes, resizable=yes");
}


// ============================================================================
// �����ϱ�
function SubmitSave() {
    if (validate(itemreg)==false) {
        return;
    }
    //��ü��۸� �ֹ����� ����.
    <% if oitem.FOneItem.Fmwdiv <> "U" then %>
   if(typeof(itemreg.itemdiv.length)!="undefined"){ 
	    if (itemreg.itemdiv[1].checked){
	        alert('�ֹ����� ��ǰ�� ��ü����ΰ�츸 �����մϴ�.');
	        itemreg.itemdiv[0].focus();
	        return;
	    }
	  }
    <% end if %>

	//��ǰ ���� �Ұ��׸� �˻�
	var cntRe = /.js["'>\s]/gi;
	if(cntRe.test(itemreg.itemcontent.value)) {
        alert('��ǰ������ js������ ���� �� �����ϴ�.');
        itemreg.itemcontent.focus();
        return;
	}
	
	//��ǰ���� ����üũ
 if (!IsDigit(itemreg.itemWeight.value)){
		alert('��ǰ���Դ�  ���ڷ� �Է��ϼ���.');
		itemreg.itemWeight.focus();
		return;
	}

	//��ǰ ǰ������
    if (!itemreg.infoDiv.value){
        alert('��ǰ�� �ش��ϴ� ǰ���� �������ֽʽÿ�.');
        itemreg.infoDiv.focus();
        return;
    } else if(itemreg.infoDiv.value=="35") {
    	if(!itemreg.itemsource.value) {
	        alert('��ǰ�� ������ �Է����ּ���.');
	        itemreg.itemsource.focus();
	        return;
    	}
    	if(!itemreg.itemsize.value) {
	        alert('��ǰ�� ũ�⸦ �Է����ּ���.');
	        itemreg.itemsize.focus();
	        return;
    	}
    }

	//������������
    if (itemreg.safetyYn[0].checked){
	    if (!itemreg.safetyDiv.value){
	        alert('�������������� �������ּ���.');
	        itemreg.safetyDiv.focus();
	        return;
	    }
	    if (!itemreg.safetyNum.value){
	        alert('����������ȣ�� �Է����ּ���.');
	        itemreg.safetyDiv.focus();
	        return;
	    }
    }

//�ؿܹ��
		if(document.itemreg.optionaddprice.value >0 && document.itemreg.deliverOverseas.checked){
			alert("�ɼǿ� �߰������� ���� ��� �ؿܹ���� �Ұ����մϴ�. �ؿܹ��üũ�� �������ּ���" );
			document.itemreg.deliverOverseas.focus();
			 return;
		}
		
		
 	if(document.itemreg.deliverOverseas.checked){
	    if(document.itemreg.itemWeight.value<=0){
	        alert("�ؿܹ�۽� ��ۺ� ������ ���� ��ǰ���Ը� �� �Է����ּ���")
	        document.itemreg.itemWeight.focus();
	        return;
	    }
	} 

	//ȭ���ݼۺ�
	try{
		if(itemreg.freight_min.value<=0||itemreg.freight_max.value<=0) {
            alert('ȭ����� ����� �Է����ּ���.');
            itemreg.freight_min.focus();
            return;
		}
	} catch(e) {}

    if(confirm("��ǰ�� �ø��ðڽ��ϱ�?") == true){
        itemreg.submit();
    }

}

function pop_10x10_person(){
	var popwin = window.open('/common/pop_10x10_person.asp','op2','width=450,height=570,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ClearVal(comp){
    comp.value = "";
}

function checkItemDiv(comp){
    var frm = comp.form;
    
    if (comp.name=="itemdiv"){
        if (frm.itemdiv[1].checked){
            frm.reqMsg.disabled=false;
        }else{
            //frm.reqMsg.checked=false;
            frm.reqMsg.disabled=true;
        }
    }
    
    //�ֹ����� ��ǰ�ΰ��.
    if (frm.itemdiv[1].checked){
        if (frm.reqMsg.checked){
            frm.itemdiv[1].value="06";
        }else{
            frm.itemdiv[1].value="16";
        }
    }
}

// ������������ ����
function chgSafetyYn(frm) {
	if(frm.safetyYn[0].checked) {
		frm.safetyDiv.disabled=false;
		frm.safetyNum.disabled=false;
	} else {
		frm.safetyDiv.disabled=true;
		frm.safetyNum.disabled=true;
	}
}

//ǰ�� ���� / ǰ�񳻿� ǥ��
function chgInfoDiv(v) {
	$("#itemInfoList").empty();

	if(v=="") {
		$("#itemInfoCont").hide();
	} else {
		$("#itemInfoCont").show();

		var str = $.ajax({
			type: "POST",
			url: "/admin/itemmaster/act_itemInfoDivForm.asp",
			data: "itemid=<%=itemid%>&ifdv="+v,
			dataType: "html",
			async: false
		}).responseText;
	
		if(str!="") {
			$("#itemInfoList").html(str);
		}
	}

	if(v=="35") {
		$("#lyItemSrc").show();
		$("#lyItemSize").show();
	} else {
		$("#lyItemSrc").hide();
		$("#lyItemSize").hide();
	}
}

//�ܼ� ���� ������
function chgInfoChk(fm) {
	$(fm).parent().parent().find('[name="infoChk"]').val($(fm).val());
}

//���� ���� ������
function chgInfoSel(fm) {
	$(fm).parent().parent().find('[name="infoChk"]').val($(fm).val());
	$(fm).parent().parent().find('[name="infoCont"]').val($(fm).attr("msg"));

	if($(fm).val()=="Y") {
		$(fm).parent().parent().find('[name="infoCont"]').removeAttr("readonly");
		$(fm).parent().parent().find('[name="infoCont"]').removeClass("text_ro");
		$(fm).parent().parent().find('[name="infoCont"]').addClass("text");
	} else {
		$(fm).parent().parent().find('[name="infoCont"]').attr("readonly", true);
		$(fm).parent().parent().find('[name="infoCont"]').addClass("text_ro");
	}
}
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
			��ϵ� ��ǰ�� �����մϴ�.<br>
			���ǻ����� ������ ���� �� ī�װ��� MD���� �����Ͻø� �˴ϴ�.
			&nbsp;&nbsp;
			<a href="javascript:pop_10x10_person();"><img src="/images/icon_arrow_link.gif" border="0" align="absbottom">&nbsp;ī�װ��� MD����ó</a> 
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
<form name="itemreg" method="post" action="do_upche_item_infomodify.asp" onsubmit="return false;">
<input type="hidden" name="itemid" value="<%= oitem.FOneItem.Fitemid %>">
<input type="hidden" name="optionaddprice" value="<%=oitem.FOneItem.fnGetOptAddPrice(oitem.FOneItem.Fitemid)%>">
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top"><img src="/images/icon_arrow_down.gif" border="0" align="absbottom">
	        	<strong>�⺻����</strong></td>
	        <td valign="top" align="right">&nbsp;</td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->



<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">��ǰ�ڵ� :</td>
  	<td bgcolor="#FFFFFF" colspan="2">
  	  <%= oitem.FOneItem.Fitemid %>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">��ü�� :</td>
  	<td bgcolor="#FFFFFF" colspan="2"><%= oitem.FOneItem.Fmakerid %></td>
  </tr>
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">ī�װ� ���� :</td>
  	<td bgcolor="#FFFFFF" colspan="2">
      <input type="hidden" name="cd1" value="<%= oitem.FOneItem.FCate_large %>">
      <input type="hidden" name="cd2" value="<%= oitem.FOneItem.FCate_mid %>">
      <input type="hidden" name="cd3" value="<%= oitem.FOneItem.FCate_small %>">
      <input type="text" name="cd1_name" value="<%= oitem.FOneItem.FCate_large_name %>" class="text" id="[on,off,off,off][ī�װ�]" size="20" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd2_name" value="<%= oitem.FOneItem.FCate_mid_name %>" class="text" id="[on,off,off,off][ī�װ�]" size="20" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd3_name" value="<%= oitem.FOneItem.FCate_small_name %>" class="text" id="[on,off,off,off][ī�װ�]" size="20" readonly style="background-color:#E6E6E6">
  	</td>
  </tr>
  <tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="����Ʈ�� ������ ī�װ�" style="cursor:help;">���� ī�װ� :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<table class=a>
		<tr>
			<td id="lyrDispList"><%=getDispOnlyCategory(oitem.FOneItem.Fitemid)%></td>
			<td valign="bottom"> </td>
		</tr>
		</table>
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
	</td>
</tr>
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">��ǰ���� :</td>
  	<td bgcolor="#FFFFFF" >
      <% if oitem.FOneItem.Fitemdiv="08" then %> 	 
     					<input type="radio" name="itemdiv" value="08" <%=chkIIF(oitem.FOneItem.Fitemdiv="08","checked","")%>  >Ƽ�ϻ�ǰ 
     				<% elseif oitem.FOneItem.Fitemdiv="09" then %> 	 
		 				<input type="radio" name="itemdiv" value="09" <%=chkIIF(oitem.FOneItem.Fitemdiv="09","checked","")%> >Present��ǰ 
		 			<% elseif oitem.FOneItem.Fitemdiv="18" then %> 	 	
	 					<input type="radio" name="itemdiv" value="18" <%=chkIIF(oitem.FOneItem.Fitemdiv="18","checked","")%>  >�����ǰ  
					<% elseif oitem.FOneItem.Fitemdiv ="82" then %>
			        <input type="radio" name="itemdiv" value="82" <%=chkIIF(oitem.FOneItem.Fitemdiv="82","checked","")%>  >���ϸ����� ��ǰ 
		 			<% elseif oitem.FOneItem.Fitemdiv ="75" then %> 
		 			<input type="radio" name="itemdiv" value="75" <%=chkIIF(oitem.FOneItem.Fitemdiv="75","checked","")%> >���ⱸ����ǰ 
		 				<% elseif oitem.FOneItem.Fitemdiv ="07" then %> 
		 			<input type="radio" name="itemdiv" value="07" <%=chkIIF(oitem.FOneItem.Fitemdiv="07","checked","")%> >�������ѻ�ǰ 
      <% else %>
		<label><input type="radio" name="itemdiv" value="01" <%=chkIIF(oitem.FOneItem.Fitemdiv ="01","checked","")%> onClick="this.form.requireMakeDay.value=0;document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">�Ϲݻ�ǰ</label>
		<br>
		<label><input type="radio" name="itemdiv" value="<%= oitem.FOneItem.Fitemdiv %>" <%=chkIIF(oitem.FOneItem.Fitemdiv="06" or oitem.FOneItem.Fitemdiv="16","checked","")%> onClick="document.getElementById('lyRequre').style.display='block';checkItemDiv(this);">�ֹ� ���ۻ�ǰ</label>
		<input type="checkbox" name="reqMsg" value="10" <%=chkIIF(oitem.FOneItem.Fitemdiv="06","checked","")%> <%=chkIIF(oitem.FOneItem.Fitemdiv="06" or oitem.FOneItem.Fitemdiv="16","","disabled")%> onClick="checkItemDiv(this);">�ֹ����� ���� �ʿ�<font color=red>(�ֹ��� �̴ϼȵ� ���۹����� �ʿ��Ѱ�� üũ)</font>
		<br>
		<%if not (oitem.FOneItem.Fitemdiv ="01" or oitem.FOneItem.Fitemdiv ="06" or oitem.FOneItem.Fitemdiv ="16") then%>
							<input type="hidden" name="itemdiv" id="itemdiv" value="<%=oitem.FOneItem.Fitemdiv%>">
							<%end if%>
      <% end if %>
  	</td>
  	<td bgcolor="#FFFFFF" >
  	    <div id="lyRequre" style="<%=chkIIF(oitem.FOneItem.Fitemdiv ="06" or oitem.FOneItem.Fitemdiv ="16","","display:none;")%>padding-left:22px;">
			�������ۼҿ��� <input type="text" name="requireMakeDay" value="<%=oitem.FOneItem.FrequireMakeDay%>" size="2" class="text" id="[off,on,off,off][�������ۼҿ���]">��
			<font color="red">(��ǰ�߼��� ��ǰ���� �Ⱓ)</font>
		</div>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">��ǰ�� :</td>
  	<td bgcolor="#FFFFFF" colspan="2">
      <%= oitem.FOneItem.Fitemname %>&nbsp;
  	</td>
  </tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ǰ�� :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<input type="text" name="itemnameEng" maxlength="64" size="60" class="text_ro" readonly id="[off,off,off,off][������ǰ��]" value="<%= oitem.FOneItem.FitemnameEng %>">&nbsp;
		<input type="button" value="���� ���� <%=chkIIF(oitem.FOneItem.FitemnameEng="" or isnull(oitem.FOneItem.FitemnameEng),"���","����")%>" class="button" onclick="popMultiLangEdit(<%= oitem.FOneItem.Fitemid %>)" />
	</td>
</tr>
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">������ :</td>
  	<td bgcolor="#FFFFFF" colspan="2">
      <input type="text" name="sourcearea" maxlength="64" size="25" class="text" id="[on,off,off,off][������]" value="<%= oitem.FOneItem.Fsourcearea %>">&nbsp;(ex:�ѱ�,�߱�,�߱�OEM,�Ϻ� �� / ��ǰ�� ��� ����: ������ �Ǵ� �ñ�����, ����: �̱���, �߱��� ��)
      <br>( ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">������ :</td>
  	<td bgcolor="#FFFFFF" colspan="2">
      <input type="text" name="makername" maxlength="32" size="25" class="text" id="[on,off,off,off][������]" value="<%= oitem.FOneItem.Fmakername %>">&nbsp;(������ü��)
  	</td>
  </tr>
  <tr align="left">
	<td height="30" width="160" bgcolor="#DDDDFF">��ǰ���� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemWeight" maxlength="12" size="8" id="[on,off,off,off][��ǰ����]" style="text-align:right" value="<%= oitem.FOneItem.Fitemweight %>">g &nbsp;(�׷������� �Է�, ex:1.5kg�� 1500) / �ؿܹ�۽� ��ۺ� ������ ���� ���̹Ƿ� ��Ȯ�� �Է�.
	</td>
</tr>
<tr align="left">
	<td height="30" width="160" bgcolor="#DDDDFF">������� :</td>
	<td   bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="deliverarea" value="" <%=chkIIF(Trim( oitem.FOneItem.Fdeliverarea)="" or IsNull( oitem.FOneItem.Fdeliverarea),"checked","")%>>�������&nbsp;
	  <input type="radio" name="deliverarea" value="C" <%=chkIIF( oitem.FOneItem.Fdeliverarea="C","checked","")%> <%if oitem.FOneItem.Fdeliverfixday<>"C" then%>disabled<%end if%>>�����ǹ��&nbsp;
	  <input type="radio" name="deliverarea" value="S" <%=chkIIF( oitem.FOneItem.Fdeliverarea="S","checked","")%> <%if oitem.FOneItem.Fdeliverfixday<>"C" then%>disabled<%end if%>>������&nbsp;
	  <label><input type="checkbox" name="deliverOverseas" value="Y" <%=chkIIF( oitem.FOneItem.FdeliverOverseas="Y","checked","")%> title="�ؿܹ���� ��ǰ���԰� �Է��� �ž� �Ϸ�˴ϴ�.">�ؿܹ��</label>
	</td>
</tr>
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">�˻�Ű���� :</td>
  	<td bgcolor="#FFFFFF" colspan="2">
      <input type="text" name="keywords" maxlength="260" size="120" class="text" id="[on,off,off,off][�˻�Ű����]" value="<%= oitem.FOneItem.Fkeywords %>">&nbsp;(�޸��α��� ex: Ŀ��,Ƽ����,����)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">��ü��ǰ�ڵ� :</td>
  	<td bgcolor="#FFFFFF" colspan="2">
  	    <input type="text" name="upchemanagecode" class="text" value="<%= oitem.FOneItem.Fupchemanagecode %>" size="30" maxlength="32" id="[off,off,off,off][��ü��ǰ�ڵ�]">
  	    (��ü���� �����ϴ� �ڵ� �ִ� 32�� - ����/���ڸ� ����)
  	</td>
  </tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ ���� :</td>
	<td bgcolor="#FFFFFF" colspan="2">
	  <input type="radio" name="usinghtml" value="N" <%=chkIIF(oitem.FOneItem.Fusinghtml="N","checked","")%>>�Ϲ�TEXT
	  <input type="radio" name="usinghtml" value="H" <%=chkIIF(oitem.FOneItem.Fusinghtml="H","checked","")%>>TEXT+HTML
	  <input type="radio" name="usinghtml" value="Y" <%=chkIIF(oitem.FOneItem.Fusinghtml="Y","checked","")%>>HTML���
	  <br>
	  <textarea name="itemcontent" rows="15" class="textarea" style="width:100%" id="[on,off,off,off][��ǰ����]"><%= oitem.FOneItem.Fitemcontent %></textarea>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ֹ��� ���ǻ��� :</td>
	<td bgcolor="#FFFFFF" colspan="2">
	  <textarea name="ordercomment" rows="5" cols="90" class="textarea" id="[off,off,off,off][���ǻ���]"><%= oitem.FOneItem.Fordercomment %></textarea><br>
	  <font color="red">Ư���� ��۱Ⱓ�̳� �ֹ��� Ȯ���ؾ߸� �ϴ� ����</font>�� �Է��Ͻø� ���Ҹ��̳� ȯ���� ���ϼ� �ֽ��ϴ�.
	</td>
</tr>
</table>

<!-- ǰ������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">ǰ������� &nbsp;<font color=gray>��ǰ����������� ���� ���� ������ ���� �Ʒ� ������ ��Ȯ�� �Է����ֽñ� �ٶ��ϴ�.</font></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">ǰ���� :</td>
	<td bgcolor="#FFFFFF">
		<select name="infoDiv" class="select" onchange="chgInfoDiv(this.value)">
		<option value="">::��ǰǰ��::</option>
		<option value="01">�Ƿ�</option>
		<option value="02">����/�Ź�</option>
		<option value="03">����</option>
		<option value="04">�м���ȭ(����/��Ʈ/�׼�����)</option>
		<option value="05">ħ����/Ŀư</option>
		<option value="06">����(ħ��/����/��ũ��/DIY��ǰ)</option>
		<option value="07">������(TV��)</option>
		<option value="08">������ ������ǰ(�����/��Ź��/�ı⼼ô��/���ڷ�����)</option>
		<option value="09">��������(������/��ǳ��)</option>
		<option value="10">�繫����(��ǻ��/��Ʈ��/������)</option>
		<option value="11">���б��(������ī�޶�/ķ�ڴ�)</option>
		<option value="12">��������(MP3/���ڻ��� ��)</option>
		<option value="14">������̼�</option>
		<option value="15">�ڵ�����ǰ(�ڵ�����ǰ/��Ÿ �ڵ�����ǰ)</option>
		<option value="16">�Ƿ���</option>
		<option value="17">�ֹ��ǰ</option>
		<option value="18">ȭ��ǰ</option>
		<option value="19">�ͱݼ�/����/�ð��</option>
		<option value="20">��ǰ(����깰)</option>
		<option value="21">������ǰ</option>
		<option value="22">�ǰ���ɽ�ǰ/ü��������ǰ</option>
		<option value="23">�����ƿ�ǰ</option>
		<option value="24">�Ǳ�</option>
		<option value="25">��������ǰ</option>
		<option value="26">����</option>
		<option value="35">��Ÿ</option>
		</select>
		<script type="text/javascript">
		document.itemreg.infoDiv.value="<%=oitem.FOneItem.FinfoDiv%>";
		</script>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:<%=chkIIF(isNull(oitem.FOneItem.FinfoDiv) or oitem.FOneItem.FinfoDiv="","none","")%>;">
	<td height="30" width="15%" bgcolor="#F8DDFF">ǰ�񳻿� :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList">
	<%
		if Not(isNull(oitem.FOneItem.FinfoDiv) or oitem.FOneItem.FinfoDiv="") then
			Server.Execute("/admin/itemmaster/act_itemInfoDivForm.asp")
		end if
	%>
	</td>
</tr>
<tr align="left">
	<td height="25" colspan="2" bgcolor="#FDFDFD"><font color="darkred">��ǰ���������� ������ ���� �Ǿ��ִ��� ��Ȯ�� �Է¹ٶ��ϴ�. ����Ȯ�ϰų� �߸��� ���� �Է½�, �׿� ���� å���� ���� ���� �ֽ��ϴ�.</font></td>
</tr>
<tr align="left" id="lyItemSrc" style="display:<%=chkIIF(oitem.FOneItem.FinfoDiv="35","","none")%>;">
  	<td height="30" width="160" bgcolor="#DDDDFF">��ǰ���� :</td>
  	<td bgcolor="#FFFFFF">
      <input type="text" name="itemsource" maxlength="64" size="50" class="text" value="<%= oitem.FOneItem.Fitemsource %>">&nbsp;(ex:�ö�ƽ,����,��,...)
  	</td>
</tr>
<tr align="left" id="lyItemSize" style="display:<%=chkIIF(oitem.FOneItem.FinfoDiv="35","","none")%>;">
  	<td height="30" width="160" bgcolor="#DDDDFF">��ǰ������ :</td>
  	<td bgcolor="#FFFFFF">
      <input type="text" name="itemsize" maxlength="64" size="50" class="text" value="<%= oitem.FOneItem.Fitemsize %>">&nbsp;(ex:7.5x15(cm))
  	</td>
</tr>
</table>
<!-- ������������ -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">������������</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">����������� :</td>
	<td bgcolor="#FFFFFF">
		<label><input type="radio" name="safetyYn" value="Y" <%=chkIIF(oitem.FOneItem.FsafetyYn="Y","checked","")%> onclick="chgSafetyYn(document.itemreg)"> ���</label>
		<label><input type="radio" name="safetyYn" value="N" <%=chkIIF(oitem.FOneItem.FsafetyYn<>"Y","checked","")%> onclick="chgSafetyYn(document.itemreg)"> ���ƴ�</label> /
		<select name="safetyDiv" <%=chkIIF(oitem.FOneItem.FsafetyYn<>"Y","disabled","")%> class="select">
		<option value="">::������������::</option>
		<option value="10" <%=chkIIF(oitem.FOneItem.FsafetyDiv="10","selected","")%>>������������(KC��ũ)</option>
		<option value="20" <%=chkIIF(oitem.FOneItem.FsafetyDiv="20","selected","")%>>�����ǰ ��������</option>
		<option value="30" <%=chkIIF(oitem.FOneItem.FsafetyDiv="30","selected","")%>>KPS �������� ǥ��</option>
		<option value="40" <%=chkIIF(oitem.FOneItem.FsafetyDiv="40","selected","")%>>KPS �������� Ȯ�� ǥ��</option>
		<option value="50" <%=chkIIF(oitem.FOneItem.FsafetyDiv="50","selected","")%>>KPS ��� ��ȣ���� ǥ��</option>
		</select>
		������ȣ <input type="text" name="safetyNum" <%=chkIIF(oitem.FOneItem.FsafetyYn<>"Y","disabled","")%> size="35" maxlength="25" class="text" value="<%=oitem.FOneItem.FsafetyNum%>" />
		
		<font color="darkred">���ƿ�ǰ�̳� �����ǰ�� ��� �ʼ� �Է�</font>
	</td>
</tr>
</table>

<%
	'ȭ����� �ݼۺ� �Է� (ȭ������� ����)
	if oitem.FOneItem.Fdeliverfixday="X" then
%>
<!-- ȭ����� �ݼۺ� �Է� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">ȭ����� ����</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">ȭ����� �ݼ� ��� :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		&nbsp;
		�ּ� <input type="text" name="freight_min" class="text" size="6" value="<%=oitem.FOneItem.Ffreight_min%>" style="text-align:right;">�� ~
		�ִ� <input type="text" name="freight_max" class="text" size="6" value="<%=oitem.FOneItem.Ffreight_max%>" style="text-align:right;">��
		<br>&nbsp; <font color="red">(��ǰ/��ȯ �� �� ���)</font>
	</td>
</tr>
</table>
<%	end if %>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="30">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="�����ϱ�" class="button" onClick="SubmitSave()">
          <input type="button" value="â �� ��" class="button" onClick="window.close()">
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
<p>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->