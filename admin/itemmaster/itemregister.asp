<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description : �¶��λ�ǰ���
' History : ������ ����
'			2017.11.27 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_v2.asp"-->
<%
CONST CBASIC_IMG_MAXSIZE = 560   'KB
CONST CMAIN_IMG_MAXSIZE = 640   'KB

dim i,j, designer, rentalItemFlag
'==============================================================================
Sub SelectBoxDesignerItem()
   dim query1 
   %><select name="designer" class="select" onchange="TnDesignerNMargineAppl(this.value);">
     <option value=''>-- ��ü���� --</option><%
   query1 = " select userid,socname_kor,defaultmargine, maeipdiv, IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(defaultDeliverPay,0) as defaultDeliverPay, IsNULL(defaultDeliveryType,'') as defaultDeliveryType from [db_user].[dbo].tbl_user_c"
   query1 = query1 + " where isusing='Y'"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           response.write("<option value='"&rsget("userid")& "," & rsget("defaultmargine") & "," & rsget("maeipdiv") & "," & rsget("defaultFreeBeasongLimit") & "," & rsget("defaultDeliverPay") & "," & rsget("defaultDeliveryType") & "'>" & rsget("userid") & "  [" & replace(db2html(rsget("socname_kor")),"'","") & "]" & "</option>")
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

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
<script type='text/javascript' src="/js/ckeditor/ckeditor.js"></script>
<script type="text/javascript">
<!-- #include file="./itemregister_javascript.asp"-->
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
		<font color="red"><strong>��ǰ���</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br><b>�Ż�ǰ�� ����մϴ�.</b>
			<!--
            <br>- ���� ȭ���� ���� ����ϼž� �����Ͽ� ���� �� ������Ʈ �˴ϴ�.
            <br>- �����̳� ������ ������ ��� ���� �źε� �� �ֽ��ϴ�.
            -->
			<br>- �⺻Ʋ������ �̿��Ͽ� ������ ��ǰ�� ����Ҽ� �ֽ��ϴ�.
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
<input type="button" class="button" value="�⺻Ʋ����" onClick="UseTemplate();"><br><br>

<form name="itemreg" method="post" action="<%= ItemUploadUrl %>/linkweb/items/itemregisterWithImage_process.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="designerid">
<input type="hidden" name="defaultmargin">
<input type="hidden" name="defaultmaeipdiv">
<input type="hidden" name="defaultFreeBeasongLimit">
<input type="hidden" name="defaultDeliverPay">
<input type="hidden" name="defaultDeliveryType">
<input type="hidden" name="DFcolorCD" value="">
<input type="hidden" name="itemoptioncode2">
<input type="hidden" name="itemoptioncode3">

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
	<td bgcolor="#FFFFFF" colspan="3"><% NewDrawSelectBoxDesignerChangeMargin "makerid", designer, "marginData", "TnDesignerNMargineAppl2" %></td>
	<% 'SelectBoxDesignerItem %> <!--(����ü�� ǥ�õ˴ϴ�)-->
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <input type="text" name="itemname" maxlength="64" size="50" class="text" id="[on,off,off,off][��ǰ��]">&nbsp;
	</td>
</tr>
<!-- ��ü��Ͻÿ��� ������(MD�� ��ϰ���) -->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰī�� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="designercomment" size="60" maxlength="128" class="text" id="[off,off,off,off][��ǰī��]"><br>
	</td>
</tr>
</table>

<!-- 2.���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>2.����</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
<input type="hidden" name="cd1" value="">
<input type="hidden" name="cd2" value="">
<input type="hidden" name="cd3" value="">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="���/���� ���� ���� ī�װ�" style="cursor:help;">���� ī�װ� :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<input type="text" name="cd1_name" value="" id="[on,off,off,off][ī�װ�]" size="20" readonly class="text_ro">
		<input type="text" name="cd2_name" value="" id="[on,off,off,off][ī�װ�]" size="20" readonly class="text_ro">
		<input type="text" name="cd3_name" value="" id="[on,off,off,off][ī�װ�]" size="20" readonly class="text_ro">
		
		<input type="button" value="ī�װ� ����" class="button" onclick="editCategory(itemreg.cd1.value,itemreg.cd2.value,itemreg.cd3.value);">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="����Ʈ�� ������ ī�װ�" style="cursor:help;">���� ī�װ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td id="lyrDispList"><table class="a" id="tbl_DispCate"></table></td>
			<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
		</tr>
		</table> 
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
	 </td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
	<td bgcolor="#FFFFFF" >
		<label><input type="radio" name="itemdiv" value="01" checked onClick="this.form.requireMakeDay.value=0;document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">�Ϲݻ�ǰ</label>
		<br>
		<label><input type="radio" name="itemdiv" value="06" onClick="document.getElementById('lyRequre').style.display='block';checkItemDiv(this);">�ֹ� ���ۻ�ǰ</label>
		<input type="checkbox" name="reqMsg" value="10" onClick="checkItemDiv(this);">�ֹ����� ���� �ʿ�<font color=red>(�ֹ��� �̴ϼȵ� ���۹����� �ʿ��Ѱ�� üũ)</font>
		<br>
		<!--��ü��Ͻÿ��� ������(MD�� ��ϰ���) -->
		<label><input type="radio" name="itemdiv" value="08" onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">Ƽ�ϻ�ǰ</label>
		<label><input type="radio" name="itemdiv" value="09" onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">Present��ǰ</label>
		<label><input type="radio" name="itemdiv" value="11" onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">��ǰ�ǻ�ǰ</label>
		<% If rentalItemFlag Then %>
			<label><input type="radio" name="itemdiv" value="30" onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">��Ż��ǰ</label>
		<% End If %>
		<label><input type="radio" name="itemdiv" value="23" onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">B2B��ǰ</label>
		<label><input type="radio" name="itemdiv" value="17" onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">�����������ǰ</label>
	</td>
    <td bgcolor="#FFFFFF">
        <div id="lyRequre" style="display:none;padding-left:22px;">
		�������ۼҿ��� <input type="text" name="requireMakeDay" value="0" size="2" class="text" id="[off,on,off,off][�������ۼҿ���]">��
		<font color="red">(��ǰ�߼��� ��ǰ���� �Ⱓ)</font>
		</div>
	</td>
</tr>
<!-- ��ü��Ͻÿ��� ������(MD�� ��ϰ���) -->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ٹ����� �������� :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<label><input type="radio" name="tenOnlyYn" value="Y" >������ǰ</label>
		<label><input type="radio" name="tenOnlyYn" value="N" checked>�Ϲݻ�ǰ</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">���� ���� ���� :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<label><input type="radio" name="adultType" value="0" checked>��ü����</label>
		<label><input type="radio" name="adultType" value="1" >���Žü�������</label>
		<label><input type="radio" name="adultType" value="2" >�̼��� ��ȸ �Ұ�</label>
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
	<td height="30" width="15%" bgcolor="#DDDDFF">���� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="margin" maxlength="32" size="5" class="text" id="[on,off,off,off][����]">%
		<input type="button" value="���ް� �ڵ����" class="button" onclick="CalcuAuto(itemreg);">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ǸŰ�(�Һ��ڰ�) :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<input type="text" name="sellcash" maxlength="16" size="12" class="text" id="[on,on,off,off][�Һ��ڰ�]" onKeyup="CalcuAuto(itemreg);">��
		<input type="hidden" name="sellvat">
	</td>
	<td width="15%" bgcolor="#DDDDFF">���ް� :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<input type="text" name="buycash" maxlength="16" size="12" class="text" id="[on,on,off,off][���ް�]" >��
		(<b>�ΰ��� ���԰�</b>)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">���ϸ��� :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="mileage" maxlength="32" size="10" id="[on,on,off,off][���ϸ���]" value="0" ReadOnly > (�ǸŰ��� 1%)
	</td>
	<td width="15%" bgcolor="#DDDDFF">����, �鼼 ���� :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="vatinclude" value="Y" checked onclick="TnGoClear(this.form);CalcuAuto(itemreg);">����</label>
		<label><input type="radio" name="vatinclude" value="N" onclick="TnGoClear(this.form);CalcuAuto(itemreg);">�鼼</label>
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
	    <input type="text" name="itemid" value="" size="20" class="text_ro" readonly id="[off,off,off,off][��ǰ�ڵ�]">
	    (��ǰ��� �Ϸ�� �ο��˴ϴ�.)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="��ǰ �� �Ӽ�" style="cursor:help;">��ǰ�Ӽ� :</td>
	<td id="lyrItemAttribAdd" bgcolor="#FFFFFF" colspan="3"></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ü��ǰ�ڵ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	    <input type="text" name="upchemanagecode" value="" size="20" maxlength="32" class="text" id="[off,off,off,off][��ü��ǰ�ڵ�]">
	    (��ü���� �����ϴ� �ڵ� �ִ� 32�� - ����/���ڸ� ����)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">ISBN :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		ISBN 13 <input type="text" name="isbn13" class="text" value="" size="13" maxlength="13">
		/ �ΰ���ȣ <input type="text" name="isbn_sub" class="text" value="" size="5" maxlength="5"><br />
		ISBN 10 <input type="text" name="isbn10" class="text" value="" size="10" maxlength="10"> (Optional)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ǰ��� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	    <input type="text" name="relateItems" value="" size="40" class="text" id="[off,off,off,off][������ǰ]">
	    (������ǰ�� �ִ� 6������ ��ϰ���, ��ǰ��ȣ�� �޸�(,)�� �����Ͽ� �Է�)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�Ǹſ��� :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="sellyn" value="Y">�Ǹ���</label>&nbsp;&nbsp;
		<label><input type="radio" name="sellyn" value="N" checked>�Ǹž���</label>
	</td>
	<td width="15%" bgcolor="#DDDDFF">��뿩�� :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="isusing" value="Y" onclick="TnChkIsUsing(this.form)">�����</label>&nbsp;&nbsp;
		<label><input type="radio" name="isusing" value="N" onclick="TnChkIsUsing(this.form)">������</label>
	</td>
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
		<input type="text" name="makername" maxlength="32" size="25" class="text" id="[on,off,off,off][������]">&nbsp;(������ü��)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	 <p> 
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="0" checked onClick="jsSetArea(this.value);"> ��ǰ ��</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="1"  onClick="jsSetArea(this.value);"> ����깰</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="2"  onClick="jsSetArea(this.value);"> ���깰</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="3"  onClick="jsSetArea(this.value);"> ��깰</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="4"  onClick="jsSetArea(this.value);"> ����갡��ǰ</span>
	  </p>
	  <p><input type="text" name="sourcearea" maxlength="64" size="64" class="text" id="[on,off,off,off][������]" /></p>
	  <div id="dvArea0" style="display:;">
	  <p><strong>ex: �ѱ�, �߱�, �߱�OEM, �Ϻ� �� </strong></BR>
	   - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	  <div id="dvArea1" style="display:none;">
	  <p><strong>������ :</strong> ����, ������ �Ǵ� �á�����, �á�����(���ѹα�, �ѱ�X)  <span style="margin-right:10px;">ex. ��(����)</span></BR>
	   <strong>���Ի� :</strong> ������� ���Ա����� <span style="margin-right:10px;">ex. ����(�߱���)</span></BR>
	   - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	  <div id="dvArea2" style="display:none;">
	  <p><strong>������ :</strong> ����,������ �Ǵ� �����ػ�(��� ���깰�� �á����� ����)   <span style="margin-right:10px;">ex. ��ġ(����), ��¡��(�����ػ�)</span> </BR>
	  	<strong>����� :</strong> ����� �Ǵ� �����(�ؿ���)   <span style="margin-right:10px;">ex. ��ġ[�����(�뼭��)]</span> </BR>
	    <strong>���Ի� :</strong> ������� ���Ա����� <span style="margin-right:10px;">ex. ���(�߱���)</span></BR>
	   - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	  <div id="dvArea3" style="display:none;">
	  <p>�Ұ���� ��� ������ ����(�ѿ�/����/���ұ���) �� ������   <span style="margin-right:10px;">ex. ����(Ⱦ���� �ѿ�), ����(ȣ�ֻ�)</span></BR>
	  - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	  <div id="dvArea4" style="display:none;">
	  <p><strong>98%�̻� ���ᰡ �ִ� ���:</strong>  �Ѱ��� ���Ḹ ǥ�� ����    <span style="margin-right:10px;">ex. ����(�̱���)</span> </BR>
	  	<strong>���� ���Ḧ ����� ���:</strong> ȥ�պ����� ���� ������ 2�� ����   <span style="margin-right:10px;">ex. ������[�а���(�̱���),���尡��(������)]</span></BR>
	  - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div> 
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemWeight" maxlength="12" size="8" id="[on,off,off,off][��ǰ����]" style="text-align:right" value="0">g &nbsp;(�׷������� �Է�, ex:1.5kg�� 1500) / �ؿܹ�۽� ��ۺ� ������ ���� ���̹Ƿ� ��Ȯ�� �Է�.
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�˻�Ű���� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="keywords" maxlength="250" size="60" class="text" id="[on,off,off,off][�˻�Ű����]">&nbsp;(�޸��α��� ex: Ŀ��,Ƽ����,����)
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
		<% DrawInfoDiv "infoDiv", "", " onchange='chgInfoDiv(this.value);'" %>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:none">
	<td height="30" width="15%" bgcolor="#F8DDFF">ǰ�񳻿� :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList"></td>
</tr>
<tr align="left">
	<td height="25" colspan="2" bgcolor="#FDFDFD"><font color="darkred">��ǰ���������� ������ ���� �Ǿ��ִ��� ��Ȯ�� �Է¹ٶ��ϴ�. ����Ȯ�ϰų� �߸��� ���� �Է½�, �׿� ���� å���� ���� ���� �ֽ��ϴ�.</font></td>
</tr>
<tr align="left" id="lyItemSrc" style="display:none;">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemsource" maxlength="64" size="50" class="text">&nbsp;(ex:�ö�ƽ,����,��,...)
	</td>
</tr>
<tr align="left" id="lyItemSize" style="display:none;">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ������ :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemsize" maxlength="64" size="50" class="text">
		<select name="unit" class="select">
		<option value="">�����Է�</option>
		<option value="mm">mm</option>
		<option value="cm" selected>cm</option>
		<option value="m��">m��</option>
		<option value="km">km</option>
		<option value="m��">m��</option>
		<option value="km��">km��</option>
		<option value="ha">ha</option>
		<option value="m��">m��</option>
		<option value="cm��">cm��</option>
		<option value="L">L</option>
		<option value="g">g</option>
		<option value="Kg">Kg</option>
		<option value="t">t</option>
		</select>
		&nbsp;(ex:7.5x15(cm))
		</td>
</tr>
</table>
<!-- 5-2.������������ -->
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
				<label><input type="radio" name="safetyYn" value="Y" checked onclick="chgSafetyYn(document.itemreg)" /> ���</label>
				<label><input type="radio" name="safetyYn" value="N" onclick="chgSafetyYn(document.itemreg)" /> ���ƴ�</label>
				<label><input type="radio" name="safetyYn" value="I" onclick="chgSafetyYn(document.itemreg)" /> ��ǰ���� ǥ��</label>
				<label><input type="radio" name="safetyYn" value="S" onclick="chgSafetyYn(document.itemreg)" /> ���������ؼ�</label>
				<input type="hidden" name="auth_go_catecode" id="auth_go_catecode" value="">
				<input type="hidden" name="real_safetydiv" id="real_safetydiv" value="">
				<input type="hidden" name="real_safetynum" id="real_safetynum" value="">
				<input type="hidden" name="real_safetyidx" id="real_safetyidx" value="">
			</td>
		</tr>
		<tr align="left">
			<td bgcolor="#FFFFFF">
				<% drawSelectBoxSafetyDivCode "safetyDiv", "", "Y", "" %>
				������ȣ <input type="text" name="safetyNum" id="[off,off,off,off][�������� ������ȣ]" size="35" maxlength="25" value="" />
				<input type="button" id="safetybtn" value="��   ��" onclick="jsSafetyAuth();" class="button">
				<input type="hidden" name="issafetyauth" id="issafetyauth" value="">
			</td>
		</tr>
		<tr align="left">
			<td bgcolor="#FFFFFF">
				<div id="safetyDivList"></div>
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
		<label><input type="radio" name="mwdiv" value="M" checked onclick="TnCheckUpcheYN(this.form);">����</label>
		<label><input type="radio" name="mwdiv" value="W" onclick="TnCheckUpcheYN(this.form);">Ư��</label>
		<label><input type="radio" name="mwdiv" value="U" onclick="TnCheckUpcheYN(this.form);">��ü���</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��۱��� :</td>
	<td width="85%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverytype" value="1" checked  onclick="TnCheckUpcheDeliverYN(this.form);">�ٹ����ٹ��</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);">��ü(����)���</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="4" onclick="TnCheckUpcheDeliverYN(this.form);">�ٹ����ٹ�����</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);">��ü���ǹ��(���� ��ۺ�ΰ�)</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);">��ü���ҹ��</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��۹�� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverfixday" value="" checked onclick="TnCheckFixday(this.form)">�ù�(�Ϲ�)</label>&nbsp;
		<label><input type="radio" name="deliverfixday" value="X" onclick="TnCheckFixday(this.form)">ȭ��</label>&nbsp;
		<label><input type="radio" name="deliverfixday" value="C" disabled onclick="TnCheckFixday(this.form)">�ö��������</label>
		<label><input type="radio" name="deliverfixday" value="G" onclick="TnCheckFixday(this.form)">�ؿ�����</label>
		<label><input type="radio" name="deliverfixday" value="L" disabled onclick="TnCheckFixday(this.form)">Ŭ����</label>
		<span id="lyrFreightRng" style="display:none;">
			<br />&nbsp;
			��ǰ/��ȯ �� ȭ����� ���(��) :
			�ּ� <input type="text" name="freight_min" class="text" size="6" value="0" style="text-align:right;">�� ~
			�ִ� <input type="text" name="freight_max" class="text" size="6" value="0" style="text-align:right;">��
		</span>
		<br>&nbsp;<font color="red">(�ö�� ��ǰ�� ��츸 �����ǹ��, ������, �ö�������� �ɼ��� ��밡���մϴ�.)</font>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverarea" value="" checked>�������</label>&nbsp;
		<label><input type="radio" name="deliverarea" value="C" disabled >�����ǹ��</label>&nbsp;
		<label><input type="radio" name="deliverarea" value="S" disabled >������</label>
		<label><input type="checkbox" name="deliverOverseas" value="Y" checked title="�ؿܹ���� ��ǰ���԰� �Է��� �ž� �Ϸ�˴ϴ�.">�ؿܹ��</label>
	</td>
</tr>
<input type="hidden" name="pojangok" value="Y">
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
	<td width="85%" bgcolor="#FFFFFF">
		<label><input type="radio" name="useoptionyn" value="Y" onClick="TnCheckOptionYN(this.form);">�ɼǻ����</label>&nbsp;&nbsp;
		<label><input type="radio" name="useoptionyn" value="N" onClick="TnCheckOptionYN(this.form);" checked>�ɼǻ�����</label>
	</td>
</tr>
<!----- �ɼǱ��� DIV ----->
<tr id="opttype" style="display:none" height="40">
    <td width="15%" bgcolor="#DDDDFF">�ɼ� ����  :</td>
    <td width="85%" bgcolor="#FFFFFF">
        <label><input type="radio" name="optlevel" value="1" onClick="TnCheckOptionYN(this.form);" checked >���� �ɼ� (�ɼ� ���� 1��)</label>
        <label><input type="radio" name="optlevel" value="2" onClick="TnCheckOptionYN(this.form);" >���� �ɼ� (�ɼ� ���� �ִ� 3��)</label><!--<font color="blue">�� ����Ư�������� ��ü����� ��츸 ���ð����մϴ�.</font> //2016.05.19 ������ ����--> 
    </td>
</tr>
<!----- ���� �ɼ� DIV ----->
<tr id="optlist" style="display:none" height="30">
    <td width="15%" bgcolor="#DDDDFF">�ɼ� ���� :</td>
  	<td width="85%" bgcolor="#FFFFFF">
      	<table width="500" border="0" cellspacing="0" cellpadding="0" class="a" >
      	<tr>
      	    <td width="100">�ɼ� ���и� :</td>
      	    <td width="400"><input type="text" name="optTypeNm" value="" size="20" maxlength="20" class="text" id="[off,off,off,off][�ɼ� ���и�]"></td>
      	</tr>
      	<tr>
      	    <td colspan="2">
              <select multiple name="realopt" class="select" style="width:400px;height:120px;"></select>
            </td>
        </tr>
        <tr>
            <td colspan="2">
              <input type="button" value="�⺻�ɼ��߰�" name="btnoptadd" class="button" onclick="popNormalOptionAdd();" >
              <input type="button" value="����ɼ��߰�" name="btnetcoptadd" class="button" onclick="popEtcOptionAdd();">
              <input type="button" value="���ÿɼǻ���" name="btnoptdel" class="button" onclick="delItemOptionAdd()" >
              <br><br>
              - �⺻�ɼ��߰� : ����, ������� �⺻������ ���ǵ� �ɼ��� �߰� �Ͻ� �� �ֽ��ϴ�.<br>
              - ����ɼ��߰� : �⺻�ɼǿ� ���ǵ��� ���� ��ǰ����ɼ��� �����Ͻ� �� �ֽ��ϴ�.<br>
              - ���ÿɼǻ��� : ���õ� �ɼ��� �����մϴ�.<br>
              - ���ǻ��� : �ѹ� ����� �ɼ��� <font color=red>������ �Ұ���</font>�մϴ�.<br>
              <br>
            </td>
        </tr>
        </table>
  	</td>
</tr>
<%
dim iMaxCols : iMaxCols = 3
dim iMaxRows : iMaxRows = 9
%>
<!----- ��Ƽ �ɼ� DIV ----->
<tr id="optlist2" style="display:none" height="30">
    <td width="15%" bgcolor="#DDDDFF">�ɼǼ��� :</td>
    <td width="85%" bgcolor="#FFFFFF">
        <table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" class="a"  bgcolor="#3d3d3d">
        <tr align="center"  bgcolor="#DDDDFF">
            <td width="100">�ɼǱ��и�</td>
            <% for j=0 to iMaxCols-1 %>
            <td>
                <input type="text" name="optionTypename<%= j+1 %>" value="" size="18" maxlength="20" class="text" id="[off,off,off,off][�ɼ� ���и�<%= j %>]">
            </td>
            <% Next %>
            <td width="80">(��Ͽ���)<br>����</td>
            <td width="80">(��Ͽ���)<br>������</td>
        </tr>
        <tr height="2" bgcolor="#FFFFFF">
            <td colspan="6"></td>
        </tr>
        <% for i=0 to iMaxRows-1 %>
        <tr align="center"  bgcolor="#FFFFFF">
            <td>�ɼǸ� <%= i+1 %></td>
            <% for j=0 to iMaxCols-1 %>
            <td>
                <input type="hidden" name="itemoption<%= j+1 %>" value="">
                <input type="text" name="optionName<%= j+1 %>" size="18" maxlength="20" class="text" id="[off,off,off,off][�ɼǸ�<%= i %><%= j %>]">
            </td>
            <% next %>
            <td>
                <% if i=0 then %>
                ����
                <% elseif i=1 then %>
                �Ķ�
                <% elseif i=2 then %>
                ���
                <% elseif i=3 then %>
                ������
                <% end if %>
            </td>
            <td>
                <% if i=0 then %>
                XL
                <% elseif i=1 then %>
                L
                <% elseif i=2 then %>
                S
                <% end if %>
            </td>
        </tr>
        <% next %>
        </table>
     </td>
</tr>

<!----- �⺻ ���� DIV ----->
<tr id="lyDFColor" height="30" style="display:;">
	<td colspan="2" bgcolor="#FFFFFF" style="padding:0px;">
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="0">
		<tr>
			<td width="15%" bgcolor="#DDDDFF">�⺻ ������ :</td>
			<td width="85%" bgcolor="#FFFFFF" style="border-left:1px solid <%= adminColor("tablebg") %>;"><%=FnSelectColorBar("",25)%></td>
		</tr>
		<tr>
			<td width="15%" rowspan="2" bgcolor="#DDDDFF" style="border-top:1px solid <%= adminColor("tablebg") %>;">���� ��ǰ�̹��� :</td>
			<td width="85%" bgcolor="#FFFFFF" style="border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;">
				<input type="file" size="40" name="imgDFColor" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text">
				<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgDFColor, 40, 1000, 1000)"> (����,1000X1000,<b><font color="red">jpg</font></b>)
			</td>
		</tr>
		<tr>
			<td width="85%" bgcolor="#FFFFFF" style="border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;">
		      - ���� �̹����� ������ ����� ���������� ��ǰ �⺻�̹����� ���˴ϴ�.
			</td>
		</tr>
		</table>
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
	<td width="15%" bgcolor="#DDDDFF" rowspan="2">�����Ǹű��� :</td>
	<td width="35%" bgcolor="#FFFFFF">   
		<label><input type="radio" name="limityn" value="N" onClick="this.form.limitno.readOnly=true; this.form.limitno.value=''; this.form.limitno.className='text_ro';document.all.dvDisp.style.display = 'none'; this.form.limitdispyn[0].checked = false; this.form.limitdispyn[1].checked = true;" checked>�������Ǹ�</label>&nbsp;&nbsp;
		<label><input type="radio" name="limityn" value="Y" onClick="this.form.limitno.readOnly=false; this.form.limitno.className='text';document.all.dvDisp.style.display = '';">�����Ǹ�</label>
		<div id="dvDisp" style="display:none;" >
			&nbsp;-> �������⿩��: 
			<input type="radio" name="limitdispyn" value="Y">���� 
			<input type="radio" name="limitdispyn" value="N" checked>�����
		</div>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">�������� :</td>
	<td width="35%" bgcolor="#FFFFFF" >
		<input type="text" name="limitno" maxlength="32" size="8" readonly class="text_ro" id="[off,on,off,off][��������]">(��)
	</td>
</tr>
<tr>
	<td colspan="3" bgcolor="#FFFFFF"><font color="red">** �ɼ��� �ִ°�� �ɼǺ��� ���������� �ϰ� �����˴ϴ�.(���������� ����� ��������)</font></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ּ�/�ִ� �Ǹż� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		�ּ�
		<input type="text" name="orderMinNum" maxlength="5" size="5" class="text" id="[off,on,off,off][�ּ��Ǹż�]" value="1">
		/ �ִ�
		<input type="text" name="orderMaxNum" maxlength="5" size="5" class="text" id="[off,on,off,off][�ִ��Ǹż�]" value="100">
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
		<!--<label><input type="radio" name="usinghtml" value="N"  >�Ϲ�TEXT</label>
		<label><input type="radio" name="usinghtml" value="H" checked>TEXT+HTML</label>
		<label><input type="radio" name="usinghtml" value="Y">HTML���</label>
		<br>
		-->
		<input type="hidden" name="usinghtml" value="Y" />
		<textarea name="itemcontent" rows="18" class="textarea" style="width:100%" id="[on,off,off,off][��ǰ����]"></textarea>
		<script>
		//
		window.onload = new function(){
			var itemContEditor = CKEDITOR.replace('itemcontent',{
				height : 450,
				// ���ε�� ���� ���
				//filebrowserBrowseUrl : '/browser/browse.asp',
				// ���� ���ε� ó�� ������
				filebrowserImageUploadUrl : '<%= ItemUploadUrl %>/linkweb/items/itemEditorContentUpload.asp'
			});
			itemContEditor.on( 'change', function( evt ) {
			    // �Է��� �� textarea ���� ����
			    document.itemreg.itemcontent.value = evt.editor.getData();
			});
		}
		</script>
		<div class="lpad10">
			�� ��ǰ�� ������ �ִ� ����(��)�� 1,000px�Դϴ�.
		</div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ ������ :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="itemvideo" rows="5" class="textarea" cols="90" id="[off,off,off,off][�����۵�����]"></textarea>
	    <br>�� Youtube, Vimeo ������ ����(Youtube : �ҽ��ڵ尪 �Է�, Vimeo : �Ӻ����� �Է�)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ֹ��� ���ǻ��� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="ordercomment" rows="5" cols="90" class="textarea" id="[off,off,off,off][���ǻ���]"></textarea><br>
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
      <input type="file" name="imgbasic" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgbasic,40, 1000, 1000)"> (<font color=red>�ʼ�</font>,1000X1000,<b><font color="red">jpg</font></b>)
      <!-- // ������ // <br><input type="checkbox" name="regimg"> ������̹������ - �̹����� <font color=red>���߿� ���</font>�Ұ�쿡�� ������̹�������� üũ�ϼ���.-->
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">����(����)�̹��� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgmask" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgmask,40, 1000, 1000)"> (����,1000X1000,<b><font color="red">jpg</font></b>)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF" title="�ٹ����ٿ����� ���ε� ������ �⺻�̹��� �Դϴ�.">�ٹ����ٱ⺻�̹��� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgtenten" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgtenten,40, 1000, 1000)"> (����,1000X1000,<b><font color="red">jpg</font></b>)
  	</td>
  </tr>
  <tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���1 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd1" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd1,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���2 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd2" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd2, 40, 1000, 1000)"> (����,1000X1000,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���3 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd3" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd3, 40, 1000, 1000)"> (����,1000X1000,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���4 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd4" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd4, 40, 1000, 1000)"> (����,1000X1000,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���5 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <input type="file" name="imgadd5" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd5, 40, 1000, 1000)"> (����,1000X1000,jpg,gif)
   	</td>
  </tr>
  <tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>�� ������ ��ǰ�����̹����� ������� �ʰ� ��ǰ�����̹����� ����մϴ�. ������ ��ϵ� ��ǰ�����̹����� ����� �ϵ� �߰� ������ �����ʰ� ������ �˴ϴ�.</strong></font>
 	</td>
 </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="imgIn">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">PC��ǰ�����̹��� #1 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40);" class="text" size="40">
      <input type="button" value="#1 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[0],40, 800, 1600)"> (����,800X1600, Max 800KB,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">PC��ǰ�����̹��� #2 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40);" class="text" size="40">
      <input type="button" value="#2 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[1],40, 800, 1600)"> (����,800X1600, Max 800KB,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">PC��ǰ�����̹��� #3 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40);" class="text" size="40">
      <input type="button" value="#3 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[2],40, 800, 1600)"> (����,800X1600, Max 800KB,jpg,gif)
  	</td>
  </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
  	<td bgcolor="#FFFFFF">
      <input type="button" value="PC��ǰ�����̹����߰�" class="button" onClick="InsertImageUp()">
  	</td>
  </tr>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="MobileimgIn">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">����ϻ�ǰ���̹��� #1 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addmoblieimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40);" class="text" size="40">
      <input type="button" value="#1 �̹��������" class="button" onClick="ClearImage2(this.form.addmoblieimgname[0],40, 640, 1200)"> (����,640X1200, Max 400KB,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">����ϻ�ǰ���̹��� #2 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addmoblieimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40);" class="text" size="40">
      <input type="button" value="#2 �̹��������" class="button" onClick="ClearImage2(this.form.addmoblieimgname[1],40, 640, 1200)"> (����,640X1200, Max 400KB,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">����ϻ�ǰ���̹��� #3 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addmoblieimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40);" class="text" size="40">
      <input type="button" value="#3 �̹��������" class="button" onClick="ClearImage2(this.form.addmoblieimgname[2],40, 640, 1200)"> (����,640X1200, Max 400KB,jpg,gif)
  	</td>
  </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>�� ����� ��ǰ�� �̹����� ������ �� �������� ��ü �˴ϴ�. html�� ������� ���� �����̿��� �������� ���ε� ���ֽñ� �ٶ��ϴ�.<br>�� ����� ��ǰ�󼼿��� �̹����� �߶� �÷��ֽñ� �ٶ��ϴ�.</strong></font>
 	</td>
 </tr>
  <tr align="left">
  	<td bgcolor="#FFFFFF">
      <input type="button" value="����ϻ�ǰ���̹����߰�" class="button" onClick="InsertMobileImageUp()">
  	</td>
  </tr>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="�����ϱ�" class="button" onClick="SubmitSave()">
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
<% if (application("Svr_Info")	= "Dev") then %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="300" height="100"></iframe>
<% else %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>

<script type="text/javascript">
	// ��������üũ. ���ȹ�
	jsSafetyCheck('','');
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->