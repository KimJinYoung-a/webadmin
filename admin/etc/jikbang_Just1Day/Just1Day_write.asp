<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/etc/jikbang_just1DayCls.asp"-->
<%
'###############################################
' PageName : Just1Day_write.asp
' Discription : ����Ʈ ������ ���/����
' History : 2008.04.08 ������ ����
'           2012.02.15 ������ : �̴ϴ޷� ��ü
'           2014.09.12 ������ : ���� �����̿����� ���� �ɿ��� ����
'###############################################

dim justDate,mode,i
mode=request("mode")
justDate=request("justDate")

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
<!--

document.domain = "10x10.co.kr";

function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,thumb,orgImgName){

	window.open("Just1Day_PopImgInput.asp?divName="+divnm+"&inputname="+iptNm+"&ImagePath="+vPath+"&maxFileSize="+Fsize+"&maxFileWidth="+Fwidth+"&makeThumbYn="+thumb+"&orgImgName="+orgImgName,'imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
}

function editcont(){	
    //���µ��� ���� ������ ���;;
    var frm=document.inputfrm;
    
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.sale_code.value="";
        frm.submit();
    }
    
}

function subcheck(){
	var frm=document.inputfrm;

	if(!frm.justDate.value) {
		alert("������ ��¥�� �������ּ���!");
		return;
	} else {
		if(frm.justDate.value<='<%=date%>') {
			//alert("��ǰ�� ����/����� ���� ������ ��¥�� �����մϴ�.");
			//return;
		}
	}

	if(!frm.itemid.value) {
		alert("����� ��ǰ�� �������ּ���!");
		return;
	}

	if(!frm.salePrice.value) {
		alert("��ǰ�� ���αݾ��� �Է����ּ���!");
		frm.salePrice.focus();
		return;
	} else {
		if(parseInt(frm.salePrice.value)>=parseInt(frm.orgPrice.value)) {
			alert("�ǸŰ����� ���ξ��� ũ�ų� ���� ���� �����ϴ�.\n���ξ��� Ȯ�����ּ���.");
			return;
		}
	}

    // ���ξ�0,���Ծ�0 �Է°����ϰ� ����
	if ((!frm.saleSuplyCash.value||frm.saleSuplyCash.value=="0")&&(frm.salePrice.value!="0")) {
		alert("��ǰ�� ���Աݾ��� �Է����ּ���!\n�ظ��Ա޾��� �ݵ�� �������԰��� �Է��ؾߵ˴ϴ�.");
		frm.saleSuplyCash.focus();
		return;
	}
    
    // ���԰��� �����ǸŰ� ���� Ŭ �� ����
    if (frm.saleSuplyCash.value*1>frm.salePrice.value*1) {
		alert("��ǰ�� ���Աݾ��� �Է����ּ���!\n�ظ��Ա޾��� �Ǹ� �ݾ� ���� Ŭ �� �����ϴ�.");
		frm.saleSuplyCash.focus();
		return;
	}
	
	if(!frm.limitNo.value) {
		alert("�������� �Ǹ��� ������ �Է����ּ���.\n\n�� �����ǸŰ� �ƴ϶�� 0�� �Է����ּ���.");
		frm.limitNo.focus();
		return;
	}

	if(frm.justDesc.value.length<=0||frm.justDesc.value.length>=240) {
		alert("��ǰ�� ���� ������ 240���̳�(4�� �̳�)�� �ۼ����ּ���.\n\n");
		frm.justDesc.focus();
		return;
	}
    
    //eastone �߰� �ǸŰ�0,���԰�0 ���ε�� ����.
    if ((frm.salePrice.value=="0")&&(frm.saleSuplyCash.value=="0")){
        if (!confirm('�����ǸŰ� 0, ���θ��԰� 0���� ��Ͻ� ���� ���� �ʽ��ϴ�. ����Ͻðڽ��ϱ�?')){
            return;
        }
    }
    
	if(frm.mode.value=="add"&&frm.itemOptCnt.value>0&&frm.limitNo.value>0) {
		if(confirm("�ɼ��� �����ϴ� ��ǰ�� �Դϴ�.\n�Է��Ͻ� ���������� �ɼǿ� �ڵ����� �ݿ����� �����Ƿ�, ���� ���� ��ǰ�������� �ɼ� ���������� ���� �Է��ϼž� ���������� �ǸŰ� �����մϴ�.")) {
			frm.submit();
		} else {
			return;
		}
	} else {
		frm.submit();
	}
}

function popItemWindow(tgf){
	var popup_item = window.open("/common/pop_singleItemSelect.asp?target=" + tgf + "&ptype=just1day", "popup_item", "width=800,height=500,scrollbars=yes,status=no");
	popup_item.focus();
}

function putPercent(){
	var pct, frm = document.inputfrm;
	if(frm.orgPrice.value==0||frm.salePrice.value==0) {
		pct = "0%";
	}
	else {
		pct = 1 - (frm.salePrice.value / frm.orgPrice.value);
		pct = pct * 100;
		pct = Math.round(pct*100) / 100
		pct = pct + "%";
	}
	frm.saleRate.value= pct;
}

function delitems(){
	var frm = document.inputfrm;
	if (confirm('�� �������� �����Ͻðڽ��ϱ�?\n\n�����ο��� ������ �Բ� �����˴ϴ�.')) {
		frm.mode.value="delete";
		frm.submit();
	}
}
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="doJust1Day_Process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="<% =mode %>">
<input type="hidden" name="itemOptCnt" value="0">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<font color="red"><b>���� ����Ʈ ������ ���/����</b></font>
	</td>
</tr>
<% if mode="add" then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��¥</td>
	<td bgcolor="#FFFFFF">
		<input id="justDate" name="justDate" class="text" size="10" maxlength="10" />
		<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="justDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "justDate", trigger    : "justDate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ǰ</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="itemid" value="" size="10" readonly>
		<input type="button" class="button" value="ã��" onClick="popItemWindow('inputfrm')">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
		���αݾ� <input type="text" class="text" name="salePrice" value="" size="10" style="text-align:right" onkeyup="putPercent()">��
		/ �ǸŰ� <input type="text" class="text_ro" name="orgPrice" value="0" size="8" readonly style="text-align:right">��,
		������ <input type="text" class="text_ro" name="saleRate" value="0%" size="4" readonly style="text-align:center">
		<br>���Աݾ� <input type="text" class="text" name="saleSuplyCash" value="" size="8" style="text-align:right">��
		<br>(���� ���Ұ�� ���αݾ�0, ���Աݾ� 0 �Է�)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="limitNo" value="0" size="4" style="text-align:right">
		<input type="hidden" name="limitYn" value="">
		(�������� 0���� ������ ������ ���� �Ǹŵ˴ϴ�.)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
		<textarea name="justDesc" class="textarea" cols="80" rows="3"></textarea>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>��� �̹���1</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('image1div','image1','i1','250','100','false','');"/>
		<input type="hidden" name="image1" value="">
		<div align="right" id="image1div"></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>��� �̹���2</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('image2div','image2','i2','600','450','true','');"/>		
		<input type="hidden" name="image2" value="">
		<div align="right" id="image2div"></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>���̹���</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('image3div','image3','i2','600','450','true','');"/>		
		<input type="hidden" name="image3" value="">
		<div align="right" id="image3div"></div>
	</td>
</tr>
<% elseif mode="edit" then %>
<%
	dim fmainitem
	set fmainitem = New Cjust1Day
	fmainitem.FCurrPage = 1
	fmainitem.FPageSize=1
	fmainitem.FRectDate=justDate
	fmainitem.Getjust1Daymodify
%>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��¥</td>
	<td bgcolor="#FFFFFF">
		<b><%=fmainitem.FItemList(0).FjustDate%></b>
		<input type="hidden" name="justDate" value="<%=fmainitem.FItemList(0).FjustDate%>">
		<input type="hidden" name="sale_code" value="<%=fmainitem.FItemList(0).Fsale_code%>">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ǰ</td>
	<td bgcolor="#FFFFFF">
		<%= "[" & fmainitem.FItemList(0).Fitemid & "] " & fmainitem.FItemList(0).Fitemname %>
		<input type="hidden" name="itemid" value="<%=fmainitem.FItemList(0).Fitemid%>">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
		���αݾ� <input type="text" class="text" name="salePrice" value="<%= fmainitem.FItemList(0).FjustSalePrice %>" size="10" style="text-align:right" onkeyup="putPercent()">��
		/ �ǸŰ� <input type="text" class="text_ro" name="orgPrice" value="<%= fmainitem.FItemList(0).ForgPrice %>" size="8" readonly style="text-align:right">��,
		������ <input type="text" class="text_ro" name="saleRate" value="<%= FormatPercent(1-(fmainitem.FItemList(0).FjustSalePrice/fmainitem.FItemList(0).ForgPrice),2) %>" size="5" readonly style="text-align:center">
		<br>���Ա޾� <input type="text" class="text" name="saleSuplyCash" value="<%= fmainitem.FItemList(0).FsaleSuplyCash %>" size="8" style="text-align:right">��
		<br>(���� ���Ұ�� ���αݾ�0, ���Աݾ� 0 �Է�)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="limitNo" value="<%= fmainitem.FItemList(0).FlimitNo %>" size="4" style="text-align:right">
		<input type="hidden" name="limitYn" value="">
		(�������� 0���� ������ ������ ���� �Ǹŵ˴ϴ�.)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
		<textarea name="justDesc" class="textarea" cols="80" rows="3"><%= fmainitem.FItemList(0).FjustDesc %></textarea>
		<input type="button" value=" ���� ���� " class="button" onclick="editcont();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>��� �̹��� 1</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('image1div','image1','i1','250','100','false','<%= fmainitem.FItemList(0).Fimg1 %>');"/>
		<input type="hidden" name="image1" value="<%= fmainitem.FItemList(0).Fimg1 %>">
		<div align="right" id="image1div"><% IF fmainitem.FItemList(0).Fimg1<>"" THEN %><img src="<%= fmainitem.FItemList(0).Fimg1 %>" width=50 height=50 ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>��� �̹��� 2</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('image2div','image2','i2','600','450','true','<%= fmainitem.FItemList(0).Fimg2 %>');"/>
		<input type="hidden" name="image2" value="<%= fmainitem.FItemList(0).Fimg2 %>">
		<div align="right" id="image2div"><% IF fmainitem.FItemList(0).Fimg2<>"" THEN %><img src="<%= fmainitem.FItemList(0).Fimg2 %>" width=50 height=50 ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>�� �̹���</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('image3div','image3','i2','600','450','true','<%= fmainitem.FItemList(0).Fimg3 %>');"/>
		<input type="hidden" name="image3" value="<%= fmainitem.FItemList(0).Fimg3 %>">
		<div align="right" id="image3div"><% IF fmainitem.FItemList(0).Fimg3<>"" THEN %><img src="<%= fmainitem.FItemList(0).Fimg3 %>" width=50 height=50 ><% End IF %></div>
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" ���� " class="button" onclick="subcheck();"> &nbsp;&nbsp;
		<% if mode="edit" then %><input type="button" value=" ���� " class="button" onclick="delitems();"> &nbsp;&nbsp;<% end if %>
		<input type="button" value=" ��� " class="button" onclick="history.back();">
	</td>
</tr>
</form>
</table>

<form name="imginputfrm" method="post" action="">
	<input type="hidden" name="divName" value="">
	<input type="hidden" name="orgImgName" value="">
	<input type="hidden" name="inputname" value="">
	<input type="hidden" name="ImagePath" value="">
	<input type="hidden" name="maxFileSize" value="">
	<input type="hidden" name="maxFileWidth" value="">
	<input type="hidden" name="makeThumbYn" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
