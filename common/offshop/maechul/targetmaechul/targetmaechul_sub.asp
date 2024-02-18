<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ��ǥ����
' History : 2013.03.15 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/maechul/targetmaechul/targetmaechul_cls.asp"-->

<!--�����¿쿩��0, ���� ���������� �θ�â ���� ���� 100% ����-->
<style type="text/css">
	<!-- 
	body {background-color:transparent; filter: Alpha(Opacity=100); margin-left: 0px; margin-top: 0px; margin-right: 0px; margin-bottom: 0px;}
	-->
</style>

<%
Dim menupos, page ,ctarget , yyyy1 ,mm1 ,i , shopid ,gubuntype
	menupos = requestcheckvar(request("menupos"),10)
	page = requestcheckvar(request("page"),10)
	gubuntype = requestcheckvar(request("gubuntype"),10)
	shopid = requestcheckvar(request("shopid"),32)
	yyyy1 = requestcheckvar(request("yyyy1"),4)
	mm1 = requestcheckvar(request("mm1"),4)
	
	if yyyy1 = "" then yyyy1 = year(now())
	if mm1 = "" then mm1 = month(now())	
	if page = "" then page = 1
	if gubuntype = "" then gubuntype = "1"

	if gubuntype = "" or shopid = "" or yyyy1 = "" or mm1 = "" then
		response.end : dbget.close()
	end if
	
set ctarget = new ctargetmaechul_list
	ctarget.FRectyyyy1 = yyyy1
	ctarget.FRectmm1 = Format00(2,mm1)
	ctarget.frectshopid = shopid
	ctarget.frectgubuntype = gubuntype
	ctarget.FPageSize = 500
	ctarget.FCurrPage = page
	
	if shopid <> "" then
		
		'/���׺���ǥ������ ���
		if gubuntype = "2" then
			ctarget.ftarget_zone
		
		'/�Ϲ� ��ǥ ����
		else
			ctarget.ftarget
		end if
	end if
%>

<script language="javascript">

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

//��ü ����
function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

//���û�ǰ �۳��� ��ǥ���� ���
function chmaechul(){

	var searchfrm = document.frm;
	
	if (!IsDigit(searchfrm.maechulpro.value)){
		alert('��ǥ��� %�� ���ڸ� �Է� �����մϴ�.');
		searchfrm.maechulpro.focus();
		return;
	}
	
	if (searchfrm.maechulpro.value<1){
		alert('��ǥ��� %�� ��Ȯ�� �Է��ϼ���.');
		searchfrm.maechulpro.focus();
		return;
	}
	
	var frm;
	var pass = false;
	
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}
				
	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				frm.targetmaechul.value = Math.round(frm.realsellsum.value * (searchfrm.maechulpro.value/100));
			}
		}
	}
}

//���û�ǰ ����
function saveArr(){

	var searchfrm = document.frm;
	
	if (searchfrm.shopid.value==''){
		alert('������ ���õ��� �ʾҽ��ϴ�');
		return;
	}
	
	var frm;
	var pass = false;
	
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}
				
	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	frmarr.mode.value = "";
	frmarr.yyyymm.value = "";
	frmarr.shopid.value = "";
	frmarr.gubuntype.value = "";
	frmarr.gubun.value ="";
	frmarr.targetmaechul.value ="";
	 
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				if (!IsDigit(frm.targetmaechul.value)){
					alert('��ǥ������ ���ڸ� �����մϴ�.');
					frm.targetmaechul.focus();
					return;
				}
				
				if (frm.targetmaechul.value<1){
					alert('��ǥ������ ��Ȯ�� �Է��ϼ���.');
					frm.targetmaechul.focus();
					return;
				}
				
				frmarr.yyyymm.value = frmarr.yyyymm.value + frm.yyyymm.value + ","
				frmarr.gubun.value = frmarr.gubun.value + frm.gubun.value + ","				
				frmarr.targetmaechul.value = frmarr.targetmaechul.value + frm.targetmaechul.value + ","

			}
		}
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frmarr.mode.value = 'tmreg';
		frmarr.shopid.value = '<%=shopid%>';
		frmarr.yyyy1.value = '<%=yyyy1%>';
		frmarr.mm1.value = '<%=mm1%>';
		frmarr.gubuntype.value = searchfrm.gubuntype.value;
		frmarr.submit();
	}
}

function frmsubmit(){
	frm.submit();
}

</script>

<Br>
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="shopid" value="<%=shopid%>">
<input type="hidden" name="yyyy1" value="<%=yyyy1%>">
<input type="hidden" name="mm1" value="<%=mm1%>">
<input type="hidden" name="gubuntype" value="<%=gubuntype%>">
<tr colspan=2>
	<td align="left">
     	�� <font color="red"><%= fnGetCommonCode("gubuntype",gubuntype) %></font>
	</td>
</tr>	
<tr>
	<td align="left">
		<% IF ctarget.fresultcount > 0 THEN %>
			��ǥ���� = �۳������ <input type="text" name="maechulpro" value="0" size=5 maxlength=6>%
			<input type="button" value="���ð��" onClick="chmaechul();" class="button">
		<% end if %>		
	</td>
	<td align="right">
		<% IF ctarget.fresultcount > 0 THEN %>
			<input type="button" value="���ü���" onClick="saveArr()" class="button">
		<% end if %>
	</td>	
</tr>
</form>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="left">
		�˻���� : <b><%=ctarget.ftotalcount%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>
		<input type="checkbox" name="ck_all" onclick="SelectCk(this)">
	</td>
	<td>
		����
	</td> 	
	<td><%=yyyy1-1%>-<%=mm1%><br>����</td>
	<td><%=yyyy1%>-<%=mm1%><br>��ǥ����</td>
	<td>��ǥ<br>��������</td>
	<td>����<br>����</td>
</tr>	
<% 
IF ctarget.fresultcount > 0 THEN
	
For i = 0 To ctarget.fresultcount -1
%>
<form name="frmBuyPrc_<%=i%>" method="get">
<input type="hidden" name="gubun" value="<%= ctarget.FItemList(i).fgubun %>">
<input type="hidden" name="realsellsum" value="<%= ctarget.FItemList(i).frealsellsum %>">
<input type="hidden" name="yyyymm" value="<%= ctarget.FItemList(i).fyyyymm %>">
<tr align="center" bgcolor=<% IF ctarget.FItemList(i).fyyyymm = "" or isnull(ctarget.FItemList(i).fyyyymm) THEN %>"#f1f1f1"<%ELSE%>"#FFFFFF"<%END IF%>> 			
    <td width=30><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
    <td align="left">
    	<%= ctarget.FItemList(i).fgubunname %>
    </td>      
    <td width=120 align="right">
    	<%= FormatNumber(ctarget.FItemList(i).frealsellsum,0) %>
    </td>      
    <td width=120 align="right">
    	<input type="text" name="targetmaechul" onKeyup="CheckThis(frmBuyPrc_<%= i %>);" value="<%= ctarget.FItemList(i).ftargetmaechul %>" size="12" maxlength="13" style="text-align:right;">
    </td>
    <td width=60>
    	<% IF ctarget.FItemList(i).fyyyymm = "" or isnull(ctarget.FItemList(i).fyyyymm) THEN %>
    		N
    	<% else %>
			Y    	
    	<% end if %>
    </td>          
    <td width=200>
    	<%= ctarget.FItemList(i).flastadminid %>
    	<Br><%= ctarget.FItemList(i).flastupdate %>
    </td>	    
</tr>
</form>
<% next %>

<% ELSE %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">��ϵ� ������ �����ϴ�.</td>
</tr>	
<% END IF %>
<form name="frmarr" method="post" action="/common/offshop/maechul/targetmaechul/targetmaechul_process.asp">
	<input type="hidden" name="mode">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="yyyy1">
	<input type="hidden" name="mm1">
	<input type="hidden" name="yyyymm">
	<input type="hidden" name="shopid">
	<input type="hidden" name="gubuntype">
	<input type="hidden" name="gubun">
	<input type="hidden" name="targetmaechul">
</form>		    
</table>

<%
set ctarget = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->