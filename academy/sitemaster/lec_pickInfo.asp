<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �ΰŽ� ����Pick ����
' History : 2012.08.08 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/lecture_pickCls.asp"-->
<%
	dim iLp, sYYYY, sMM, sCDL, sLevel, page, i
	sYYYY = RequestCheckvar(Request("yyyy"),4)
	sMM = RequestCheckvar(Request("mm"),2)
	sCDL = RequestCheckvar(Request("cdl"),3)
	sLevel = RequestCheckvar(Request("level"),10)

	if sYYYY="" then sYYYY=year(date)
	if sMM="" then sMM=Num2Str(Month(date),2,"0","R")

	page = RequestCheckvar(request("page"),10)
	if page="" then page=1

	dim oLecPick
	set oLecPick = new CLecPick
		oLecPick.FCurrPage = page
		oLecPick.FPageSize=20

		oLecPick.FRectCDL = sCDL
		oLecPick.FRectLecLevel = sLevel
		if sYYYY<>"" and sMM<>"" then oLecPick.FRectYYYYMM = sYYYY & sMM

		oLecPick.GetLecPickList
%>
<script type="text/javascript">
function goPage(page){
	document.frm.method="GET";
	document.frm.action="";
	document.frm.page.value= page;
	document.frm.submit();
}

// �ű� ���� ���
function fnNewLecReg() {
	var f = document.frm;

	if(!f.arrLecIdx.value) {
		alert("����Ͻ� ���¹�ȣ�� �Է����ּ���.\n\n�ذ��¹�ȣ�� �޸�(,)�� �����ؼ� �������¸� ���ÿ� ����Ͻ� �� �ֽ��ϴ�.");
		f.arrLecIdx.focus();
		return;
	}

	if(!f.cdl.value) {
		alert("����Ͻ� ī�װ��� �������ּ���.");
		f.cdl.focus();
		return;
	}

	if(!f.level.value) {
		alert("����Ͻ� ������ ���̵��� �������ּ���.");
		f.level.focus();
		return;
	}

	f.mode.value="add";
	f.method="POST";
	f.action="doLecPick.asp";
	f.submit();
}

// ���õ� ���� ����
function fnSelectLecDel() {
	var f = document.frm;
	var l = document.frmList;

	if(l.chkSel=="undefined") return;

	var arrIdx="", chk=false;
	if(!l.chkSel.length) {
		if(l.chkSel.checked) {
			chk=true;
			arrIdx = l.chkSel.value;
		}
	} else {
		for(var i=0;i<l.chkSel.length;i++) {
			if(l.chkSel[i].checked) {
				chk=true;
				if(arrIdx=="") {
					arrIdx = l.chkSel[i].value;
				} else {
					arrIdx += "," + l.chkSel[i].value;
				}
			}
		}
	}
	if(!chk) {
		alert("���õ� ���°� �����ϴ�.");
		return;
	}

	if(confirm("���õ� ���¸� �����Ͻðڽ��ϱ�?\n������ �Ϸ�Ǹ� ���� �� �� �����ϴ�.")) {
		f.mode.value="del";
		f.arrSn.value=arrIdx;
		f.method="POST";
		f.action="doLecPick.asp";
		f.submit();
	}
}
</script>
<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page">
<input type="hidden" name="mode">
<input type="hidden" name="arrSn">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
	<td align="left">
<!--
		ȸ�� :
		<select name="yyyy" class="select">
		<% for iLp=2012 to year(dateadd("yyyy",1,date)) %>
			<option value="<%=iLp%>"><%=iLp%></option>
		<% next %>
		</select>��
		<select name="mm" class="select">
		<% for iLp=1 to 12 %>
			<option value="<%=Num2Str(iLp,2,"0","R")%>"><%=iLp%></option>
		<% next %>
		</select>��
		&nbsp;
-->
		 ī�װ� :
		<select name="cdl" class="select">
			<option value="">::����::</option>
			<option value="10">������</option>
			<option value="20">��ű�</option>
			<option value="30">�ٹ̱�</option>
			<option value="40">������</option>
			<option value="50">�׸���</option>
			<option value="60">����</option>
		</select>
		&nbsp;
		/ ���̵� :
		<select name="level" class="select">
			<option value="">::����::</option>
			<option value="L">�ʱ�</option>
			<option value="M">�߱�</option>
			<option value="H">���</option>
		</select>
		<script type="text/javascript">
			//document.frm.yyyy.value="<%=sYYYY%>";
			//document.frm.mm.value="<%=sMM%>";
			document.frm.cdl.value="<%=sCDL%>";
			document.frm.level.value="<%=sLevel%>";
		</script>
	</td>
	<td width="50" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</table>
<!-- �˻� �� -->
<p>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="���û���" onClick="fnSelectLecDel()">
	</td>
	<td align="right">
		�ű� ���
		<input type="text" name="arrLecIdx" size="60" style="text">
		<input type="button" class="button" value="����" onClick="fnNewLecReg()">
	</td>
</tr>
</table>
</form>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<form name="frmList" style="margin:0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr height="25" bgcolor="#F6F6F6">
	<td colspan="6">
		�˻���� : <b><%= formatNumber(oLecPick.FTotalCount,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%=adminColor("sky")%>">
	<td width="40"></td>
	<td width="100">ī�װ�</td>
  	<td width="100">���̵�</td>
  	<td width="100">���¹�ȣ</td>
  	<td>���¸�</td>
  	<td width="100">�����</td>
</tr>
<%
	if oLecPick.FresultCount>0 then
		for iLp=0 to oLecPick.FResultCount - 1
%>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="chkSel" value="<%=oLecPick.FItemList(iLp).FpickSn%>"></td>
	<td><%=oLecPick.FItemList(iLp).FcdlNm%></td>
  	<td><%=oLecPick.FItemList(iLp).FlecLvName%></td>
  	<td><%=oLecPick.FItemList(iLp).FlecIdx%></td>
  	<td align="left"><%=oLecPick.FItemList(iLp).FlecTitle%></td>
  	<td><%=left(oLecPick.FItemList(iLp).Fregdate,10)%></td>
</tr>
<%		Next %>
<tr height="25" bgcolor="#F6F6F6">
	<td colspan="6" align="center">
	<% if oLecPick.HasPreScroll then %>
		<a href="javascript:goPage('<%= oLecPick.StartScrollPage-1 %>')">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oLecPick.StartScrollPage to oLecPick.FScrollCount + oLecPick.StartScrollPage - 1 %>
		<% if i>oLecPick.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oLecPick.HasNextScroll then %>
		<a href="javascript:goPage('<%= i %>')">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% else %>
<tr height="160" bgcolor="<%=adminColor("pink")%>">
	<td colspan="6" align="center">��ϵ� ���°� �����ϴ�. ����� �űԵ������ ���¸� ������ּ���.</td>
</tr>
<% end if %>
</table>
</form>
<%	set oLecPick = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->