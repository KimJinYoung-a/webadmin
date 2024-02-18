<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ��Ÿ�������>>�����̺�Ʈ
' History : 2015.05.27 ������ ����
'			2023.04.26 �ѿ�� ����(����¡�� �ӽ� ����)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/event/etcsongjangcls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%

dim dateGubun, chkDate
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim fromDate ,toDate, tmpDate
dim fromDate2 ,toDate2
dim onesongjang,i, notfinish, research, page, inputdatetype
dim searchtype, searchvalue, deleteyn,isupchebeasong, gubuncd, jungsanYN
dim isFinish, isInput
	inputdatetype = requestCheckvar(request("inputdatetype"),32)
	'notfinish = requestCheckvar(request("notfinish"),10)
	research = requestCheckvar(request("research"),10)
	page = requestCheckvar(request("page"),10)
	searchtype = requestCheckvar(request("searchtype"),10)
	searchvalue = request("searchvalue")
	deleteyn = requestCheckvar(request("deleteyn"),10)
	gubuncd = requestCheckvar(request("gubuncd"),10)
	isupchebeasong	= requestCheckvar(request("isupchebeasong"),10)
	jungsanYN = requestCheckvar(request("jungsanYN"),10)
	isFinish = requestCheckvar(request("isFinish"),1)
	isInput = requestCheckvar(request("isInput"),1)

    dateGubun   = requestCheckvar(request("dateGubun"),32)
    chkDate   = requestCheckvar(request("chkDate"),32)
	yyyy1   = requestCheckvar(request("yyyy1"),32)
	mm1     = requestCheckvar(request("mm1"),32)
	dd1     = requestCheckvar(request("dd1"),32)
	yyyy2   = requestCheckvar(request("yyyy2"),32)
	mm2     = requestCheckvar(request("mm2"),32)
	dd2     = requestCheckvar(request("dd2"),32)

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())+1), 1)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	tmpDate = DateAdd("d", -1, toDate)
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))

	fromDate2 = fromDate
	toDate2 = toDate
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)
end if

if (research="") and (inputdatetype="") then inputdatetype = "3MBEFORE"
'if (research="") and (notfinish="") then notfinish = "on"
if (research="") and (deleteyn="") then deleteyn = "N"
if (research="") and (isFinish="") then isFinish = "N"
if (research="") and (isInput="") then isInput = "Y"

if page="" then page=1

set onesongjang = new CEventsBeasong
	onesongjang.FPageSize = 1000
	onesongjang.FCurrPage = page
	onesongjang.FRectOnlySongjangNotInput = notfinish
	'onesongjang.FRectOnlyMisend = notfinish
	onesongjang.FRectSearchType = searchtype
	onesongjang.FRectSearchValue = searchvalue
	onesongjang.FRectDeleteyn = deleteyn
	onesongjang.FRectGubuncd = gubuncd
	onesongjang.FRectIsupchebeasong	= isupchebeasong
	onesongjang.FRectJungsanYN = jungsanYN
	onesongjang.FRectinputdatetype = inputdatetype
	onesongjang.FRectIsFinish = isFinish
	onesongjang.FRectIsInput = isInput

	if (chkDate = "Y") then
        onesongjang.FRectDateGubun = dateGubun
		onesongjang.FRectStartdate = fromDate
		onesongjang.FRectEndDate = toDate
	end if

	onesongjang.getEventBeasongInfoList

%>
<script language='javascript'>

function NextPage(page){
	frm.page.value=page;
	frm.submit();
}

function ShowSongJangDetail(frm){
	frm.submit();
}

function AnCheckNSongjangView(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('���� ������ �����ϴ�.');
		return;
	}

	var ret = confirm('���� �������� ���������� �ۼ��Ͻðڽ��ϱ�?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.idarr.value = upfrm.idarr.value + "|" + frm.id.value;
				}
			}
		}
		upfrm.submit();
	}

	upfrm.target = 'popsongjangmaker';
    upfrm.action="/admin/etcsongjang/popsongjangmaker_event.asp"
	upfrm.submit();
}

function saveSongjang(frm){
	if (frm.txsongjang.value.length<1){
		alert('�����ȣ�� �Է��ϼ���');
		frm.txsongjang.focus();
		return;
	}
	frm.action = 'dosongjangmaker_event.asp';
	frm.submit();
}

function delThis(){
	var frm;
	var pass = false;
	var upfrm = document.frmSubmit;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('���� ������ �����ϴ�.');
		return;
	}

	var ret = confirm('���� ������ ������ ���� �Ͻðڽ��ϱ�?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					if (frm.issended.value=="Y" || frm.issended.value==""){
						alert('�߼� �Ϸ���� ���� �ϽǼ� �����ϴ�.');
						frm.cksel.focus();
						return;
					}

					upfrm.idarr.value = upfrm.idarr.value + "|" + frm.id.value;
				}
			}
		}
	}

	upfrm.mode.value = "delarr";
	upfrm.target = 'svc';
    upfrm.action="/admin/etcsongjang/lib/doeventbeasonginfo.asp"
	upfrm.submit();
}

function saveMiChulgo(iid){
    if (confirm('���� ����½� ����ϴ� �޴� �Դϴ�.\n\n ��ȯ �Ͻðڽ��ϱ�?')){
        frmSubmit.action = "dosongjangmaker_event.asp";
        frmSubmit.id.value = iid;
        frmSubmit.mode.value = "michulgo";
        frmSubmit.submit();
    }
}

function AddEtcSongjangAdd(){
	var popwin = window.open('popsongjangadd.asp','popsongjangadd','width=600,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function AddEtcSongjangBatchAdd(){
	var popwin = window.open('popbatchsongjangadd.asp','popbatchsongjangadd','width=1700,height=300,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function AddEtcSongjangBatchAdd_GsType(){
	var popwin = window.open('popbatchsongjangadd_GsType.asp','popbatchsongjangadd','width=600,height=300,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popSongjangFile_GsType(){
    if (frm.sjYYYYMMDD.value.length<10){
        alert('��¥�� �Է��ϼ���.');
        return;
    }

    alert('���� Ŭ�� �ٸ��̸����� ���� -> Excel 97-2003���չ����� �����ϼ���.');
    document.all.svc.src="popChildrenSongjangGsType.asp?sjYYYYMMDD=" + frm.sjYYYYMMDD.value;
}

function EditDeliverInfo(iid){
	var popwin = window.open('popeventsongjangedit.asp?id=' + iid,'popeventsongjangedit','width=600,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="T">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<label>����:
		<select name="gubuncd" class="select">
			<option value="">��ü
			<option value="ev" <% if gubuncd="ev" then response.write "selected" %> >�̺�Ʈ
			<option value="4" <% if gubuncd="4" then response.write "selected" %> >���Ľ����̼�
			<option value="1" <% if gubuncd="1" then response.write "selected" %> >�������ΰŽ�
			<option value="90" <% if gubuncd="90" then response.write "selected" %> >��ǰ
<!--
			<option value="96" <% if gubuncd="96" then response.write "selected" %> >��
			<option value="97" <% if gubuncd="97" then response.write "selected" %> >29cm��
-->
			<option value="98" <% if gubuncd="98" then response.write "selected" %> >����
			<option value="99" <% if gubuncd="99" then response.write "selected" %> >��Ÿ
			<option value="80" <% if gubuncd="80" then response.write "selected" %> >CS���
			<option value="70" <% if gubuncd="70" then response.write "selected" %> >�������
		</select>
		</label>
		&nbsp;&nbsp;
		<label>��������:
		<select name="deleteyn" class="select">
		<option value="">��ü
		<option value="N" <% if deleteyn="N" then response.write "selected" %> >���󳻿�
		<option value="Y" <% if deleteyn="Y" then response.write "selected" %> >��������
		</select>
		</label>
		&nbsp;&nbsp;
		<label>��۱���:
		<select name="isupchebeasong" class="select">
		<option value="">��ü
		<option value="N" <% if isupchebeasong="N" then response.write "selected" %> >�ٹ�
		<option value="Y" <% if isupchebeasong="Y" then response.write "selected" %> >��ü
		</select>
		</label>
		&nbsp;&nbsp;
		<label>���꿩��:
		<select name="jungsanYN" class="select">
		<option value="">��ü
		<option value="Y" <% if jungsanYN="Y" then response.write "selected" %> >������
		<option value="N" <% if jungsanYN="N" then response.write "selected" %> >�������
		</select>
		</label>
		&nbsp;&nbsp;
		<select name="searchtype" class="select">
			<option value="">�˻����� ����</option>
			<option value="eCode" <% if searchtype="eCode" then response.write "selected" %> >�̺�Ʈ�ڵ�</option>
			<option value="username" <% if searchtype="username" then response.write "selected" %> >����</option>
			<option value="reqname" <% if searchtype="reqname" then response.write "selected" %> >�����θ�</option>
			<option value="gubun" <% if searchtype="gubun" then response.write "selected" %> >���и�</option>
			<option value="userid" <% if searchtype="userid" then response.write "selected" %> >���̵�</option>
			<option value="dlvMkrid" <% if searchtype="dlvMkrid" then response.write "selected" %> >��۾�üID</option>
			<option value="songjangno" <% if searchtype="songjangno" then response.write "selected" %> >�����ȣ</option>
		</select>
		<input type="text" name="searchvalue" value="<%= searchvalue %>" size="32" maxlength="32">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="ShowSongJangDetail(frm);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label>�����:
		<select name="isFinish" class="select">
		<option value="">��ü</option>
		<option value="Y" <%=chkIIF(isFinish="Y","selected","")%>>���Ϸ�</option>
		<option value="N" <%=chkIIF(isFinish="N","selected","")%>>�����</option>
		</select>
		</label>
		&nbsp;&nbsp;
		<label>������Է�:
		<select name="isInput" class="select">
		<option value="">��ü</option>
		<option value="Y" <%=chkIIF(isInput="Y","selected","")%>>�Է¿Ϸ�</option>
		<option value="N" <%=chkIIF(isInput="N","selected","")%>>�Է�����</option>
		</select>
		</label>
        &nbsp;
        <label>
		    <input type="checkbox" name="chkDate" value="Y" <% if (chkDate = "Y") then %>checked<% end if %> >
            <select class="select" name="dateGubun">
                <option value="reqDeliverDate" <%= CHKIIF(dateGubun="reqDeliverDate", "selected", "") %>>����û��</option>
                <option value="senddate" <%= CHKIIF(dateGubun="senddate", "selected", "") %>>�����</option>
            </select>
		    <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
        </label>
		<!--
		&nbsp;&nbsp;
		<label><input type="checkbox" name="notfinish" value="on" <%=chkIIF(notfinish="on","checked","")%>>�����</label>
		-->
		&nbsp;&nbsp;
		<label><input type="checkbox" name="inputdatetype" value="3MBEFORE" <%=chkIIF(inputdatetype="3MBEFORE","checked","")%>>�Է�����(3��������)�����Ⱥ���</label>

	</td>
</tr>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="�������" onclick="AddEtcSongjangAdd()">
		&nbsp;
		<input type="button" class="button" value="�ϰ����" onclick="AddEtcSongjangBatchAdd()">
<!--
		<input type="button" value="���̺��ĥ�己 ���" onclick="AddEtcSongjangBatchAdd_GsType()">
		&nbsp;&nbsp;
		<input type="text" name="sjYYYYMMDD" value="<%= Left(Now(),10) %>" maxlength="10" size=10>
		<input type="button" value="ĥ�己����" onclick="popSongjangFile_GsType()">
-->
	</td>
	<td align="right">
		<input type="button" value="���û��׻���" onclick="delThis();" class="button">
		&nbsp;
		<input type="button" value="���û��׼������Ϻ���" onclick="AnCheckNSongjangView();" class="button">
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="17">
		�˻���� : <b><%= onesongjang.FResultCount %>/<%= onesongjang.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= onesongjang.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="30"><input type="checkbox" name="cksel" onClick="AnSelectAllFrame(true);"></td>
	<td width="70">����</td>
	<td>�̺�Ʈ�ڵ�</td>
	<td>�̺�Ʈ��(���и�)</td>
	<td width="70">���̵�</td>
	<td width="60">����</td>
	<td width="60">������</td>
	<td>��ǰ��</td>
	<td width="80">�����</td>
	<td width="80">�����<br>�Է���</td>
	<td width="80">������Է�<br>������</td>
	<td width="50">���<br>����</td>
	<td width="50">��������</td>
	<td width="50">�����</td>
	<td width="100">������ȣ</td>
	<td width="70">����û��</td>
	<td width="70">�����</td>
</tr>
<% if onesongjang.FResultCount>0 then %>
<% for i=0 to onesongjang.FResultCount-1 %>
<form name="frmBuyPrc_<%= onesongjang.FItemList(i).Fid %>" action="post">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="id" value="<%= onesongjang.FItemList(i).Fid %>">

<% if onesongjang.FItemList(i).Fdeleteyn="Y" then %>
	<tr bgcolor="#CCCCCC" >
<% else %>
	<tr bgcolor="#FFFFFF" >
<% end if %>

	<td align="center">
	    <% if IsNULL(onesongjang.FItemList(i).Finputdate) then %>
			<input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  ><!-- ?? �̰� ���ʿ��Ѱ�? �켱 ���� disabled -->
	    <% else %>
			<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
	    <% end if %>
	</td>
	<td align="center"><%=onesongjang.FItemList(i).getEventKind%></td>
	<td align="center"><%=onesongjang.FItemList(i).Fgetcode%></td>
	<td align="center"><a href="javascript:EditDeliverInfo('<%= onesongjang.FItemList(i).Fid %>');"><%= onesongjang.FItemList(i).Fgubunname %></a></td>
	<td align="center"><%= printUserId(onesongjang.FItemList(i).FUserId, 2, "*") %></td>
	<td align="center"><%= onesongjang.FItemList(i).Fusername %></td>
	<td align="center"><%= onesongjang.FItemList(i).FReqName %></td>
	<td align="center"><%= onesongjang.FItemList(i).getPrizeTitle %></td>
	<td align="center"><%= Left(onesongjang.FItemList(i).Fregdate,10) %></td>
	<td align="center">
		<% if IsNULL(onesongjang.FItemList(i).Finputdate) then %>
			<font color="red">(�Է�����)</font>
		<% else %>
			<%= FormatDateTime(onesongjang.FItemList(i).Finputdate,2) %>
		<% end if %>
	</td>
	<td align="center"><%= Left(onesongjang.FItemList(i).Fevtprize_enddate,10) %></td>
	<td align="center">
	    <% if (onesongjang.FItemList(i).FIsUpchebeasong="Y") then %>
			<font color="blue">��ü</font>
	    <% elseif onesongjang.FItemList(i).FIsUpchebeasong="N" then %>
			�ٹ�
	    <% else %>

	    <% end if %>
	</td>
	<td align="center" ><%=onesongjang.FItemList(i).FjungsanYN%></td>
	<td align="right" >
	    <% if NOT isNULL(onesongjang.FItemList(i).Fjungsan) then %>
			<%= FormatNumber(onesongjang.FItemList(i).Fjungsan,0)%>
	    <% end if %>
	</td>
	<td align="center" nowrap>
	    <!--
	        <input type="text" name="txsongjang" value="<%= onesongjang.FItemList(i).FSongjangNo %>" size=12 maxlength=32>
			<input class="button" type="button" value="����" onClick="saveSongjang(frmBuyPrc_<%= onesongjang.FItemList(i).Fid %>)">
		-->
		<%= onesongjang.FItemList(i).Fdivname %><br>
	    <%= onesongjang.FItemList(i).FSongjangNo %>
	</td>
	<td align="center">
		<% if onesongjang.FItemList(i).FreqDeliverDate<> "" then %>
			<%= onesongjang.FItemList(i).FreqDeliverDate %>
		<% end if %>
	</td>
	<td align="center">
		<input type="hidden" name="issended" value="<%= onesongjang.FItemList(i).Fissended %>">

		<% if onesongjang.FItemList(i).Fsenddate <> "" then %>
		    <% = FormatDateTime(onesongjang.FItemList(i).Fsenddate,2) %>
		    <% if (onesongjang.FItemList(i).FIsSended="Y") and (onesongjang.FItemList(i).FIsUpchebeasong<>"Y")  then %>
			<br><input class="button" type="button" value="�����" onClick="saveMiChulgo('<%= onesongjang.FItemList(i).Fid %>')">
			<% end if %>
		<% else %>
			&nbsp;
		<% end if %>
	</td>
</tr>
</form>
<%
if i mod 300 = 0 then
	Response.Flush		' ���۸��÷���
end if

next
%>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="17" align="center">
    	<% if onesongjang.HasPreScroll then %>
			<a href="javascript:NextPage('<%= onesongjang.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + onesongjang.StartScrollPage to onesongjang.FScrollCount + onesongjang.StartScrollPage - 1 %>
			<% if i>onesongjang.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
			<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if onesongjang.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="17" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<form name="frmArrupdate" method="post" >
<input type="hidden" name="idarr" value="">
</form>
<form name="frmSubmit" method="post" action="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="id" value="">
<input type="hidden" name="idarr" value="">
</form>
<iframe id="svc" name="svc" src="" frameborder="0" width="0" height="0" marginwidth="0" marginheight="0" ></iframe>

<%
set onesongjang = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
