<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->
<%
'#######################################################
'	History	:  2008.10.23 �ѿ�� ����
'	Description : ���ų�����
'#######################################################
%>
<%
dim mode,diaryid
mode=request("mode")
diaryid= request("diaryid")

dim objDiary ,YearUse
'set objDiary = new organizerCls
'objDiary.getDiaryItem diaryid
'YearUse = "2009"
''YearUse = objDiary.DiaryPrd.FYear
'set objDiary = nothing


dim objInfo ,intLoop
set objInfo = new organizerCls
objinfo.FYearUse = YearUse
objInfo.getDiaryInfo diaryid


%>
<script language="javascript">

function changeselect(id_value,idx){
search_frm.idx.value = +idx;
search_frm.search_order.value = +id_value;
search_frm.submit();
}

//�����߰�
function new_reg(){

	if (newinsert_frm.info_name_newinsert.value ==''){
		alert('������������ �Է����ּ���');
		newinsert_frm.info_name_newinsert.focus();
	}else{
		newinsert_frm.mode_newinsert.value= 'newinsert';
		newinsert_frm.action = "/admin/organizer/lib/organizerregproc.asp";
		newinsert_frm.submit();	
	}
	
}

// ���� ����
function id_delete(infoGubun){
	var aa = confirm('���� �����Ͻðڽ��ϱ�?');
	
	if (aa) {
	newinsert_frm.mode_newinsert.value= 'vardelete';
	newinsert_frm.info_gubun_delete.value = infoGubun;
	newinsert_frm.action = "/admin/organizer/lib/organizerregproc.asp";
	newinsert_frm.submit();	
	}	
}

function delItem(info_idx,info_gubun){
	document.delFrm.info_idx.value=info_idx;
	document.delFrm.info_gubun.value=info_gubun;
	document.delFrm.submit();
}

function checkPageCnt(str) {

	if(str.value < '0' || str.value.length < 0){
		str.value='0';
	}
}

function showimage(img){
	var pop = window.open('viewImage.asp?imageUrl='+img,'imgview','width=600,height=600,resizable=yes');
}

function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,thumb){

	window.open('','imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imginput';
	document.imginputfrm.action='organizer_img_input.asp';
	document.imginputfrm.submit();
}

function jsImgDel(divnm,iptNm,vPath){

	window.open('','imgdel','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imginput';
	document.imginputfrm.action='http://upload.10x10.co.kr/linkweb/organizer/organizer_image_del_proc.asp';
	document.imginputfrm.submit();
}

function subchk(){

	var infoname ='';
	var infogubun ='';
	var infoImage ='';
	var infocnt ='';
	for(var i=0;i<document.regfrm.elements.length;i++){


		if(document.regfrm.elements[i].name.substr(0,9)=="info_name"){
			infoname = infoname + document.regfrm.elements[i].value + ',';
		}
		if(document.regfrm.elements[i].name.substr(0,9)=="infogubun"){
			infogubun = infogubun + document.regfrm.elements[i].value + ',';
		}

		if(document.regfrm.elements[i].name.substr(0,7)=="infoimg"){
			infoImage = infoImage + document.regfrm.elements[i].value + ',';
		}
		if(document.regfrm.elements[i].name.substr(0,12)=="info_pageCnt"){
			infocnt = infocnt + document.regfrm.elements[i].value + ',';
		}
	}
	document.realregfrm.mode.value=document.regfrm.mode.value;
	document.realregfrm.infoname.value=infoname;
	document.realregfrm.infogubun.value= infogubun;
	document.realregfrm.infoImage.value=infoImage;
	document.realregfrm.infocnt.value=infocnt;

	//document.realregfrm.TotalPageName.value	= document.regfrm.TotalPageName.value;
	//document.realregfrm.TotalPagepageCnt.value	= document.regfrm.TotalPagepageCnt.value;
	//document.realregfrm.etcname.value		= document.regfrm.etcname.value;


	document.realregfrm.submit();
}

document.domain="10x10.co.kr"

</script>

<!-- ��� �κ� ���� -->
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#9d9d9d">
	<form name="regfrm" method="post" action="">
	<input type="hidden" name="mode" value="edit">
	<tr bgcolor="#FFFFFF">
		<td align="center">��������</td>
		<td align="center" >�̹���</td>
		<td align="center" >�˻�page<br>����</td>
		<td></td>
		<td align="center">Pages</td>
		<td align="center">����</td>
	</tr>
<%
dim temp
	temp = 0	 
%>
<% if objInfo.FResultCount>0 then %>
	<% for intLoop =0 to objInfo.FResultCount -1 %>
	<%if objInfo.FItemList(intLoop).FinfoGubun<>0  then %>

	<%
	'//�߰� ���������� ���� �������а� �ְ� �����Ѵ�.
	if temp < objInfo.FItemList(intLoop).FinfoGubun then
		temp = objInfo.FItemList(intLoop).FinfoGubun
	end if
	%>
	<tr bgcolor="#FFFFFF">
		<td><input type="text" name="info_name" value="<%= objInfo.FItemList(intLoop).Finfoname %>"></td>
		<td>
			<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('infodiv<%= intLoop %>','infoimg<%= intLoop %>','info','200','400','false');"/>
			<input type="button" class="button" size="30" value="�̹��� ����" onclick="jsImgDel('infodiv<%= intLoop %>','infoimg<%= intLoop %>','info');"/>

			<input type="hidden" name="infogubun" value="<%= objInfo.FItemList(intLoop).FinfoGubun %>">
			<input type="hidden" name="infoimg<%= intLoop %>" value="<%= objInfo.FItemList(intLoop).Finfoimg %>">
		</td>
		<td align="center">
			<%= objInfo.FItemList(intLoop).fsearch_view %>
		</td>
		<td>
			<div align="center" id="infodiv<%= intLoop %>">
				<% if (not isnull(objInfo.FItemList(intLoop).Finfoimg)) and (trim(objInfo.FItemList(intLoop).Finfoimg)<>"") then%>
				<img src="<%= objInfo.FItemList(intLoop).getInfoImgUrl %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= objInfo.FItemList(intLoop).getInfoImgUrl %>');">
				<% end if %>
			</div>
		</td>
		<td><input type="text" name="info_pageCnt" size="3" value="<%= objInfo.FItemList(intLoop).FinfoPageCnt %>"  />��</td>
		<td><input type="button" value="����" class="button" onclick="id_delete('<%=objInfo.FItemList(intLoop).FinfoGubun %>');"></td>
	</tr>

	<% end if %>
	<% next %>



<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="6" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>		

<% end if %>
<% 
		temp = temp +1
%>	
	</form>

	<tr bgcolor="#FFFFFF">
		<td colspan=6>
			<input type="button" class="button" value="Ȯ��" onclick="subchk();"/>&nbsp;&nbsp;
			<input type="button" class="button" value="���" onclick="window.close();"/>
		</td>
	</tr>
</table>
<!-- ��� �κ� ��-->

<br>
<!-- �����߰� ����-->
<table border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#9d9d9d">
<form name="newinsert_frm" method="post" action="">
	<input type="hidden" name="mode_newinsert" value="">
	<input type="hidden" name="Diaryid_newinsert" value="<%= diaryid %>">
	<input type="hidden" name="info_gubun_newinsert" value="<%= temp %>">
	<input type="hidden" name="info_gubun_delete" value="">	
	<tr bgcolor="<%= adminColor("topbar") %>">
		<td >�����߰�</td>
		<td bgcolor="FFFFFF">
			<input type="text" name ="info_name_newinsert" size="30">
		</td>
		<td bgcolor="FFFFFF">
			<input type="button" value="�߰�" class="button" onclick="new_reg();">
		</td>		
	</tr>
		
</form>
</table>
<!-- �����߰� ��-->

<!-- ����Ʈ �˻������� �������� ���Ľ���-->
<%				
dim oip_search,i , a															
	set oip_search = new organizerCls
	oip_search.fsearch_list()
%>
<br>
�� ����Ʈ ���ų����� �˻��������� ��� ���� ���ļ��� �����Դϴ�. ����� �ٷ� ���� �˴ϴ�.<br>
���ڰ� �������� ��ܿ� ����˴ϴ�.

<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#9d9d9d">
<form name="search_frm" method="get" action="search_reg.asp">
<input type="hidden" name="idx">
<input type="hidden" name="search_order">
<input type="hidden" name="Diaryid_search" value="<%= diaryid %>">
<% for i = 0 to oip_search.ftotalcount -1 %>
	<tr bgcolor="#FFFFFF">
		<td><%= oip_search.fitemlist(i).finfo_name %></td>
		<td>
			<select  onchange="javascript:changeselect(this.value,<%=oip_search.fitemlist(i).fidx %>);">
				<% for a = 1 to 50 %>
				<option value="<%=a%>" <% if oip_search.fitemlist(i).fsearch_order = a then response.write " selected"%>><%=a%></option>
				<% next %>
			</select>
		</td>
	</tr>
<% next %>	
</form>
</table>
<!-- ����Ʈ �˻������� �������� ���ĳ�-->

<form name="realregfrm" method="post" action="proc_diary_Info.asp">
<input type="hidden" name="diaryid" value="<%= diaryid %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="infoname" value="">
<input type="hidden" name="infogubun" value="">
<input type="hidden" name="infoImage" value="">
<input type="hidden" name="infocnt" value="">

<input type="hidden" name="TotalPageName" value="">
<input type="hidden" name="TotalPagepageCnt" value="">
<input type="hidden" name="etcname" value="">
</form>
<form name="imginputfrm" method="post" action="">
<input type="hidden" name="YearUse" value="<%= YearUse %>">
<input type="hidden" name="divName" value="">
<input type="hidden" name="orgImgName" value="">
<input type="hidden" name="inputname" value="">
<input type="hidden" name="ImagePath" value="">
<input type="hidden" name="maxFileSize" value="">
<input type="hidden" name="maxFileWidth" value="">
<input type="hidden" name="makeThumbYn" value="">
</form>
<% set objInfo = nothing %>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->