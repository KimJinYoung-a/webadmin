<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������
' Hieditor : 2010.11.15 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim oforecast,i , oforecast_detail ,page , idx , image_url , forecastgubun, link_url, couponidx
dim cardidx ,startdate ,enddate ,isusing ,regdate , contents
	cardidx = requestcheckvar(request("cardidx"),10)
	idx = requestcheckvar(request("idx"),10)
	page = request("page")
	if page = "" then page = 1
			
if cardidx = "" then
	response.write "<script language='javascript'>alert('�ش� ��ȣ�� �����ϴ�');self.close();</script>"
end if
	
'//��
set oforecast_detail = new cforecast_list
	oforecast_detail.frectidx = idx
	
	'//�����ϰ�쿡�� ����
	if idx <> "" then
		oforecast_detail.fcarddetail_oneitem()
	end if
	
	if oforecast_detail.ftotalcount > 0 then
		idx = oforecast_detail.FOneItem.fidx
		cardidx = oforecast_detail.FOneItem.fcardidx
		forecastgubun = oforecast_detail.FOneItem.fforecastgubun
		image_url = oforecast_detail.FOneItem.fimage_url
		contents = oforecast_detail.FOneItem.fcontents
		isusing = oforecast_detail.FOneItem.fisusing
		link_url = oforecast_detail.FOneItem.flink_url
		couponidx = oforecast_detail.FOneItem.fcouponidx
	end if

'// ����Ʈ
set oforecast = new cforecast_list
	oforecast.FPageSize = 20
	oforecast.FCurrPage = page
	oforecast.frectcardidx = cardidx
	oforecast.fcard_detaillist()	
%>

<script language="javascript">

document.domain = "10x10.co.kr";

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
	document.imginputfrm.action='PopImgInput.asp';
	document.imginputfrm.submit();
}

function jsImgDel(divnm,iptNm,vPath){

	window.open('','imgdel','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imgdel';
	document.imginputfrm.action='PopImgInput.asp';
	document.imginputfrm.submit();
}

//����
function reg(){
	if (frm.forecastgubun.value==''){
	alert('ī�屸���� �Է����ּ���');
	frm.forecastgubun.focus();
	return;
	}
	if (frm.contents.value==''){
	alert('�󼼼����� �Է����ּ���');
	frm.contents.focus();
	return;
	}				
	if (frm.isusing.value==''){
	alert('��뿩�θ� �������ּ���');
	return;
	}
	
	frm.action='/admin/momo/forecast/card_process.asp';
	frm.mode.value='detailadd';
	frm.submit();
}

function cardedit(idx){
	frm.idx.value=idx;
	frm.submit();
}
	
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="mode">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��ȣ</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= cardidx %><input type="hidden" name="cardidx" value="<%= cardidx %>">		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>ī���ȣ</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= idx %><input type="hidden" name="idx" value="<%= idx %>">		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>ī�屸��</td>
	<td bgcolor="#FFFFFF" align="left">
		<% if idx = "" then %>
			<% drawforecastgubun "forecastgubun", forecastgubun , ""%>
		<% else %>
			<%= getforecastgubun(forecastgubun) %><input type="hidden" name="forecastgubun" value="<%=forecastgubun%>">
		<% end if %>			
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�󼼼���</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="contents" style="width:450px; height:100px;"><%=contents%></textarea>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�̹���<br>235x309</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('image_urldiv','image_url','image','2000','235','true');"/>		
		<input type="hidden" name="image_url" value="<%= image_url %>">
		<div align="right" id="image_urldiv"><% IF image_url<>"" THEN %><img src="<%=webImgUrl%>/momo/forecast/image/<%= image_url %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>URL</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="link_url" value="<%=link_url%>" size="60">	
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>���ʽ����� ID</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="couponidx" value="<%=couponidx%>" size="5">	
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��뿩��</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="isusing" value="<%=isusing%>">
			<option value="" <% if isusing = "" then response.write " selected" %>>��뿩��</option>
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>			
	</td>
</tr>
<tr align="center" bgcolor="FFFFFF">
	<td colspan=2><input type="button" onclick="reg();" value="����" class="button"></td>
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

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">			
		</td>
		<td align="right">		
			<input type="button" onclick="cardedit('');" value="�űԵ��" class="button">					
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oforecast.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oforecast.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= oforecast.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">	
	<td align="center">��ȣ</td>
	<td align="center">ī���ȣ</td>
	<td align="center">����</td>	
	<td align="center">��뿩��</td>
	<td align="center">���</td>
</tr>
<% for i=0 to oforecast.FresultCount-1 %>

<% if oforecast.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='FFFFaa';>
<% end if %>
	<td align="center">
		<%= oforecast.FItemList(i).fcardidx %>
	</td>
	<td align="center">
		<%= oforecast.FItemList(i).fidx %>
	</td>				
	<td align="center">
		<%= getforecastgubun(oforecast.FItemList(i).fforecastgubun) %>
	</td>		
	<td align="center">
		<%= oforecast.FItemList(i).fisusing %>
	</td>
		
	<td align="center">
		<input type="button" onclick="cardedit(<%= oforecast.FItemList(i).fidx %>);" class="button" value="����">			
	</td>			
</tr>   

<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if oforecast.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= oforecast.StartScrollPage-1 %>&isusing=<%=isusing%>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oforecast.StartScrollPage to oforecast.StartScrollPage + oforecast.FScrollCount - 1 %>
			<% if (i > oforecast.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oforecast.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?page=<%= i %>&isusing=<%=isusing%>>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if oforecast.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>

<%
	set oforecast_detail = nothing
	set oforecast = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->