<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : showbanner_insert.asp
' Discription : ����� showbanner
' History : 2014.03.13 ����ȭ
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/showBannerCls.asp" -->
<%
	Dim idx , listimg , state , reservationdate , stitle , viewtext , playcate
	Dim viewno , textimg , worktext , partMDid,partWDid
	Dim simg1, simg2 ,simg3 ,simg4 ,simg5
	Dim salt1, salt2 ,salt3 ,salt4 ,salt5
	Dim surl1, surl2 ,surl3 ,surl4 ,surl5
	Dim oShowbanner , mainTopBGColor , subtitle
	idx = request("idx")	
	'//db 1row
	set oShowbanner = new CShowBannerContents
		 oShowbanner.FRectIdx = idx
		
		if idx <> "" Then
			oShowbanner.GetOneRowShowBanner()

			if oShowbanner.FResultCount > 0 then			
				simg1					= oShowbanner.FOneItem.Fsimg1
				simg2					= oShowbanner.FOneItem.Fsimg2
				simg3					= oShowbanner.FOneItem.Fsimg3
				simg4					= oShowbanner.FOneItem.Fsimg4
				simg5					= oShowbanner.FOneItem.Fsimg5
				stitle					= oShowbanner.FOneItem.Fstitle
				reservationdate			= oShowbanner.FOneItem.Freservationdate
				state					= oShowbanner.FOneItem.Fstate
				worktext				= oShowbanner.FOneItem.Fworktext
				partMDid				= oShowbanner.FOneItem.FpartMDid
				partWDid				= oShowbanner.FOneItem.FpartWDid
				salt1					= oShowbanner.FOneItem.Fsalt1
				salt2					= oShowbanner.FOneItem.Fsalt2
				salt3					= oShowbanner.FOneItem.Fsalt3
				salt4					= oShowbanner.FOneItem.Fsalt4
				salt5					= oShowbanner.FOneItem.Fsalt5
				surl1					= oShowbanner.FOneItem.Fsurl1
				surl2					= oShowbanner.FOneItem.Fsurl2
				surl3					= oShowbanner.FOneItem.Fsurl3
				surl4					= oShowbanner.FOneItem.Fsurl4
				surl5					= oShowbanner.FOneItem.Fsurl5
				mainTopBGColor			= oShowbanner.FOneItem.Fcolorcode
				viewno					= oShowbanner.FOneItem.Fviewno
				subtitle				= oShowbanner.FOneItem.Fsubtitle
			end if	
		end if
	set oShowbanner = Nothing

	Dim oSubItemList
	set oSubItemList = new CShowBannerContents
		oSubItemList.FPageSize = 100
		oSubItemList.FRectIdx = idx
		If idx <> "" then
			oSubItemList.GetContentsItemList()
		End If 
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/js/jqueryui/css/evol.colorpicker.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/evol.colorpicker.min.js"></script>
<script type="text/javascript">
	function jsgolist(){
		self.location.href="/admin/mobile/showbanner/";
	}

	function jsPopCal(sName){
		var winCal;

		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//�̹��� Ȯ��ȭ�� ��â���� �����ֱ�
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/admin/sitemaster/play/lib/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}

	function jsSetImg(sImg, sName, sSpan){	
		document.domain ="10x10.co.kr";
		
		var winImg;
		winImg = window.open('/admin/mobile/lib/pop_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	function subcheck(){
		var frm=document.inputfrm;

		if (!frm.stitle.value){
			alert('�������� ������ּ���');
			frm.stitle.focus();
			return;
		}

		if (!frm.subtitle.value){
			alert('���������� ������ּ���');
			frm.stitle.focus();
			return;
		}

		if (!frm.worktext.value){
			alert('�۾������� ������ּ���');
			frm.worktext.focus();
			return;
		}

		if (!frm.reservationdate.value){
			alert('���¿������� ������ּ���');
			frm.reservationdate.focus();
			return;
		}

		if(!frm.state.value){
			alert("���¸� �������ּ���");
			frm.state.focus();
			return;
		}

		if(!frm.partmdid.value){
			alert("��� MD�� �����ϼ���.");
			frm.partmdid.focus();
			return;
		}

		if(!frm.partwdid.value){
			alert("��� WD�� �����ϼ���.");
			frm.partwdid.focus();
			return;
		}

		frm.submit();
	}

	function putLinkText(key,gubun) {
		var frm = document.inputfrm;
		var urllink
		switch(gubun) {
			case '1':
				urllink = frm.surl1;
				break;
			case '2':
				urllink = frm.surl2;
				break;
			case '3':
				urllink = frm.surl3;
				break;
			case '4':
				urllink = frm.surl4;
				break;
			case '5':
				urllink = frm.surl5;
				break;
		}

		switch(key) {
			case 'search':
				urllink.value='/search/search_result.asp?rect=�˻���';
				break;
			case 'event':
				urllink.value='/event/eventmain.asp?eventid=�̺�Ʈ��ȣ';
				break;
			case 'itemid':
				urllink.value='/category/category_itemprd.asp?itemid=��ǰ�ڵ�';
				break;
			case 'category':
				urllink.value='/category/category_list.asp?cdl=ī�װ�';
				break;
			case 'brand':
				urllink.value='/street/street_brand.asp?makerid=�귣����̵�';
				break;
		}
	}

	$(function(){
		//������ư
		$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");

		$( "#subList" ).sortable({
			placeholder: "ui-state-highlight",
			start: function(event, ui) {
				ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
			},
			stop: function(){
				var i=99999;
				$(this).parent().find("input[name^='sort']").each(function(){
					if(i>$(this).val()) i=$(this).val()
				});
				if(i<=0) i=1;
				$(this).parent().find("input[name^='sort']").each(function(){
					$(this).val(i);
					i++;
				});
			}
		});
		//�÷���Ŀ
		$("input[name='mainTopBGColor']").colorpicker();
	});
	//����
	function popSubEdit(subidx) {
	<% if idx <>"" then %>
		var popwin = window.open('popSubItemEdit.asp?listIdx=<%=idx%>&subIdx='+subidx,'popTemplateManage','width=800,height=300,scrollbars=yes,resizable=yes');
		popwin.focus();
	<% else %>
		alert("���ø� ������ ������ ���� ������ּ���.");
	<% end if %>
	}

	// ��ǰ�˻� �ϰ� ���
	function popRegSearchItem() {
	<% if idx <> "" then %>
		var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/mobile/showbanner/doSubRegItemCdArray.asp?listidx=<%=idx%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		popwin.focus();
	<% else %>
		alert("���ø� ������ ������ ���� ������ּ���.");
	<% end if %>
	}

	// ��ǰ�ڵ� �ϰ� ���
	function popRegArrayItem() {
	<% if idx<>"" then %>
		var popwin = window.open('popSubRegItemCdArray.asp?listIdx=<%=idx%>','popRegArray','width=600,height=300,scrollbars=yes,resizable=yes');
		popwin.focus();
	<% else %>
		alert("���ø� ������ ������ ���� ������ּ���.");
	<% end if %>
	}

	function chkAllItem() {
		if($("input[name='chkIdx']:first").attr("checked")=="checked") {
			$("input[name='chkIdx']").attr("checked",false);
		} else {
			$("input[name='chkIdx']").attr("checked","checked");
		}
	}

	function saveList() {
		var chk=0;
		$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
			if($(this).attr("checked")) chk++;
		});
		if(chk==0) {
			alert("�����Ͻ� ���縦 �������ּ���.");
			return;
		}
		if(confirm("�����Ͻ� ����� ���� ������ �����Ͻðڽ��ϱ�?")) {
			document.frmList.action="doListModify.asp";
			document.frmList.submit();
		}
	}
</script>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="showbannerProc.asp">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="simg1" value="<%=simg1%>">
<input type="hidden" name="simg2" value="<%=simg2%>">
<input type="hidden" name="simg3" value="<%=simg3%>">
<input type="hidden" name="simg4" value="<%=simg4%>">
<input type="hidden" name="simg5" value="<%=simg5%>">
<tr height="30">
	<td colspan="4" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>Mobile&gt;&gt;Showbanner ���/����</b></font><br/><br/>
		�� ���̹��� ������ �ִ� 5����� �����ϸ� �̹��� ������ ��!!!! �������� ��� ���ּž� �մϴ�.��<br/>
		�� ������ ��� ��ư�� �̹��� ������ �����˴ϴ� ��
	</td>
</tr>
<% If idx <> "0" Then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">idx</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<b><%=idx%></b>
	</td>
</tr>
<% End If %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���ļ���</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text" name="viewno" value="<%=chkiif(viewno="","0",viewno)%>" size="5"/>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���� ����</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text" name="stitle" value="<%=stitle%>" size="50"/>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���� ����</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text" name="subtitle" value="<%=subtitle%>" size="50"/>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
	<td colspan="3" bgcolor="#FFFFFF">
   		<%IF state = "9" THEN%>
   			<%=reservationdate%><input type="hidden" name="reservationdate" size=20 maxlength=10 value="<%=reservationdate%>"/>
   		<%ELSE%>
   			<input type="text" name="reservationdate" size=20 maxlength=10 value="<%=reservationdate%>" onClick="jsPopCal('reservationdate');"  style="cursor:pointer;"/>
   		<%END IF%>
		��) (<%=Left(Now(),10)%>)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<% Draweventstate2 "state" , state ,"" %> �� ������ �ؼ� �����Ͽ��� ������ =< ���� �̾�߸� ������ �˴ϴ�.
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����MD</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<% sbGetpartid "partmdid",partmdid,"","11" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">���WD</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<% sbGetpartid "partwdid",partwdid,"","12" %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�۾� ���� ����</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<textarea name="worktext" rows="8" cols="80"><%=worktext%></textarea>
	</td>
</tr>
<% If idx <> "0" Then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̹��� �� �κ� �÷�</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" name="mainTopBGColor" value="<%=mainTopBGColor%>" class="text" style="width:80px;" /><br>�� ����� �÷� �ڵ带 ���� �ֽ� ��� #�� �� �ٿ��ּ���.
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="10%">���̹��� 1</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="button" name="btnsimg1" value="�̹������" onClick="jsSetImg('<%=simg1%>','simg1','simgdiv1')" class="button"/>
		<div id="simgdiv1" style="padding: 5 5 5 5">
			<%IF simg1 <> "" THEN %>			
				<img src="<%=simg1%>" border="0" height=100 onclick="jsImgView('<%=simg1%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('simg1','simgdiv1');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				�̹��� �ּ� : <%=simg1%>
			<%END IF%>
		</div>
	</td>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̹�����ũ&����</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="#FFFFFF">
			<tr>
				<td>ALT : <input type="text" class="text" name="salt1" value="<%=salt1%>" size="40"/></td>
			</tr>
			<tr>
				<td>
					URL : <input type="text" class="text" name="surl1" value="<%=surl1%>" size="80"/><br/><br/>
					<div style="padding-left:25px;">
					<font color="#707070">
					- <span style="cursor:pointer" onClick="putLinkText('search','1')">�˻���� ��ũ : /search/search_result.asp?rect=<font color="darkred">�˻���</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('event','1')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('itemid','1')">��ǰ�ڵ� ��ũ : /category/category_itemprd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('category','1')">ī�װ� ��ũ : /category/category_list.asp?cdl=<font color="darkred">ī�װ�</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('brand','1')">�귣����̵� ��ũ : /street/street_brand.asp?makerid=<font color="darkred">�귣����̵�</font></span>
					</font>
					</div>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���̹��� 2</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnsimg2" value="�̹������" onClick="jsSetImg('<%=simg2%>','simg2','simgdiv2')" class="button"/>
		<div id="simgdiv2" style="padding: 5 5 5 5">
			<%IF simg2 <> "" THEN %>			
				<img src="<%=simg2%>" border="0" height="100" onclick="jsImgView('<%=simg2%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('simg2','simgdiv2');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				�̹��� �ּ� : <%=simg2%>
			<%END IF%>
		</div>
	</td>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̹�����ũ&����</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="#FFFFFF">
			<tr>
				<td>ALT : <input type="text" class="text" name="salt2" value="<%=salt2%>" size="40"/></td>
			</tr>
			<tr>
				<td>
					URL : <input type="text" class="text" name="surl2" value="<%=surl2%>" size="80"/><br/><br/>
					<div style="padding-left:25px;">
					<font color="#707070">
					- <span style="cursor:pointer" onClick="putLinkText('search','2')">�˻���� ��ũ : /search/search_result.asp?rect=<font color="darkred">�˻���</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('event','2')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('itemid','2')">��ǰ�ڵ� ��ũ : /category/category_itemprd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('category','2')">ī�װ� ��ũ : /category/category_list.asp?cdl=<font color="darkred">ī�װ�</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('brand','2')">�귣����̵� ��ũ : /street/street_brand.asp?makerid=<font color="darkred">�귣����̵�</font></span>
					</font>
					</div>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���̹��� 3</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnsimg3" value="�̹������" onClick="jsSetImg('<%=simg3%>','simg3','simgdiv3')" class="button"/>
		<div id="simgdiv3" style="padding: 5 5 5 5">
			<%IF simg3 <> "" THEN %>			
				<img src="<%=simg3%>" border="0" height="100" onclick="jsImgView('<%=simg3%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('simg3','simgdiv3');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				�̹��� �ּ� : <%=simg3%>
			<%END IF%>
		</div>
	</td>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̹�����ũ&����</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="#FFFFFF">
			<tr>
				<td>ALT : <input type="text" class="text" name="salt3" value="<%=salt3%>" size="40"/></td>
			</tr>
			<tr>
				<td>
					URL : <input type="text" class="text" name="surl3" value="<%=surl3%>" size="80"/><br/><br/>
					<div style="padding-left:25px;">
					<font color="#707070">
					- <span style="cursor:pointer" onClick="putLinkText('search','3')">�˻���� ��ũ : /search/search_result.asp?rect=<font color="darkred">�˻���</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('event','3')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('itemid','3')">��ǰ�ڵ� ��ũ : /category/category_itemprd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('category','3')">ī�װ� ��ũ : /category/category_list.asp?cdl=<font color="darkred">ī�װ�</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('brand','3')">�귣����̵� ��ũ : /street/street_brand.asp?makerid=<font color="darkred">�귣����̵�</font></span>
					</font>
					</div>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���̹��� 4</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnsimg4" value="�̹������" onClick="jsSetImg('<%=simg4%>','simg4','simgdiv4')" class="button"/>
		<div id="simgdiv4" style="padding: 5 5 5 5">
			<%IF simg4 <> "" THEN %>			
				<img src="<%=simg4%>" border="0" height="100" onclick="jsImgView('<%=simg4%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('simg4','simgdiv4');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				�̹��� �ּ� : <%=simg4%>
			<%END IF%>
		</div>
	</td>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̹�����ũ&����</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="#FFFFFF">
			<tr>
				<td>ALT : <input type="text" class="text" name="salt4" value="<%=salt4%>" size="40"/></td>
			</tr>
			<tr>
				<td>
					URL : <input type="text" class="text" name="surl4" value="<%=surl4%>" size="80"/><br/><br/>
					<div style="padding-left:25px;">
					<font color="#707070">
					- <span style="cursor:pointer" onClick="putLinkText('search','4')">�˻���� ��ũ : /search/search_result.asp?rect=<font color="darkred">�˻���</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('event','4')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('itemid','4')">��ǰ�ڵ� ��ũ : /category/category_itemprd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('category','4')">ī�װ� ��ũ : /category/category_list.asp?cdl=<font color="darkred">ī�װ�</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('brand','4')">�귣����̵� ��ũ : /street/street_brand.asp?makerid=<font color="darkred">�귣����̵�</font></span>
					</font>
					</div>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���̹��� 5</td>
	<td bgcolor="#FFFFFF" >
		<input type="button" name="btnsimg5" value="�̹������" onClick="jsSetImg('<%=simg5%>','simg5','simgdiv5')" class="button"/>
		<div id="simgdiv5" style="padding: 5 5 5 5">
			<%IF simg5 <> "" THEN %>			
				<img src="<%=simg5%>" border="0" height="100" onclick="jsImgView('<%=simg5%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('simg5','simgdiv5');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				�̹��� �ּ� : <%=simg5%>
			<%END IF%>
		</div>
	</td>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̹�����ũ&����</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="#FFFFFF">
			<tr>
				<td>ALT : <input type="text" class="text" name="salt5" value="<%=salt5%>" size="40"/></td>
			</tr>
			<tr>
				<td>
					URL : <input type="text" class="text" name="surl5" value="<%=surl5%>" size="80"/><br/><br/>
					<div style="padding-left:25px;">
					<font color="#707070">
					- <span style="cursor:pointer" onClick="putLinkText('search','5')">�˻���� ��ũ : /search/search_result.asp?rect=<font color="darkred">�˻���</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('event','5')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('itemid','5')">��ǰ�ڵ� ��ũ : /category/category_itemprd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('category','5')">ī�װ� ��ũ : /category/category_list.asp?cdl=<font color="darkred">ī�װ�</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('brand','5')">�귣����̵� ��ũ : /street/street_brand.asp?makerid=<font color="darkred">�귣����̵�</font></span>
					</font>
					</div>
				</td>
			</tr>
		</table>
	</td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
	<td colspan="4" align="center">
		<input type="button" value=" ���� " class="button" onclick="subcheck();"/> &nbsp;&nbsp;
		<input type="button" value=" ���(main) " class="button" onclick="jsgolist();"/>
	</td>
</tr>
</form>
</table>
<%
	If idx <> "0" then
%>
<a id="itemlist"></a>
<br/><br/><p><b>�� ���� ����</b><br/></p>
<!-- // ��ϵ� ���� ��� --------->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="900" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="10">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
		    <td align="left">
		    	�� <%=oSubItemList.FTotalCount%> �� /
		    	<input type="button" value="��ü����" class="button" onClick="chkAllItem()">
		    	<input type="button" value="��������" class="button" onClick="saveList()" title="ǥ�ü��� �� ��뿩�θ� �ϰ������մϴ�.">
		    </td>
		    <td align="right">
		    	<input type="button" value="��ǰ�ڵ�� ���" class="button" onClick="popRegArrayItem()" />
		    	<input type="button" value="��ǰ �߰�" class="button" onClick="popRegSearchItem()" />
		    	<img src="/images/icon_new_registration.gif" border="0" onclick="popSubEdit('')" style="cursor:pointer;" align="absmiddle">
		    </td>
		</tr>
		</table>
	</td>
</tr>
<col width="30" />
<col width="60" />
<col span="3" width="0*" />
<col width="70" />
<col width="80" />
<col width="80" />
<tr align="center" bgcolor="#DDDDFF">
    <td>&nbsp;</td>
    <td>�����ȣ</td>
    <td>�̹���</td>
    <td>��ǰ�ڵ�</td>
    <td>ǥ�ü���</td>
    <td>��뿩��</td>
</tr>
<tbody id="subList">
<%	
	Dim lp
	For lp=0 to oSubItemList.FResultCount-1 
%>
<tr align="center" bgcolor="<%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","#FFFFFF","#F3F3F3")%>#FFFFFF">
    <td><input type="checkbox" name="chkIdx" value="<%=oSubItemList.FItemList(lp).FIdxsub%>" /></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FIdxsub%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).FIdxsub%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FIdxsub%>)" style="cursor:pointer;">
    <%
    	if Not(oSubItemList.FItemList(lp).FsmallImage="" or isNull(oSubItemList.FItemList(lp).FsmallImage)) then
    		Response.Write "<img src='" & oSubItemList.FItemList(lp).FsmallImage & "' height='50' />"
    	end if
    %>
    </td>
    <td>
    <%
    	if Not(oSubItemList.FItemList(lp).FItemid="0" or isNull(oSubItemList.FItemList(lp).FItemid) or oSubItemList.FItemList(lp).FItemid="") then
    		Response.Write "[" & oSubItemList.FItemList(lp).FItemid & "]" & oSubItemList.FItemList(lp).Fitemname
    	end if
    %>
    </td>
    <td><input type="text" name="sort<%=oSubItemList.FItemList(lp).FIdxsub%>" size="3" class="text" value="<%=oSubItemList.FItemList(lp).Fsortnum%>" style="text-align:center;" /></td>
    <td>
		<span class="rdoUsing">
		<input type="radio" name="use<%=oSubItemList.FItemList(lp).FIdxsub%>" id="rdoUsing<%=lp%>_1" value="Y" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","checked","")%> /><label for="rdoUsing<%=lp%>_1">���</label><input type="radio" name="use<%=oSubItemList.FItemList(lp).FIdxsub%>" id="rdoUsing<%=lp%>_2" value="N" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="N","checked","")%> /><label for="rdoUsing<%=lp%>_2">����</label>
		</span>
    </td>
</tr>
<% Next %>
</tbody>
</table>
</form>
<div style="padding-bottom:200px;"></div>
<%
	End If 
	set oSubItemList = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
