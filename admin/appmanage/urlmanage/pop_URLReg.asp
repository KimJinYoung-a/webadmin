<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2014-08-18 - ����ȭ ����
' Discription : APPURL ���
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/appURLCls.asp"-->
<%
	Dim idx, urldiv, urltitle, urlcontent, isusing , dispCate , tmpappurl

	If Date() >= "2014-10-01" Then
		tmpappurl = "/apps/appCom/wish/web2014/"
	Else
		tmpappurl = "/apps/appCom/wish/webview/"
	End If 

	idx		= requestCheckVar(request("idx"),10)
	dispCate = requestCheckvar(request("disp"),16)

	if idx<>"" then
		dim oAppurl
		set oAppurl = New APPURL
		oAppurl.FCurrPage = 1
		oAppurl.FPageSize=1
		oAppurl.FRectidx = idx
		oAppurl.getappurl
	
		if oAppurl.FResultCount>0 then
			urldiv		= oAppurl.FItemList(0).Furldiv
			urltitle	= oAppurl.FItemList(0).Furltitle
			urlcontent	= oAppurl.FItemList(0).Furlcontent
			isusing		= oAppurl.FItemList(0).Fisusing
			dispCate	= oAppurl.FItemList(0).Fcatecode
		end if
	
		set oAppurl = Nothing
	end If

	Function URLDecode(sConvert)
		Dim aSplit
		Dim sOutput
		Dim I
		If IsNull(sConvert) Then
		   URLDecode = ""
		   Exit Function
		End If

		If sConvert <> "" then

			' convert all pluses to spaces
			sOutput = REPLACE(sConvert, "+", " ")

			' next convert %hexdigits to the character
			aSplit = Split(sOutput, "%")

			If IsArray(aSplit) Then
			  sOutput = aSplit(0)
			  For I = 0 to UBound(aSplit) - 1
				sOutput = sOutput & _
				  Chr("&H" & Left(aSplit(i + 1), 2)) &_
				  Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
			  Next
			End If
		End If 

		URLDecode = sOutput
	End Function

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
<!--
	//�ڵ� ���
	function jsRegCode(){
		var frm = document.frmReg;
		if(!frm.urltitle.value) {
			alert("URL���� �Է��� �ּ���");
			frm.urltitle.focus();
			return false;
		}
		if(!frm.urldiv.value) {
			alert("URL������ ������ �ּ���");
			frm.urldiv.focus();
			return false;
		}
		if(!frm.urlcontent.value) {
			alert("URL������ �Է��� �ּ���");
			frm.urlcontent.focus();
			return false;
		}

		if(confirm("�Է��� ������ ��Ȯ�մϱ�?")) {
			return true;
		}
		return false;
	}

	//url �ڵ� ����
	function chklink(v){
		if (v == "1"){
			document.frmReg.urlcontent.value = "<%=tmpappurl%>category/category_itemprd.asp?itemid=��ǰ�ڵ�";
			$("#bestseltr").css("display","none");
			$("#catesel").css("display","none");
			$("#oDispCate").prop('disabled',false);
		}else if (v == "2"){
			document.frmReg.urlcontent.value = "<%=tmpappurl%>event/eventmain.asp?eventid=�̺�Ʈ�ڵ�&rdsite=rdsite��(�ʼ��ƴ�)";
			$("#bestseltr").css("display","none");
			$("#catesel").css("display","none");
			$("#oDispCate").prop('disabled',false);
		}else if (v == "3"){
			document.frmReg.urlcontent.value = "makerid=�귣���";
			$("#bestseltr").css("display","none");
			$("#catesel").css("display","none");
			$("#oDispCate").prop('disabled',false);
		}else if (v == "4"){
			document.frmReg.urlcontent.value = "cd1=&nm1=";
			$("#bestseltr").css("display","none");
			$("#catesel").css("display","block");
			$("#oDispCate").attr('readonly','readonly');
		}else if (v == "9"){
			document.frmReg.urlcontent.value = "<%=tmpappurl%>today/index.asp";
			$("#bestseltr").css("display","none");
			$("#catesel").css("display","none");
			$("#oDispCate").prop('disabled',false);
		}else if (v == "8"){ //' �ܺ� URL - �ܺ� ��ü �α� ������
			document.frmReg.urlcontent.value = "";
			$("#bestseltr").css("display","none");
			$("#catesel").css("display","none");
			$("#oDispCate").prop('disabled',false);
		}else if (v == "10"){ //' ����Ʈ
			document.frmReg.urlcontent.value = "<%=tmpappurl%>award/awarditem.asp?atype=ne";
			$("#bestseltr").css("display","block");
			$("#catesel").css("display","none");
			$("#oDispCate").prop('disabled',false);
		}else if (v == "11"){ //' ��ٱ���
			document.frmReg.urlcontent.value = "<%=tmpappurl%>inipay/ShoppingBag.asp";
			$("#bestseltr").css("display","none");
			$("#catesel").css("display","none");
			$("#oDispCate").prop('disabled',false);
		}else{
			document.frmReg.urlcontent.value = "APP URL ������ ���� ���ּ���.";
			$("#bestseltr").css("display","none");
			$("#catesel").css("display","none");
			$("#oDispCate").prop('disabled',false);
		}
	}
//-->
</script>
<script>
function chgDispCate(dc) {
	$.ajax({
		url: "/admin/appmanage/urlmanage/dispCateSelectBox_response.asp?disp="+dc,
		cache: false,
		async: false,
		success: function(message) {
			// ���� �ֱ�
			$("#lyrDispCtBox").empty().html(message);
			if (dc.length == 3){
				document.frmReg.urlcontent.value = "cd1="+$("#dispcateval1 option:selected").val()+"&nm1="+$("#dispcateval1 option:selected").text();
				$("#catecode").val(dc);
			}else if (dc.length == 6){
				document.frmReg.urlcontent.value = "cd1="+$("#dispcateval1 option:selected").val()+"&cd2="+$("#dispcateval2 option:selected").val()+"&nm1="+$("#dispcateval1 option:selected").text()+"&nm2="+$("#dispcateval2 option:selected").text();
				$("#catecode").val(dc);
			}else if (dc.length == 9){
				document.frmReg.urlcontent.value = "cd1="+$("#dispcateval1 option:selected").val()+"&cd2="+$("#dispcateval2 option:selected").val()+"&cd3="+$("#dispcateval3 option:selected").val()+"&nm1="+$("#dispcateval1 option:selected").text()+"&nm2="+$("#dispcateval2 option:selected").text()+"&nm3="+$("#dispcateval3 option:selected").text();
				$("#catecode").val(dc);
			}else{
				
			}
		}
	});
}

function chgBestSel(v)
{
	$("#oDispCate").val("<%=tmpappurl%>award/awarditem.asp?atype="+v);
	
}
$(function(){
	chgDispCate('<%=dispCate%>');
});
</script>
<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a" >
<tr>
	<td colspan="2"><!--//�ڵ� ��� �� ����-->	
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a" >
		<form name="frmReg" method="post" action="doappurl.asp" onSubmit="return jsRegCode();">
		<input type="hidden" name="idx" value="<%=idx%>">
		<input type="hidden" name="catecode" id="catecode">
		<tr>			
			<td><b>APP URL ��� �� ����</b></td>
		</tr>	
		<tr>
			<td>	
				<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
				<% IF idx <> "" THEN%>	
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">��ũ��ȣ</td>
					<td bgcolor="#FFFFFF"><%=idx%></td>
				</tr>
				<%END IF%>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">Ÿ��Ʋ</td>
					<td bgcolor="#FFFFFF"><input type="text" size="32" maxlength="64" name="urltitle" value="<%=urltitle%>" class="text"></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">APP URL ����</td>
					<td bgcolor="#FFFFFF"><% DrawSelectBoxAppUrlDiv "urldiv", urldiv %></td>
				</tr>
				<tr id="catesel" style="display:<%=chkiif(idx<>"" And urldiv = "4","block","none")%>">
					<td bgcolor="#EFEFEF" width="100" align="center">����ī�װ� ����</td>
					<td bgcolor="#FFFFFF">
						<span id="lyrDispCtBox"></span>
					</td>
				</tr>
				<tr id="bestseltr" style="display:<%=chkiif(idx<>"" And urldiv = "10","block","none")%>">
					<td bgcolor="#EFEFEF" width="100" align="center">����Ʈ ����</td>
					<td bgcolor="#FFFFFF">
						<select name="bestsel" id="bestsel" onchange="chgBestSel(this.value);">
							<option value="ne">�Ż�ǰ����Ʈ</option>
							<option value="be">����Ʈ����</option>
							<option value="st">���׵𼿷�</option>
							<option value="br">����Ʈ�귣��</option>
							<option value="vi">VIP����Ʈ</option>
						</select>
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">�ڵ峻��</td>
					<td bgcolor="#FFFFFF"><textarea name="urlcontent" class="textarea" id="oDispCate" style="width:100%; height:40px;"><%=URLDecode(urlcontent)%></textarea></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" align="center">��뿩��</td>
					<td bgcolor="#FFFFFF">
						<label><input type="radio" value="Y" name="isusing" onfocus="this.blur()" <%IF isusing="Y" or isusing="" THEN%>checked<%END IF%>>���</label>
						<label><input type="radio" value="N" name="isusing" onfocus="this.blur()" <%IF isusing="N" THEN%>checked<%END IF%>>������</label>
					</td>
				</tr>
				</table>		
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td align="left"><a href="javascript:self.close()"><img src="/images/icon_cancel.gif" border="0"></a></td>
					<td align="right"><input type="image" src="/images/icon_save.gif"></td>
				</tr>
				</table>
			</td>
		</tr>	
		</form>
		</table>
	</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->