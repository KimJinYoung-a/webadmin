<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ��ǰ���
' History : 
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 

<!-- #include virtual="/lib/classes/items/itemcls_v2.asp"-->
<!-- #include virtual="/lib/classes/event/eventPartnerWaitCls.asp"-->
<!-- #include virtual="/partner/lib/function/incPageFunction.asp" -->
<%
dim evtCode
evtCode =    requestCheckVar(Request("eC"),10)

if evtCode = "" then
		Call Alert_close ("���԰�ο� ������ ������ϴ�.  ")
end if

dim sdiv
sdiv =   requestCheckVar(Request("sdiv"),1)
dim menupos
dim itemid, makerid, itemname, waititemid
dim sellyn, isusing, danjongyn, limityn, mwdiv,sailyn
dim page, cdl, cdm, cds, dispCate
dim infodivYn, itemdiv,overseaYN
dim sSort,iPageSize


itemid  	= requestCheckvar(request("itemid"),255)
makerid 	= RequestCheckVar(request("makerid"),32)
itemname 	= RequestCheckVar(request("itemname"),32)

sellyn		= RequestCheckVar(request("sellyn"),10)
isusing 	= RequestCheckVar(request("isusing"),10)
danjongyn = RequestCheckVar(request("danjongyn"),10)
limityn 	= RequestCheckVar(request("limityn"),10)
mwdiv 		= RequestCheckVar(request("mwdiv"),10)
sailyn 		= RequestCheckVar(request("sailyn"),10)
page 			= RequestCheckVar(request("page"),10)
  

dispCate 	= requestCheckvar(request("disp"),16)
infodivYn = requestCheckvar(request("infodivYn"),10) 
itemdiv 	= requestCheckvar(request("itemdiv"),2)
overseaYN	= requestCheckvar(request("overseaYN"),1)
sSort			= requestCheckvar(request("sSort"),10)
if (sellyn="") then sellyn="A"

if (page="") then page=1
if sSort = "" then sSort = "ID"

isusing="Y"

'��ǰ�ڵ� ��ȿ�� �˻�(2008.08.01;������)
if itemid<>"" then
	dim iA ,arrTemp,arrItemid
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

'==============================================================================


dim oitem
iPageSize = 5

dim i, arrItem, iItemCnt,iTotPage
dim ClsEvt, arrList , intLoop,iTotCnt

set ClsEvt = new CEvent
ClsEvt.FevtCode = evtCode
makerid =  ClsEvt.fnGetMakerid 
clsEvt.FSdiv = sdiv

 
ClsEvt.FRectItemid = itemid
ClsEvt.FRectItemName = itemname
ClsEvt.FPSize = iPageSize
ClsEvt.FCPage = page
ClsEvt.FRectDispCate		= dispCate
ClsEvt.FRectSailYn       = sailyn
ClsEvt.FRectSort      =  sSort
if (sellyn <> "A") then
    ClsEvt.FRectSellYN = sellyn
end if
 
arrItem = ClsEvt.fnGetProductList
iItemCnt = ClsEvt.FitemTotCnt
 
arrList = ClsEvt.fnGetEventItemBanner
iTotCnt = ClsEvt.FTotCnt 
set ClsEvt = nothing
iTotPage 	=  int((iItemCnt-1)/iPageSize) +1  '��ü ������ ��	


%>
<html>
	<head>
<link rel="stylesheet" type="text/css" href="/css/adminPartnerCommon.css" />

<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script type="text/javascript">
	function SubmitSearch(){
	document.frm.submit();
}


function tnCheckAll(bool, comp){
    var frm = comp.form;
    if (!comp.length){
        comp.checked = bool;
    }else{
        for (var i=0;i<comp.length;i++){
            comp[i].checked = bool;
        }
    }
}

function tnCheckOne(itemid, comp){
	if(comp.value==itemid){
		if(comp.checked){
			comp.checked = false;
		}else{
			comp.checked = true;
		}
	}else{
		for (var i=0;i<comp.length;i++){
			if(comp[i].value==itemid){
				if(comp[i].checked){
					comp[i].checked = false;
				}else{
					comp[i].checked = true;
				}

			}
		}
	}
}

function TnSelectItemReg(rfrm){	
		$("#sortable input:checkbox[name='cksel']").attr('checked', true);
		rfrm.target="FrameCKP";
		rfrm.submit();
}

function TnSelectItemDel(rfrm){ 
		rfrm.target="FrameCKP";
		rfrm.submit();
	 
}

function TnSaveThemeItemBanner(){
	opener.document.frmEvt.target='FrameCKP';
	opener.document.frmEvt.upback.value='Y';
	opener.document.frmEvt.submit();
	//self.close();
}

function TnDelThemeItemBanner(){
	opener.document.frmEvt.target='FrameCKP';
	opener.document.frmEvt.upback.value='Y';
	opener.document.frmEvt.submit();
	$("#sortable input:checkbox[name='cksel']").attr('checked', false);
	location.reload();
}

function jsSetBanner(rfrm){
	rfrm.hidM.value = "TS";
	//rfrm.target="FrameCKP";
	rfrm.submit(); 
}
</script>
</head>
<body>	 
<!-- �˾� ������ : �ּ� 1100*750 -->
		<div class="popWinV17">	
			<h1>��ǰ ����</h1>		 
			<div class="popContainerV17 noScrl">
				<div class="contL">
					<h2>��ǰ ����</h2>						
					<div id="unitType01" class="unitPannel">
						<form name="frm" method=get>						
						<input type="hidden" name="page" >
						<input type='hidden' name='eC' id='eC'value="<%=evtCode%>">
						<input type='hidden' name='sdiv' id='sdiv'value="<%=sdiv%>">
						<div class="searchWrap" style="border-top:none;">
							<div class="search rowSum1">								
								<ul>
									<li>
										<label class="formTit" for="pdtName">��ǰ�� :</label>
										<input type="text" class="formTxt" id="pdtName" style="width:170px" placeholder="��ǰ�� �Է�" name="itemname" value="<%= itemname %>" size="20" />
									</li>
								</ul>
								<ul>
									<li>
										<label class="formTit" for="ctgy1">ī�װ� :</label>
										<!-- #include virtual="/common/module/dispCateSelectBox_upche.asp"-->
									</li>
								</ul>
								<div class="floating1">
									<label class="formTit" for="pdtCode">��ǰ�ڵ� :</label>
									<textarea class="formTxtA" rows="3" id="pdtCode" style="width:120px" placeholder="��ǰ�ڵ� �Է�" name="itemid" ><%=replace(itemid,",",chr(10))%></textarea>
								</div>
							</div> 
							<dfn class="line"></dfn>
							<div class="search">
								<ul>
									<li>
										<label class="formTit" for="sell">�Ǹ� :</label>
										<% drawSelectBoxSellYN "sellyn", sellyn %>
									</li>
									<li>
							   		<label class="formTit" for="sale">���� :</label><% drawSelectBoxSailYN "sailyn", sailyn %>
							   	</li>
							</div>		
							<input type="button" class="schBtn" value="�˻�" onClick="javascript:SubmitSearch();"/>
							
						</div>					
						<div class="tbListWrap tMar15">
						<div class="ftRt pad10">
								<span>�˻���� : <strong><%=formatnumber(iItemCnt,0) %> </strong></span> <span class="lMar10">������ : <strong><%=formatnumber(page,0)%> / <%=formatnumber(iTotPage,0)%></strong></span>
							</div>
							<ul class="thDataList">
								<li>
									<p class="cell05"><input type="checkbox" onClick="tnCheckAll(this.checked,frm.cksel);" /></p>
									<p class="cell10">��ǰ ID</p>
									<p class="cell10">�̹���</p>
									<p>��ǰ��</p>
									<p class="cell15">�ǸŰ�</p>
									<p class="cell15">���ް�</p>
									<p class="cell10">�Ǹſ���</p>
								</li>
							</ul>											    
							<ul class="tbDataList"  id="items">
						<% if isArray(arrItem) then %>
						    <% for i=0 to ubound(arrItem,2)%>			
								<li>
									<p class="cell05"><input type="checkbox" name="cksel" id="cksel" value="<%= arrItem(0,i)%>" /></p>
									<p class="cell10"><a href="<%=wwwUrl%>/<%= arrItem(0,i)%>" target="_blank"><%= arrItem(0,i)%></a></p>
									<p class="cell10"><img src="<%=webImgUrl& "/image/small/" & GetImageSubFolderByItemid(arrItem(0,i)) & "/" & arrItem(12,i)%>" width="50" height="50" border="0" alt=""></p>
									<p class="lt"><%= arrItem(2,i)%><input type="hidden" id="itemname<%=i%>" value="<%= arrItem(2,i)%>"></p>
									<p class="cell15">  
										<%= FormatNumber(arrItem(15,i),0) %>
										<%
										'���ΰ�
												if arrItem(19,i)="Y" then
													Response.Write "<br><font color=#F08050>("&CLng((arrItem(15,i)-arrItem(17,i))/arrItem(15,i)*100) & "%��)" & FormatNumber(arrItem(17,i),0) & "</font>"
												end if
												'������
												if arrItem(20,i)="Y" then
													Select Case arrItem(22,i)
														Case "1"
															Response.Write "<br><font color=#5080F0>(��)" &  FormatNumber(arrItem(3,intLoop)-(CLng(arrItem(23,intLoop)*arrItem(3,intLoop)/100)),0) & "</font>"
														Case "2"
															Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(arrItem(3,intLoop)-arrItem(23,intLoop),0) & "</font>"
													end Select
												end if
									    %>
									   </p>
									  <p class="cell15" >
									  	<%= FormatNumber(arrItem(16,i),0) %>
											<%
											'���ΰ�
												if arrItem(19,i)="Y" then
													Response.Write "<br><font color=#F08050>" & FormatNumber(arrItem(18,i),0) & "</font>"
												end if
												'������
												if arrItem(20,i)="Y" then
													if arrItem(22,i)="1" or arrItem(22,i)="2" then
														if arrItem(24,i)=0 or isNull(arrItem(24,i)) then
															Response.Write "<br><font color=#5080F0>" & FormatNumber(arrItem(16,i),0) & "</font>"
														else
															Response.Write "<br><font color=#5080F0>" & FormatNumber(arrItem(24,i),0) & "</font>"
														end if
													end if
												end if
											%>  
									  </p>
									<p class="cell10">
										<%if arrItem(5,i)="Y" then%>
										<span class="cBl1">�Ǹ���</span>
										<%elseif  arrItem(5,i)="S" then%>
										<span class="cRd1">�Ͻ�ǰ��</span>
										<%else%>
										<span class="cRd1">�Ǹž���</span>
										<%end if%>
									</p>
								</li>
							<%next%>	
						<%else%>								
								<li>
									<p>��ϵ� ��ǰ�� �����ϴ�.</p>									 
								</li>
						<%end if%>		
							</ul>
							</form>
							<div class="ct tPad20 cBk1">
								<%sbDisplayPaging "page", page, iItemCnt, iPageSize, 10,menupos %>
							</div>
						</div>
					</div>
				</div>

				<input type="button" class="btnMove" id="additem" title="�����ؼ� ���" />
			<form name="frmArrupdate" method="post" action="procEvent.asp">
					<input type='hidden' name='item_temp'>
					<input type='hidden' name='hidM' value="TB">
					<input type="hidden" name="check" id="check" value="<% If iTotCnt > 0 Then Response.write iTotCnt Else Response.write 0 %>">
					<input type='hidden' name='checkcnt' id='checkcnt'>
					<input type='hidden' name='eC' id='eC'value="<%=evtCode%>">
					<input type='hidden' name='sdiv' id='sdiv'value="<%=sdiv%>">
					<input type='hidden' name='upback' value="N">
					<input type="hidden" name="delid" value="">
					<input type="hidden" name="page" value="<%=page%>">
					<input type="hidden" name="itemid" value="<%=itemid%>">
					<input type="hidden" name="disp" value="<%=dispcate%>">
					<input type="hidden" name="itemname" value="<%=itemname%>">
					<input type="hidden" name="sellyn" value="<%=sellyn%>">
					<input type="hidden" name="sailyn" value="<%=sailyn%>">
				<div class="contR">
					<h2 style="margin-left:-1px;">PC ���� ����</h2>
					<!--<div class="pad10 ftRt">
						<select class="formSlt" id="sorting" title="���� ����">
							<option>5���� ��ǰ ����</option>
							<option>10���� ��ǰ ����</option>
						</select>
					</div>-->
					<div class="tbListWrap">
						<ul class="thDataList">
							<li>
								<p class="cell10"><input type="checkbox" onClick="tnCheckAll(this.checked,frmArrupdate.cksel);" /></p>
								<p class="cell15">��ǰ ID</p>
								<p>��ǰ��</p>
								<!--<p class="cell20">�ǸŰ���</p>							-->
							</li>
						</ul>
						<div id="sitem">
						<ul id="sortable" class="tbDataList">
							<%if isArray(arrList ) then							  
									for intLoop = 0 To UBound(arrList,2)
								%>
							<li id='del<%= arrList(1,intLoop)%>'>
								<p class="cell10"><input type='checkbox' name='cksel' id="cksel<%=intLoop%>" value='<%= arrList(1,intLoop) %>' /><input type='hidden' name='sitemname' id='sitemname' value='<%=arrList(3,intLoop) %>'></p>
								<p class="cell15"><%=arrList(1,intLoop)%></p>
								<p class="lt"><span><%=arrList(3,intLoop)%></span></p>
						<!--		<p class="cell20"><%=FormatNumber(arrList(7,intLoop),0)%>
									<%
										'���ΰ�
												if arrList(4,intLoop)="Y" then
													Response.Write "<br><font color=#F08050>("&CLng((arrList(7,intLoop)-arrList(9,intLoop))/arrList(7,intLoop)*100) & "%��)" & FormatNumber(arrList(9,intLoop),0) & "</font>"
												end if
												'������
												if arrList(11,intLoop)="Y" then
													Select Case arrList(12,intLoop)
														Case "1"
															Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(arrList(5,intLoop)-(CLng(arrList(13,intLoop)*arrList(5,intLoop)/100)),0) & "</font>"
														Case "2"
															Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(arrList(5,intLoop)-arrList(13,intLoop),0) & "</font>"
													end Select
												end if
									%>
								</p>-->
							</li>
							<% next
						end if
							 %>
						</ul>
						</div>
					</form>
						<div class="pad10 rt">
							<input type="button" class="btn" value="���û���" onclick="" id="delitem"/> 
						</div>
					</div>
				</div>
			</div>
			<div class="popBtnWrap">
				<input type="button" value="����" onclick="jsSetBanner(frmArrupdate);" class="cRd1" style="width:100px; height:30px;" />
			<!--	<input type="button" value="���" onclick="self.close();" style="width:100px; height:30px;" />-->
			</div>
		</div> 
	</div> 
</div>	 
<script>
$(function(){
	$("#additem").click(function(){
		
		var CheckOverlap=false;
		$("#items input:checkbox[name='cksel']").each(function(i){
			if($("#items input:checkbox[name='cksel']:eq(" + i + ")").is(":checked")){
				$("#sortable input:checkbox[name='cksel']").each(function(j){
					//alert($("#items input:checkbox[name='cksel']:eq(" + i + ")").val() + " / " + $("#sortable input:checkbox[name='cksel']:eq(" + j + ")").val());
					if($("#items input:checkbox[name='cksel']:eq(" + i + ")").val() == $("#sortable input:checkbox[name='cksel']:eq(" + j + ")").val()){
						CheckOverlap=true;
						$("#items input:checkbox[name='cksel']:eq(" + i + ")").attr('checked', false);
						return false;
					}
					else{
						CheckOverlap=false;
						return;
					}
				});
			}
		});
		//alert((Number($("#check").val())+Number($("#items input:checkbox[name='cksel']:checked").length)));
		
		if(CheckOverlap){
			alert("���õ� ��ǰ�� ���� ��ǰ�� �ֽ��ϴ�.");
			return false;
		}else if($("#items input:checkbox[name='cksel']:checked").length<1){
			alert("���õ� ��ǰ�� �����ϴ�.");
			return false;
		}else if((Number($("#check").val())+Number($("#items input:checkbox[name='cksel']:checked").length))>3){
			alert("��ǰ�� �ִ� 3������ ������ �� �ֽ��ϴ�..");
			return false;
		}else{
			// ���߰�
			var oRow;
			$("#items input:checkbox[name='cksel']").each(function(i){
				if($("#items input:checkbox[name='cksel']:eq(" + i + ")").is(":checked")){
					oRow = "					<li id='del" + $("#items input:checkbox[name='cksel']:eq(" + i + ")").val() + "'>"
					oRow += "						<p class='cell10'><input type='checkbox' name='cksel' id='cksel'" + i + " value='" + $("#items input:checkbox[name='cksel']:eq(" + i + ")").val() + "' /><input type='hidden' name='sitemname' id='sitemname'  value='" + $("#itemname"+i).val() + "'></p>"
					oRow += "						<p class='cell15'>" + $("#items input:checkbox[name='cksel']:eq(" + i + ")").val() + "</p>"
					oRow += "						<p class='lt'><span class='textOverflow'>" + string_cut($("#itemname"+i).val(), 20, "...") + "</span></p>"
					oRow += "					</li>"
					//alert(oRow);
					$("#sitem ul").append(oRow);
					$("#check").val(Number($("#check").val())+1);//�ɼ� �߰� ���� ī��Ʈ
				}
			});
		}
		tnCheckAll(true,frmArrupdate.cksel);
		document.frmArrupdate.hidM.value="TB";
		document.frmArrupdate.upback.value="Y";
		TnSelectItemReg(frmArrupdate);
	});

var delid="" ;
	$("#delitem").click(function(){
		if($("#sortable input:checkbox[name='cksel']:checked").length<1){
			alert("���õ� ��ǰ�� �����ϴ�.");
		}else{
				$("#sortable input:checkbox[name='cksel']").each(function(i){
					if($("#sortable input:checkbox[name='cksel']:eq(" + i + ")").is(":checked")){
					//	$("#del"+ $("#sortable input:checkbox[name='cksel']:eq(" + i + ")").val()).empty();
						if(delid==""){
							delid = $("#sortable input:checkbox[name='cksel']:eq(" + i + ")").val();
						}else{
							delid = delid +"," + $("#sortable input:checkbox[name='cksel']:eq(" + i + ")").val();
						}
						$("#check").val(Number($("#check").val())-1);//�ɼ� �߰� ���� ī��Ʈ
					}
				
			document.frmArrupdate.hidM.value="TD";
			document.frmArrupdate.upback.value="Y";
			document.frmArrupdate.delid.value = delid;
			TnSelectItemDel(frmArrupdate);
			});
		}
	});
});
function string_cut(str, len, tail) {
	return str.substr(0, len)+tail;
}
</script>
<script>
$(function() {
	$("#sortable").sortable({
		placeholder:"handling"
	}).disableSelection();

	$(".tab li").click(function() {
		$(".tab li").removeClass('selected');
		$(this).addClass('selected');
		$('.unitPannel').hide();
		var activeTab = $(this).find("a").attr("href");
		$(activeTab).show();
		return false;
	});
});

</script>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe> 
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->