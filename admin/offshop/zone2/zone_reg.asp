<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 오프라인 조닝 삽별구역설정
' Hieditor : 2010.12.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone2/zone_cls.asp"-->

<%
Dim ozone,idx , i , shopid,zonename,racktype,unit,orderno,regdate ,isusing , zonegroup ,menupos
dim omanager ,managershopyn
	idx = requestCheckVar(request("idx"),10)
	menupos = requestCheckVar(request("menupos"),10)

set ozone = new czone_list
	ozone.frectidx = idx

set omanager = new czone_list
	omanager.frectzoneidx = idx
	
	'//수정시에만 쿼리
	if idx <> "" then		
		ozone.fzone_oneitem()
		
		if ozone.ftotalcount >0 then			
			shopid = ozone.FOneItem.fshopid		
			zonename = ozone.FOneItem.fzonename
			unit = ozone.FOneItem.funit			
			regdate = ozone.FOneItem.fregdate
			isusing = ozone.FOneItem.fisusing
			managershopyn = ozone.FOneItem.fmanagershopyn
			
			if managershopyn = "Y" then
				omanager.Getshopzonemanager()
			end if
		end if
	end if
	
%>

<script type="text/javascript">

	window.resizeTo(800, 500);
	
	function reg(){
		if (frm.shopid.value=='') {
			alert('매장을 선택해 주세요');
			frm.zonename.focus();
			return;
		}
		
		if (frm.zonename.value=='') {
			alert('조닝명을 입력해 주세요');
			frm.zonename.focus();
			return;
		}
		
		if (frm.unit.value=='') {
			alert('조닝 크기를 입력해 주세요');
			frm.unit.focus();			
			return;
		}
		
		if(frm.unit.value!=''){
			if (!IsDouble(frm.unit.value)){
				alert('조닝 크기는 숫자만 가능합니다.');
				frm.unit.focus();
				return;
			}
		}	

		if (frm.isusing.value=='') {
			alert('사용여부를 선택해 주세요');
			frm.isusing.focus();			
			return;
		}
		
		frm.action='/admin/offshop/zone2/zone_process.asp';
		frm.mode.value = "zonereg";
		frm.submit();
	}

    // 매장 선택 팝업
	function popmanagerSelect(){
		var popmanagerSelect = window.open("/admin/offshop/zone2/pop_managerSelect.asp", "popmanagerSelect","width=600,height=400,scrollbars=yes,resizable=yes");
		popmanagerSelect.focus();
	}

	//팝업에서 매니져 선택 추가
	function addSelectedmanager(empno,username){
		var lenRow = tablemanager.rows.length;

		// 기존에 값에 중복값 여부 검사
		if(lenRow>1)	{
			for(l=0;l<document.all.empno.length;l++)	{
				if(document.all.empno[l].value==empno) {
					alert("이미 지정된 담당자 입니다");
					return;
				}
			}
		}
		else {
			if(lenRow>0) {
				if(document.all.empno.value==empno) {
					alert("이미 지정된 담당자 입니다");
					return;
				}
			}
		}

		// 행추가
		var oRow = tablemanager.insertRow(lenRow);
		oRow.onmouseover=function(){tablemanager.clickedRowIndex=this.rowIndex};

		// 셀추가 (이름,삭제버튼)
		var oCell1 = oRow.insertCell(0);		
		var oCell3 = oRow.insertCell(1);

		oCell1.innerHTML = username + "<input type='hidden' name='empno' value='" + empno + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delSelectdmanager()' align=absmiddle>";
	}

	// 선택삭제
	function delSelectdmanager(){
	    
		if(confirm("선택한 담당자를 삭제하시겠습니까?"))
			tablemanager.deleteRow(tablemanager.clickedRowIndex);
	}
			
</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		※ <font color="red">[중요] </font>매장내 조닝이 변경되거나 없어지면,
		<br>기존 조닝을 수정하지 마시고, 사용안함 돌리신후, 새로 등록하세요.
		<br>기존 조닝을 현재 변경될 조닝으로 수정후 사용 하실경우,
		<br>기존 조닝으로 등록되어진 상품들이 모두 현재 조닝으로 변경되는 문제가 발생됩니다
	</td>
	<td align="right">
		<input type="button" value="신규등록" class="button" onclick="location.href='?menupos=<%=request("menupos")%>';">&nbsp;&nbsp;
		<input type="button" value="창닫기" class="button" onclick="window.close();">	
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<form name="frm" method="post" style="margin:0px;" >
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
	<td align="center">번호<br></td>
	<td>
		<%=idx%><input type="hidden" name="idx" value="<%=idx%>">
	</td>
</tr>	
<tr bgcolor="#FFFFFF">
	<td align="center">SHOP</td>
	<td>
		<% drawSelectBoxOffShop "shopid",shopid %>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td align="center">조닝명</td>
	<td>
		<input type="text" name="zonename" value="<%=zonename%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">조닝크기</td>
	<td>
		<input type="text" name="unit" value="<%=unit%>" size=5 maxlength=5> ex)1
		<p>※ 해당지역의 평수로 사용하시거나, 유동적으로 편하신대로 지정해서 사용하시면 됩니다
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">매장내담당자<br></td>
	<td>
		<table border="0" cellspacing="0" class="a">
		<tr>
			<td>
			    <table name='tablemanager' id='tablemanager' class=a>
			    <% if managershopyn = "Y" then %>
			        <% for i=0 to omanager.FResultCount-1 %>
			        <tr onMouseOver='tablemanager.clickedRowIndex=this.rowIndex'>
				    	<td>
				    	    <%= omanager.FItemList(i).fusername %>
				    	    <input type='hidden' name='empno' value='<%= omanager.FItemList(i).fempno %>'></td>  
				    	<td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delSelectdmanager()' align=absmiddle></td>   
			        </tr>
			        <% next %>
			    <% end if %>
			    </table>
			</td>
			<td valign="bottom"><input type="button" class='button' value="추가" onClick="popmanagerSelect()"></td>
		<tr>
	    </table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">사용여부<br></td>
	<td>
		<select name="isusing">
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan=2>
		<input type="button" value="저장" class="button" onclick="reg();">
	</td>
</tr>
</table>	
</form>

<%
set ozone = nothing
set omanager = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
