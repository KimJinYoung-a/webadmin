<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/pop_event_winner.asp
' Description :  이벤트 당첨등록
' History : 2007.02.22 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/eventWinner_function.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<script language="javascript">
window.resizeTo(600,460);
<!--
	function jsChType(iVal){
		var frm = document.all;
		if(iVal == "2"){
			frm.div1.style.display = "none";
			frm.div2.style.display = "";
		}else if	(iVal == "3"){
			frm.div1.style.display = "none";
			frm.div2.style.display = "none";
		}else{
			frm.div1.style.display = "";
			frm.div2.style.display = "none";
		}
	}

	//-- jsPopCal : 달력 팝업 --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}


	function jsWinnerSubmit(frm){
		if(!frm.sR.value){
			alert("등수를 입력해주세요");
			frm.sR.focus();
			return false;
		}

		if(!IsDigit(frm.sR.value)){
			alert("등수는 숫자만 입력가능합니다.");
			frm.sR.focus();
			return false;
		}

		if(!frm.sW.value){
			alert("당첨자를 입력해주세요");
			frm.sW.focus();
			return false;
		}

		if(frm.selType.value == "1"){
			if(!frm.sGN.value){
				alert("사은품명을 입력해주세요");
				frm.sGN.focus();
				return false;
			}

			if (frm.reqdeliverdate.value.length<1){
			    frm.reqdeliverdate.focus();
			    alert('출고 요청일을 선택하세요.');
			    return false;
			}

			if ((!frm.isupchebeasong[0].checked)&&(!frm.isupchebeasong[1].checked)){
        		alert('배송 구분을 선택하세요.');
        		return false;
        	}

            if ((frm.isupchebeasong[1].checked)&&(frm.makerid.value.length<1)){
                alert('업체 아이디를 선택하세요.');
        		return false;
            }
		}

		if(frm.selType.value == "2"){
			if(!frm.couponvalue.value){
				alert("쿠폰금액 또는 할인율을 입력해주세요!");
				frm.couponvalue.focus();
				return false;
			}

			if(!frm.minbuyprice.value){
				alert("최소금액을 입력해주세요!");
				frm.minbuyprice.focus();
				return false;
			}

			 if(!frm.sDate.value || !frm.eDate.value ){
			  	alert("기간을 입력해주세요");
			  	frm.sDate.focus();
			  	return false;
			  }

			  if(frm.sDate.value > frm.eDate.value){
			  	alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
			  	frm.sDate.focus();
			  	return false;
			  }
		}

		if(confirm("등록하신 내용은 수정 또는 삭제가 불가능하며 고객에게 바로 적용됩니다.\n\n등록 하시겠습니까? ")){
			return true;
		}else{
		    return false;
		}
	}

	function disabledBox(comp){
        var frm = comp.form;
        if (comp.value=="Y"){
            frm.makerid.disabled = false;
        }else{
            frm.makerid.selectedIndex = 0;
            frm.makerid.disabled = true;
        }
    }
//-->
</script>
<%
Dim eCode : eCode = Request("eC")
dim arridx : arridx = chkarray(request("arridx"))
%>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 당첨자 등록</div>
<table width="580" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
<form name="frmWin" method="post" action="eventprize_process.asp" onSubmit="return jsWinnerSubmit(this);">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="mode" value="I">
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">구분</td>
				<td bgcolor="#FFFFFF">
					<select name="selType" onChange="jsChType(this.value);">
					<option value="1">사은품배송</option>
					<option value="2">쿠폰발급</option>
					<option value="3">Only View</option>
					</select>
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">등수</td>
				<td bgcolor="#FFFFFF"><input type="text" size="2" name="sR"></td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">등수별칭</td>
				<td bgcolor="#FFFFFF"><input type="text" name="sRN" size="20"></td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">당첨자</td>
				<td bgcolor="#FFFFFF">
					콤머로 구분, 공백없이 (예: aaa,bbb,ccc)<br>
					<textarea name="sW" rows="2" cols="60"><%= arridx %></textarea>
				</td>
			</tr>
		</table>
	</td>

</tr>
<tr>
	<td>
		<div id="div1" style="display:;">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td align="center" width="100"  bgcolor="<%= adminColor("tabletop") %>">배송지 등록구분</td>
				<td bgcolor="#FFFFFF">
					<input type=radio name=rdgubun value="U">User가 배송지 입력
					<input type=radio name=rdgubun value="F" checked>User 기본 주소 사용 <font color="blue">[가능한 기본 주소지 사용]</font>
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">사은품명</td>
				<td bgcolor="#FFFFFF"><input type="text" name="sGN" size="20"></td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">사은품 상품번호</td>
				<td bgcolor="#FFFFFF"><input type="text" name="itemid" size="10"></td>
			</tr>
			<!-- 배송 구분 추가 : 서동석 -->
			<tr>
            	<td align="center" bgcolor="<%= adminColor("tabletop") %>">출고요청일</td>
            	<td bgcolor="#FFFFFF">
            		<input type="text" name="reqdeliverdate" size="10" maxlength="10"  value="" >
		            <a href="javascript:jsPopCal('reqdeliverdate');"><img src="/images/calicon.gif" border="0" align="absmiddle"></a>
            	</td>
            </tr>
			<tr>
            	<td align="center" bgcolor="<%= adminColor("tabletop") %>">배송구분</td>
            	<td bgcolor="#FFFFFF">
            		<input type=radio name=isupchebeasong value="N" onClick="disabledBox(this);">텐바이텐배송
            		<input type=radio name=isupchebeasong value="Y" onClick="disabledBox(this);">업체직접배송
            	</td>
            </tr>
            <tr>
            	<td align="center" bgcolor="<%= adminColor("tabletop") %>">업체배송시<br>업체ID</td>
            	<td bgcolor="#FFFFFF">
            	    <% drawSelectBoxDesignerwithName "makerid","" %>
            	    <script language='javascript'>
            	    document.frmWin.makerid.disabled=true;
            	    </script>
            	</td>
            </tr>

		</table>
		</div>
		<div id="div2" style="display:;">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">쿠폰타입</td>
				<td bgcolor="#FFFFFF">
					<input type=text name=couponvalue maxlength=7 size=10>
					<input type=radio name=coupontype value="1" onclick="alert('% 할인 쿠폰입니다.');">%할인
					<input type=radio name=coupontype value="2" checked >원할인
					(금액 또는 % 할인)
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">최소구매금액</td>
				<td bgcolor="#FFFFFF"><input type=text name=minbuyprice maxlength=7 size=10>원 이상 구매시 사용가능(숫자)</td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">유효기간</td>
				<td bgcolor="#FFFFFF">
					<input type="text" name="sDate" value="<%= left(now(),10) %>"  size="10" maxlength="10" onClick="jsPopCal('sDate');" style="cursor:hand;">
					~<input type="text" name="eDate" size="10"  maxlength="10" onClick="jsPopCal('eDate');" style="cursor:hand;">
				</td>
			</tr>
		</table>
		</div>
	</td>

</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right" height="40">
		<input type="image" src="/images/icon_confirm.gif">
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</form>
</table>

<!-- #include virtual="/lib/db/dbclose.asp" -->