<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 이세로 전자계산서 선택
' History : 2012.02.07 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/tax/EseroTaxCls.asp"-->
<%
Dim clsEsero, arrList, intLoop
Dim iTotCnt,iPageSize, iTotalPage,page
Dim dSDate,dEDate,ssearchText,itaxsellType,itaxModiType,itaxType, iMapTpYn, iMapTp
Dim totSum, tgType
	iPageSize = 50
	page = requestCheckvar(Request("page"),10)
	if page="" then page=1


	dSDate = requestCheckvar(Request("dSD"),10)
	dEDate = requestCheckvar(Request("dED"),10)
	ssearchText = requestCheckvar(Request("sST"),200)
	itaxsellType = requestCheckvar(Request("iTST"),10)
	itaxModiType = requestCheckvar(Request("iTMT"),10)
	itaxType = requestCheckvar(Request("iTT"),10)
    iMapTpYn   = requestCheckvar(Request("iMapTpYn"),10)
    iMapTp     = requestCheckvar(Request("iMapTp"),10)
    totSum     = requestCheckvar(Request("totSum"),10)
    tgType     = requestCheckvar(Request("tgType"),20)

    if (itaxsellType="") then itaxsellType="0"

Set clsEsero = new CEsero
  clsEsero.FSDate      =dSDate
	clsEsero.FEDate      =dEDate
	clsEsero.FsearchText =ssearchText
	clsEsero.FtaxsellType=itaxsellType
	clsEsero.FtaxModiType=itaxModiType
	clsEsero.FtaxType    =itaxType
	clsEsero.FMappingTypeYN = iMapTpYn
	clsEsero.FMappingType   = iMapTp
	clsEsero.FtotSum     =totSum
	clsEsero.FCurrPage 	= page
	clsEsero.FPageSize 	= iPageSize
	arrList = clsEsero.fnGetEseroTaxList
	iTotCnt = clsEsero.FTotCnt
Set clsEsero = nothing
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<script type="text/javascript" src="/admin/approval/eapp/eapp.js"></script>
<script language="javascript">
<!--
// 페이지 이동
function jsGoPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}

	//검색
	function jsSearch(){
		document.frm.submit();
	}

	//달력보기
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	function jsSetTax(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP){
		opener.document.all.dView1.style.display = "";
		opener.document.frm.sEK.value= eTax;
		for(i=0;i<opener.document.frm.rdoDK.length;i++){
			opener.document.frm.rdoDK[i].checked = false;
			if(iDK==9){
				if(opener.document.frm.rdoDK[i].value ==2){
					opener.document.frm.rdoDK[i].checked= true;
				}
			}else{
				if(opener.document.frm.rdoDK[i].value ==1){
					opener.document.frm.rdoDK[i].checked= true;
				}
			}
		}


			if(iVK==1){
				opener.document.frm.sVK.value = "과세(부가세 10%)";
				opener.document.frm.rdoVK.value = 0;
			}else if(iVK==2)	{
				opener.document.frm.sVK.value = "영세";
				opener.document.frm.rdoVK.value = 3;
			}else{
				opener.document.frm.sVK.value = "면세";
				opener.document.frm.rdoVK.value = 2;
			}

		opener.document.frm.dID.value= dID;
		opener.document.frm.sINm.value= sInm;
		opener.document.frm.mTP.value= jsSetComma(mTP);
		opener.document.frm.mSP.value= jsSetComma(mSP);
		opener.document.frm.mVP.value= jsSetComma(mVP);
		
		opener.jsTexSetting();

		self.close();
	}

	function jsSetTaxNormal(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP){
	    opener.fillTaxInfo(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP);
	    self.close();
	}

	function jsSetTaxWithpayreq(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP,prIdx){
	    opener.fillTaxInfoWithPayreqIdx(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP,prIdx);
	    self.close();
	}

function popErpSending(itaxkey){
    var winD = window.open("popRegfileHand.asp?taxkey="+itaxkey,"popErpSending","width=860, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}

//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
	<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="page" value="">
			<input type="hidden" name="tgType" value="<%= tgType %>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td  rowspan="2" width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left">
					<input type="radio" name="iTST" value="0" <%= CHKIIF(itaxsellType="0","checked","") %> >매입
					<input type="radio" name="iTST" value="1" <%= CHKIIF(itaxsellType="1","checked","") %> >매출&nbsp;&nbsp;
					 작성일:
					<input type="text" name="dSD" size="10" value="<%=dSDate%>" onClick="jsPopCal('dSD');"  style="cursor:hand;">
					-
					<input type="text" name="dED" size="10" value="<%=dEDate%>" onClick="jsPopCal('dED');"  style="cursor:hand;">
					&nbsp;&nbsp;검색어:
					<input type="text" name="sST" value="<%=ssearchText%>" size="30"><font color="Gray">(사업자등록번호,상호,품목명)</font>
				</td>
				<td  rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="jsSearch();">
				</td>
			</tr>
			<tr>
			    <td  bgcolor="#FFFFFF">
			        매칭상태 :
			        <select Name="iMapTpYn">
			        <option value="">전체
			        <option value="Y" <%= CHKIIF(iMapTpYn="Y","selected","") %> >매칭
			        <option value="N" <%= CHKIIF(iMapTpYn="N","selected","") %> >비매칭
			        </select>
			        &nbsp;&nbsp;
			        매칭구분 :
			        <select Name="iMapTp">
			        <option value="">전체
			        <option value="1" <%= CHKIIF(iMapTp="1","selected","") %> >온라인 매입
			        <option value="2" <%= CHKIIF(iMapTp="2","selected","") %> >오프라인 매입
			        <option value="9" <%= CHKIIF(iMapTp="9","selected","") %> >기타 매입
			        <option value="11" <%= CHKIIF(iMapTp="11","selected","") %> >매출
			        </select>
			        &nbsp;&nbsp;
			        계산서구분:
			        <select Name="iTMT">
			        <option value="">전체
			        <option value="0" <%= CHKIIF(itaxModiType="0","selected","") %> >전자(일반)
			        <option value="1" <%= CHKIIF(itaxModiType="1","selected","") %> >전자(수정)
			        <option value="9" <%= CHKIIF(itaxModiType="9","selected","") %> >기타(수기)
			        </select>
			        &nbsp;&nbsp;
			        과세구분:
			        <select Name="iTT">
			        <option value="">전체
			        <option value="1" <%= CHKIIF(itaxType="1","selected","") %> >과세
			        <option value="2" <%= CHKIIF(itaxType="2","selected","") %> >영세
			        <option value="3" <%= CHKIIF(itaxType="3","selected","") %> >면세
			        </select>
			        &nbsp;&nbsp;
			        금액:
			        <input type="text" name="totSum" value="<%= totSum %>" maxlength="9" size="10">
			    </td>
			</tr>
			</form>
		</table>
	</td>
</tr>
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="FFFFFF">
				<td colspan="19">
					검색결과 : <b><%=iTotCnt%></b> &nbsp;
					페이지 : <b><%= page %> / <%=iTotalPage%></b>
				</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td rowspan="2">작성일자</td>
				<td rowspan="2">승인번호</td>
				<td colspan="2"><%IF itaxsellType="0" THEN%>공급자<%ELSE%>공급받는자<%END IF%></td>
				<td rowspan="2">합계금액</td>
				<td rowspan="2">공급가액</td>
				<td rowspan="2">세액</td>
				<td rowspan="2">분류</td>
				<td rowspan="2">종류</td>
				<td rowspan="2">품목명</td>
				<td rowspan="2">매핑<br>상태</td>
				<td rowspan="2">매핑<br>구분</td>
				<td rowspan="2">사업부문</td>
				<td rowspan="2">ERP<br>전송상태</td>
				<td rowspan="2">선택</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td>사업자등록번호</td>
				<!-- td>종</td -->
				<td>상호</td>
			</tr>
			<%
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
				%>
			<tr align="center" bgcolor="#FFFFFF">
			    <td><%= arrList(1,intLoop) %></td>
			    <td><a href="javascript:popErpSending('<%= arrList(0,intLoop) %>')"><%= arrList(0,intLoop) %></a></td>
			    <% if arrList(15,intLoop)=1 then %>
			    <td><a href="javascript:popHandMapping('<%= arrList(15,intLoop) %>','<%= arrList(1,intLoop) %>','<%= arrList(0,intLoop) %>','<%= arrList(7,intLoop) %>')"><%= arrList(7,intLoop) %></a></td>
			    <td><%= arrList(9,intLoop) %></td>
			    <% else %>
			    <td><a href="javascript:popHandMapping('<%= arrList(15,intLoop) %>','<%= arrList(1,intLoop) %>','<%= arrList(0,intLoop) %>','<%= arrList(2,intLoop) %>')"><%= arrList(2,intLoop) %></a></td>
			    <td><%= arrList(4,intLoop) %></td>
			    <% end if %>
			    <td align="right"><%= FormatNumber(arrList(12,intLoop),0) %></td>
			    <td align="right"><%= FormatNumber(arrList(13,intLoop),0) %></td>
			    <td align="right"><%= FormatNumber(arrList(14,intLoop),0) %></td>
			    <td><%= getSellTypeName(arrList(15,intLoop)) %></td>
			    <td><%= gettaxModiTypeName(arrList(16,intLoop)) %>/<%= gettaxTypeName(arrList(17,intLoop)) %></td>
			    <td><%= arrList(22,intLoop) %></td>
			    <td><%= getMatchStateName(arrList(31,intLoop)) %></td>
			    <td>
			        <% if (tgType="NRM") and Not IsNULL(arrList(31,intLoop)) and (arrList(29,intLoop)="9") then %>
			        <a href="javascript:jsSetTaxWithpayreq('<%=arrList(0,intLoop)%>','<%=arrList(16,intLoop)%>','<%=arrList(17,intLoop)%>','<%=arrList(1,intLoop)%>','<%=sReSearchText%>','<%=arrList(12,intLoop)%>','<%=arrList(13,intLoop)%>','<%=arrList(14,intLoop)%>','<%= arrList(30,intLoop) %>');"><%= getMatchTypeName(arrList(29,intLoop)) %><br><%= arrList(30,intLoop) %></a>
			        <% else %>
			        <%= getMatchTypeName(arrList(29,intLoop)) %><br><%= arrList(30,intLoop) %>
			        <% end if %>
			    </td>
			    <td><%= getbizSecCDName(arrList(32,intLoop)) %>
			    <% if arrList(35,intLoop)>0 then %>
			    외 <%= arrList(35,intLoop) %>
			    <% end if %>
			    </td>
			    <td>
			        <% if Not IsNULL(arrList(33,intLoop)) then %>
    			    [<%= arrList(33,intLoop) %>]
    			    <%= arrList(34,intLoop) %>
			        <% end if %>
			    </td>
			   <td><%Dim sReSearchText
			    	sReSearchText = replace(arrList(22,intLoop),"'","\'")
			    	sReSearchText = replace(sReSearchText,"""","")
			    	%>
			        <input <%= chkIIF(not IsNULL(arrList(31,intLoop)),"disabled","") %> type="button" class="button" value="선택" onClick="<%= CHKIIF(tgType="NRM","jsSetTaxNormal","jsSetTax") %>('<%=arrList(0,intLoop)%>','<%=arrList(16,intLoop)%>','<%=arrList(17,intLoop)%>','<%=arrList(1,intLoop)%>','<%=sReSearchText%>','<%=arrList(12,intLoop)%>','<%=arrList(13,intLoop)%>','<%=arrList(14,intLoop)%>')">

					<%
						if (C_ADMIN_AUTH) or (C_MngPart) then
							if (not IsNULL(arrList(31,intLoop))) and (not IsNULL(arrList(29,intLoop))) then
					%>
						<input type="button" class="button_auth" value="변경(관리자)" onClick="<%= CHKIIF(tgType="NRM","jsSetTaxNormal","jsSetTax") %>('<%=arrList(0,intLoop)%>','<%=arrList(16,intLoop)%>','<%=arrList(17,intLoop)%>','<%=arrList(1,intLoop)%>','<%=sReSearchText%>','<%=arrList(12,intLoop)%>','<%=arrList(13,intLoop)%>','<%=arrList(14,intLoop)%>')">
			        <%
							end if
						end if
					%>
			    </td>
			</tr>
		<%	Next
			ELSE%>
			<tr height=30 align="center" bgcolor="#FFFFFF">
				<td colspan="19">등록된 내용이 없습니다.</td>
			</tr>
			<%END IF%>
		</table>
	</td>
</tr>
<!-- 페이지 시작 -->
		<%
		Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10

		iStartPage = (Int((page-1)/iPerCnt)*iPerCnt) + 1

		If (page mod iPerCnt) = 0 Then
			iEndPage = page
		Else
			iEndPage = iStartPage + (iPerCnt-1)
		End If
		%>
			<tr height="25" >
				<td colspan="15" align="center">
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
					    <tr valign="bottom" height="25">
					        <td valign="bottom" align="center">
					         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
							<% else %>[pre]<% end if %>
					        <%
								for ix = iStartPage  to iEndPage
									if (ix > iTotalPage) then Exit for
									if Cint(ix) = Cint(page) then
							%>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
							<%		else %>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
							<%
									end if
								next
							%>
					    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
							<% else %>[next]<% end if %>
					        </td>
					    </tr>
					</table>
				</td>
			</tr>
</table>
<!-- 페이지 끝 -->
</body>
</html>




