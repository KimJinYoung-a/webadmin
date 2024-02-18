<% if (IsDisplayCSMaster = true) then %>
	<%
	dim jupsugubun, jupsudefaulttitle

	jupsugubun = GetCSCommName("Z001", divcd)
	jupsudefaulttitle = ogiftcardordermaster.FOneItem.GetAccountdivName + " " + ogiftcardordermaster.FOneItem.GetJumunDivName + " 상태중 주문취소"

	%>
<tr >
    <td >
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr>
            <td bgcolor="<%= adminColor("topbar") %>" width="80" align="center">접수구분</td>
            <td bgcolor="#FFFFFF">
			    	<font style='line-height:100%; font-size:15px; color:blue; font-family:돋움; font-weight:bold'><%= jupsugubun %></font>
			    	&nbsp;
                <% if (Not IsStatusRegister) then %>
			    	<font style='line-height:100%; font-size:15px; color:#CC3333; font-family:돋움; font-weight:bold'>[<%= ocsaslist.FOneItem.GetCurrstateName %>]</font>
			    	<% if ocsaslist.FOneITem.FDeleteyn<>"N" then %>
						<font style='line-height:100%; font-size:15px; color:#FF0000; font-family:돋움; font-weight:bold'>- 삭제된 내역</font>
			    	<% end if %>
		    	<% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" width="80" align="center">주문번호</td>
            <td bgcolor="#FFFFFF" width="200" >
                <%= giftorderserial %>
                [<font color="<%= ogiftcardordermaster.FOneItem.CancelYnColor %>"><%= ogiftcardordermaster.FOneItem.CancelYnName %></font>]
                [<font color="<%= ogiftcardordermaster.FOneItem.IpkumDivColor %>"><%= ogiftcardordermaster.FOneItem.GetJumunDivName %></font>]
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">접수자</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsStatusRegister) then %>
                    <%= session("ssbctid") %>
                <% else %>
                    <%= ocsaslist.FOneItem.Fwriteuser %>
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">주문자ID</td>
            <td bgcolor="#FFFFFF">
                <%= ogiftcardordermaster.FOneItem.FUserID %>(<font color="<%= ogiftcardordermaster.FOneItem.GetUserLevelColor %>"><%= ogiftcardordermaster.FOneItem.GetUserLevelName %></font>)
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">접수일시</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsStatusRegister) then %>
                	<%= now() %>
                <% else %>
                	<%= ocsaslist.FOneItem.Fregdate %>
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">주문자정보</td>
            <td bgcolor="#FFFFFF">
                <%= ogiftcardordermaster.FOneItem.FBuyname %>
                 &nbsp;
                 [<%= ogiftcardordermaster.FOneItem.FBuyHp %>]
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">접수제목</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsStatusRegister) then %>
                	<input <% if IsStatusFinishing then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= jupsudefaulttitle %>" size="56" maxlength="56">
                <% else %>
                	<input <% if IsStatusFinishing then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= ocsaslist.FOneItem.Ftitle %>" size="56" maxlength="56">
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">수령인정보</td>
            <td bgcolor="#FFFFFF">
                 [<%= ogiftcardordermaster.FOneItem.FReqHp %>]
            </td>
        </tr>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">사유구분</td>
            <td bgcolor="#FFFFFF">
                <input type="hidden" name="gubun01" value="<%= ocsaslist.FOneItem.Fgubun01 %>">
                <input type="hidden" name="gubun02" value="<%= ocsaslist.FOneItem.Fgubun02 %>">
                <input class="text_ro" type="text" name="gubun01name" value="<%= ocsaslist.FOneItem.Fgubun01name %>" size="16" Readonly >
                &gt;
                <input class="text_ro" type="text" name="gubun02name" value="<%= ocsaslist.FOneItem.Fgubun02name %>" size="16" Readonly >
                <input class="csbutton" type="button" value="선택" onClick="divCsAsGubunSelect(frmaction.gubun01.value, frmaction.gubun02.value, frmaction.gubun01.name, frmaction.gubun02.name, frmaction.gubun01name.name, frmaction.gubun02name.name,'frmaction','causepop');">
                <div id="causepop" style="position:absolute;"></div>

                <!-- 일부 사유 미리 표시 -->
                <%
                '참조쿼리
				'select top 100 m.comm_cd, m.comm_name, d.comm_cd, d.comm_name
				'from
				'	db_cs.dbo.tbl_cs_comm_code m
				'	left join db_cs.dbo.tbl_cs_comm_code d
				'	on
				'		m.comm_cd = d.comm_group
				'where
				'	1 = 1
				'	and m.comm_group = 'Z020'
				'	and m.comm_isdel <> 'Y'
				'	and d.comm_isdel <> 'Y'
				'order by m.comm_cd, d.comm_cd
                %>
                <% if (ocsaslist.FOneItem.IsCancelProcess) then %>
	                [<a href="javascript:selectGubun('C004','CD01','공통','단순변심','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">단순변심</a>]
	                [<a href="javascript:selectGubun('C004','CD05','공통','품절','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">품절</a>]
	                [<a href="javascript:selectGubun('C004','CD99','공통','기타','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">기타</a>]

                <% elseif (ocsaslist.FOneItem.IsReturnProcess) then %>
	                [<a href="javascript:selectGubun('C004','CD01','공통','단순변심','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">단순변심</a>]
	                [<a href="javascript:selectGubun('C005','CE01','상품관련','상품불량','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">상품불량</a>]
	                [<a href="javascript:selectGubun('C006','CF01','물류관련','오발송','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">오배송</a>]
                    [<a href="javascript:selectGubun('C004','CD04','공통','사이즈교환','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">사이즈교환</a>]
                    [<a href="javascript:selectGubun('C004','CD06','공통','사이즈 안맞음','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">사이즈 안맞음(고객변심)</a>]
                <% elseif (divcd="A009") or (divcd="A006") or (divcd="A700") or (divcd="A900") then %>
                	[<a href="javascript:selectGubun('C004','CD99','공통','기타','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">기타</a>]

                <% elseif (divcd="A001") then %>
                	[<a href="javascript:selectGubun('C006','CF03','물류관련','구매상품누락','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">상품누락</a>]

                <% elseif (divcd="A002") then %>
	                [<a href="javascript:selectGubun('C006','CF04','물류관련','사은품누락','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">(물류)사은품누락</a>]
	                [<a href="javascript:selectGubun('C005','CE05','상품관련','이벤트오등록','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">(MD)이벤트오등록</a>]

                <% elseif (divcd="A000") then %>
	                [<a href="javascript:selectGubun('C005','CE01','상품관련','상품불량','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">상품불량</a>]
	                [<a href="javascript:selectGubun('C006','CF01','물류관련','오발송','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">오발송</a>]
	                [<a href="javascript:selectGubun('C006','CF02','물류관련','상품파손','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">상품파손</a>]
	                [<a href="javascript:selectGubun('C004','CD04','공통','사이즈교환','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">사이즈교환</a>]
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">결제정보</td>
            <td bgcolor="#FFFFFF">
            	<%= FormatNumber(ogiftcardordermaster.FOneItem.Fsubtotalprice,0) %>원
            	&nbsp;
                [<%= ogiftcardordermaster.FOneItem.GetAccountdivName %>]
            </td>
        </tr>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center" rowspan="2">접수내용</td>
            <td bgcolor="#FFFFFF" rowspan="2">
            	<textarea <% if IsStatusFinishing then response.write "class='textarea_ro' ReadOnly" else response.write "class='textarea'" end if %> name="contents_jupsu" cols="68" rows="6"><%= ocsaslist.FOneItem.Fcontents_jupsu %></textarea>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">배송지정보</td>
            <td bgcolor="#FFFFFF" valign="top">
            	[<%= ogiftcardordermaster.FOneItem.FReqEmail %>]<br>
            </td>
        </tr>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center"></td>
            <td bgcolor="#FFFFFF" valign="top">

            </td>
        </tr>
        <% if (IsStatusFinishing) or (IsStatusFinished) then %>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">처리내용</td>
            <td bgcolor="#FFFFFF">
            	<textarea class='textarea' name="contents_finish" cols="68" rows="7"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
            </td>
            <td bgcolor="<%= adminColor("pink") %>" align="center">처리관련<br>고객오픈<br>내용입력</td>
            <td bgcolor="#FFFFFF">
            	<table border="0" cellspacing="0" cellpadding="0" class="a" valign="top">
            	<tr>
				    <td>
				    	<input class="text" type="text" name="opentitle" value="<%= ocsaslist.FOneItem.Fopentitle %>" size="48" maxlength="60" readonly>
				    </td>
				</tr>
				<tr>
				    <td>
				    	<textarea class="textarea" name="opencontents" cols="48" rows="5" readonly><%= ocsaslist.FOneItem.Fopencontents %></textarea>
				    </td>
				</tr>
				</table>
			</td>
        </tr>
        <% end if %>
        </table>
	</td>
</tr>
<% end if %>
