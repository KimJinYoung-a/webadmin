<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  출고지시서 도식화 메뉴
' History : 이상구 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/tenbalju.asp"-->
<%

dim research, yyyy1,mm1,dd1,yyyymmdd,nowdate
dim pagesize

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")

pagesize = request("pagesize")

if yyyy1="" then
	nowdate = CStr(Now)
	nowdate = DateSerial(Left(nowdate,4), CLng(Mid(nowdate,6,2))-2,Mid(nowdate,9,2))
	yyyy1 = Left(nowdate,4)
	mm1 = Mid(nowdate,6,2)
	dd1 = Mid(nowdate,9,2)
end if

if (pagesize="") then pagesize=200

dim ojumun
set ojumun = new CTenBalju

ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
ojumun.FPageSize = pagesize

ojumun.GetBaljuItemHierachyProc

dim i, j, k

%>

<script>

/**
 * 수직방향으로 같은 값을 가진 cell 을 merge 한다.
 *

 * [IE 6.0], [FireFox 2.0]

 *
 * <입력 파라미터>
 * - table : Table 객체
 * - startRowIdx : 테이블의 몇 번째 row 에서부터 merge 를 수행할 지 결정하는 row's Index
 * - cellIdx : merge 하기 위한 테이블의 cell's Index
 *
 * <반환값>
 * - 없음
 *
 * ex) var table = document.getElementById("tbl");
 *     mergeVerticalCell(table, 0, 0);
 *
 */
function mergeVerticalCell(table, startRowIdx, cellIdx) {
  var rows            = table.getElementsByTagName("tr");
  var numRows         = rows.length;
  var numRowSpan      = 1;
  var currentRow      = null;
  var currentCell     = null;
  var currentCellData = null;
  var nextRow         = null;
  var nextCell        = null;
  var nextCellData    = null;

  for (var i = startRowIdx; i < (numRows-1); i++) {   // i 는 row's index

    // 새롭게 cell merge 를 해야하면,
    // 현재(비교의 기준이 되는..) row, cell, data 구함
    if (numRowSpan <= 1) {
      currentRow      = table.getElementsByTagName('tr')[i];
      currentCell     = currentRow.getElementsByTagName('td')[cellIdx];
      currentCellData = currentCell.childNodes[0].data;
    }


    if (i < numRows-1) {  // 현재 row 가 마지막 row 가 아니면

      // 다음 row, cell, data 구함
      if (table.getElementsByTagName('tr')[i+1]) {
        nextRow       = table.getElementsByTagName('tr')[i+1];
        nextCell      = nextRow.getElementsByTagName('td')[cellIdx];
        nextCellData  = nextCell.childNodes[0].data;

        // 현재 cell == 다음 cell 이면, merge
        if (currentCellData == nextCellData) {
          numRowSpan              += 1;
          currentCell.rowSpan     = numRowSpan;
          nextCell.style.display  = 'none';

        // 현재 cell != 다음 cell 이면,
        // 새로운 현재(비교의 기준이 되는..) cell 을 구할 수 있도록 초기화
        } else {
          numRowSpan = 1;

        }
      }
    }
  }
}


/**
 * 수직방향으로 같은 값을 가진 cell 을 merge 한다.
 * 단, mergeVerticalCell() 함수를 통해서 먼저 위탁 cell 들이 merge 된 이후,
 * merge 된 cell 을 기준으로 merge 된 cell 의 범위 내에 포함되는 row 의 cell 에 대해서만 merge 한다.
 *

 * [IE 6.0], [FireFox 2.0]

 *
 * <입력 파라미터>
 * - table : Table 객체
 * - startRowIdx : 테이블의 몇 번째 row 에서부터 merge 를 수행할 지 결정하는 row's Index
 * - basicCellIdx : 이미 merge 된 cell 중에서 기준이 되는 cell's index
 * - cellIdx : merge 하기 위한 테이블의 cell's Index
 *
 * <반환값>
 * - 없음
 *
 * ex) var table = document.getElementById("tbl");
 *     mergeVerticalCell(table, 0, 0);
 *     mergeDependentVerticalCell(table, 0, 0, 1);
 *
 */
function mergeDependentVerticalCell(table, startRowIdx, basicCellIdx, cellIdx) {
  var rows                  = table.getElementsByTagName("tr");
  var numRows               = rows.length;
  var numRowSpan            = 1;  // 초기화
  var countLoopInBasicMerge = 1;  // 초기화   merge 된 cell 내에서의 반복루프 처리 횟수
  var currentRow            = null;
  var currentCell           = null;
  var currentCellData       = null;
  var nextRow               = null;
  var nextCell              = null;
  var nextCellData          = null;
  var basicRowSpan          = null;

  for (var i = startRowIdx; i < (numRows-1); i++) {   // i 는 row's index

    // 기준 rowSpan 값 설정
    // basicCellIdx 에 해당하는 cell 의 rowSpan 값이 기준 rowSpan 범위가 됨.
    if (i == startRowIdx || (countLoopInBasicMerge== 1 && numRowSpan == 1)) {
      basicRowSpan  = table.getElementsByTagName('tr')[i].getElementsByTagName("td")[basicCellIdx].rowSpan;
    }

    // 새롭게 cell merge 를 해야하면,
    // 현재(비교의 기준이 되는..) row, cell, data 구함
    if (numRowSpan <= 1) {
      currentRow      = table.getElementsByTagName('tr')[i];
      currentCell     = currentRow.getElementsByTagName('td')[cellIdx];
      currentCellData = currentCell.childNodes[0].data;
    }


    if (i < numRows-1) {  // 현재 row 가 마지막 row 가 아니면

      if (countLoopInBasicMerge < basicRowSpan) {  // 기준 row 의 rowSpan 값을 초과해서 merge 할 수 없음.
        // 다음 row, cell, data 구함
        if (table.getElementsByTagName('tr')[i+1]) {
          nextRow       = table.getElementsByTagName('tr')[i+1];
          nextCell      = nextRow.getElementsByTagName('td')[cellIdx];
          nextCellData  = nextCell.childNodes[0].data;

          // 현재 cell == 다음 cell 이면, merge
          if (currentCellData == nextCellData) {
            numRowSpan              += 1;
            currentCell.rowSpan     = numRowSpan;
            nextCell.style.display  = 'none';

          // 현재 cell != 다음 cell 이면,
          // 새로운 현재(비교의 기준이 되는..) cell 을 구할 수 있도록 초기화
          } else {
            numRowSpan = 1;

          }
        }

        countLoopInBasicMerge++;

      // 기준 rowSpan 범위 이상이면,
      // 새로운 rowSpan 을 설정할 수 있도록 값을 초기화

      } else {
        countLoopInBasicMerge = 1;
        numRowSpan = 1;

      }
    }
  }
}

window.onload = function() {
    var table = document.getElementById("tbl");

    mergeVerticalCell(table, 1, 0);
    for (var i = 1; i < 14; i++) {
        mergeDependentVerticalCell(table, 1, (i - 1), i);
    }
}

function popOpenBaljuMaker(sitename, before15hour, excItem, danpumYN, boxGubun) {
    var frm = document.frm;
    var yyyy1, mm1, dd1;

    if (before15hour == undefined) {
        before15hour = '';
    }

    if (excItem == undefined) {
        excItem = '';
    }

    if (danpumYN == undefined) {
        danpumYN = '';
    }

    if (boxGubun == undefined) {
        boxGubun = '';
    }

    yyyy1 = frm.yyyy1.value;
    mm1 = frm.mm1.value;
    dd1 = frm.dd1.value;

	var popwin = window.open("/admin/ordermaster/_newbaljumaker.asp?extSiteName=" + sitename + "&yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&dd1=" + dd1 + "&before15hour=" + before15hour + "&excItem=" + excItem + "&danpumYN=" + danpumYN + "&boxGubun=" + boxGubun,"popOpenBaljuMaker","width=1700 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function jsPopNoSize() {
    var popwin = window.open("/admin/dataanalysis/report/simpleQry.asp?menupos=4116&qryidx=218","jsPopNoSize","width=600 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<form name="frm" method="get" >
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#F4F4F4" >
	    <td rowspan="2" width="50" bgcolor="#EEEEEE">검색<br>조건</td>
        <td align="left">
            * 기간 : <% DrawOneDateBox yyyy1,mm1,dd1 %> ~ 현재
            <!--
            &nbsp;
            * 텐바이텐배송 건수 :
			<select class="select" name="pagesize" >
				<option value="10" <% if pagesize="10" then response.write "selected" %> >10</option>
				<option value="20" <% if pagesize="20" then response.write "selected" %> >20</option>
				<option value="50" <% if pagesize="50" then response.write "selected" %> >50</option>
				<option value="100" <% if pagesize="100" then response.write "selected" %> >100</option>
				<option value="120" <% if pagesize="120" then response.write "selected" %> >120</option>
				<option value="150" <% if pagesize="150" then response.write "selected" %> >150</option>
				<option value="200" <% if pagesize="200" then response.write "selected" %> >200</option>
				<option value="250" <% if pagesize="250" then response.write "selected" %> >250</option>
				<option value="300" <% if pagesize="300" then response.write "selected" %> >300</option>
				<option value="400" <% if pagesize="400" then response.write "selected" %> >400</option>
				<option value="500" <% if pagesize="500" then response.write "selected" %> >500</option>
				<option value="600" <% if pagesize="600" then response.write "selected" %> >600</option>
				<option value="800" <% if pagesize="800" then response.write "selected" %> >800</option>
				<option value="1000" <% if pagesize="1000" then response.write "selected" %> >1000</option>
				<option value="2000" <% if pagesize="2000" then response.write "selected" %> >2000</option>
			</select>
            -->
        </td>
        <td rowspan="2" width="50" bgcolor="#EEEEEE">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p />

<input type="button" class="button" value="사이즈 미지정 상품/사은품 목록" onClick="jsPopNoSize()">

<p />

<table id="tbl" width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">-</td>
        <td width="100">건수</td>

        <td width="100">회사별</td>
        <td width="100">건수</td>

        <td width="150">시간구분</td>
        <td width="100">건수</td>

        <td width="100">주문구분</td>
        <td width="100">건수</td>

        <td width="100">SKU별</td>
        <td width="100">건수</td>

        <td width="100">박스구분</td>
        <td width="100">건수</td>

        <td width="100">박스상세</td>
        <td width="100">건수</td>

        <td>비고</td>
	</tr>
<% if (ojumun.FResultCount<1) then %>
	<tr bgcolor="#FFFFFF" height="31"><td colspan="25" align="center">내역이 없습니다.</td></tr>
<% else %>
<% for i=0 to ojumun.FResultCount-1 %>
    <tr bgcolor="#FFFFFF" align="center">
        <td>전체</td>
        <td><%= ojumun.FItemList(i).FtotOrderCnt %></td>

        <td><%= ojumun.FItemList(i).Fsitename %></td>
        <td><a href="javascript:popOpenBaljuMaker('<%= ojumun.FItemList(i).Fsitename %>')"><%= ojumun.FItemList(i).FtotSitenameOrderCnt %></a></td>

        <td>
            <%
            select case ojumun.FItemList(i).Fbefore15hour
                case "Y":
                    response.write "15시 이전"
                case "N":
                    response.write "15시 이후"
                case "B":
                    response.write "전일 15시 이전"
                case else:
                    response.write "ERR"
            end select
            %>
        </td>
        <td>
            <a href="javascript:popOpenBaljuMaker('<%= ojumun.FItemList(i).Fsitename %>', '<%= ojumun.FItemList(i).Fbefore15hour %>')"><%= ojumun.FItemList(i).FtotBefore15hourOrderCnt %></a>
        </td>

        <td>
            <%
            select case ojumun.FItemList(i).FexcItem
                case "Y":
                    response.write "제외주문"
                case "N":
                    response.write "정상주문"
                case else:
                    response.write "ERR"
            end select
            %>
        </td>
        <td><a href="javascript:popOpenBaljuMaker('<%= ojumun.FItemList(i).Fsitename %>', '<%= ojumun.FItemList(i).Fbefore15hour %>', '<%= ojumun.FItemList(i).FexcItem %>')"><%= ojumun.FItemList(i).FtotexcItemCnt %></a></td>

        <td>
            <%
            select case ojumun.FItemList(i).FdanpumYN
                case "Y":
                    response.write "단품"
                case "N":
                    response.write "합포장"
                case else:
                    response.write "ERR"
            end select
            %>
        </td>
        <td><a href="javascript:popOpenBaljuMaker('<%= ojumun.FItemList(i).Fsitename %>', '<%= ojumun.FItemList(i).Fbefore15hour %>', '<%= ojumun.FItemList(i).FexcItem %>', '<%= ojumun.FItemList(i).FdanpumYN %>')"><%= ojumun.FItemList(i).FtotdanpumYNCnt %></a></td>

        <td>
            <%= ojumun.FItemList(i).FboxGubun %>
        </td>
        <td>
            <a href="javascript:popOpenBaljuMaker('<%= ojumun.FItemList(i).Fsitename %>', '<%= ojumun.FItemList(i).Fbefore15hour %>', '<%= ojumun.FItemList(i).FexcItem %>', '<%= ojumun.FItemList(i).FdanpumYN %>', '<%= ojumun.FItemList(i).FboxGubun %>')">
                <%= ojumun.FItemList(i).FtotboxGubunCnt %>
            </a>
        </td>

        <td>
            <%= ojumun.FItemList(i).FboxGubunDetail %>
        </td>
        <td>
            <a href="javascript:popOpenBaljuMaker('<%= ojumun.FItemList(i).Fsitename %>', '<%= ojumun.FItemList(i).Fbefore15hour %>', '<%= ojumun.FItemList(i).FexcItem %>', '<%= ojumun.FItemList(i).FdanpumYN %>', '<%= ojumun.FItemList(i).FboxGubunDetail %>')">
                <%= ojumun.FItemList(i).FtotboxGubunDetailCnt %>
            </a>
        </td>

        <td></td>
	</tr>
<% next %>
<% end if %>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
