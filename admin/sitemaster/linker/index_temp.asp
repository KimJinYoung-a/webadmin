<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : LINKER ����
'	History		: 2021.10.14 ������ ����
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
</p>
<link rel="stylesheet" type="text/css" href="/css/linker.css">
<link rel="stylesheet" href="https://cdn.materialdesignicons.com/3.6.95/css/materialdesignicons.min.css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">

<div class="container">
    <header class="linker-title">
        <h3>LINKER ����</h3>
        <div class="nickname-btn-area">
            <a href="#">�г��� ����</a>
            <a href="#">�г��� ��Ӿ� ����</a>
        </div>
    </header>

    <div class="linker-content">
        <div class="forum-list-container">
            <div class="title">
                <h5>����</h5>
                <div class="btn-area">
                    <button class="linker-btn">���� ���</button>
                    <button class="linker-btn">���� ��������</button>
                </div>
            </div>

            <!-- region ���� ����Ʈ -->
            <ul class="forum-list">
                <li class="on">
                    <p class="number">03</p>
                    <div class="info">
                        <strong>�ٹ����� 20�ֳ��� �������ּ���!</strong>
                        <span>2021-09-09 ~ 2021-09-31 / ����</span>
                    </div>
                </li>
                <li>
                    <p class="number">02</p>
                    <div class="info">
                        <strong>��Ų��������! ���� �޾ư�����!</strong>
                        <span>2021-09-09 ~ 2021-09-31 / ���¾���</span>
                    </div>
                </li>
            </ul>
            <!-- endregion -->

        </div>
        <div class="forum-content">

            <!-- region ��� ���� ���� -->
            <div class="forum-content-top">
                <div class="title-control">
                    <div class="title">
                        <p>�ƴ� ����?</p>
                        <h5>�ٹ����� 20�ֳ��� �������ּ���!</h5>
                    </div>
                    <div class="btn-area">
                        <button class="linker-btn">���� ����</button>
                        <button class="linker-btn">�������� ���</button>
                        <button class="linker-btn">���� ����</button>
                    </div>
                </div>
                <div class="title-info">
                    <span>��Ⱓ : 2021-09-09 ~ 2021-09-31</span>
                    <span>����Ʈ ���¿��� : ����</span>
                    <span>���� ���� : 2</span>
                </div>
            </div>
            <!-- endregion -->

            <div class="forum-info">
                <div class="title">
                    <div>
                        <h3>���� �ȳ�</h3>
                        <span>���� �ȳ��� 5�������� ����� �� �ֽ��ϴ�.</span>
                    </div>
                    <div>
                        <button class="linker-btn">���� �ȳ� ���</button>
                        <button class="linker-btn">���ļ���</button>
                        <button class="linker-btn">���� �׸� ����</button>
                    </div>
                </div>

                <table class="forum-list-tbl">
                    <colgroup>
                        <col style="width: 50px;">
                        <col style="width: 100px;">
                        <col style="width: 300px;">
                        <col>
                    </colgroup>
                    <thead>
                        <tr>
                            <th><input type="checkbox"></th>
                            <th>�������</th>
                            <th>�ȳ�����</th>
                            <th></th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td><input type="checkbox"></td>
                            <td>1</td>
                            <td class="tl" colspan="2">���� �� �̾߱� #2 �ٹ����ٰ� �Բ� �ڶ�� '��ġ����Ŀ'</td>
                        </tr>
                    </tbody>
                </table>
            </div>

            <div class="forum-content-bottom">

                <div class="title">
                    <h3>������ ����Ʈ</h3>
                    <a>�Ű� ������ ����</a>
                </div>

                <!-- region �˻� -->
                <div class="search">
                    <div>
                        <div class="search-group">
                            <label>ȸ������:</label>
                            <select>
                                <option>��ü</option>
                                <option>Host</option>
                                <option>Guest</option>
                                <option>User</option>
                            </select>
                        </div>
                        <div class="search-group">
                            <label>ȸ�����:</label>
                            <select>
                                <option>��ü</option>
                                <option>STAFF</option>
                                <option>VVIP</option>
                                <option>VIP GOLD</option>
                            </select>
                        </div>
                        <div class="search-group">
                            <select>
                                <option>�г���</option>
                            </select>
                            :
                            <input type="text">
                        </div>
                        <div class="search-group">
                            <label>�������:</label>
                            <input type="text" class="date" readonly> ~
                            <input type="text" class="date" readonly>
                        </div>
                    </div>
                    <button class="linker-btn">�˻�</button>
                </div>
                <!-- endregion -->

                <div class="forum-posting-result">
                    <div class="forum-posting-top">
                        <p>�˻���� : <span>1,312</span></p>
                        <div>
                            <button class="linker-btn">���� �׸� ����</button>
                            <button class="linker-btn">���� ������ ����</button>
                        </div>
                    </div>

                    <!-- region ������ ����Ʈ -->
                    <table class="forum-list-tbl">
                        <colgroup>
                            <col style="width: 50px;">
                            <col style="width: 70px;">
                            <col style="width: 180px;">
                            <col>
                            <col style="width: 110px;">
                            <col style="width: 150px;">
                            <col style="width: 170px;">
                        </colgroup>
                        <thead>
                            <tr>
                                <th><input type="checkbox"></th>
                                <th>idx</th>
                                <th>�ۼ��� ����</th>
                                <th>�ۼ�����</th>
                                <th>��� ��������</th>
                                <th>�ۼ�����</th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><input type="checkbox"></td>
                                <td>1758</td>
                                <td>Host / STAFF / ������¯</td>
                                <td>��ġ�ھ�� �쿬�� ġ������ �Ϳ����� ã�ƹ��ȴµ� �Ʊ���� ��������. �� �Ϳ���� ��¥ �ְ�� XD</td>
                                <td class="posting-red">Y</td>
                                <td>2021-09-08 15:13:15</td>
                                <td>
                                    <button class="linker-btn">����</button>
                                    <button class="linker-btn">����</button>
                                </td>
                            </tr>
                            <tr>
                                <td><input type="checkbox"></td>
                                <td>1757</td>
                                <td>User / VIP GOLD / ������¯</td>
                                <td>��ġ�ھ�� �쿬�� ġ������ �Ϳ����� ã�ƹ��ȴµ� �Ʊ���� ��������. �� �Ϳ���� ��¥ �ְ�� XD</td>
                                <td>N</td>
                                <td>2021-09-08 15:13:15</td>
                                <td>
                                    <button class="linker-btn">����</button>
                                    <button class="linker-btn">����</button>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                    <!-- endregion -->

                    <!-- region ����¡ -->
                    <ul class="pagination">
                        <li class="disabled"><a>&lt;</a></li>
                        <li class="on"><a>1</a></li>
                        <li><a>2</a></li>
                        <li><a>3</a></li>
                        <li><a>&gt;</a></li>
                    </ul>
                    <!-- endregion -->

                </div>
            </div>
        </div>
    </div>


    <!-- region ���� �űԵ�� ��� -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>���� �űԵ��</h3>
            </div>
            <div class="modal-container">
                <div>
                    <table class="modal-write-tbl">
                        <colgroup>
                            <col style="width:100px;">
                            <col>
                        </colgroup>
                        <tbody>
                            <tr>
                                <th>����</th>
                                <td><input type="text" placeholder="���� ������ �Է����ּ���"></td>
                            </tr>
                            <tr>
                                <th>������</th>
                                <td><input type="text" placeholder="���� �������� �Է����ּ���"></td>
                            </tr>
                            <tr>
                                <th>����</th>
                                <td><textarea placeholder="���� ������ �Է����ּ���"></textarea></td>
                            </tr>
                            <tr>
                                <th>��׶���<br>PC</th>
                                <td>
                                    <p class="radio-area">
                                        <input id="backPcImage" type="radio" checked>
                                        <label for="backPcImage">�̹���</label>
                                        <input id="backPcVideo" type="radio">
                                        <label for="backPcVideo">������</label>
                                    </p>
                                    <button class="linker-btn">�̹��� ÷��</button>
                                </td>
                            </tr>
                            <tr>
                                <th>��׶���<br>M</th>
                                <td>
                                    <p class="radio-area">
                                        <input id="backPcImage" type="radio" checked>
                                        <label for="backPcImage">�̹���</label>
                                        <input id="backPcVideo" type="radio">
                                        <label for="backPcVideo">������</label>
                                    </p>
                                    <input type="text" placeholder="���� URL�� �Է����ּ���">
                                </td>
                            </tr>
                            <tr>
                                <th>��Ⱓ</th>
                                <td>
                                    <span class="datepicker">
                                        <label for="datepicker1">
                                            <strong>������</strong>
                                            <span class="mdi mdi-calendar-month"></span>
                                        </label>
                                        <input type="text" id="datepicker1" readonly>

                                        <label for="datepicker2">
                                            <strong>������</strong>
                                            <span class="mdi mdi-calendar-month"></span>
                                        </label>
                                        <input type="text" id="datepicker2" readonly>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <th>����Ʈ<br>���⿩��</th>
                                <td>
                                    <p class="radio-area">
                                        <input id="showY" type="radio" checked>
                                        <label for="showY">Y</label>
                                        <input id="showN" type="radio">
                                        <label for="showN">N</label>
                                    </p>
                                </td>
                            </tr>
                            <tr>
                                <th>���ļ���</th>
                                <td><input type="text" style="width: 100px;"></td>
                            </tr>
                            <tr>
                                <th>���</th>
                                <td><textarea></textarea></td>
                            </tr>
                        </tbody>
                    </table>

                    <div class="modal-btn-area">
                        <button class="linker-btn">����</button>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region ���� ���� �������� ��� -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap" style="width: 1100px;">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>���� ���� ���� ����</h3>
            </div>
            <div class="modal-container">
                <div class="forum-posting-result">
                    <div class="forum-posting-top">
                        <p><span>1,312</span>��</p>
                        <div>
                            <button class="linker-btn">���ļ��� ����</button>
                        </div>
                    </div>

                    <table class="forum-list-tbl">
                        <colgroup>
                            <col style="width: 30px;">
                            <col style="width: 60px;">
                            <col>
                            <col style="width: 200px;">
                            <col style="width: 120px;">
                            <col style="width: 70px;">
                            <col style="width: 160px;">
                            <col style="width: 140px;">
                        </colgroup>
                        <thead>
                            <tr>
                                <th><input type="checkbox"></th>
                                <th>NO.</th>
                                <th>���� ����</th>
                                <th>���� ������</th>
                                <th>����Ʈ ���� ����</th>
                                <th>���ļ���</th>
                                <th>��Ⱓ</th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><input type="checkbox"></td>
                                <td>1758</td>
                                <td>��Ų�� �Բ��ϴ� ��Ų��������!</td>
                                <td>���� 20�ֳ� �̺�Ʈ</td>
                                <td>����</td>
                                <td><input type="text" class="forum-sort"></td>
                                <td>2021-07-13 ~ 2021-07-31</td>
                                <td>
                                    <button class="linker-btn">����</button>
                                    <button class="linker-btn">����</button>
                                </td>
                            </tr>
                            <tr>
                                <td><input type="checkbox"></td>
                                <td>1758</td>
                                <td>��Ų�� �Բ��ϴ� ��Ų��������!</td>
                                <td>���� 20�ֳ� �̺�Ʈ</td>
                                <td>����</td>
                                <td><input type="text" class="forum-sort"></td>
                                <td>2021-07-13 ~ 2021-07-31</td>
                                <td>
                                    <button class="linker-btn">����</button>
                                    <button class="linker-btn">����</button>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region ���� �ȳ� ��� -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap" style="width: 750px;">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>���� �ȳ�</h3>
            </div>
            <div class="modal-container">
                <div>
                    <table class="modal-write-tbl">
                        <colgroup>
                            <col style="width:120px;">
                            <col>
                        </colgroup>
                        <tbody>
                            <tr>
                                <th>���� �ȳ� ����</th>
                                <td>
                                    <p class="radio-area">
                                        <input id="descrCount1" type="radio" name="descrCount" checked>
                                        <label for="descrCount1">1��</label>
                                        <input id="descrCount2" type="radio" name="descrCount">
                                        <label for="descrCount2">2��</label>
                                        <input id="descrCount3" type="radio" name="descrCount">
                                        <label for="descrCount3">3��</label>
                                        <input id="descrCount4" type="radio" name="descrCount">
                                        <label for="descrCount4">4��</label>
                                        <input id="descrCount5" type="radio" name="descrCount">
                                        <label for="descrCount5">5��</label>
                                    </p>
                                    <p class="descr">���� ������ 2�� �̻��� ���, �˾����� ǥ��˴ϴ�</p>
                                </td>
                            </tr>
                            <tr>
                                <th>���� �ȳ�1</th>
                                <td>
                                    <p class="forum-descr-title">
                                        <input type="text" placeholder="���� �ȳ� 1�� ������ �Է����ּ���">
                                        <button class="linker-btn">�̸�����</button>
                                    </p>
                                    <textarea class="forum-descr-code" rows="8" placeholder="���� ������ �ڵ�� �Է����ּ���"></textarea>
                                </td>
                            </tr>
                            <tr>
                                <th>���� �ȳ�2</th>
                                <td>
                                    <p class="forum-descr-title">
                                        <input type="text" placeholder="���� �ȳ� 2�� ������ �Է����ּ���">
                                        <button class="linker-btn">�̸�����</button>
                                    </p>
                                    <textarea class="forum-descr-code" rows="8" placeholder="���� ������ �ڵ�� �Է����ּ���"></textarea>
                                </td>
                            </tr>
                            <tr>
                                <th>�����ڵ�</th>
                                <td>
                                    <textarea class="forum-descr-sample" rows="6" wrap="off" readonly></textarea>
                                </td>
                            </tr>
                        </tbody>
                    </table>

                    <div class="modal-btn-area">
                        <button class="linker-btn">����</button>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region ���� �ȳ� �̸����� ��� -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap forum-descr-preview" style="width: 400px;">
            <div class="modal-header"></div>
            <div class="modal-body">
                <div class="modal-cont">
                    <div class="ex_img">
                        <img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_01.gif" alt="since 2001">
                        <img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_02.jpg?v=2" alt="�ٹ�����, �ӱ��ſ� �ϳ� �����ΰ�?">
                        <img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_03.gif" alt="�ӱ���">
                        <img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_04.jpg?v=2" alt="20��° �ӱ��� ���� ����!">
                        <img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_05.gif?v=2.1" alt="�ӱ��ŵ��� ������ �����?">
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region ������ ���� ��� -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>������ ����</h3>
            </div>
            <div class="modal-container">
                <div>
                    <table class="modal-write-tbl">
                        <colgroup>
                            <col style="width:120px;">
                            <col>
                        </colgroup>
                        <tbody>
                            <tr>
                                <th>idx</th>
                                <td class="content">1234</td>
                            </tr>
                            <tr>
                                <th>ȸ������</th>
                                <td class="content">User</td>
                            </tr>
                            <tr>
                                <th>ȸ�����</th>
                                <td class="content">VIP</td>
                            </tr>
                            <tr>
                                <th>�г���</th>
                                <td class="content">����Ѵٶ���#123</td>
                            </tr>
                            <tr>
                                <th>�ۼ�����</th>
                                <td><textarea rows="5">��ġ�ھ�� �쿬�� ġ������ �Ϳ����� ã�ƹ��ȴµ� �Ʊ���� ��������. �� �Ϳ���� ��¥ �ְ�� XD</textarea></td>
                            </tr>
                            <tr>
                                <th>����������</th>
                                <td><button class="linker-btn">�̺�Ʈ : 348282</button></td>
                            </tr>
                            <tr>
                                <th>��� ���� ����</th>
                                <td>
                                    <input id="fixPostingY" type="radio" name="fixPosting" checked>
                                    <label for="fixPostingY">����</label>
                                    <input id="fixPostingN" type="radio" name="fixPosting">
                                    <label for="fixPostingN">��������</label>
                                </td>
                            </tr>
                            <tr>
                                <th>�ۼ��Ͻ�</th>
                                <td class="content">
                                    <strong>2021-09-08 15:13:15</strong>
                                    <span class="posting-update">2021-09-10 14:11:12 ����</span>
                                </td>
                            </tr>
                        </tbody>
                    </table>

                    <div class="modal-btn-area">
                        <button class="linker-btn">����</button>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region ������ ���� ��� -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>������ ����</h3>
                <p class="add-descr">���� �׸��� �������� ��쿡�� �ϰ� ����˴ϴ�.</p>
            </div>
            <div class="modal-container">
                <div>

                    <table class="modal-write-tbl">
                        <colgroup>
                            <col style="width:120px;">
                            <col>
                        </colgroup>
                        <tbody>
                            <tr>
                                <th><span class="required">���� ����<i></i></span></th>
                                <td><input type="text" style="width: 50px;"></td>
                            </tr>
                            <tr>
                                <th><span class="required">���� �Ⱓ<i></i></span></th>
                                <td>
                                    <span class="datepicker">
                                        <label for="datepicker3">
                                            <strong>������</strong>
                                            <span class="mdi mdi-calendar-month"></span>
                                        </label>
                                        <input type="text" id="datepicker3" readonly>

                                        <label for="datepicker4">
                                            <strong>������</strong>
                                            <span class="mdi mdi-calendar-month"></span>
                                        </label>
                                        <input type="text" id="datepicker4" readonly>
                                    </span>
                                </td>
                            </tr>
                        </tbody>
                    </table>

                    <div class="modal-btn-area">
                        <button class="linker-btn">����</button>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region ���� ������ ���� ��� -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap" style="width: 1100px;">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>���� ������ ����</h3>
            </div>
            <div class="modal-container">
                <div class="forum-posting-result">
                    <div class="forum-posting-top">
                        <p><span>1,312</span>��</p>
                        <div>
                            <button class="linker-btn">���� ����</button>
                            <button class="linker-btn">������� ����</button>
                        </div>
                    </div>

                    <table class="forum-list-tbl">
                        <colgroup>
                            <col style="width: 30px;">
                            <col style="width: 60px;">
                            <col style="width: 170px;">
                            <col>
                            <col style="width: 100px;">
                            <col style="width: 160px;">
                            <col style="width: 170px;">
                        </colgroup>
                        <thead>
                            <tr>
                                <th><input type="checkbox"></th>
                                <th>idx</th>
                                <th>�ۼ��� ����</th>
                                <th>�ۼ�����</th>
                                <th>���� ���� ����</th>
                                <th>���� �Ⱓ</th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><input type="checkbox"></td>
                                <td>1758</td>
                                <td>Host / STAFF / ������</td>
                                <td>��ġ�ھ�� �쿬�� ġ������ �Ϳ����� ã�ƹ��ȴµ� �Ʊ���� ��������. �� �Ϳ���� ��¥ �ְ�� XD</td>
                                <td><input type="text" class="forum-sort" value="1"></td>
                                <td>2021-07-13 ~ 2021-07-31</td>
                                <td>
                                    <button class="linker-btn">����</button>
                                    <button class="linker-btn long">���� ����</button>
                                </td>
                            </tr>
                            <tr>
                                <td><input type="checkbox"></td>
                                <td>1758</td>
                                <td>User / VIP / ������¯</td>
                                <td>��ġ�ھ�� �쿬�� ġ������ �Ϳ����� ã�ƹ��ȴµ� �Ʊ���� ��������. �� �Ϳ���� ��¥ �ְ�� XD</td>
                                <td><input type="text" class="forum-sort" value="2"></td>
                                <td>2021-07-13 ~ 2021-07-31</td>
                                <td>
                                    <button class="linker-btn">����</button>
                                    <button class="linker-btn long">���� ����</button>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region �Ű� ������ ���� ��� -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap" style="width: 1100px;">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>�Ű� ������ ����</h3>
            </div>
            <div class="modal-container">
                <div class="forum-posting-result">
                    <div class="forum-posting-top">
                        <p><span>1,312</span>��</p>
                        <div>
                            <button class="linker-btn">��� ����</button>
                            <button class="linker-btn">���� ������ ����</button>
                        </div>
                    </div>

                    <table class="forum-list-tbl">
                        <colgroup>
                            <col style="width: 30px;">
                            <col style="width: 60px;">
                            <col style="width: 80px;">
                            <col style="width: 170px;">
                            <col>
                            <col style="width: 170px;">
                            <col style="width: 200px;">
                        </colgroup>
                        <thead>
                            <tr>
                                <th><input type="checkbox"></th>
                                <th>idx</th>
                                <th>������ �̹���</th>
                                <th>�ۼ��� ����</th>
                                <th>�ۼ�����</th>
                                <th>����� ������</th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><input type="checkbox"></td>
                                <td>1758</td>
                                <td><img src="//fiximage.10x10.co.kr/web2015/common/img_profile_04.png" class="thumb"></td>
                                <td>Host / STAFF / ������</td>
                                <td>��ġ�ھ�� �쿬�� ġ������ �Ϳ����� ã�ƹ��ȴµ� �Ʊ���� ��������. �� �Ϳ���� ��¥ �ְ�� XD</td>
                                <td><button class="linker-btn link">�̺�Ʈ : 348282</button></td>
                                <td>
                                    <button class="linker-btn long">��� ����</button>
                                    <button class="linker-btn long">������ ����</button>
                                </td>
                            </tr>
                            <tr>
                                <td><input type="checkbox"></td>
                                <td>1758</td>
                                <td><img src="//fiximage.10x10.co.kr/web2015/common/img_profile_04.png" class="thumb"></td>
                                <td>Host / STAFF / ������</td>
                                <td>��ġ�ھ�� �쿬�� ġ������ �Ϳ����� ã�ƹ��ȴµ� �Ʊ���� ��������. �� �Ϳ���� ��¥ �ְ�� XD</td>
                                <td><button class="linker-btn link">�ܺ� URL</button></td>
                                <td>
                                    <button class="linker-btn long">��� ����</button>
                                    <button class="linker-btn long">������ ����</button>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region �г��� ���� ��� -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap" style="width: 900px;">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>�г��� ����</h3>
            </div>
            <div class="modal-container">
                <div>
                    <div class="search">
                        <div>
                            <div class="search-group">
                                <select>
                                    <option>�ܾ�1</option>
                                    <option>�ܾ�2</option>
                                </select>
                                :
                                <input type="text">
                            </div>
                        </div>
                        <button class="linker-btn">�˻�</button>
                    </div>

                    <div class="modal-nicknames-area">
                        <div class="modal-nicknames-content">
                            <div class="nicknames-btn-area">
                                <button class="linker-btn">�űԵ��</button>
                                <button class="linker-btn">����</button>
                            </div>

                            <table class="forum-list-tbl">
                                <colgroup>
                                    <col style="width: 50px;">
                                    <col style="width: 60px;">
                                    <col>
                                    <col style="width: 150px;">
                                </colgroup>
                                <thead>
                                    <tr>
                                        <th><input type="checkbox"></th>
                                        <th>NO.</th>
                                        <th>�ܾ�1</th>
                                        <th></th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td><input type="checkbox"></td>
                                        <td>112</td>
                                        <td>�����</td>
                                        <td>
                                            <button class="linker-btn">����</button>
                                            <button class="linker-btn">����</button>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td><input type="checkbox"></td>
                                        <td>112</td>
                                        <td>�ǵ�Ÿ���</td>
                                        <td>
                                            <button class="linker-btn">����</button>
                                            <button class="linker-btn">����</button>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td><input type="checkbox"></td>
                                        <td>112</td>
                                        <td>������</td>
                                        <td>
                                            <button class="linker-btn">����</button>
                                            <button class="linker-btn">����</button>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>

                        <div class="modal-nicknames-content">
                            <div class="nicknames-btn-area">
                                <button class="linker-btn">�űԵ��</button>
                                <button class="linker-btn">����</button>
                            </div>

                            <table class="forum-list-tbl">
                                <colgroup>
                                    <col style="width: 50px;">
                                    <col style="width: 60px;">
                                    <col>
                                    <col style="width: 150px;">
                                </colgroup>
                                <thead>
                                    <tr>
                                        <th><input type="checkbox"></th>
                                        <th>NO.</th>
                                        <th>�ܾ�2</th>
                                        <th></th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td><input type="checkbox"></td>
                                        <td>112</td>
                                        <td>�ٶ���</td>
                                        <td>
                                            <button class="linker-btn">����</button>
                                            <button class="linker-btn">����</button>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td><input type="checkbox"></td>
                                        <td>112</td>
                                        <td>�ܽ���</td>
                                        <td>
                                            <button class="linker-btn">����</button>
                                            <button class="linker-btn">����</button>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td><input type="checkbox"></td>
                                        <td>112</td>
                                        <td>������</td>
                                        <td>
                                            <button class="linker-btn">����</button>
                                            <button class="linker-btn">����</button>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>

                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region �ܾ� ��� ��� -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>�ܾ�1 ���</h3>
            </div>
            <div class="modal-container">
                <div>
                    <table class="modal-write-tbl">
                        <colgroup>
                            <col style="width:120px;">
                            <col>
                        </colgroup>
                        <tbody>
                            <tr>
                                <th>�ܾ�1</th>
                                <td>
                                    <textarea rows="3"></textarea>
                                    <p class="descr">���� �ܾ �߰��� ��� ',(��ǥ)'�� �����Ͽ� �Է����ּ���.</p>
                                </td>
                            </tr>
                        </tbody>
                    </table>

                    <div class="modal-btn-area">
                        <button class="linker-btn">����</button>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->

    <!-- region �г��� ��Ӿ� ���� ��� -->
    <div class="modal" style="display: none;">
        <div class="modal-overlay"></div>
        <div class="modal-wrap" style="width: 900px;">
            <button class="modal-close-btn"></button>
            <div class="modal-title">
                <h3>�г��� ��Ӿ� ����</h3>
            </div>
            <div class="modal-container">
                <div>
                    <div class="search">
                        <div>
                            <div class="search-group">
                                <label>��Ӿ�:</label>
                                <input type="text">
                            </div>
                        </div>
                        <button class="linker-btn">�˻�</button>
                    </div>

                    <div>
                        <div class="nicknames-btn-area">
                            <button class="linker-btn">�űԵ��</button>
                            <button class="linker-btn">����</button>
                        </div>

                        <table class="forum-list-tbl">
                            <colgroup>
                                <col style="width: 50px;">
                                <col style="width: 60px;">
                                <col>
                                <col style="width: 150px;">
                            </colgroup>
                            <thead>
                                <tr>
                                    <th><input type="checkbox"></th>
                                    <th>NO.</th>
                                    <th>��Ӿ�</th>
                                    <th></th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr>
                                    <td><input type="checkbox"></td>
                                    <td>112</td>
                                    <td>^&%&</td>
                                    <td>
                                        <button class="linker-btn">����</button>
                                        <button class="linker-btn">����</button>
                                    </td>
                                </tr>
                                <tr>
                                    <td><input type="checkbox"></td>
                                    <td>112</td>
                                    <td>^%&^&^</td>
                                    <td>
                                        <button class="linker-btn">����</button>
                                        <button class="linker-btn">����</button>
                                    </td>
                                </tr>
                                <tr>
                                    <td><input type="checkbox"></td>
                                    <td>112</td>
                                    <td>@*#(*</td>
                                    <td>
                                        <button class="linker-btn">����</button>
                                        <button class="linker-btn">����</button>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </div>

                </div>
            </div>
        </div>
    </div>
    <!-- endregion -->


    <script>
        document.querySelector('body').classList.add('noscroll');
        $(function(){
            for( let i=1 ; i<=4 ; i++ ) {
                $('#datepicker' + i).datepicker( {
                    inline: true,
                    showOtherMonths: true,
                    showMonthAfterYear: true,
                    monthNames: [ '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12' ],
                    dayNamesMin: ['��', '��', 'ȭ', '��', '��', '��', '��'],
                    dateFormat: 'yy-mm-dd',
                });
            }

            document.querySelector('.forum-descr-sample').value = ''
                + '<div class="ex_img">\n'
                +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_01.gif" alt="since 2001">\n'
                +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_02.jpg" alt="�ٹ�����, �ӱ��ſ� �ϳ� �����ΰ�?">\n'
                +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_03.gif" alt="�ӱ���">\n'
                +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_04.jpg" alt="20��° �ӱ��� ���� ����!">\n'
                +   '\t<img src="http://fiximage.10x10.co.kr/web2021/anniv2021/m/forum_history02_05.gif" alt="�ӱ��ŵ��� ������ �����?">\n'
                + '</div>';
        });
    </script>

</div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->