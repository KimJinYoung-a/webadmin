<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>회의실안내</title>
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0" />
    <meta http-equiv="refresh" content="60" />
    <script type="text/javascript" src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
    <script>window.jQuery || document.write(decodeURIComponent('%3Cscript src="js/jquery.min.js"%3E%3C/script%3E'))</script>
    <link rel="stylesheet" type="text/css" href="https://cdn3.devexpress.com/jslib/21.2.6/css/dx.common.css" />
    <link rel="stylesheet" type="text/css" href="https://cdn3.devexpress.com/jslib/21.2.6/css/dx.darkmoon.css" />
    <!--<link rel="stylesheet" type="text/css" href="https://cdn3.devexpress.com/jslib/21.2.6/css/dx.light.css" />-->
    <script src="https://cdn3.devexpress.com/jslib/21.2.6/js/dx.all.js"></script>
    <script src="index.js"></script>
    <style>
      .dx-scheduler-work-space-week .dx-scheduler-header-panel-cell,
      .dx-scheduler-work-space-work-week .dx-scheduler-header-panel-cell {
        text-align: center;
        vertical-align: middle;
      }

      .dx-scheduler-work-space .dx-scheduler-header-panel-cell .name {
        font-size: 13px;
        line-height: 15px
      }

      .dx-scheduler-work-space .dx-scheduler-header-panel-cell .number {
        font-size: 15px;
        line-height: 15px
      }
      .appointment-content {
        width: 360px;
      }

      .dx-popup-content .appointment-content {
        height: 40px;
        line-height: 20px;
      }

      .appointment-badge {
        text-align: center;
        float: left;
        margin-right: 12px;
        color: white;
        width: 42px;
        height: 42px;
        font-size: 20px;
        line-height: 42px;
        border-radius: 42px;
        margin-top: 19px;
        margin-left: 2px;
        display: flex;
        justify-content: center;
        flex-direction: column;
        padding-bottom: 2px;
      }

      .appointment-dates {
        color: #c6bbc5;
        font-size: 12px;
        text-align: left;
        float: left;
      }

      .appointment-text {
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
        width: 270px;
        font-size: 15px;
        text-align: left;
        float: left;
      }
    .delete-appointment.dx-state-hover,
    .dx-list-item.dx-state-hover .dx-button {
      box-shadow: none;
      background-color: inherit;
    }

    .delete-appointment .dx-icon-trash {
      color: #337ab7 !important;
      font-size: 23px !important;
    }
    </style>
</head>
<body class="dx-viewport">
    <div class="demo-container">
        <div id="scheduler"></div>
    </div>
</body>
</html>