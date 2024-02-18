// ��Ʈ ���� ����
// ���� �����ϸ� + 0.0.1
// ����� �߰��Ǹ� + 0.1
// ������ �ɶ��� + 1
var chartjsVersion = "0.7.3";
var tChart = {};

// debug�� �ɼ�
tChart.debug = false;
tChart.out = function(args) {
    if ( tChart.debug ) {
        console.log(args);
    }
};
tChart.global = {
    // save draw chart instance, data table, elementid,...
    // get chart
    // re draw
    // options resettings
    //   line on/off resettings
    // {
    //     chart: chart,
    //     dataTable: dataTable,
    //     options: options,
    //     elementId: elementId,
    //     element: ele
    // }
    data: {},
    putChart: function(params) {
        tChart.global.data[params.elementId] = params;
    },
    getChart: function(id) {
        return tChart.global.data[id];
    },
    setLinesHidden: function(elementId, columnsIndex) {
        var chartData = tChart.global.getChart(elementId);
        _.each(columnsIndex, function(i) {
            chartData.options.series[i].lineWidth = 0;
        });
        chartData.chart.draw(chartData.dataTable, chartData.options);
    },
    setLinesShown: function(elementId, columnsIndex) {
        var chartData = tChart.global.getChart(elementId);
        _.each(columnsIndex, function(i) {
            chartData.options.series[i].lineWidth = chartData.options.defaultLineWidth;
        });
        chartData.chart.draw(chartData.dataTable, chartData.options);
    }
};

tChart.hook = {
    makeStringToDate: function(pos) {
        return function(row) {
            var date = new Date(row[pos]);
            if ( !_.isNaN(date.getDay()) ) {
                // new Date("2016-02-01")�� ��� �ð��� 9�ð� �ȴ�.
                // �׷��� ��ġ�� ��߳���.
                date.setHours(0);
                row[pos] = date;
            }
            return row;
        };
    }
};

tChart.loader = {
    load: function(url, complete) {
        $.ajax({
            url: url,
            jsonpCallback: 'callback',
            contentType: "application/json",
            dataType: "jsonp",
            success: function( response ) {
                complete(response, null);
            },
            error: function( error ) {
                complete(null, error);
            }
        });
    },
    loadWithCache: function(url, complete) {
        var localStorage = window.localStorage;
        if ( localStorage && localStorage.getItem(url) ) {
            console.log('cache hit', url);
            f(JSON.parse(localStorage.getItem(url)));
        } else {
            this.load(url, function(data, error) {
                if ( data ) {
                    localStorage.setItem(url, JSON.stringify(data));
                }
                f(data, error);
            });
        }
    },
    loadWithLocalProxy: function(url, f) {
        $.post("/api", {'url': url}, function(data) {
            f(data, null);
        }).fail(function(err) {
            console.log(err);
        });
    },
    // �׽�Ʈ �뵵�θ� ����Ұ�.
    loadWithCacheProxy: function(url, f) {
        var localStorage = window.localStorage;
        if ( localStorage && localStorage.getItem(url) ) {
            console.log('cache hit', url);
            f(JSON.parse(localStorage.getItem(url)));
        } else {
            this.loadWithLocalProxy(url, function(data, error) {
                if ( data ) {
                    localStorage.setItem(url, JSON.stringify(data));
                }
                f(data, error);
            });
        }
    }
};

tChart.chart = {
    globalOptions: function(type) {
        var defaultOptions = {'tooltip' : { 'isHtml': true }};

        var options = {
            'line' : {
                'hAxis' : {
                    'format' : 'yyyy-MM-dd'
                },
                'focusTarget' : 'category'
            },
            'pie': {},
            'bar': {
                'hAxis' : {
                    'format' : 'yyyy-MM-dd'
                },
                'focusTarget' : 'category'
            }
        };

        var selectedOptions = options[type];

        if ( !selectedOptions ) {
            selectedOptions = {};
        }

        return _.merge(defaultOptions, selectedOptions);
    },
    toType: function(s) {
        var ls = s.toLowerCase();
        if ( ls == 'string') {
            return 'string';
        } else if ( ls == 'integer' || ls == 'float' ) {
            return 'number';
        } else {
            return 'string';
        }
    },
    readyChart: function(end) {
        google.charts.load('current', {'packages':['corechart', 'table', 'line']});
        google.charts.setOnLoadCallback(function() {
            // ��Ʈ���� �ε� �Ϸ�Ǹ� f()�� ȣ����.
            end();
        });
    },
    /**
     * draw google table chart
     * deprecated �Ǿ���.
     * @param {DataTable} dataTable
     * @param {String} elementId
     * @param {Object} options
     * @returns {google.visualization.Table}
     */
    googleTableChart: function(dataTable, elementId, options) {
        var defaultOptions = { info: false, paging: false, searching: false, scrollY: 300 };

        // scrollY height
        var height = $(element(elementId)).height();
        if ( height > 0 ) {
            defaultOptions.scrollY = height;
        }

        if ( _.get(options, "paging", false) ) {
            // paging = true?
            delete defaultOptions.scrollY;
        }
        var mergedOptions = _.merge(defaultOptions, options);
        return tChart.chart.dataTable(dataTable, elementId, mergedOptions);

        // if ( tChart.converter.isDate(dataTable, 0) ) {
        //     setDataTableDateFormat(dataTable, 0, options.pattern);
        // }
        // // cleanup
        // var ele = element(elementId);
        // $(ele).html('');

        // // table View
        // var chart = new google.visualization.Table(ele);
        // chart.frozenColumns = 1;

        // var defaultOptions = {
        //     'allowHtml': true,
        //     'width': '100%', height: '100%'
        // };

        // var mergedOptions = _.merge(defaultOptions, options);
        // chart.draw(dataTable, mergedOptions);

        // tChart.global.putChart({
        //     chart: chart,
        //     dataTable: dataTable,
        //     options: mergedOptions,
        //     elementId: elementId,
        //     element: ele
        // });

        // return chart;
    },
    /**
     * draw google line chart
     * @param {DataTable} dataTable google DataTable
     * @param {String} elementId set element id whitout #
     * @param {Object} options google chart options
     * @returns {google.visualization.LineChart}
     */
    googleLineChart: function(dataTable, elementId, options) {
        if ( tChart.converter.isDate(dataTable, 0) ) {
            setDataTableDateFormat(dataTable, 0, options.pattern);
        }

        // line width
        var defaultLineWidth = options.defaultLineWidth || 2;
        var defaultDashStyle = options.defaultDashStyle || [4, 1];

        // get html element
        var ele = element(elementId);

        // chart container cleanup
        $(ele).html('');

        var chart = new google.visualization.LineChart(ele);
        var defaultOptions = {
            'legend': { 'position': 'bottom' },
            'series': {},
            hAxis: {
                format: options.pattern || 'yyyy-MM-dd'
            },
            'defaultLineWidth': defaultLineWidth
        };

        var length = dataTable.getNumberOfColumns();
        for ( var i = 0; i < length; i++ ) {
            defaultOptions.series[i] = {
                lineWidth: defaultLineWidth
            };
        }
        // set dashs with index
        _.each(options.dashsWithIndex, function(index) {
            _.merge(defaultOptions.series[index], {lineDashStyle: defaultDashStyle});
        });
        // set dashs
        _.each(options.dashs, function(b, index) {
            if ( b ) {
                _.merge(defaultOptions.series[index], {lineDashStyle: defaultDashStyle});
            }
        });

        var mergedOptions = _.merge({}, tChart.chart.globalOptions('line'), defaultOptions, options);
        // set chart line hidden
        if ( _.isArray(mergedOptions.columnsHiddenIndex) ) {
            var columns = mergedOptions.columnsHiddenIndex;
            _.each(columns, function(n) {
                mergedOptions.series[n].lineWidth = 0;
            });
        }

        chart.draw(dataTable, mergedOptions);

        // ������Ʈ���� Ư�� �÷��� ���õǸ� �� ���� ������.
        if ( !mergedOptions.lineOnOff ) {
            google.visualization.events.addListener(chart, 'select', function() {
                if ( chart.getSelection()[0].column === null ) {
                    return;
                }
                var columnIndex = chart.getSelection()[0].column - 1;
                if  ( mergedOptions.series[columnIndex].lineWidth == defaultLineWidth ) {
                    mergedOptions.series[columnIndex].lineWidth = 0.0;
                } else {
                    mergedOptions.series[columnIndex].lineWidth = defaultLineWidth;
                }
                chart.draw(dataTable, mergedOptions);
            });
        }


        tChart.global.putChart({
            chart: chart,
            dataTable: dataTable,
            options: mergedOptions,
            elementId: elementId,
            element: ele
        });

        return chart;
    },
    /**
     * draw google pie chart
     * @param {DataTable} dataTable
     * @param {String} elementId
     * @param {Object} options
     * @returns {google.visualization.PieChart} 
     */
    googlePieChart: function(dataTable, elementId, options) {
        if ( tChart.converter.isDate(dataTable, 0) ) {
            setDataTableDateFormat(dataTable, 0, options.pattern);
        }
        var ele = element(elementId);

        // cleanup
        $(ele).html('');
        var chart = new google.visualization.PieChart(ele);
        var defaultOptions = {
        };

        var mergedOptions = _.merge({}, defaultOptions, tChart.chart.globalOptions('pie'), options);

        chart.draw(dataTable, mergedOptions);
        tChart.global.putChart({
            chart: chart,
            dataTable: dataTable,
            element: ele,
            elementId: elementId,
            options: mergedOptions
        });
        return chart;
    },
    googleBarChart: function(dataTable, elementId, options) {
        if ( tChart.converter.isDate(dataTable, 0) ) {
            setDataTableDateFormat(dataTable, 0, options.pattern);
        }
        var ele = element(elementId);
        $(ele).html('');
        var defaultOptions = {
        };
        var chart = new google.visualization.BarChart(element(elementId));

        var mergedOptions = _.merge({}, defaultOptions, tChart.chart.globalOptions('bar'), options);

        chart.draw(dataTable, mergedOptions);

        tChart.global.putChart({
            chart: chart,
            options: mergedOptions,
            dataTable: dataTable,
            element: ele,
            elementId: elementId
        });

        return chart;
    },
    rawChart: function(rawData, elementId) {
        $('#' + elementId).html('');
        $('<textarea style="width:100%; height:100%">' + JSON.stringify(rawData, null, 2) + '</textarea>').appendTo('#' + elementId);
    },
    dataTable: function(data, elementId, options) {
        // reduce
        var groupSum = function(total, row) {
            // only number type
            return _.chain(_.zip(total, row))
                .map(function(item) {
                    if ( _.every(item, _.isNumber) ) {
                        // �������� ��� ������ ��� ���ؼ� ����
                        return _.sum(item);
                    } else {
                        // ���ڰ� �ƴ� ��� "-"�� ����
                        return "-";
                    }
                }).value();
        };
        var groupAvg = function(item) {
            if ( _.isNumber(item) ) {
                return item / data.row.length;
            } else {
                return "-";
            }
        };
        var sumRow = _.reduce(data.raw.rows, groupSum);
        // 0�� ��ҿ� ���� �̸�
        sumRow[0] = "�հ�";
        var avgRow = _.map(sumRow, groupAvg);
        avgRow[0] = "���";

        // formatNumber
        var formatNumberFunc = function(item) {
            if ( _.isNumber(item) ) {
                return {name: formatNumber(item), attr: 'style="text-align: right; padding: 8px 10px;"'};
            } else {
                return item;
            }
        };
        var fixedFunc = function(item) {
            if ( fixed != null && _.isNumber(fixed) && isFloat(item) ) {
                return parseFloat(item.toFixed(fixed));
            } else {
                return item;
            }
        };

        // fixed
        var fixed = _.get(options, "group.fixed", null);
        sumRow = _.map(sumRow, fixedFunc);
        sumRow = _.map(sumRow, formatNumberFunc);
        avgRow = _.map(avgRow, fixedFunc);
        avgRow = _.map(avgRow, formatNumberFunc);

        var groupType = _.get(options, "group.type", null);
        var footRow = "";
        // foot row �� ����.
        if ( groupType ) {
            footRow += "<tfoot><tr>";
            var toTableRowTag = function(item) {
                return "<th>" + item + "</th>";
            };
            if ( groupType === "avg" ) {
                // footRow += _.map(avgRow, toTableRowTag).join("");
                footRow += tagWithChilds("th", avgRow);
            } else if ( groupType === "sum" ) {
                // footRow += _.map(sumRow, toTableRowTag).join("");
                footRow += tagWithChilds("th", sumRow);
            }
            footRow += "</tr></tfoot>";
        }

        var ele = element(elementId);
        // cleanup
        $(ele).html('');
        // http://www.css-prefix.com/
        // bootstrap prefix ttbt(tenbyten bootstrap)
        // 10x10 css��� �浹�ϱ� ������ bootstrap�ڵ带 Ŀ������.
        $(ele).addClass('ttbt');
        var tableId = elementId + "_dataTable";
        var html = "";
        var head = data.head;
        html += "<table id='" + tableId + "' width='100%' class='table-striped table-bordered'><thead>";
        // head
        _.each(head, function(eachHead) {
            html += "<tr>";
            html += tagWithChilds("th", eachHead);
            html += "</tr>";
        });
        html += "</thead>";

        // foot
        html += footRow;

        var rows = data.row;
        html += "<tbody>";
        _.each(rows, function(row) {
            html += "<tr>" + tagWithChilds("td", row) + "</tr>";
        });
        html += "</tbody>";
        html += "</table>";
        $(ele).html(html);
        var dataTable = $(element(tableId)).DataTable(options);
        $(ele).css("height", "");
        tChart.global.putChart({
            chart: dataTable,
            dataTable: data,
            element: element,
            elementId: elementId,
            options: options
        });
        return dataTable;
    }
};

tChart.converter = {
    isDate: function(dataTable, columnIndex) {
        var dataType = Object.prototype.toString.call(dataTable.getValue(columnIndex, 0));
        var dateDataType = Object.prototype.toString.call(new Date());

        return dataType === dateDataType;
    },
    makeTable: function(data) {
        var newData = deepCopy(data);
        // columnHeaders���� name�� �̾Ƽ� �迭�� ����.
        var head = _.map(newData.columnHeaders, function(h) { return h.name; });
        var body = newData.rows;

        body.splice(0, 0, head);
        return body;
    },
    selectColumns: function(data, columnIndexs, hook) {
        if ( !columnIndexs ) {
            return data;
        }
        var newData = deepCopy(data);
        // groupname ó���� ������.
        var head = _.map(newData.columnHeaders, function(h) {
            if ( newData.isGroupHeader ) {
                return h.groupName + " " + h.name;
            } else {
                return h.name;
            }
        });
        var body = newData.rows;

        //body.splice(0, 0, head);
        if ( !newData.rows ) {
            throw new Error("not found rows in api data");
        }
        newData.rows.splice(0, 0, head);

        // ���õ� �÷��鸸���� �̷���� �迭�� ����.
        // op
        // func some(arr)
        var mergedHeaderFunc = function(arr) {
            return _.reduce(arr, function(total, n) {
                return total + " + " + n;
            });
        };
        var adderFunc = function(arr) {
            return _.reduce(arr, function(total, n) {
                return total + n;
            });
        };
        var selectedData = _.map(newData.rows, function(eachRow) {
            var result = [];
            for ( var i = 0; i < columnIndexs.length; i++ ) {
                var eachIndex = columnIndexs[i];
                if ( _.isArray(eachIndex) ) {
                    // ù��°�� op
                    var op = _.first(eachIndex);
                    var indexList = _.rest(eachIndex);

                    var value = null;

                    var filteredValues = _.filter(eachRow, function(n, nidx) {
                            return _.includes(indexList, nidx);
                    });
                    if ( !_.isNumber(filteredValues[0]) ) {
                        value = mergedHeaderFunc(filteredValues);
                    } else {
                        // number
                        if ( op === "+" ) {
                            value = adderFunc(filteredValues);
                        } else if ( _.isFunction(op) ) {
                            value = op(filteredValues);
                        } else {
                            value = 0;
                        }
                    }

                    result.push(value);
                } else if ( _.isNumber(eachIndex) ) {
                    result.push(eachRow[eachIndex]);
                }
            }
            if ( hook ) {
                return hook(result);
            } else {
                return result;
            }
        });

        return selectedData;
    },
    // * deprecated �Ǿ���.
    googleTableChart: function(data, columnsIndex) {
        var result = tChart.converter.dataTable(data, columnsIndex);
        result.dataTable = result;
        return result;
        // var body = makeTable(data);
        // var dataTable = google.visualization.arrayToDataTable(body);

        // return {
        //     raw: data,
        //     dataTable: dataTable
        // };
    },
    googleLineChart: function(data, columnIndexs, hook) {
        var columns = selectColumns(data, columnIndexs, hook);
        var dataTable = google.visualization.arrayToDataTable(columns); 

        return {
            raw: data,
            dataTable: dataTable
        };
    },
    googlePieChart: function(data, columnIndexs) {
        var dataTable = google.visualization.arrayToDataTable(selectColumns(data, columnIndexs));

        return {
            raw: data,
            dataTable: dataTable
        };
    },
    googlePieChartWithSum: function(data, columnIndexs) {
        if ( columnIndexs.length < 2 ) {
            console.log('�÷��� �ΰ� �̻� �����ؾ� �մϴ�.');
        }
        var selectedData = selectColumns(data, columnIndexs);
        var foldData = _.reduce(_.rest(selectedData), function(total, row) {
            for ( var i = 0; i < total.length; i++ ) {
                total[i] += row[i];
            }
            return total;
        });
        var resultData = _.zip(_.first(selectedData), foldData);
        resultData = _([["�׸�", "�հ�"]]).concat(resultData).value();

        var dataTable = google.visualization.arrayToDataTable(resultData);

        return {
            raw: data,
            dataTable: dataTable
        };
    },
    googleBarChart: function(data, columnIndexs, hook) {
        var dataTable = google.visualization.arrayToDataTable(selectColumns(data, columnIndexs, hook));

        return {
            raw: data,
            dataTable: dataTable
        };
    },
    dataTable: function(data, columnIndexs) {
        // http://www.datatables.net/examples/index
        var newData = deepCopy(data);
        // var newData = selectColumns(data, columnIndexs);
        var result = {head: [], row: []};
        result.raw = newData;

        if ( newData.isGroupHeader ) {
            var names = _.map(newData.columnHeaders, function(v) {
                return {groupName: v.groupName, name: v.name};
            });
            var groupWithGroupName = _.groupBy(newData.columnHeaders, "groupName");

            var groupNames = _.map(names, function(name) {
                var length = groupWithGroupName[name.groupName].length;
                var attr = "colspan='" + length + "'";
                if ( length == 1 ) {
                     attr = "rowspan='2'";
                }
                return {name: name.groupName, length: length, attr: attr };
            });

            result.head.push(_.uniq(groupNames, "name"));

            var bottomRow = _.chain(names)
                            .map(function(name) {
                                var length = groupWithGroupName[name.groupName].length;
                                if ( length > 1 ) {
                                    return name.name;
                                } else {
                                    return null;
                                }
                            })
                            .filter(function(o) { return !_.isNull(o); })
                            .value();

            result.head.push(bottomRow);
        } else {
            result.head.push(_.map(newData.columnHeaders, function(eachHeader) {
                return {name: eachHeader.name};
            }));
        }

        _.each(newData.rows, function(eachRow) {
            var row = _.map(eachRow, function(item) {
                if ( _.isNumber(item) ) {
                    return {"attr": 'align="right"', "name": formatNumber(item)};
                } else {
                    return item;
                }
            });
            result.row.push(row);
        });

        return result;
    }
};

tChart.transform = {
    withGroupName: function(data, sep) {
        var newData = deepCopy(data);
        sep = sep || " ";

        _.each(newData.columnHeaders, function(v) {
            if ( v.groupName && v.groupName.length > 0 ) {
                v.name = v.groupName + sep + v.name;
            }
        });

        return newData;
    }
};

// chain(data).convert('linechart', [0, 1, 2], hook).chart('linechart', options).into('viewId');
// �Լ� ���ձ�
tChart.Chain = function(data) {
    var self = this;
    this.data = data;
    this.convert = function(type, p1, p2) {
        if ( !tChart.convert[type] ) {
            throw "not found type error";
        }
        self.convert = {
            f: tChart.converter[type],
            p1: p1,
            p2: p2
        };
        return self;
    };
    this.chart = function(type, p1) {
        if ( !tChart.chart[type] ) {
            throw "not found type error";
        }
        self.chart = {
            f: tChart.chart[type],
            p1: p1
        };
        return self;
    };
    this.into = function(viewId) {
        var convertedData  = self.convert.f(self.data, self.convert.p1, self.convert.p2);
        self.chart.f(convertedData, viewId, self.chart.p1);
    };
};


// -----------------------------------------------
// util functions

// http://stackoverflow.com/questions/3885817/how-do-i-check-that-a-number-is-float-or-integer
function isInt(n) {
    return Number(n) === n && n % 1 === 0;
}

function isFloat(n) {
    return Number(n) === n && n % 1 !== 0;
}

function chain(data) {
    // chain(data).convert('linechart').draw('linechart').into('viewid');
    return new tChart.chain(data);
}

/**
 * google datatable date format
 * https://developers.google.com/chart/interactive/docs/reference
 * @param {DataTable} dataTable
 * @param {Number} columnIndex
 * @param {String} pattern
 * @returns {DataTable} 
 */
function setDataTableDateFormat(dataTable, columnIndex, pattern) {
    var dateFormatter = new google.visualization.DateFormat({
        pattern: pattern || "yyyy-MM-dd"
    });
    dateFormatter.format(dataTable, columnIndex);

    return dataTable;
}

// 3�ڸ����� �޸��� ����
function formatNumber(n) {
    if ( _.isNumber(n) && isFloat(n) ) {
        return Number(n).toLocaleString('en');
    } else {
        // float�� �ƴҰ�� �Ҽ����� �������� �Ҽ����� �ڸ���.
        return Number(n).toLocaleString('en').split('.')[0];
    }
}

function tagWithChild(name, child) {
    return "<" + name + " " + (child.attr || "") + ">" + (child.name || child) + "</" + name + ">";
}

function tagWithChilds(name, childs) {
    var result = "";
    _.each(childs, function(child) {
        result += tagWithChild(name, child);
    });
    return result;
}

function deepCopy(obj) {
    return $.extend(true, {}, obj);
}

function element(id) {
    return document.getElementById(id.replace('#', ''));
}


// -----------------------------------------------
// interface functions

function load(url, complete) {
    tChart.loader.load(url, complete);
}

function loadWithLocalProxy(url, f) {
    tChart.loader.loadWithLocalProxy(url, f);
}
function loadWithCacheProxy(url, f) {
    tChart.loader.loadWithCacheProxy(url, f);
}

// ���� ��Ʈ
function toType(s) {
    return tChart.chart.toType(s);
}

function readyChart(end) {
    tChart.chart.readyChart(end);
}

function drawGoogleChartTable(dataTable, elementId, options) {
    return tChart.chart.googleTableChart(dataTable, elementId, options);
}

// api doc
// https://developers.google.com/chart/interactive/docs/gallery/linechart
function drawGoogleChartLine(dataTable, elementId, options) {
    return tChart.chart.googleLineChart(dataTable, elementId, options);
}

function drawGoogleChartPie(dataTable, elementId, options) {
    return tChart.chart.googlePieChart(dataTable, elementId, options);
}

function drawGoogleChartBar(dataTable, elementId, options) {
    return tChart.chart.googleBarChart(dataTable, elementId, options);
}

function drawRaw(rawData, elementId) {
    tChart.chart.rawChart(rawData, elementId);
}

function drawDataTable(data, elementId, options) {
    return tChart.chart.dataTable(data, elementId, options);
}


function makeTable(data) {
    return tChart.converter.makeTable(data);
}

function selectColumns(data, columnIndexs, hook) {
    return tChart.converter.selectColumns(data, columnIndexs, hook);
}

function convertDataForGoogleChartTable(data) {
    return tChart.converter.googleTableChart(data);
}

function convertDataForGoogleChartLine(data, columnIndexs, hook) {
    return tChart.converter.googleLineChart(data, columnIndexs, hook);
}

function convertDataForGoogleChartPie(data, columnIndexs) {
    return tChart.converter.googlePieChart(data, columnIndexs);
}

function convertDataForGoogleChartPieWithSum(data, columnIndexs) {
    return tChart.converter.googlePieChartWithSum(data, columnIndexs);
}

function convertDataForGoogleChartBar(data, columnIndexs, hook) {
    return tChart.converter.googleBarChart(data, columnIndexs, hook);
}

function convertDataForDataTable(data, columnIndexs) {
    return tChart.converter.dataTable(data, columnIndexs);
}

function transformWithGroupName(data, sep) {
    return tChart.transform.withGroupName(data, sep);
}

function googleChartRedraw(chart, dataTable, options) {
    chart.draw(dataTable, options);
}

/**
 * ���� ������Ʈ�� ������ ���̰� ��.
 * @param {String} containerId ��Ʈ �׸��� ��� �����̳� ���̵�
 * @param {Array} linesIndex ���� ������Ʈ�� ���� �ε�����
 * @example turnOnLineChartLines("lineContainer", [0])
 */
function turnOnLineChartLines(containerId, linesIndex) {
    tChart.global.setLinesShown(containerId, linesIndex);
}

/**
 * ���� ������Ʈ�� ������ �Ⱥ��̰� ��.
 * @param {String} containerId ��Ʈ �׸��� ��� �����̳� ���̵�
 * @param {Array} linesIndex ���� ������Ʈ�� ���� �ε�����
 * @example turnOffLineChartLines("lineContainer", [0])
 */
function turnOffLineChartLines(containerId, linesIndex) {
    tChart.global.setLinesHidden(containerId, linesIndex);
}


/**
 * ��Ʈ�� Date ���� ����. convert �Լ��� 3��° ���ڷ� �ַ� ����. �Ʒ� �Լ��� ������.
 * function hook(row) {
 *   row[0] = new Date(row[0]);
 *   return row;
 * }
 * @returns hook
 */
function hookDate() {
    return tChart.hook.makeStringToDate(0);
}
