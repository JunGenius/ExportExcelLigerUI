/*
 *  采用jquery模板插件——jQuery Boilerplate
 *
 *  Made by QuJun
 *  2017/01/10
 */
//table2excel.js
; (function ($, window, document, undefined) {


    var pluginName = "table2excel",

    rootPath = "http://" + window.location.host,


    currentElement = null,

    defaults = {
        exclude: ".noExl",  // 不导出excel 样式为.noExl(默认)
        name: "Table2Excel",
        islageruigrid: false, // 表格是否是ligerGrid
        sheetName: "Sheet1",// sheetName
        ligerui: {
            isGetAll: false, //  是否获取ligerGrid全部数据(前提需配置isligerGrid:true) ,如果获取全部数据失败，则导出当前页表格数据
            gridInstance: null // ligerGrid 对象 (如果为null,则获取当前页数据)
        },
        css: [ // css 样式
            "/RefLib/css/ligerui_grid_export_excel.css"
        ],
    };

    // 构造函数
    function Plugin(element, options) {
        this.element = element;
        // jQuery has an extend method which merges the contents of two or
        // more objects, storing the result in the first object. The first object
        // is generally empty as we don't want to alter the default options for
        // future instances of the plugin
        //
        this.settings = $.extend({}, defaults, options);
        this._defaults = defaults;
        this._name = pluginName;
        this.init();
    }

    Plugin.prototype = {
        init: function () {
            var e = this;

            var utf8Heading = "<meta http-equiv=\"content-type\" content=\"application/vnd.ms-excel; charset=UTF-8\">";

            e.template = {
                head: "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns=\"http://www.w3.org/TR/REC-html40\">" + utf8Heading + "<head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets>",
                sheet: {
                    head: "<x:ExcelWorksheet><x:Name>",
                    tail: "</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>"
                },
                mid: "</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->",
                tail: "</head><body>",
                table: {
                    head: "<table><tbody>",
                    tail: "</tbody></table>"
                },
                css: {
                    head: "<style>",
                    tail: "</style>"
                },
                foot: "</body></html>"
            };

            e.lagerui_template = {

                boby: {
                    head: "<table class = 'l-grid-table'>",
                    tail: "</table>",
                }
            };


            e.tableRows = [];

            // 是否是lageruigrid 表格
            if (e.settings.islageruigrid) {
                // 是否获取全部数据
                if (e.settings.ligerui.isGetAll) {
                    // lageruigrid 对象
                    var grid = e.settings.ligerui.gridInstance;
                    if (grid == null) {
                        e.gridToExcel(0);
                    } else {
                        grid.loadAllServerData(function (data) {
                            if (data == null) {
                                e.gridToExcel(0);
                            } else {
                                e.gridToExcel(1, grid, data);
                            }
                        });

                    }

                } else {
                    e.gridToExcel(0);
                }

            } else {
                // 导出当前节点下数据
                $(e.element).children().not(e.settings.exclude).each(function (i, o) {
                    var tempRows = "";

                    $(o).each(function (i, o) {
                        tempRows += $(o).prop("outerHTML");
                    });

                    e.tableRows.push(tempRows);
                });

                e.tableToExcel(e.tableRows, e.settings.name, e.settings.sheetName);
            }


        },


        // type : 1:导出全部数据 2.导出当前数据
        gridToExcel: function (type, grid, data) {

            var e = this;

            var g = $(e.element).parent().parent();

            var title = $(g).find(".l-layout-header").text();

            e.tableRows.push(e.lagerui_template.boby.head);


            var _getGridTitle = function () {

                var _arr = [];

                var columnCount;

                var c = $(g).find(".l-grid2 .l-grid-header-table");

                var c1 = $(g).find(".l-grid1 .l-grid-header-table");

                $(c1).find("tr").each(function (i, o) {

                    var columns = "";

                    columnCount = 0;

                    $(o).find("td:visible").each(function (i, o) {

                        columns += "<th>" + $(o).text() + "</th>";

                        columnCount++;
                    });

                    var _dom = $(c).find("tr").eq(i);

                    $(_dom).find("td:visible").each(function (i, o) {

                        columns += "<th>" + $(o).text() + "</th>";

                        columnCount++;
                    });

                    _arr.push("<tr>" + columns + "</tr>");

                });

                e.tableRows.push("<tr><th colspan = '" + columnCount + "'>" + title + "</th></tr>");

                if (_arr.length > 0) {
                    for (var i = 0 ; i < _arr.length; i++) {
                        e.tableRows.push(_arr[i]);
                    }
                }
            }


            var _getGridBody = function () {
                // 内容
                var t = $(g).find(".l-grid2 .l-grid-body-table");

                var t1 = $(g).find(".l-grid1 .l-grid-body-table");

                $(t1).find("tr").each(function (i, o) {

                    var info = "";

                    $(o).find("td:visible").each(function (i, o) {

                        info += "<td>" + $.trim($(o).text()) + "</td>";
                    });

                    $(t).find("tr").eq(i).find("td:visible").each(function (i, o) {

                        info += "<td>" + $.trim($(o).text()) + "</td>";
                    });

                    e.tableRows.push("<tr>" + info + "</tr>");

                });
            }

            var _getGridBodyFromServer = function () {

                var cols = grid.getOptions().columns;

                var rows = data.Rows;

                for (var i = 0; i < rows.length; i++) {

                    var info = "";

                    for (var j = 0; j < cols.length; j++) {
                        var h = cols[j].hide;
                        if (h) {
                            continue;
                        }

                        var name = cols[j].name;

                        var rowName;

                        var rowdata = rows[i];

                        if (cols[j].render) {

                            rowName = e._nvl(cols[j].render(rowdata, i));

                            if (e._isHtml(rowName)) {
                                rowName = $(rowName).html();
                            }

                        } else if (name === undefined || name == "") {

                            continue;

                        } else {
                            rowName = e._nvl(rows[i][name]);
                        }

                        if (String(rowName).indexOf("/Date") != -1) {
                            rowName = e.ChangeDateFormat_Date(rowName);
                        }

                        info += "<td>" + rowName + "</td>";
                    }

                    e.tableRows.push("<tr>" + "<td>" + (i + 1) + "</td>" + info + "</tr>");
                }

            }

            _getGridTitle();


            if (type == 1) {
                _getGridBodyFromServer();
            } else {
                _getGridBody();
            }

            e.tableRows.push(e.lagerui_template.boby.tail);

            e.tableToExcel(e.tableRows, e.settings.name, e.settings.sheetName);
        },

        // 导出excel
        tableToExcel: function (table, name, sheetName) {
            var e = this, fullTemplate = "", i, link, a;

            e.uri = "data:application/vnd.ms-excel;base64,";
            e.base64 = function (s) {
                return window.btoa(unescape(encodeURIComponent(s)));
            };
            e.format = function (s, c) {
                return s.replace(/{(\w+)}/g, function (m, p) {
                    return c[p];
                });
            };

            sheetName = typeof sheetName === "undefined" ? "Sheet" : sheetName;

            e.ctx = {
                worksheet: name || "Worksheet",
                table: table,
                sheetName: sheetName,
            };

            var fullTemplate = e.template.head;



            var css_boby = "";

            var cssStyles = e.settings.css;

            // 添加css样式
            if (cssStyles.length != 0) {
                $.ajaxSettings.async = false;
                css_boby = e.template.css.head;
                for (var i = 0; i < cssStyles.length; i++) {
                    $.get(rootPath + cssStyles[i], function (result) {
                        css_boby += result;
                    });
                }
                css_boby += e.template.css.tail;

                $.ajaxSettings.async = true;//设置getJson异步
            }


            //if ($.isArray(table)) {
            //    for (i in table) {
            //        //fullTemplate += e.template.sheet.head + "{worksheet" + i + "}" + e.template.sheet.tail;
            //        fullTemplate += e.template.sheet.head + sheetName + i + e.template.sheet.tail;
            //    }
            //}
            // 设置Excell sheet
            fullTemplate += e.template.sheet.head + sheetName + e.template.sheet.tail;

            fullTemplate += e.template.mid + css_boby + e.template.tail;

            if ($.isArray(table)) {
                for (i in table) {
                    //fullTemplate += e.template.table.head + "{table" + i + "}" + e.template.table.tail;
                    fullTemplate += "{table" + i + "}";
                }
            }

            fullTemplate += e.template.foot;

            for (i in table) {
                e.ctx["table" + i] = table[i];
            }
            delete e.ctx.table;

            if ($.browser.msie

                || $.browser.version == "11.0")      // If Internet Explorer
            {
                if (typeof Blob !== "undefined") {
                    // Must be replaced
                    for (var i in table) {
                        //var reg = eval("/{table[" + i + "]}/");
                        fullTemplate = fullTemplate.replace("{table" + i + "}", table[i]);
                    }

                    //use blobs if we can
                    fullTemplate = [fullTemplate];
                    //convert to array
                    var blob1 = new Blob(fullTemplate, { type: "text/html" });
                    window.navigator.msSaveBlob(blob1, getFileName(e.settings));
                } else {

                    var iframe = e.AppendIframeToBoby();
                    iframeArea.document.open("text/html", "replace");
                    iframeArea.document.write(e.format(fullTemplate, e.ctx));
                    iframeArea.document.close();
                    //focus();
                    sa = iframeArea.document.execCommand("SaveAs", true, getFileName(e.settings));
                }

            } else {
                link = e.uri + e.base64(e.format(fullTemplate, e.ctx));
                a = document.createElement("a");
                a.download = getFileName(e.settings);
                a.href = link;

                document.body.appendChild(a);

                a.click();

                document.body.removeChild(a);
            }

            return true;
        },

        // 添加iframe
        AppendIframeToBoby: function () {
            var iframe = document.createElement('iframe');
            iframe.id = "iframeArea";
            iframe.width = "200";
            iframe.height = "200";
            iframe.scrolling = "no";
            iframe.frameBorder = "0";
            iframe.style.display = "none";
            document.body.appendChild(iframe);
            return iframe;
        },

        ChangeDateFormat_Date: function (data) {
            if (data != null) {
                var date = new Date(parseInt(data.replace("/Date(", "").replace(")/", ""), 10));
                var month = date.getMonth() + 1 < 10 ? "0" + (date.getMonth() + 1) : date.getMonth() + 1;
                var currentDate = date.getDate() < 10 ? "0" + date.getDate() : date.getDate();
                var dateStr = date.getFullYear() + "-" + month + "-" + currentDate;
                if (dateStr == '1900-01-01') {
                    return "";
                }
                return dateStr;
            }
            return "";
        },

        _nvl: function (value) {
            if (value == null || value == undefined) {
                return "";
            }
            return value;
        },
        _isHtml: function (value) {
            var reg = new RegExp('^<([^>\s]+)[^>]*>(.*?<\/\\1>)?$');
            return reg.test(value);
        }
    };

    function getFileName(settings) {
        return (settings.filename ? settings.filename : "table2excel") +
               (settings.fileext ? settings.fileext : ".xlsx");
    }


    // 对构造函数的一个轻量级封装，
    // 防止产生多个实例
    $.fn[pluginName] = function (options) {
        var e = this;
        e.each(function () {
            if (!$.data(e, "plugin_" + pluginName)) {
                $.data(e, "plugin_" + pluginName, new Plugin(this, options));
            }
        });

        // chain jQuery functions
        return e;
    };

})(jQuery, window, document);
