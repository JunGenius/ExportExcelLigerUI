/*!
 * 
 * tableexport.js  
 * tableExport() 导出excel
 * 依赖 FileSaver.js(导出) , jszip.js(压缩) , xlsx.js(支持xlsx导出格式), table2excel.js(xls格式导出)
 * xls格式, 默认加载ligerui_grid_export_excel.css 样式
 * 2017/1/16 QuJun
 */

; (function (window, undefined) {

    /*--- GLOBALS ---*/
    var $ = window.jQuery;

    $.extend($.ligerMethos.Grid,
          {
              /*
                  获取ligeruiGrid 配置项 (导出excel表格使用) 
                  QuJun 2017/01/10
              */
              getOptions: function () {
                  return this.options;
              },
              /*
                 获取ligeruiGrid 全部数据 (ajax请求配置参数等同于前台配置ligerGrid参数（只更改usePage属性false）)
                 callback:回调函数 返回数据 请求失败返回null
                 QuJun 2017/01/10
              */
              loadAllServerData: function (callback) {
                  var g = this, p = this.options;

                  var param = [];
                  if (p.parms) {
                      var parms = $.isFunction(p.parms) ? p.parms() : p.parms;
                      if (parms.length) {
                          $(parms).each(function () {
                              param.push({ name: this.name, value: this.value });
                          });
                      }
                      else if (typeof parms == "object") {
                          for (var name in parms) {
                              param.push({ name: name, value: parms[name] });
                          }
                      }
                  }
                  if (p.dataAction == "server") {
                      if (p.sortName) {
                          param.push({ name: p.sortnameParmName, value: p.sortName });
                          param.push({ name: p.sortorderParmName, value: p.sortOrder });
                      }
                  };


                  var ajaxOptions = {
                      type: p.method,
                      url: p.url,
                      data: param,
                      async: p.async,
                      dataType: 'json',
                      beforeSend: function () {
                          if (g.hasBind('ExportExcelLoading')) {
                              g.trigger('ExportExcelLoading');
                          }
                          else {
                              g._setLoadingMessage("导出中");
                              g.toggleLoading(true);
                          }

                      },
                      success: function (data) {

                          if (callback != null) {
                              callback(data);
                          }

                      },
                      complete: function () {
                          g.trigger('complete', [g]);
                          if (g.hasBind('loaded')) {
                              g.trigger('loaded', [g]);
                          }
                          else {
                              g.toggleLoading.ligerDefer(g, 10, [false]);
                          }

                      },
                      error: function (XMLHttpRequest, textStatus, errorThrown) {
                          if (g.hasBind('loaded')) {
                              g.trigger('loaded', [g]);
                          }
                          else {
                              g.toggleLoading.ligerDefer(g, 10, [false]);
                          }
                          if (callback != null) {
                              callback(null);
                          }
                      }
                  };

                  if (p.contentType) ajaxOptions.contentType = p.contentType;
                  $.ajax(ajaxOptions);
              }
          }
      );

    var settings;

    $.fn.tableExport = function (options) {

        e = this;

        settings = $.extend({}, $.fn.tableExport.defaults, options),
            rowD = $.fn.tableExport.rowDel,
        // 该功能暂时不支持
            ignoreRows = settings.ignoreRows instanceof Array ? settings.ignoreRows : [settings.ignoreRows],
            ignoreCols = settings.ignoreCols instanceof Array ? settings.ignoreCols : [settings.ignoreCols];

        var fileName = settings.fileName;

        var fileType = settings.fileType;

        // 不支持Blob 则使用 tableToExcel.js 导出excel 如果为xlsx格式 则改成xls格式
        if (typeof Blob === "undefined") {
            $.fn.tableExport.exportMethods._tableToExcel(e);
            return;
        }

        // xls格式  使用table2excel导出excel 需传入css样式
        if (fileType == "xls") {
            $.fn.tableExport.exportMethods._tableToExcel(e);
            return;
        }

        // 支持四种导出格式
        exporters = {
            xlsx: function (data, name) {
                var mimeType = $.fn.tableExport.xlsx.mimeType,
                    fileExtension = $.fn.tableExport.xlsx.fileExtension;
                $.fn.tableExport.exportMethods._export2file(data, mimeType, name, fileExtension, settings);
            },
            xls: function (data, name) {

                var mimeType = $.fn.tableExport.xls.mimeType,
                    fileExtension = $.fn.tableExport.xls.fileExtension;
                $.fn.tableExport.exportMethods._export2file(data, mimeType, name, fileExtension, settings);
            },
            csv: function (data, name) {
                var mimeType = $.fn.tableExport.csv.mimeType,
                    fileExtension = $.fn.tableExport.csv.fileExtension;
                $.fn.tableExport.exportMethods._export2file($.fn.tableExport.exportMethods._getSeqData(data, $.fn.tableExport.csv.separator, $.fn.tableExport.rowDel), mimeType, name, fileExtension, settings);
            },
            txt: function (data, name) {
                var mimeType = $.fn.tableExport.txt.mimeType,
                    fileExtension = $.fn.tableExport.txt.fileExtension;
                $.fn.tableExport.exportMethods._export2file($.fn.tableExport.exportMethods._getSeqData(data, $.fn.tableExport.txt.separator, $.fn.tableExport.rowDel), mimeType, name, fileExtension, settings);
            }
        }

        // 是否是lageruigrid 表格
        if (settings.islageruigrid) {
            // 是否获取全部数据
            if (settings.ligerui.isGetAll) {
                // lageruigrid 对象
                var grid = settings.ligerui.gridInstance;
                if (grid == null) {
                    exporters[fileType]($.fn.tableExport.exportMethods._ligerGridData(0), fileName);
                } else {

                    if (settings.ligerui.source) {
                        exporters[fileType]($.fn.tableExport.exportMethods._ligerGridData(1, grid, settings.ligerui.source), fileName);
                    } else {
                        grid.loadAllServerData(function (data) {
                            exporters[fileType]($.fn.tableExport.exportMethods._ligerGridData(1, grid, data), fileName);
                        });
                    }
                }

            } else {
                exporters[fileType]($.fn.tableExport.exportMethods._ligerGridData(0), fileName);
            }

        } else {
            // 导出当前节点下数据 
            exporters[fileType]($.fn.tableExport.exportMethods._getData(), fileName);
        }

    };


    $.fn.tableExport.defaults = {
        title: "", // 标题为空 则不设置标题
        exclude: ".noExl", // 不是xls格式： 只有加在table tr 或 td 上class 有效  xls格式;任意
        fileType: "xlsx",  // xls格式:导出使用table2excel.js (需传入css样式,否则无样式)
        fileName: "Excel Document Name", // 文件名 同sheet名
        islageruigrid: false, // 表格是否是ligerGrid (ligeruigrid 必须指定true,否则显示不正确)
        ligerui: {
            isGetAll: false, //  是否获取ligerGrid全部数据(前提需配置isligerGrid:true) ,如果获取全部数据失败，则导出当前页表格数据
            gridInstance: null, // ligerGrid 对象 (如果为null,则获取当前页数据)
            serialNumberWidth: "27", // 序号宽度
            RenderColumns: {} // (废除)自定义计算列{列名:{render:callback(回调方法 参数:data(该行数据) return: 计算结果)}} 
            //source: {}   // 本地数据指定数据源 导出全部数据 否则只导出当前页面数据. 若不是本地数据 则无需指定.(优先级大于服务段获取全部数据)
        },
        frozen: null, // 冻结信息 （c:冻结列, r：冻结行）
        // xls格式 样式  (如果xls格式并且是ligeruigrid 或者 浏览器不支持blob , 前台未传入css样式,则默认加载ligerui_grid_export_excel.css 样式 ，若路径改变则需重新指定 )
        // xlsx格式  如果浏览器不支持blob,则自动转换xls格式,需传入css样式
        css: [],
        ignoreRows: null,// TODO
        ignoreCols: null  // TODO
    };

    $.fn.tableExport.charset = "charset=utf-8";

    $.fn.tableExport.xlsx = {
        defaultClass: "xlsx",
        buttonContent: "Export to xlsx",
        mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        fileExtension: ".xlsx"
    };

    $.fn.tableExport.xls = {
        defaultClass: "xls",
        buttonContent: "Export to xls",
        separator: "\t",
        mimeType: "application/vnd.ms-excel",
        fileExtension: ".xls"
    };

    $.fn.tableExport.csv = {
        defaultClass: "csv",
        buttonContent: "Export to csv",
        separator: ",",
        mimeType: "application/csv",
        fileExtension: ".csv"
    };

    $.fn.tableExport.txt = {
        defaultClass: "txt",
        buttonContent: "Export to txt",
        separator: "  ",
        mimeType: "text/plain",
        fileExtension: ".txt"
    };

    $.fn.tableExport.titleHeight = "35";

    $.fn.tableExport.defaultFileName = "myDownload";

    $.fn.tableExport.defaultButton = "button-default";

    $.fn.tableExport.bootstrap = ["btn", "btn-default", "btn-toolbar"];

    $.fn.tableExport.rowDel = "\r\n";

    $.fn.tableExport.entityMap = { "&": "&#38;", "<": "&#60;", ">": "&#62;", "'": '&#39;', "/": '&#47' };

    $.fn.tableExport.Style = [{
        fill: {
            patternType: "none", // none / solid
            fgColor: { rgb: "#CCFFFF" },
            bgColor: { rgb: "#666666" }
        },
        font: {
            name: 'Times New Roman',
            sz: 16,
            color: { rgb: "#666666" },
            bold: true,
            italic: false,
            underline: false
        },
        border: {
            top: { style: "thin", color: { auto: 1 } },
            right: { style: "thin", color: { auto: 1 } },
            bottom: { style: "thin", color: { auto: 1 } },
            left: { style: "thin", color: { auto: 1 } }
        }
    }
    ]

    $.fn.tableExport.exportMethods = {

        data: [],

        colsW: [],// 列宽 自定义表格（2个以上table） 以最后一行宽度为准

        rowsH: [],

        rowNum: 0,

        merges: [],

        columnCount: 0,

        colsIndex: 0,

        frozen: null, // 冻结信息


        _reset: function () {

            this.data = [];

            this.colsW = [];

            this.rowsH = [];

            this.rowNum = 0;

            this.merges = [];

            this.columnCount = 0;

            this.colsIndex = 0;

            this.frozen = null; // 冻结信息
        },

        _getData: function () {

            var com = this;

            com._reset();

            var _getGridTitle = function () {

                if (settings.title != "" && com.rowNum == 0) {
                    com.data[0] = [];
                    com.merges[0] = "";
                    com.rowsH.push({ wpx: $.fn.tableExport.titleHeight });
                    com.rowNum++;
                }

                //_getGridTitleByDom();

                if (settings.title != "") {
                    for (var k = 0; k < com.colsIndex; k++) {
                        if (k == 0) {
                            com.data[0][k] = { st: '4', v: settings.title };
                        } else {
                            com.data[0][k] = { st: '4', v: "" };
                        }
                    }
                    com.merges[0] = { s: { c: 0, r: 0 }, e: { c: com.colsIndex - 1, r: 0 } };
                }
            }

            $(e).find("table").each(function (i, o) {

                $(o).find("tr").not(settings.exclude).each(function (i, o) {

                    com.colsW = [];

                    com.colsIndex = 0;

                    //if (settings.title != "" && com.rowNum == 0) {
                    //    com.data[0] = [];
                    //    com.merges[0] = "";
                    //    com.rowsH.push({ wpx: $.fn.tableExport.titleHeight });
                    //    com.rowNum++;
                    //}

                   

                    com.rowsH.push({ wpx: $(o).height()});

                    if (com.data[com.rowNum] == null) {
                        com.data[com.rowNum] = [];
                    }

                    var len = $(o).find("th:visible").length;

                    $(o).find("th:visible").not(settings.exclude).each(function (i, o) {

                        var rowspan = $(o).attr("rowspan");
                        // 是否合并列
                        if (rowspan !== undefined && parseInt(rowspan) > 1) {
                            // 合并信息 ( s :起始单元格(c:列 , r: 行) e:结束单元格(同上))
                            com.merges.push({ s: { c: com.colsIndex, r: com.rowNum }, e: { c: com.colsIndex, r: com.rowNum + parseInt(rowspan) - 1 } });

                            for (var k = 1; k < rowspan; k++) {

                                // 如果合并行 ， 则自动将行数据补全（如果合并两行 ， 则将下一行该列数据赋值为:"",保持每行数据数一致）
                                if (com.data[com.rowNum + k] == null) {
                                    com.data[com.rowNum + k] = [];
                                }
                                // {}:st:样式 v:内容
                                com.data[com.rowNum + k][com.data[com.rowNum + k].length] = { st: '2', v: "" };
                            }
                        }


                        var colspan = $(o).attr("colspan");
                        var w;
                        if (colspan !== undefined && parseInt(colspan) > 1) {

                            com.merges.push({ s: { c: com.colsIndex, r: com.rowNum }, e: { c: com.colsIndex + parseInt(colspan) - 1, r: com.rowNum } });

                            for (var j = 0; j < colspan; j++) {
                                w = $(o).width() / colspan;
                                // 每一列宽度
                                com.colsW[com.colsW.length] = { wpx: w };
                                if (j == 0) {
                                    com.data[com.rowNum][com.data[com.rowNum].length] = { st: '2', v: $.trim($.trim($(o).text())) };
                                } else {
                                    com.data[com.rowNum][com.data[com.rowNum].length] = { st: '2', v: "" };
                                }
                            }
                            com.colsIndex += parseInt(colspan);
                        } else {
                            com.colsW[com.colsW.length] = { wpx: $(o).width() };
                            com.data[com.rowNum][com.data[com.rowNum].length] = { st: '2', v: $.trim($.trim($(o).text())) };
                            com.colsIndex++;
                        }



                    });
                    $(o).find("td:visible").not(settings.exclude).each(function (i, o) {

                        var rowspan = $(o).attr("rowspan");

                        if (rowspan !== undefined && parseInt(rowspan) > 1) {

                            com.merges.push({ s: { c: com.colsIndex, r: com.rowNum }, e: { c: com.colsIndex, r: com.rowNum + parseInt(rowspan) - 1 } });

                            for (var k = 1; k < parseInt(rowspan) ; k++) {

                                if (com.data[com.rowNum + k] == null) {
                                    com.data[com.rowNum + k] = [];
                                }

                                com.data[com.rowNum + k][com.data[com.rowNum + k].length] = { st: '1', v: "" };;
                            }
                        }

                        var colspan = $(o).attr("colspan");
                        if (colspan !== undefined && parseInt(colspan) > 1) {
                            com.merges.push({ s: { c: com.colsIndex, r: com.rowNum }, e: { c: com.colsIndex + parseInt(colspan) - 1, r: com.rowNum } });
                            for (var j = 0; j < colspan; j++) {
                                w = $(o).width() / colspan;
                                com.colsW[com.colsW.length] = { wpx: w};
                                if (j == 0) {
                                    if ($(o).find("img").length > 0) {
                                        com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: $(o).find("img") };
                                    } else {
                                        com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: $.trim($.trim($(o).text())) };
                                    }
                                    //com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: $.trim($.trim($(o).text())) };
                                } else {
                                    com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: "" };
                                }
                            }
                            com.colsIndex += parseInt(colspan);
                        } else {
                            com.colsW[com.colsW.length] = { wpx: $(o).width() };
                            //if ($(o).is('img')) {
                            //    com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: $(o) };
                            //} else {
                            //    com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: $.trim($.trim($(o).text())) };
                            //}

                            if ($(o).find("img").length > 0) {
                                com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: $(o).find("img") };
                            } else {
                                com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: $.trim($.trim($(o).text())) };
                            }

                            //com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: $.trim($.trim($(o).text())) };
                            com.colsIndex++;
                        }

                    });

                    //if (settings.title != "" && i == 0) {
                    //    for (var k = 0; k < com.colsIndex - 1; k++) {
                    //        if (k == 0) {
                    //            com.data[0][k] = { st: '4', v: settings.title };
                    //        } else {
                    //            com.data[0][k] = { st: '4', v: "" };
                    //        }
                    //    }
                    //    com.merges[0] = { s: { c: 0, r: 0 }, e: { c: com.colsIndex - 1, r: 0 } };
                    //}
                    com.rowNum++;
                });
            });

            _getGridTitle();

            com._getFrozenInfo();

            var d = {
                data: com.data,
                colsW: com.colsW,
                rowsH: com.rowsH,
                merges: com.merges
            }

            if (com.frozen) {
                d.frozen = com.frozen;
            }

            return d;
        },

        _ligerGridData: function (type, grid, dataSource) {

            this._reset();

            var g = $(e).parent().parent();

            var title = $(g).find(".l-layout-header").text();

            var c = $(g).find(".l-grid2 .l-grid-header-table");

            var c1 = $(g).find(".l-grid1 .l-grid-header-table");

            var com = this;

            var _getGridTitle = function () {

                if (settings.title != "" && com.rowNum == 0) {
                    com.data[0] = [];
                    com.merges[0] = "";
                    com.rowsH.push({ wpx: $.fn.tableExport.titleHeight });
                    com.rowNum++;
                }

                //com.colsW = [];

                //com.colsIndex = 0;

                //_getGridTitleByDom(c1);

                _getGridTitleByDom();

                //com.rowNum++;

                if (settings.title != "") {
                    for (var k = 0; k < com.colsIndex; k++) {
                        if (k == 0) {
                            com.data[0][k] = { st: '4', v: settings.title };
                        } else {
                            com.data[0][k] = { st: '4', v: "" };
                        }
                    }
                    com.merges[0] = { s: { c: 0, r: 0 }, e: { c: com.colsIndex - 1, r: 0 } };
                }
            }

            var _getGridTitleByDom = function () {

                var c = $(g).find(".l-grid2 .l-grid-header-table");

                var c1 = $(g).find(".l-grid1 .l-grid-header-table");

                var _com = this;

                var _temp = 0; // 用户临时存储表格列数 (用于 双层表头，取列数最多为标题列数)
                // 标题
                $(c1).find("tr").each(function (i, o) {

                    com.colsW = [];

                    com.colsIndex = 0;

                    _getGridTdData(o);

                    var _oo = $(c).find("tr").eq(i);

                    _getGridTdData(_oo);

                    com.rowNum++;

                    if (_temp > com.colsIndex) {
                        com.colsIndex = _temp;
                    } else {
                        _temp = com.colsIndex;
                    }
                });

                function _getGridTdData(o) {

                    com.rowsH.push({ wpx: $(o).height()});

                    if (com.data[com.rowNum] == null) {
                        com.data[com.rowNum] = [];
                    }

                    // 第一个单元格
                    //com.colsW[com.colsW.length] = { wpx: settings.ligerui.serialNumberWidth === undefined ? "27" : settings.ligerui.serialNumberWidth };
                    //com.data[com.rowNum][com.data[com.rowNum].length] = { st: '2', v: "" };
                    //com.colsIndex++;

                    $(o).find("td:visible").each(function (i, o) {

                        var rowspan = $(o).attr("rowspan");

                        if (rowspan !== undefined && parseInt(rowspan) > 1) {

                            com.merges.push({ s: { c: com.colsIndex, r: com.rowNum }, e: { c: com.colsIndex, r: com.rowNum + parseInt(rowspan) - 1 } });

                            for (var k = 1; k < rowspan; k++) {

                                if (com.data[com.rowNum + k] == null) {
                                    com.data[com.rowNum + k] = [];
                                }

                                com.data[com.rowNum + k][com.data[com.rowNum + k].length] = { st: '2', v: "" };
                            }
                        }


                        var colspan = $(o).attr("colspan");
                        var w;
                        if (colspan !== undefined && parseInt(colspan) > 1) {

                            com.merges.push({ s: { c: com.colsIndex, r: com.rowNum }, e: { c: com.colsIndex + parseInt(colspan) - 1, r: com.rowNum } });

                            for (var j = 0; j < colspan; j++) {
                                w = $(o).width() / colspan;
                                com.colsW[com.colsW.length] = { wpx: w };
                                if (j == 0) {
                                    com.data[com.rowNum][com.data[com.rowNum].length] = { st: '2', v: $.trim($.trim($(o).text())) };
                                } else {
                                    com.data[com.rowNum][com.data[com.rowNum].length] = { st: '2', v: "" };
                                }
                            }
                            com.colsIndex += parseInt(colspan);
                        } else {
                            com.colsW[com.colsW.length] = { wpx: $(o).width() };
                            com.data[com.rowNum][com.data[com.rowNum].length] = { st: '2', v: $.trim($.trim($(o).text())) };
                            com.colsIndex++;
                        }
                    });
                }
            }

            var _getGridBoby = function () {
                // 内容
                var t = $(g).find(".l-grid2 .l-grid-body-table");

                var t1 = $(g).find(".l-grid1 .l-grid-body-table");


                com.colsW = [];

                com.colsIndex = 0;

                $(t1).find("tr").each(function (i, o) {
                    _getTableBody(o);
                    var _o = $(t).find("tr").eq(i);
                    _getTableBody(_o);
                    com.rowNum++;
                });
            }

            var _getTableBody = function (o) {

                com.rowsH.push({ wpx: $(o).height()});

                if (com.data[com.rowNum] == null) {
                    com.data[com.rowNum] = [];
                }

                //// 序号
                //com.colsW[com.colsW.length] = { wpx: "27" };
                //com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: i + 1 };
                //com.colsIndex++;

                $(o).find("td:visible").each(function (i, o) {

                    com.colsW[com.colsW.length] = { wpx: $(o).width() };

                    if ($(o).find("img").length > 0) {
                        com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: $(o).find("img") };
                    } else {
                        com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: $.trim($.trim($(o).text())) };
                    }
                      
                    //if ($(o).is('img')) {
                    //    com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: $(o) };
                    //} else {
                    //    com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: $.trim($.trim($(o).text())) };
                    //}

                    //com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: $.trim($(o).text()) };
                    com.colsIndex++;
                });
            }

            var _getGridBodyFromServer = function () {

                var cols = grid.columns;

                var rows = dataSource.Rows;

                for (var i = 0; i < rows.length; i++) {

                    com.colsIndex = 0;


                    if (com.data[com.rowNum] == null) {
                        com.data[com.rowNum] = [];
                    }

                    // 序号
                    com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: i + 1 };
                    com.colsIndex++;

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

                            rowName = com._nvl(cols[j].render(rowdata, i));

                            if (com._isHtml(rowName)) {

                                if ($(rowName).is('img')) {
                                    rowName = $(rowName);
                                } else {
                                    rowName = $(rowName).html();
                                }
                            }

                        } else if (name === undefined || name == "") {

                            continue;

                        } else {
                            rowName = com._nvl(rows[i][name]);
                        }



                        //var rowName = com._getRowValue(rows[i], name);

                        if (String(rowName).indexOf("/Date") != -1) {
                            rowName = com._ChangeDateFormat_Date(rowName);
                        }

                        com.data[com.rowNum][com.data[com.rowNum].length] = { st: '3', v: rowName };
                        com.colsIndex++;
                    }
                    com.rowNum++;
                }
            }

            _getGridTitle();

            if (type == 1 && dataSource != null) {
                _getGridBodyFromServer();
            } else {
                _getGridBoby();
            }

            com._getFrozenInfo();

            var d = {
                data: com.data,
                colsW: com.colsW,
                rowsH: com.rowsH,
                merges: com.merges,
                frozen: com.frozen
            }
            return d;
        },

        _getFrozenInfo: function () {

            if (!settings.frozen) {
                return;
            }

            this.frozen = { topLeftCell: { c: settings.frozen.c, r: settings.frozen.r }, ySplit: settings.frozen.r, xSplit: settings.frozen.c };

            //this.frozen = { topLeftCell: { c: 0, r: 2 }, ySplit:2 };

        },

        _getRowValue: function (row, name) {
            if (settings.ligerui.RenderColumns !== undefined
                && settings.ligerui.RenderColumns[name] !== undefined
                && settings.ligerui.RenderColumns[name].render !== undefined) {
                return settings.ligerui.RenderColumns[name].render(row);
            }
            if (row[name] === undefined || row[name] == null) {
                return "";
            } else {
                return row[name];
            }
        },

        // 分割数据 (用于cvs , txt , xls格式数据)
        _getSeqData: function (data, seq1, seq2) {
            if (data == null) {
                return;
            }
            var d = data.data;

            var ar = [];
            for (var i = 0; i < d.length; i++) {
                var arr = [];
                for (var j = 0; j < d[i].length; j++) {
                    arr.push(d[i][j].v);
                }
                ar.push(arr.join(seq1));
            }
            return ar.join(seq2);
        },

        _tableToExcel: function (e) {
            var css = [];
            if (settings.islageruigrid && (settings.css === undefined || settings.css == null || settings.css.length == 0)) {
                css = ["/RefLib/excel/css/ligerui_grid_export_excel.css"];
            } else {
                css = settings.css;
            }

            e.table2excel({
                exclude: settings.exclude,
                name: "Excel Document Name",
                filename: settings.fileName,
                sheetName: settings.fileName,
                fileext: ".xls",
                css: css,
                islageruigrid: settings.islageruigrid,
                ligerui: settings.ligerui
            });
        },

        _escapeHtml: function (string) {
            return String(string).replace(/[&<>'\/]/g, function (s) {
                return $.fn.tableExport.entityMap[s];
            });
        },

        _dateNum: function (v, date1904) {
            if (date1904) v += 1462;
            var epoch = Date.parse(v);
            return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
        },

        // 创建sheet页 (xlsx)
        _createSheet: function (d, opts) {
            var ws = {};
            var cols = [];
            var images = [];
            var data = d.data;
            var range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
            for (var R = 0; R != data.length; ++R) {
                for (var C = 0; C != data[R].length; ++C) {
                    if (range.s.r > R) range.s.r = R;
                    if (range.s.c > C) range.s.c = C;
                    if (range.e.r < R) range.e.r = R;
                    if (range.e.c < C) range.e.c = C;
                    var cell = { v: data[R][C] };
                    if (cell.v == null) continue;
                    var cell_ref = XLSX.utils.encode_cell({ c: C, r: R });

                    //var reg = /^[+-]?(\d|[0-9]\d|\d)(.\d{1,2})?%$/;

                    var reg = /^[+-]?((\d+\.?\d*)|(\d*\.\d+))\%$/;
                    if (typeof cell.v.v === 'number' || !isNaN(Number(cell.v.v))) {
                        cell.v.t = 'n';
                    } else if (cell.v.v instanceof jQuery && cell.v.v.prop('tagName') == 'IMG') {
                        images.push({
                            c: C,
                            r: R,
                            element: cell.v.v
                        })
                        cell.v.v = '';
                    } else if (typeof cell.v.v === 'boolean') {
                        cell.v.t = 'b';
                    } else if (cell.v.v instanceof Date || !isNaN(Date.parse(cell.v.v))) {
                        cell.v.t = 's';
                        cell.v.z = XLSX.SSF._table[23];
                        //cell.v.v = dateNum(Date.parse(cell.v.v));
                    } else if (cell.v.v.match(reg) != null) {
                        cell.v.t = 'p';
                    } else {
                        cell.v.t = 's';
                    }
                    //if (cell.v.st == '4') {
                    //    cell.v.s = $.fn.tableExport.Style[0];
                    //}


                    ws[cell_ref] = cell.v;
                }
            }
            if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);


            ws["!images"] = [];

            $.each(images, function (index, image) {

                ws["!images"].push({
                    name: 'image' + index + '.png',
                    data: imgToDataUrl(image.element[0]),
                    opts: { base64: true },
                    type: "png",
                    position: {
                        type: 'twoCellAnchor',
                        attrs: { editAs: 'oneCell' },
                        from: { col: image.c, row: image.r },
                        to: { col: image.c + 1, row: image.r + 1 }
                    }
                });
            });

            ws['!cols'] = d.colsW;
            ws['!rows'] = d.rowsH;
            ws['!merges'] = d.merges;
            if (d.frozen != null) {
                ws['!frozen'] = d.frozen;
            }
            return ws;
        },

        _string2ArrayBuffer: function (s) {
            var buf = new ArrayBuffer(s.length);
            var view = new Uint8Array(buf);
            for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
            return buf;
        },

        _ChangeDateFormat_Date: function (data) {
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

        _export2file: function (data, mime, name, extension, opts) {
            if (extension === ".xlsx") {
                var wb = new Workbook(),
                    ws = this._createSheet(data, opts);

                wb.SheetNames.push(name);
                wb.Sheets[name] = ws;

                var wopts = { bookType: 'xlsx', bookSST: false, type: 'binary' },
                    wbout = XLSX.write(wb, wopts);

                data = this._string2ArrayBuffer(wbout);
            }

            saveAs(new Blob([data],
                { type: mime + ";" + $.fn.tableExport.charset }),
                name + extension);
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
    }

    function Workbook() {
        if (!(this instanceof Workbook)) return new Workbook();
        this.SheetNames = [];
        this.Sheets = {};
    }

    function imgToDataUrl(img) {
        var canvas = document.createElement('canvas');
        canvas.width = img.naturalWidth; // or 'width' if you want a special/scaled size
        canvas.height = img.naturalHeight; // or 'height' if you want a special/scaled size

        canvas.getContext('2d').drawImage(img, 10, 10);
        return canvas.toDataURL('image/png').replace(/^data:image\/(png|jpg);base64,/, '');
    }

}(window));