# ExportExcelLigerUI
使用:

               $("#").tableExport({
                                title: "excel标题名称",
                                fileName: "excel导出文件名称",
                                fileType: "xlsx",  // xls格式:导出使用table2excel.js (需传入css样式,否则无样式,默认导出xlsx)
                                islageruigrid: true, // 是否基于ligerui
                                ligerui: {
                                    isGetAll: true, // 是否全部导出数据
                                    gridInstance: g, // ligerGrid实例
                                    serialNumberWidth: "50"
                                },
                                frozen: { c: 1, r: 2 } //冻结列 (c:列 r:行)
                            });


默认属性:
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
