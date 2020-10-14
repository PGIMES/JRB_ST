<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Report_ST.aspx.cs" Inherits="Report_ST" %>

<%@ Register Assembly="DevExpress.Web.v17.2, Version=17.2.4.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web" TagPrefix="dx" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>渗透实验明细</title>
    <link href="/css/Site.css" rel="stylesheet" type="text/css" />
    <link href="/css/bootstrap.min.css" rel="stylesheet" type="text/css" />
    <link href="/css/font-awesome.min.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .area {
            height: 90px;
            width: 90px;
            vertical-align: middle;
            text-align: center;
            padding-top: 3px;
            font-size: smaller;
        }

        .area_lg {
            height: 100px;
            width: 100px;
            vertical-align: middle;
            text-align: center;
            padding-top: 3px;
        }

        .area_x_lg {
            height: 100px;
            width: 120px;
            vertical-align: middle;
            text-align: left;
            padding-top: 2px;
            font-size: 11px;
        }

        .area_border_gray {
            border: 2px solid gray;
        }

        .area_block {
            float: left;
            padding-left: 5px;
            padding-bottom: 3px;
        }

        .font-info {
            font-size: x-small;
        }

        .btn-gray {
            background: gray;
        }

        .btn-yellow {
            background: #FF9933;
        }

        .btn-red {
            background: Red;
        }

        .row-container {
            padding-left: 5px;
            padding-right: 5px;
        }

        .form-input {
            display: float;
            width: 50%;
        }

        .btn-group > .btn:first-child:not(:last-child):not(.dropdown-toggle) {
            border-top-right-radius: 0;
            border-bottom-right-radius: 0;
        }

        .btn-group > .btn:first-child {
            margin-left: 0;
        }

        .btn-group-vertical > .btn, .btn-group > .btn {
            position: relative;
            float: left;
        }

        .btn-primary {
            color: #fff;
            background-color: #337ab7;
            border-color: #2e6da4;
        }

        .btn-padding-s {
            display: inline-block;
            padding: 3px 3px;
            margin-bottom: 0;
            font-size: 14px;
            font-weight: 400;
            line-height: 1.42857143;
            text-align: center;
            white-space: nowrap;
            vertical-align: middle;
            -ms-touch-action: manipulation;
            touch-action: manipulation;
            cursor: pointer;
            -webkit-user-select: none;
            -moz-user-select: none;
            -ms-user-select: none;
            user-select: none;
            background-image: none;
            border: 1px solid transparent;
            border-radius: 4px;
        }

        .area_drop {
            width: 20px;
            height: 20px;
            font-size: 12px;
        }

        .head {
            font-size: 25px;
            text-align: left;
            vertical-align: bottom;
        }

        .input-edit {
            background-color: Yellow;
        }
    </style>
    <script src="/js/jquery.min.js" type="text/javascript"></script>
    <script src="/js/bootstrap.min.js" type="text/javascript"></script>
    <script src="/js/layer/layer.js"></script>

    <script type="text/javascript">
        $(document).ready(function () {
            $("#mestitle").html("【渗透实验明细】");

            setHeight();
            $(window).resize(function () {
                setHeight();
            });

            $('#btn_upload').click(function () {
                var url = "ST_upload.aspx";

                layer.open({
                    title: '上传渗透数据',
                    closeBtn: 2,
                    type: 2,
                    area: ['600px', '350px'],
                    fixed: false, //不固定
                    maxmin: true, //开启最大化最小化按钮
                    content: url
                });
            });

        });
        
        function setHeight() {
            $("div[class=dxgvCSD]").css("height", ($(window).height() - $("#div_p").height() - 220) + "px");
        }
        	
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <div>
        <div class="mainTop">
            <div class="mainTopLeft" style="float: left;">
                <table>
                    <tr>
                        <td>
                            <div id="divImg"></div>
                            <div class="h3" style="margin-left: 10px" id="headTitle">
                                PGI管理系统<div class="btn-group">
                                    <div class="area_drop" data-toggle="dropdown">
                                        <span class="caret"></span>
                                    </div>
                                </div>
                                <span id="mestitle"></span>
                            </div>
                        </td>
                    </tr>
                </table>
            </div>
        </div>
    </div>
    <div class="panel-body" id="div_p">
        <div class="col-sm-12">
            <table style="line-height:40px">
                <tr>                    
                    <td style="text-align:right;">日期：</td>
                    <td>
                        <asp:TextBox ID="txt_ym" runat="server" class="form-control" Width="120px" />
                    </td>
                    <td style="text-align:right;">&nbsp;&nbsp;机台：</td>
                    <td>
                        <asp:TextBox ID="txt_jq" runat="server" class="form-control" Width="120px" />
                    </td>   
                    <td> 
                        &nbsp;
                        <asp:Button ID="Bt_select" runat="server" Text="查询" class="btn btn-large btn-primary" OnClick="Bt_select_Click" Width="70px" /> 
                        &nbsp;
                        <asp:Button ID="Bt_Export" runat="server" class="btn btn-large btn-primary" OnClick="Bt_Export_Click" Text="导出" Width="70px" /> 
                        &nbsp;
                        <button id="btn_upload" type="button" class="btn btn-primary btn-large">上传渗透数据</button>
                    </td> 
                </tr>
            </table>
        </div>
    </div>
    <div runat="server" id="DIV1" style="margin-left: 5px; margin-right: 5px; margin-bottom: 10px">
        <table>
            <tr>
                <td>
                    <dx:ASPxGridView ID="GV_PART" ClientInstanceName="grid" runat="server" KeyFieldName="id" AutoGenerateColumns="False"  
                            OnCustomCellMerge="GV_PART_CustomCellMerge" OnPageIndexChanged="GV_PART_PageIndexChanged" 
                            OnExportRenderBrick="GV_PART_ExportRenderBrick">
                        <ClientSideEvents EndCallback="function(s, e) { setHeight(); }" />
                        <SettingsBehavior AllowDragDrop="TRUE" AllowFocusedRow="false" AllowSelectByRowClick="false" ColumnResizeMode="Control" AutoExpandAllGroups="true" MergeGroupsMode="Always" SortMode="Value" 
                                 AllowCellMerge="true"/>
                        <SettingsPager PageSize="100"></SettingsPager>
                        <Settings ShowFilterRow="True" ShowFilterRowMenu="True"   VerticalScrollBarMode="Visible" VerticalScrollBarStyle="Standard" VerticalScrollableHeight="500"
                                 ShowFilterRowMenuLikeItem="True"  ShowFooter="True"  HorizontalScrollBarMode="Auto" />
                        <Columns>
                            <dx:GridViewDataTextColumn Caption="机台" FieldName="jq" Width="60px" VisibleIndex="1" />
                            <dx:GridViewDataDateColumn Caption="日期" FieldName="date" Width="90px" VisibleIndex="2">
                                <PropertiesDateEdit DisplayFormatString="yyyy/MM/dd"></PropertiesDateEdit>
                            </dx:GridViewDataDateColumn>
                            <dx:GridViewDataTextColumn Caption="材料颜色" FieldName="material_color" Width="100px" VisibleIndex="3" PropertiesTextEdit-EncodeHtml="false" />
                            <dx:GridViewDataTextColumn Caption="产品结构" FieldName="product_code" Width="130px" VisibleIndex="4" PropertiesTextEdit-EncodeHtml="false" />
                            <dx:GridViewDataTextColumn Caption="批号" FieldName="batch_no" Width="70px" VisibleIndex="5" />
                            <dx:GridViewDataTextColumn Caption="容器" FieldName="cavity" Width="40px" VisibleIndex="6" />
                            <dx:GridViewDataTextColumn Caption="开始重量" FieldName="start_weight" Width="70px" VisibleIndex="7" />

                            <dx:GridViewDataTextColumn Caption="2D重量" FieldName="D2_weight" Width="70px" VisibleIndex="8" /> 
                            <dx:GridViewDataTextColumn Caption="2D失重" FieldName="D2_loss_weight" Width="70px" VisibleIndex="9" /> 
                            <dx:GridViewDataTextColumn Caption="2D失重率" FieldName="D2_loss_rate" Width="70px" VisibleIndex="10" > 
                                <PropertiesTextEdit DisplayFormatString="{0:P2}"></PropertiesTextEdit>
                            </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="2D备注" FieldName="D2_remark" Width="50px" VisibleIndex="11" /> 

                            <dx:GridViewDataTextColumn Caption="4D重量" FieldName="D4_weight" Width="70px" VisibleIndex="12" /> 
                            <dx:GridViewDataTextColumn Caption="4D失重" FieldName="D4_loss_weight" Width="70px" VisibleIndex="13" />
                            <dx:GridViewDataTextColumn Caption="4D失重率" FieldName="D4_loss_rate" Width="70px" VisibleIndex="14" > 
                                <PropertiesTextEdit DisplayFormatString="{0:P2}"></PropertiesTextEdit>
                            </dx:GridViewDataTextColumn> 
                            <dx:GridViewDataTextColumn Caption="4D备注" FieldName="D4_remark" Width="50px" VisibleIndex="15" /> 

                            <dx:GridViewDataTextColumn Caption="7D重量" FieldName="D7_weight" Width="70px" VisibleIndex="16" /> 
                            <dx:GridViewDataTextColumn Caption="7D失重" FieldName="D7_loss_weight" Width="70px" VisibleIndex="17" /> 
                            <dx:GridViewDataTextColumn Caption="7D失重率" FieldName="D7_loss_rate" Width="70px" VisibleIndex="18" > 
                                <PropertiesTextEdit DisplayFormatString="{0:P2}"></PropertiesTextEdit>
                            </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="7D备注" FieldName="D7_remark" Width="50px" VisibleIndex="19" /> 

                            <dx:GridViewDataTextColumn Caption="超标处理结果" FieldName="deal_result" Width="80px" VisibleIndex="20" >
                                <HeaderStyle BackColor="#F0E68C" />

                            </dx:GridViewDataTextColumn>
                            <dx:GridViewDataDateColumn Caption="上传日期" FieldName="upload_time" Width="90px" VisibleIndex="21" >
                                <PropertiesDateEdit DisplayFormatString="yyyy/MM/dd"></PropertiesDateEdit>
                                <HeaderStyle BackColor="#F0E68C" />
                            </dx:GridViewDataDateColumn>
                            <dx:GridViewDataTextColumn Caption="文件名称" FieldName="ori_filename" Width="250px" VisibleIndex="22">
                                <HeaderStyle BackColor="#F0E68C" />
                                <DataItemTemplate>
                                    <dx:ASPxHyperLink ID="hpl_ori_filename" runat="server" Text='<%# Eval("ori_filename")%>' Cursor="pointer"
                                        NavigateUrl='<%# "/UploadFile/ST/"+ Eval("new_filename") %>'  
                                        Target="_blank">                                        
                                    </dx:ASPxHyperLink>
                                </DataItemTemplate> 
                                <Settings AllowAutoFilterTextInputTimer="False" />
                            </dx:GridViewDataTextColumn> 
                            <dx:GridViewDataTextColumn Caption="id" FieldName="id" VisibleIndex="99" Width="0px"                                
                                 HeaderStyle-CssClass="hidden" CellStyle-CssClass="hidden" FooterCellStyle-CssClass="hidden"></dx:GridViewDataTextColumn>
                        </Columns>
                        <Styles>
                            <Header BackColor="#99CCFF"></Header>
                            <FocusedRow BackColor="#99CCFF" ForeColor="#0000CC"></FocusedRow>
                            <Footer HorizontalAlign="Right"></Footer>
                            <AlternatingRow Enabled="true" />
                            <Footer ForeColor="Red" Font-Bold="true"></Footer>
                        </Styles>
                    </dx:ASPxGridView>

                </td>
            </tr>
        </table>
    </div>
    <div class="footer"></div>
    </form>
</body>
</html>
