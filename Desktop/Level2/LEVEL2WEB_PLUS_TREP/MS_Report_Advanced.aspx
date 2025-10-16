<%@ Register TagPrefix="Ghnnam" TagName="NavBar" Src="NavBar.ascx" %>
<%@ Page Language="vb" AutoEventWireup="false" CodeFile="MS_Report_Advanced.aspx.vb" Inherits="LEVEL2WEB.MS_Report_Advanced" %>
<%@ Register Assembly="System.Web.DataVisualization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" Namespace="System.Web.UI.DataVisualization.Charting" TagPrefix="asp" %>
<%@ Register Assembly="RJS.Web.WebControl.PopCalendar" Namespace="RJS.Web.WebControl" TagPrefix="rjs" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head id="Head1" runat="server">
    <title>Melt Shop Advanced Report</title>
    <meta content="True" name="vs_showGrid">
    <style type="text/css">
        body { font-family: Verdana, Arial, sans-serif; background-color: #DCE8F2; }
        .filter-panel { background-color: #0099ff; color: #ffffff; padding: 12px 18px; border-radius: 4px; margin-bottom: 12px; }
        .filter-panel label { font-weight: bold; margin-right: 6px; }
        .filter-row { display: flex; flex-wrap: wrap; gap: 18px; align-items: center; }
        .filter-row .field { display: flex; flex-direction: column; min-width: 140px; }
        .filter-row .field span { font-size: 0.8em; margin-bottom: 4px; }
        .filter-row .field input[type="text"] { padding: 4px 6px; border: 1px solid #dddddd; border-radius: 3px; font-size: 0.8em; }
        .filter-row .field .aspNetDisabled { background-color: #f0f0f0; }
        .filter-actions { margin-top: 10px; display: flex; gap: 12px; align-items: center; }
        .filter-actions .status { font-size: 0.8em; font-weight: bold; color: #ffe49f; }
        .content-container { padding: 16px 20px; background-color: #ffffff; border-radius: 6px; box-shadow: 0 2px 6px rgba(0, 0, 0, 0.12); }
        .summary-banner { margin-bottom: 16px; font-size: 0.85em; color: #0f3f63; display: flex; gap: 20px; flex-wrap: wrap; }
        .summary-cards { display: flex; flex-wrap: wrap; gap: 16px; margin-bottom: 18px; }
        .summary-card { flex: 1 1 220px; background-color: #f5fbff; border: 1px solid #bbd8f0; border-radius: 6px; padding: 12px 14px; min-width: 200px; }
        .summary-card h3 { font-size: 0.95em; margin: 0 0 6px 0; color: #005a9b; }
        .summary-card ul { list-style: none; padding: 0; margin: 0; font-size: 0.8em; }
        .summary-card ul li { margin-bottom: 4px; }
        .chart-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 18px; margin-bottom: 18px; }
        .chart-wrapper { background-color: #f8fbff; border: 1px solid #d0e4f5; border-radius: 6px; padding: 10px; }
        .chart-wrapper h4 { margin: 0 0 8px 0; font-size: 0.85em; color: #074a7c; }
        .grid-wrapper { margin-top: 12px; }
        .grid-wrapper h4 { font-size: 0.85em; color: #074a7c; margin-bottom: 6px; }
        .detail-grid { width: 100%; border-collapse: collapse; font-size: 0.78em; }
        .detail-grid th, .detail-grid td { border: 1px solid #c5d8ea; padding: 6px 8px; text-align: center; }
        .detail-grid th { background-color: #3399ff; color: #ffffff; }
        .detail-grid tr:nth-child(even) { background-color: #f3f7fb; }
        .menu-container { margin-bottom: 10px; }
        .status-message { color: #b30000; font-size: 0.8em; font-weight: bold; }
        .legend-note { font-size: 0.75em; color: #415b76; margin-top: -8px; margin-bottom: 12px; }
        .spacer { height: 12px; }
        .tab-header { font-size: 0.9em; color: #174468; margin-bottom: 10px; font-weight: bold; }
    </style>
</head>
<body>
    <form id="Form1" method="post" runat="server">
        <table id="TableLayout" style="z-index: 104; left: -3px; position: absolute; top: -2px;" cellspacing="0" cellpadding="0" width="100%" align="left" border="0">
            <tr>
                <td valign="top" align="left" colspan="4" style="height: 82px; background-image: url('images/Top_jpg/New_Header.jpg');"></td>
            </tr>
            <tr>
                <td style="width: 89px" valign="top" align="left" width="89" bgcolor="#83d7ff">
                    <p align="left"><Ghnnam:NavBar ID="Navbar1" runat="server" /></p>
                </td>
                <td valign="top" align="left" colspan="3">
                    <div class="content-container">
                        <div class="filter-panel">
                            <div class="filter-row">
                                <div class="field">
                                    <span>Period</span>
                                    <asp:RadioButtonList ID="rblPeriod" runat="server" RepeatDirection="Horizontal">
                                        <asp:ListItem Text="Day" Value="Day" Selected="True"></asp:ListItem>
                                        <asp:ListItem Text="Month" Value="Month"></asp:ListItem>
                                        <asp:ListItem Text="Quarter" Value="Quarter"></asp:ListItem>
                                        <asp:ListItem Text="Year" Value="Year"></asp:ListItem>
                                        <asp:ListItem Text="Range" Value="Range"></asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div class="field">
                                    <span>From</span>
                                    <asp:TextBox ID="txtFromDate" runat="server" ReadOnly="True"></asp:TextBox>
                                    <rjs:PopCalendar ID="pcFrom" runat="server" Control="txtFromDate" Font-Names="Verdana" Font-Size="0.8em" From-Date="1999-04-13" />
                                </div>
                                <div class="field">
                                    <span>To</span>
                                    <asp:TextBox ID="txtToDate" runat="server" ReadOnly="True"></asp:TextBox>
                                    <rjs:PopCalendar ID="pcTo" runat="server" Control="txtToDate" Font-Names="Verdana" Font-Size="0.8em" From-Date="1999-04-13" />
                                </div>
                                <div class="field">
                                    <span>Shift</span>
                                    <asp:RadioButtonList ID="rblShift" runat="server" RepeatDirection="Horizontal">
                                        <asp:ListItem Text="All" Value="All" Selected="True"></asp:ListItem>
                                        <asp:ListItem Text="Shift 1 (12AM-8AM)" Value="Shift1"></asp:ListItem>
                                        <asp:ListItem Text="Shift 2 (8AM-4PM)" Value="Shift2"></asp:ListItem>
                                        <asp:ListItem Text="Shift 3 (4PM-12AM)" Value="Shift3"></asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </div>
                            <div class="filter-actions">
                                <asp:Button ID="btnApply" runat="server" Text="Apply Filter" CssClass="apply-button" />
                                <span class="status">
                                    <asp:Label ID="lblFilterInfo" runat="server"></asp:Label>
                                </span>
                            </div>
                            <asp:Label ID="lblStatus" runat="server" CssClass="status-message"></asp:Label>
                        </div>

                        <div class="summary-banner">
                            <span><strong>Heat Range:</strong> <asp:Literal ID="litHeatRange" runat="server" /></span>
                            <span><strong>Period:</strong> <asp:Literal ID="litPeriodDescription" runat="server" /></span>
                            <span><strong>Shift:</strong> <asp:Literal ID="litShiftDescription" runat="server" /></span>
                        </div>

                        <div class="menu-container">
                            <asp:Menu ID="menuAreas" runat="server" Orientation="Horizontal" Font-Names="Verdana" Font-Size="0.8em" ForeColor="Black" OnMenuItemClick="menuAreas_MenuItemClick">
                                <Items>
                                    <asp:MenuItem Text="Summary" Value="0" Selected="True"></asp:MenuItem>
                                    <asp:MenuItem Text="EAF" Value="1"></asp:MenuItem>
                                    <asp:MenuItem Text="CCM" Value="2"></asp:MenuItem>
                                    <asp:MenuItem Text="LRF" Value="3"></asp:MenuItem>
                                </Items>
                                <StaticMenuStyle BackColor="#83D7FF" />
                                <StaticHoverStyle BackColor="#0099FF" BorderColor="Black" BorderStyle="Solid" BorderWidth="1px" />
                            </asp:Menu>
                        </div>

                        <asp:MultiView ID="mvAreas" runat="server" ActiveViewIndex="0">
                            <asp:View ID="viewSummary" runat="server">
                                <div class="tab-header">Overall Melt Shop Overview</div>
                                <div class="summary-cards">
                                    <div class="summary-card">
                                        <h3>EAF Highlights</h3>
                                        <ul>
                                            <li>Total Scrap: <asp:Literal ID="litSummaryEafScrap" runat="server" /></li>
                                            <li>DRI (All): <asp:Literal ID="litSummaryEafDri" runat="server" /></li>
                                            <li>Liquid Steel: <asp:Literal ID="litSummaryEafLiquid" runat="server" /></li>
                                            <li>EAF Yield: <asp:Literal ID="litSummaryEafYield" runat="server" /></li>
                                        </ul>
                                    </div>
                                    <div class="summary-card">
                                        <h3>CCM Highlights</h3>
                                        <ul>
                                            <li>Billet 12M: <asp:Literal ID="litSummaryCcm12m" runat="server" /></li>
                                            <li>Billet 8M: <asp:Literal ID="litSummaryCcm8m" runat="server" /></li>
                                            <li>Short Billets: <asp:Literal ID="litSummaryCcmShort" runat="server" /></li>
                                            <li>Total CCM Wgt: <asp:Literal ID="litSummaryCcmWeight" runat="server" /></li>
                                        </ul>
                                    </div>
                                    <div class="summary-card">
                                        <h3>LRF Highlights</h3>
                                        <ul>
                                            <li>Total Lime: <asp:Literal ID="litSummaryLrfLime" runat="server" /></li>
                                            <li>Alloying (SiMn/FeMn): <asp:Literal ID="litSummaryLrfAlloys" runat="server" /></li>
                                            <li>LF Power: <asp:Literal ID="litSummaryLrfPower" runat="server" /></li>
                                            <li>Avg Stir Time: <asp:Literal ID="litSummaryLrfStir" runat="server" /></li>
                                        </ul>
                                    </div>
                                    <div class="summary-card">
                                        <h3>Plant Yield</h3>
                                        <ul>
                                            <li>CCM Yield: <asp:Literal ID="litSummaryCcmYield" runat="server" /></li>
                                            <li>Plant Yield: <asp:Literal ID="litSummaryPlantYield" runat="server" /></li>
                                            <li>Heats Count: <asp:Literal ID="litSummaryCcmHeats" runat="server" /></li>
                                        </ul>
                                    </div>
                                </div>
                                <div class="chart-grid">
                                    <div class="chart-wrapper">
                                        <h4>Production Weight by Area</h4>
                                        <asp:Chart ID="chartSummaryProduction" runat="server" Height="260px" Width="380px">
                                            <ChartAreas>
                                                <asp:ChartArea Name="ProductionArea"></asp:ChartArea>
                                            </ChartAreas>
                                            <Legends>
                                                <asp:Legend Name="Legend1"></asp:Legend>
                                            </Legends>
                                        </asp:Chart>
                                    </div>
                                    <div class="chart-wrapper">
                                        <h4>Energy Footprint</h4>
                                        <asp:Chart ID="chartSummaryEnergy" runat="server" Height="260px" Width="380px">
                                            <ChartAreas>
                                                <asp:ChartArea Name="EnergyArea"></asp:ChartArea>
                                            </ChartAreas>
                                            <Legends>
                                                <asp:Legend Name="Legend2"></asp:Legend>
                                            </Legends>
                                        </asp:Chart>
                                    </div>
                                </div>
                            </asp:View>

                            <asp:View ID="viewEaf" runat="server">
                                <div class="tab-header">Electric Arc Furnace (EAF)</div>
                                <div class="summary-cards">
                                    <div class="summary-card">
                                        <h3>Charge Mix</h3>
                                        <ul>
                                            <li>Scrap: <asp:Literal ID="litEafTotalScrap" runat="server" /></li>
                                            <li>DRI (Pellet + Feed): <asp:Literal ID="litEafTotalDri" runat="server" /></li>
                                            <li>HBI &amp; Pig Iron: <asp:Literal ID="litEafHbiPig" runat="server" /></li>
                                        </ul>
                                    </div>
                                    <div class="summary-card">
                                        <h3>Power &amp; Gas</h3>
                                        <ul>
                                            <li>Total Power (MWh): <asp:Literal ID="litEafTotalKwh" runat="server" /></li>
                                            <li>KWh / t Liquid: <asp:Literal ID="litEafKwhPerTon" runat="server" /></li>
                                            <li>Oxygen / t: <asp:Literal ID="litEafOxygenPerTon" runat="server" /></li>
                                        </ul>
                                    </div>
                                    <div class="summary-card">
                                        <h3>Process</h3>
                                        <ul>
                                            <li>Avg Tap-to-Tap: <asp:Literal ID="litEafAvgTapTime" runat="server" /></li>
                                            <li>Avg Power On: <asp:Literal ID="litEafAvgPowerOn" runat="server" /></li>
                                            <li>DRI / Charge: <asp:Literal ID="litEafDriRatio" runat="server" /></li>
                                        </ul>
                                    </div>
                                </div>
                                <div class="chart-grid">
                                    <div class="chart-wrapper">
                                        <h4>Charge Materials Mix</h4>
                                        <asp:Chart ID="chartEafMaterials" runat="server" Height="260px" Width="380px">
                                            <ChartAreas>
                                                <asp:ChartArea Name="MaterialsArea"></asp:ChartArea>
                                            </ChartAreas>
                                            <Legends>
                                                <asp:Legend Name="MaterialsLegend"></asp:Legend>
                                            </Legends>
                                        </asp:Chart>
                                    </div>
                                    <div class="chart-wrapper">
                                        <h4>Energy Distribution</h4>
                                        <asp:Chart ID="chartEafEnergy" runat="server" Height="260px" Width="380px">
                                            <ChartAreas>
                                                <asp:ChartArea Name="EafEnergyArea"></asp:ChartArea>
                                            </ChartAreas>
                                            <Legends>
                                                <asp:Legend Name="EafEnergyLegend"></asp:Legend>
                                            </Legends>
                                        </asp:Chart>
                                    </div>
                                    <div class="chart-wrapper">
                                        <h4>Liquid Steel Trend</h4>
                                        <asp:Chart ID="chartEafTrend" runat="server" Height="260px" Width="380px">
                                            <ChartAreas>
                                                <asp:ChartArea Name="EafTrendArea"></asp:ChartArea>
                                            </ChartAreas>
                                            <Legends>
                                                <asp:Legend Name="EafTrendLegend"></asp:Legend>
                                            </Legends>
                                        </asp:Chart>
                                        <div class="legend-note">Trend plotted per heat using tapping start time.</div>
                                    </div>
                                </div>
                                <div class="grid-wrapper">
                                    <h4>EAF Heat Detail</h4>
                                    <asp:GridView ID="gvEafDetail" runat="server" AutoGenerateColumns="False" CssClass="detail-grid">
                                        <Columns>
                                            <asp:BoundField DataField="Heat" HeaderText="Heat" />
                                            <asp:BoundField DataField="Shift" HeaderText="Shift" />
                                            <asp:BoundField DataField="StartTime" HeaderText="Start" />
                                            <asp:BoundField DataField="Grade" HeaderText="Grade" />
                                            <asp:BoundField DataField="LiquidSteel" HeaderText="Liq Steel (t)" DataFormatString="{0:N2}" />
                                            <asp:BoundField DataField="Kwh" HeaderText="Power (MWh)" DataFormatString="{0:N2}" />
                                            <asp:BoundField DataField="KwhPerTon" HeaderText="KWh/t" DataFormatString="{0:N2}" />
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </asp:View>

                            <asp:View ID="viewCcm" runat="server">
                                <div class="tab-header">Continuous Casting Machine (CCM)</div>
                                <div class="summary-cards">
                                    <div class="summary-card">
                                        <h3>Production</h3>
                                        <ul>
                                            <li>Heats: <asp:Literal ID="litCcmHeats" runat="server" /></li>
                                            <li>Total Weight: <asp:Literal ID="litCcmWeight" runat="server" /></li>
                                            <li>Average T/H: <asp:Literal ID="litCcmProdRate" runat="server" /></li>
                                        </ul>
                                    </div>
                                    <div class="summary-card">
                                        <h3>Billet Count</h3>
                                        <ul>
                                            <li>12 Meters: <asp:Literal ID="litCcm12m" runat="server" /></li>
                                            <li>8 Meters: <asp:Literal ID="litCcm8m" runat="server" /></li>
                                            <li>Short: <asp:Literal ID="litCcmShort" runat="server" /></li>
                                        </ul>
                                    </div>
                                    <div class="summary-card">
                                        <h3>Process Times</h3>
                                        <ul>
                                            <li>Total Duration: <asp:Literal ID="litCcmDuration" runat="server" /></li>
                                            <li>Ladle Metal: <asp:Literal ID="litCcmLadleMetal" runat="server" /></li>
                                            <li>Yield: <asp:Literal ID="litCcmYield" runat="server" /></li>
                                        </ul>
                                    </div>
                                </div>
                                <div class="chart-grid">
                                    <div class="chart-wrapper">
                                        <h4>Billet Length Distribution</h4>
                                        <asp:Chart ID="chartCcmBillets" runat="server" Height="260px" Width="380px">
                                            <ChartAreas>
                                                <asp:ChartArea Name="CcmBilletArea"></asp:ChartArea>
                                            </ChartAreas>
                                            <Legends>
                                                <asp:Legend Name="CcmBilletLegend"></asp:Legend>
                                            </Legends>
                                        </asp:Chart>
                                    </div>
                                    <div class="chart-wrapper">
                                        <h4>Shift Contribution</h4>
                                        <asp:Chart ID="chartCcmShift" runat="server" Height="260px" Width="380px">
                                            <ChartAreas>
                                                <asp:ChartArea Name="CcmShiftArea"></asp:ChartArea>
                                            </ChartAreas>
                                            <Legends>
                                                <asp:Legend Name="CcmShiftLegend"></asp:Legend>
                                            </Legends>
                                        </asp:Chart>
                                    </div>
                                    <div class="chart-wrapper">
                                        <h4>Billet Weight Trend</h4>
                                        <asp:Chart ID="chartCcmTrend" runat="server" Height="260px" Width="380px">
                                            <ChartAreas>
                                                <asp:ChartArea Name="CcmTrendArea"></asp:ChartArea>
                                            </ChartAreas>
                                            <Legends>
                                                <asp:Legend Name="CcmTrendLegend"></asp:Legend>
                                            </Legends>
                                        </asp:Chart>
                                    </div>
                                </div>
                                <div class="grid-wrapper">
                                    <h4>CCM Heat Detail</h4>
                                    <asp:GridView ID="gvCcmDetail" runat="server" AutoGenerateColumns="False" CssClass="detail-grid">
                                        <Columns>
                                            <asp:BoundField DataField="Heat" HeaderText="Heat" />
                                            <asp:BoundField DataField="Shift" HeaderText="Shift" />
                                            <asp:BoundField DataField="Start" HeaderText="Start" />
                                            <asp:BoundField DataField="End" HeaderText="End" />
                                            <asp:BoundField DataField="Duration" HeaderText="Duration (min)" />
                                            <asp:BoundField DataField="BilletWgt" HeaderText="Weight (t)" DataFormatString="{0:N2}" />
                                            <asp:BoundField DataField="ProdRate" HeaderText="t/hr" DataFormatString="{0:N2}" />
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </asp:View>

                            <asp:View ID="viewLrf" runat="server">
                                <div class="tab-header">Ladle Refining Furnace (LRF)</div>
                                <div class="summary-cards">
                                    <div class="summary-card">
                                        <h3>Additives</h3>
                                        <ul>
                                            <li>Lime: <asp:Literal ID="litLrfTotalLime" runat="server" /></li>
                                            <li>Alloy Additions: <asp:Literal ID="litLrfAlloyTotals" runat="server" /></li>
                                            <li>Coke: <asp:Literal ID="litLrfTotalCoke" runat="server" /></li>
                                        </ul>
                                    </div>
                                    <div class="summary-card">
                                        <h3>Energy &amp; Gas</h3>
                                        <ul>
                                            <li>Total Power (MWh): <asp:Literal ID="litLrfTotalPower" runat="server" /></li>
                                            <li>Power / t Liquid: <asp:Literal ID="litLrfPowerPerTon" runat="server" /></li>
                                            <li>N2 Nm<sup>3</sup>: <asp:Literal ID="litLrfTotalN2" runat="server" /></li>
                                        </ul>
                                    </div>
                                    <div class="summary-card">
                                        <h3>Process Times</h3>
                                        <ul>
                                            <li>Avg Stirring: <asp:Literal ID="litLrfAvgStir" runat="server" /></li>
                                            <li>EAF-LRF ΔT: <asp:Literal ID="litLrfDeltaT" runat="server" /></li>
                                        </ul>
                                    </div>
                                </div>
                                <div class="chart-grid">
                                    <div class="chart-wrapper">
                                        <h4>LRF Additive Consumption</h4>
                                        <asp:Chart ID="chartLrfAdditives" runat="server" Height="260px" Width="380px">
                                            <ChartAreas>
                                                <asp:ChartArea Name="LrfAdditiveArea"></asp:ChartArea>
                                            </ChartAreas>
                                            <Legends>
                                                <asp:Legend Name="LrfAdditiveLegend"></asp:Legend>
                                            </Legends>
                                        </asp:Chart>
                                    </div>
                                    <div class="chart-wrapper">
                                        <h4>Alloy Distribution</h4>
                                        <asp:Chart ID="chartLrfDistribution" runat="server" Height="260px" Width="380px">
                                            <ChartAreas>
                                                <asp:ChartArea Name="LrfDistributionArea"></asp:ChartArea>
                                            </ChartAreas>
                                            <Legends>
                                                <asp:Legend Name="LrfDistributionLegend"></asp:Legend>
                                            </Legends>
                                        </asp:Chart>
                                    </div>
                                    <div class="chart-wrapper">
                                        <h4>LRF Energy Trend</h4>
                                        <asp:Chart ID="chartLrfTrend" runat="server" Height="260px" Width="380px">
                                            <ChartAreas>
                                                <asp:ChartArea Name="LrfTrendArea"></asp:ChartArea>
                                            </ChartAreas>
                                            <Legends>
                                                <asp:Legend Name="LrfTrendLegend"></asp:Legend>
                                            </Legends>
                                        </asp:Chart>
                                    </div>
                                </div>
                                <div class="grid-wrapper">
                                    <h4>LRF Heat Detail</h4>
                                    <asp:GridView ID="gvLrfDetail" runat="server" AutoGenerateColumns="False" CssClass="detail-grid">
                                        <Columns>
                                            <asp:BoundField DataField="Heat" HeaderText="Heat" />
                                            <asp:BoundField DataField="Ladle" HeaderText="Ladle" />
                                            <asp:BoundField DataField="Stir" HeaderText="Stir (min)" />
                                            <asp:BoundField DataField="Power" HeaderText="Power (MWh)" DataFormatString="{0:N2}" />
                                            <asp:BoundField DataField="Lime" HeaderText="Lime (kg)" DataFormatString="{0:N0}" />
                                            <asp:BoundField DataField="Alloys" HeaderText="Alloys (kg)" DataFormatString="{0:N0}" />
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </asp:View>
                        </asp:MultiView>
                    </div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
