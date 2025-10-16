Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Globalization
Imports System.Web.UI.DataVisualization.Charting

Namespace LEVEL2WEB
    Partial Class MS_Report_Advanced
        Inherits System.Web.UI.Page

        Private Const ChartPalette As ChartColorPalette = ChartColorPalette.BrightPastel

        Private Class FilterContext
            Public Property FromDate As DateTime
            Public Property ToDate As DateTime
            Public Property ShiftCode As String
            Public Property ShiftLabel As String
            Public Property PeriodDescription As String
        End Class

        Private Class HeatRange
            Public Property StartHeat As Integer?
            Public Property EndHeat As Integer?
            Public Property HeatCount As Integer
        End Class

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
            If Not IsPostBack Then
                If MS_Orcl_Con Is Nothing Then
                    MS_Connect_me()
                End If
                InitializeFilters()
                LoadAllData()
            End If
        End Sub

        Private Sub InitializeFilters()
            Dim defaultDate As DateTime = Date.Today.AddDays(-1)
            txtFromDate.Text = defaultDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture)
            txtToDate.Text = defaultDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture)
            rblPeriod.SelectedValue = "Day"
            rblShift.SelectedValue = "All"
            lblFilterInfo.Text = "Defaulting to the previous production day."
            ms_rate(defaultDate)
        End Sub

        Protected Sub btnApply_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnApply.Click
            LoadAllData()
        End Sub

        Protected Sub menuAreas_MenuItemClick(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.MenuEventArgs) Handles menuAreas.MenuItemClick
            Dim targetIndex As Integer
            If Integer.TryParse(e.Item.Value, targetIndex) Then
                mvAreas.ActiveViewIndex = targetIndex
            End If
        End Sub

        Private Sub LoadAllData()
            lblStatus.Text = String.Empty

            If MS_Orcl_Con Is Nothing Then
                MS_Connect_me()
            End If

            Dim context As FilterContext = ResolveFilter()
            If context Is Nothing Then
                ClearVisuals()
                Return
            End If

            ms_rate(context.FromDate)

            Dim range As HeatRange = GetHeatRange(context)
            If range Is Nothing OrElse Not range.StartHeat.HasValue OrElse Not range.EndHeat.HasValue Then
                lblStatus.Text = "No CCM heats were found for the selected filter."
                litHeatRange.Text = "N/A"
                litPeriodDescription.Text = context.PeriodDescription
                litShiftDescription.Text = context.ShiftLabel
                ClearVisuals()
                Return
            End If

            litHeatRange.Text = String.Format(CultureInfo.InvariantCulture, "{0} - {1} ({2} heats)", range.StartHeat.Value, range.EndHeat.Value, range.HeatCount)
            litPeriodDescription.Text = context.PeriodDescription
            litShiftDescription.Text = context.ShiftLabel

            Dim eafSummary As DataTable = LoadEafSummary(range, context)
            Dim lrfSummary As DataTable = LoadLrfSummary(range, context)
            Dim ccmSummary As DataTable = LoadCcmSummary(range, context)
            Dim detailTable As DataTable = LoadCombinedDetail(range, context)

            BindSummary(context, range, eafSummary, lrfSummary, ccmSummary)
            BindEafSection(detailTable, eafSummary)
            BindCcmSection(detailTable, ccmSummary, context)
            BindLrfSection(detailTable, lrfSummary)
        End Sub

        Private Function ResolveFilter() As FilterContext
            Dim context As New FilterContext()

            Dim fromDate As DateTime
            Dim toDate As DateTime

            If Not DateTime.TryParseExact(txtFromDate.Text, "MM/dd/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, fromDate) Then
                lblStatus.Text = "Invalid 'From' date. Please select a valid date."
                Return Nothing
            End If

            If Not DateTime.TryParseExact(txtToDate.Text, "MM/dd/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, toDate) Then
                toDate = fromDate
            End If

            Select Case rblPeriod.SelectedValue
                Case "Day"
                    context.FromDate = fromDate.Date
                    context.ToDate = fromDate.Date.AddDays(1).AddSeconds(-1)
                    context.PeriodDescription = fromDate.ToString("dd MMM yyyy", CultureInfo.InvariantCulture)
                Case "Month"
                    Dim startOfMonth As New DateTime(fromDate.Year, fromDate.Month, 1)
                    context.FromDate = startOfMonth
                    context.ToDate = startOfMonth.AddMonths(1).AddSeconds(-1)
                    context.PeriodDescription = startOfMonth.ToString("MMMM yyyy", CultureInfo.InvariantCulture)
                Case "Quarter"
                    Dim quarterStartMonth As Integer = ((fromDate.Month - 1) \ 3) * 3 + 1
                    Dim startOfQuarter As New DateTime(fromDate.Year, quarterStartMonth, 1)
                    context.FromDate = startOfQuarter
                    context.ToDate = startOfQuarter.AddMonths(3).AddSeconds(-1)
                    Dim quarterLabel As Integer = ((fromDate.Month - 1) \ 3) + 1
                    context.PeriodDescription = String.Format(CultureInfo.InvariantCulture, "Q{0} {1}", quarterLabel, fromDate.Year)
                Case "Year"
                    Dim startOfYear As New DateTime(fromDate.Year, 1, 1)
                    context.FromDate = startOfYear
                    context.ToDate = startOfYear.AddYears(1).AddSeconds(-1)
                    context.PeriodDescription = fromDate.Year.ToString(CultureInfo.InvariantCulture)
                Case "Range"
                    If fromDate > toDate Then
                        lblStatus.Text = "The 'From' date must be earlier than or equal to the 'To' date."
                        Return Nothing
                    End If
                    context.FromDate = fromDate.Date
                    context.ToDate = toDate.Date.AddDays(1).AddSeconds(-1)
                    context.PeriodDescription = String.Format(CultureInfo.InvariantCulture, "{0} - {1}", fromDate.ToString("dd MMM yyyy", CultureInfo.InvariantCulture), toDate.ToString("dd MMM yyyy", CultureInfo.InvariantCulture))
                Case Else
                    context.FromDate = fromDate.Date
                    context.ToDate = fromDate.Date.AddDays(1).AddSeconds(-1)
                    context.PeriodDescription = fromDate.ToString("dd MMM yyyy", CultureInfo.InvariantCulture)
            End Select

            Select Case rblShift.SelectedValue
                Case "Shift1"
                    context.ShiftCode = "Shift A"
                    context.ShiftLabel = "Shift 1 (00:00 - 08:00)"
                Case "Shift2"
                    context.ShiftCode = "Shift B"
                    context.ShiftLabel = "Shift 2 (08:00 - 16:00)"
                Case "Shift3"
                    context.ShiftCode = "Shift C"
                    context.ShiftLabel = "Shift 3 (16:00 - 00:00)"
                Case Else
                    context.ShiftCode = String.Empty
                    context.ShiftLabel = "All shifts"
            End Select

            Return context
        End Function

        Private Function GetHeatRange(ByVal context As FilterContext) As HeatRange
            Dim startStr As String = context.FromDate.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture)
            Dim endStr As String = context.ToDate.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture)

            Dim query As String = "SELECT MIN(heat_num) AS min_heat, MAX(heat_num) AS max_heat, COUNT(heat_num) AS heat_count " &
                                  "FROM CCM_FINISHED_HEAT " &
                                  "WHERE heat_tap_finish >= TO_DATE('" & startStr & "', 'mm/dd/yyyy hh24:mi:ss') " &
                                  "AND heat_tap_finish <= TO_DATE('" & endStr & "', 'mm/dd/yyyy hh24:mi:ss')"

            If Not String.IsNullOrEmpty(context.ShiftCode) Then
                query &= " AND SHIFT = '" & context.ShiftCode & "'"
            End If

            Dim ds As New DataSet()
            Using da As New OleDbDataAdapter(query, MS_Orcl_Con)
                da.Fill(ds, "bounds")
            End Using

            If ds.Tables("bounds").Rows.Count = 0 Then
                Return Nothing
            End If

            Dim row As DataRow = ds.Tables("bounds").Rows(0)
            If row.IsNull("min_heat") OrElse row.IsNull("max_heat") Then
                Return Nothing
            End If

            Dim result As New HeatRange()
            result.StartHeat = Convert.ToInt32(row("min_heat"), CultureInfo.InvariantCulture)
            result.EndHeat = Convert.ToInt32(row("max_heat"), CultureInfo.InvariantCulture)
            result.HeatCount = Convert.ToInt32(row("heat_count"), CultureInfo.InvariantCulture)
            Return result
        End Function

        Private Function LoadEafSummary(ByVal range As HeatRange, ByVal context As FilterContext) As DataTable
            Dim query As String = "SELECT Sum(chg_wgt) AS tot_scrap, Sum(DRI_WGT) AS Tot_dri, Sum(LIME) AS tot_lime, NVL(SUM(DRI_CONT_FEED), 0) AS Tot_dri_Feed, " &
                                  "Sum(Coke) AS tot_coke, Sum(Coal) AS Tot_Coal, Sum(Si_mn) AS tot_Si_mn, Sum(FE_MN) AS tot_Fe_mn, Sum(FE_si) AS tot_Fe_si, " &
                                  "Sum(Kwh) AS Tot_kwh, Sum(Mod_o2) AS TOT_o2, Sum(Mod_ch4) AS TOT_ch4, Sum(ACT_Liq_WGT) AS Tot_Liq, " &
                                  "AVG(Power_on_time) AS AVG_Power_on_time, AVG(Heat_duration) AS AVR_DUR, " &
                                  "NVL(SUM(hbi_tot_cons), 0) AS hbi_sum, NVL(SUM(pig_iron_cons), 0) AS pig_iron_sum " &
                                  "FROM EAF_FINISHED_HEAT_SUM " &
                                  "WHERE HEAT_NUM >= " & range.StartHeat.Value & " AND HEAT_NUM <= " & range.EndHeat.Value

            If Not String.IsNullOrEmpty(context.ShiftCode) Then
                query &= " AND SHIFT = '" & context.ShiftCode & "'"
            End If

            Dim ds As New DataSet()
            Using da As New OleDbDataAdapter(query, MS_Orcl_Con)
                da.Fill(ds, "EAF_SUM")
            End Using
            Return ds.Tables("EAF_SUM")
        End Function

        Private Function LoadLrfSummary(ByVal range As HeatRange, ByVal context As FilterContext) As DataTable
            Dim query As String = "SELECT Sum(LRF_LIME) AS tot_lime, Sum(LRF_COKE) AS tot_coke, Sum(LRF_Si_mn) AS tot_Si_mn, " &
                                  "Sum(LRF_FE_MN) AS tot_Fe_mn, Sum(LRF_FE_si) AS tot_Fe_si, Sum(LRF_Kwh) AS Tot_kwh, " &
                                  "Sum(N2_nm3) AS TOT_N2, AVG(Stirring_time) AS AVG_stirring_time, AVG(LRF_START_TMP) AS AVG_START_TMP " &
                                  "FROM LRF_FINISHED_HEAT " &
                                  "WHERE HEAT_NUM >= " & range.StartHeat.Value & " AND HEAT_NUM <= " & range.EndHeat.Value

            If Not String.IsNullOrEmpty(context.ShiftCode) Then
                query &= " AND SHIFT = '" & context.ShiftCode & "'"
            End If

            Dim ds As New DataSet()
            Using da As New OleDbDataAdapter(query, MS_Orcl_Con)
                da.Fill(ds, "LRF_SUM")
            End Using
            Return ds.Tables("LRF_SUM")
        End Function

        Private Function LoadCcmSummary(ByVal range As HeatRange, ByVal context As FilterContext) As DataTable
            Dim query As String = "SELECT Sum(heat_duration) AS tot_Dur, Sum(Ladle_Start_wgt) AS tot_Ldl_Wgt, " &
                                  "Sum(Act_No_Billets_12m) AS tot_12m, Sum(Act_No_Billets_8m) AS tot_8m, NVL(Sum(Act_No_Billets_short), 0) AS tot_billet_short, " &
                                  "Count(heat_num) AS count_heat_no " &
                                  "FROM CCM_FINISHED_HEAT " &
                                  "WHERE BACK_CHARGED = 0 " &
                                  "AND HEAT_NUM >= " & range.StartHeat.Value & " AND HEAT_NUM <= " & range.EndHeat.Value

            If Not String.IsNullOrEmpty(context.ShiftCode) Then
                query &= " AND SHIFT = '" & context.ShiftCode & "'"
            End If

            Dim ds As New DataSet()
            Using da As New OleDbDataAdapter(query, MS_Orcl_Con)
                da.Fill(ds, "CCM_SUM")
            End Using
            Return ds.Tables("CCM_SUM")
        End Function

        Private Function LoadCombinedDetail(ByVal range As HeatRange, ByVal context As FilterContext) As DataTable
            Dim query As String = "SELECT E.HEAT_NUM, E.SHIFT AS EAF_SHIFT, E.HEAT_START_TIME, NVL(E.ACT_LIQ_WGT, 0) AS ACT_LIQ_WGT, NVL(E.KWH, 0) AS EAF_KWH, " &
                                  "NVL(E.POWER_ON_TIME, 0) AS POWER_ON_TIME, NVL(E.AVR_KWH, 0) AS AVR_KWH, NVL(E.MOD_O2, 0) AS MOD_O2, NVL(E.MOD_CH4, 0) AS MOD_CH4, " &
                                  "NVL(E.CHG_WGT, 0) AS CHG_WGT, NVL(E.DRI_WGT, 0) AS DRI_WGT, NVL(E.DRI_CONT_FEED, 0) AS DRI_FEED, " &
                                  "NVL(E.HBI_TOT_CONS, 0) AS HBI_CONS, NVL(E.PIG_IRON_CONS, 0) AS PIG_CONS, " &
                                  "NVL(L.LRF_KWH, 0) AS LRF_KWH, NVL(L.LRF_LIME, 0) AS LRF_LIME, NVL(L.LRF_COKE, 0) AS LRF_COKE, " &
                                  "NVL(L.LRF_SI_MN, 0) AS LRF_SI_MN, NVL(L.LRF_FE_MN, 0) AS LRF_FE_MN, NVL(L.LRF_FE_SI, 0) AS LRF_FE_SI, " &
                                  "NVL(L.N2_NM3, 0) AS LRF_N2, NVL(L.STIRRING_TIME, 0) AS LRF_STIR_TIME, NVL(L.LRF_START_TMP, 0) AS LRF_START_TMP, " &
                                  "C.SHIFT AS CCM_SHIFT, C.HEAT_TAP_START, C.HEAT_TAP_FINISH, NVL(C.HEAT_DURATION, 0) AS HEAT_DURATION, " &
                                  "NVL(C.LADLE_START_WGT, 0) AS LADLE_START_WGT, NVL(C.ACT_NO_BILLETS_12M, 0) AS BILLET_12M, " &
                                  "NVL(C.ACT_NO_BILLETS_8M, 0) AS BILLET_8M, NVL(C.ACT_NO_BILLETS_SHORT, 0) AS BILLET_SHORT, " &
                                  "NVL(C.NO_STRANDS, 0) AS NO_STRANDS, NVL(C.LDL_NUM, 0) AS LADLE_NUM, NVL(C.TND_NUM, 0) AS TND_NUM, NVL(C.BACK_CHARGED, 0) AS BACK_CHARGED, " &
                                  "G.GRADE_NAME " &
                                  "FROM EAF_FINISHED_HEAT_SUM E " &
                                  "JOIN MS_PROD_ORDER O ON E.PROD_ORDER_NUM = O.PROD_ORDER_NUM " &
                                  "JOIN IC_ITEM_MST_LUT I ON O.ITEM_ID = I.ITEM_ID " &
                                  "JOIN MS_STEEL_GRADE G ON I.GRADE_ID = G.GRADE_ID " &
                                  "LEFT JOIN LRF_FINISHED_HEAT L ON E.HEAT_NUM = L.HEAT_NUM " &
                                  "LEFT JOIN CCM_FINISHED_HEAT C ON E.HEAT_NUM = C.HEAT_NUM " &
                                  "WHERE E.HEAT_NUM >= " & range.StartHeat.Value & " AND E.HEAT_NUM <= " & range.EndHeat.Value

            If Not String.IsNullOrEmpty(context.ShiftCode) Then
                query &= " AND E.SHIFT = '" & context.ShiftCode & "'"
            End If

            query &= " ORDER BY E.HEAT_NUM"

            Dim ds As New DataSet()
            Using da As New OleDbDataAdapter(query, MS_Orcl_Con)
                da.Fill(ds, "DETAIL")
            End Using
            Return ds.Tables("DETAIL")
        End Function
        Private Sub BindSummary(ByVal context As FilterContext, ByVal range As HeatRange, ByVal eafSummary As DataTable, ByVal lrfSummary As DataTable, ByVal ccmSummary As DataTable)
            If eafSummary.Rows.Count = 0 OrElse ccmSummary.Rows.Count = 0 OrElse lrfSummary.Rows.Count = 0 Then
                lblStatus.Text = "Insufficient data to build summary for the selected range."
                ClearVisuals()
                Return
            End If

            Dim eafRow As DataRow = eafSummary.Rows(0)
            Dim ccmRow As DataRow = ccmSummary.Rows(0)
            Dim lrfRow As DataRow = lrfSummary.Rows(0)

            Dim totScrap As Double = GetDoubleValue(eafRow, "tot_scrap")
            Dim totDri As Double = GetDoubleValue(eafRow, "Tot_dri")
            Dim totDriFeed As Double = GetDoubleValue(eafRow, "Tot_dri_Feed")
            Dim hbiTot As Double = GetDoubleValue(eafRow, "hbi_sum")
            Dim pigTotKg As Double = GetDoubleValue(eafRow, "pig_iron_sum")

            Dim totEafChargeKg As Double = (totScrap + totDri + totDriFeed) * 1000.0
            totEafChargeKg += hbiTot * 1000.0
            totEafChargeKg += pigTotKg
            Dim totEafChargeTon As Double = If(totEafChargeKg > 0, totEafChargeKg / 1000.0, 0.0)

            Dim tot12m As Double = GetDoubleValue(ccmRow, "tot_12m")
            Dim tot8m As Double = GetDoubleValue(ccmRow, "tot_8m")
            Dim totShort As Double = GetDoubleValue(ccmRow, "tot_billet_short")
            Dim totHeats As Integer = Convert.ToInt32(GetDoubleValue(ccmRow, "count_heat_no"), CultureInfo.InvariantCulture)

            Dim ccmWeight As Double = CDbl(tot8m * bill_8m_rate + tot12m * bill_12m_rate + totShort * bill_short_rate)
            Dim eafLiquid As Double = If(ccmWeight > 0, ccmWeight / 0.985, 0.0)

            litSummaryEafScrap.Text = FormatTons(totScrap)
            litSummaryEafDri.Text = FormatTons(totDri + totDriFeed)
            litSummaryEafLiquid.Text = FormatTons(eafLiquid)
            litSummaryEafYield.Text = FormatPercent(If(totEafChargeTon > 0, (eafLiquid * 100.0) / totEafChargeTon, 0.0))

            litSummaryCcm12m.Text = FormatNumber(tot12m, "###,##0")
            litSummaryCcm8m.Text = FormatNumber(tot8m, "###,##0")
            litSummaryCcmShort.Text = FormatNumber(totShort, "###,##0")
            litSummaryCcmWeight.Text = FormatTons(ccmWeight)
            litSummaryCcmHeats.Text = totHeats.ToString(CultureInfo.InvariantCulture)

            Dim ccmYield As Double = If(eafLiquid > 0, (ccmWeight * 100.0) / eafLiquid, 0.0)
            litSummaryCcmYield.Text = FormatPercent(ccmYield)
            litSummaryPlantYield.Text = FormatPercent(If(totEafChargeTon > 0, (ccmWeight * 100.0) / totEafChargeTon, 0.0))

            Dim lrfLime As Double = GetDoubleValue(lrfRow, "tot_lime")
            Dim lrfSiMn As Double = GetDoubleValue(lrfRow, "tot_Si_mn")
            Dim lrfFeMn As Double = GetDoubleValue(lrfRow, "tot_Fe_mn")
            Dim lrfFeSi As Double = GetDoubleValue(lrfRow, "tot_Fe_si")
            Dim lrfPowerMWh As Double = GetDoubleValue(lrfRow, "Tot_kwh") / 1000.0
            Dim lrfStir As Double = GetDoubleValue(lrfRow, "AVG_stirring_time")

            litSummaryLrfLime.Text = FormatNumber(lrfLime, "###,##0")
            litSummaryLrfAlloys.Text = FormatNumber(lrfSiMn + lrfFeMn + lrfFeSi, "###,##0")
            litSummaryLrfPower.Text = FormatNumber(lrfPowerMWh, "###,##0.00")
            litSummaryLrfStir.Text = FormatNumber(lrfStir, "###,##0.0")

            PrepareSummaryCharts(totEafChargeTon, eafLiquid, ccmWeight, GetDoubleValue(eafRow, "Tot_kwh") / 1000.0, lrfPowerMWh)
        End Sub

        Private Sub PrepareSummaryCharts(ByVal eafChargeTons As Double, ByVal eafLiquidTons As Double, ByVal ccmWeightTons As Double, ByVal eafPowerMWh As Double, ByVal lrfPowerMWh As Double)
            chartSummaryProduction.Series.Clear()
            chartSummaryProduction.ChartAreas("ProductionArea").AxisX.Interval = 1
            chartSummaryProduction.ChartAreas("ProductionArea").AxisY.Title = "Tons"
            chartSummaryProduction.Palette = ChartPalette

            Dim prodSeries As Series = chartSummaryProduction.Series.Add("Production")
            prodSeries.ChartType = SeriesChartType.Column
            prodSeries.IsValueShownAsLabel = True
            prodSeries.Points.AddXY("EAF Charge", eafChargeTons)
            prodSeries.Points.AddXY("Liquid Steel", eafLiquidTons)
            prodSeries.Points.AddXY("CCM Billets", ccmWeightTons)

            chartSummaryEnergy.Series.Clear()
            chartSummaryEnergy.Legends("Legend2").Docking = Docking.Bottom
            chartSummaryEnergy.Palette = ChartPalette

            Dim energySeries As Series = chartSummaryEnergy.Series.Add("Energy")
            energySeries.ChartType = SeriesChartType.Pie
            energySeries.Points.AddXY("EAF Power (MWh)", eafPowerMWh)
            energySeries.Points.AddXY("LRF Power (MWh)", lrfPowerMWh)
        End Sub
        Private Sub BindEafSection(ByVal detailTable As DataTable, ByVal eafSummary As DataTable)
            chartEafMaterials.Series.Clear()
            chartEafEnergy.Series.Clear()
            chartEafTrend.Series.Clear()

            If eafSummary.Rows.Count = 0 Then
                HideEafControls()
                Return
            End If

            Dim eafRow As DataRow = eafSummary.Rows(0)
            Dim totScrap As Double = GetDoubleValue(eafRow, "tot_scrap")
            Dim totDri As Double = GetDoubleValue(eafRow, "Tot_dri")
            Dim totDriFeed As Double = GetDoubleValue(eafRow, "Tot_dri_Feed")
            Dim hbiTot As Double = GetDoubleValue(eafRow, "hbi_sum")
            Dim pigTotKg As Double = GetDoubleValue(eafRow, "pig_iron_sum")
            Dim totKwhMWh As Double = GetDoubleValue(eafRow, "Tot_kwh") / 1000.0
            Dim totLiq As Double = GetDoubleValue(eafRow, "Tot_Liq")
            Dim totO2 As Double = GetDoubleValue(eafRow, "TOT_o2")
            Dim totCh4 As Double = GetDoubleValue(eafRow, "TOT_ch4")
            Dim avgTap As Double = GetDoubleValue(eafRow, "AVR_DUR")
            Dim avgPowerOn As Double = GetDoubleValue(eafRow, "AVG_Power_on_time")

            litEafTotalScrap.Text = FormatTons(totScrap)
            litEafTotalDri.Text = FormatTons(totDri + totDriFeed)
            litEafHbiPig.Text = String.Format(CultureInfo.InvariantCulture, "{0} t / {1} t", FormatTons(hbiTot), FormatTons(pigTotKg / 1000.0))
            litEafTotalKwh.Text = FormatNumber(totKwhMWh, "###,##0.00")

            Dim kwhPerTon As Double = If(totLiq > 0, (GetDoubleValue(eafRow, "Tot_kwh") / totLiq), 0.0)
            litEafKwhPerTon.Text = FormatNumber(kwhPerTon, "###,##0.0")
            litEafOxygenPerTon.Text = FormatNumber(If(totLiq > 0, totO2 / totLiq, 0.0), "###,##0.0")
            litEafAvgTapTime.Text = FormatNumber(avgTap, "###,##0.0")
            litEafAvgPowerOn.Text = FormatNumber(avgPowerOn, "###,##0.0")
            litEafDriRatio.Text = FormatPercent(If((totScrap + totDri + totDriFeed) > 0, (totDri + totDriFeed) * 100.0 / (totScrap + totDri + totDriFeed), 0.0))

            PrepareChart(chartEafMaterials, "MaterialsArea", ChartPalette)
            Dim materialsSeries As Series = chartEafMaterials.Series.Add("Materials")
            materialsSeries.ChartType = SeriesChartType.Column
            materialsSeries.IsValueShownAsLabel = True
            materialsSeries.Points.AddXY("Scrap", totScrap)
            materialsSeries.Points.AddXY("DRI (Pellet)", totDri)
            materialsSeries.Points.AddXY("DRI Feed", totDriFeed)
            materialsSeries.Points.AddXY("HBI", hbiTot)
            materialsSeries.Points.AddXY("Pig Iron", pigTotKg / 1000.0)

            PrepareChart(chartEafEnergy, "EafEnergyArea", ChartPalette)
            Dim energySeries As Series = chartEafEnergy.Series.Add("Energy")
            energySeries.ChartType = SeriesChartType.Pie
            energySeries.Points.AddXY("Power (MWh)", totKwhMWh)
            energySeries.Points.AddXY("Oxygen (Nm3)", totO2)
            energySeries.Points.AddXY("Methane (Nm3)", totCh4)

            PrepareChart(chartEafTrend, "EafTrendArea", ChartPalette)
            chartEafTrend.ChartAreas("EafTrendArea").AxisX.LabelStyle.Format = "MM-dd HH:mm"
            chartEafTrend.ChartAreas("EafTrendArea").AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount
            chartEafTrend.ChartAreas("EafTrendArea").AxisY.Title = "kWh / ton"

            Dim trendSeries As Series = chartEafTrend.Series.Add("KWh per ton")
            trendSeries.ChartType = SeriesChartType.Line
            trendSeries.XValueType = ChartValueType.DateTime
            trendSeries.MarkerStyle = MarkerStyle.Circle
            trendSeries.MarkerSize = 5

            Dim detailView As DataRow() = detailTable.Select(String.Empty, "HEAT_START_TIME ASC")
            Dim tableForGrid As DataTable = BuildEafDetailTable()

            For Each row As DataRow In detailView
                Dim heatNum As Integer = Convert.ToInt32(row("HEAT_NUM"), CultureInfo.InvariantCulture)
                Dim startTime As DateTime = If(row.IsNull("HEAT_START_TIME"), DateTime.MinValue, Convert.ToDateTime(row("HEAT_START_TIME"), CultureInfo.InvariantCulture))
                Dim liq As Double = GetDoubleValue(row, "ACT_LIQ_WGT")
                Dim kwh As Double = GetDoubleValue(row, "EAF_KWH")
                Dim kwhPerTonPoint As Double = If(liq > 0, kwh / liq, 0.0)

                If startTime > DateTime.MinValue Then
                    trendSeries.Points.AddXY(startTime, kwhPerTonPoint)
                End If

                Dim newRow As DataRow = tableForGrid.NewRow()
                newRow("Heat") = heatNum
                newRow("Shift") = FormatShiftName(row("EAF_SHIFT").ToString())
                newRow("StartTime") = If(startTime = DateTime.MinValue, String.Empty, startTime.ToString("MM-dd HH:mm", CultureInfo.InvariantCulture))
                newRow("Grade") = row("GRADE_NAME").ToString()
                newRow("LiquidSteel") = liq
                newRow("Kwh") = kwh / 1000.0
                newRow("KwhPerTon") = kwhPerTonPoint
                tableForGrid.Rows.Add(newRow)
            Next

            gvEafDetail.DataSource = tableForGrid
            gvEafDetail.DataBind()
        End Sub
        Private Sub BindCcmSection(ByVal detailTable As DataTable, ByVal ccmSummary As DataTable, ByVal context As FilterContext)
            chartCcmBillets.Series.Clear()
            chartCcmShift.Series.Clear()
            chartCcmTrend.Series.Clear()

            If ccmSummary.Rows.Count = 0 Then
                HideCcmControls()
                Return
            End If

            Dim ccmRow As DataRow = ccmSummary.Rows(0)
            Dim tot12m As Double = GetDoubleValue(ccmRow, "tot_12m")
            Dim tot8m As Double = GetDoubleValue(ccmRow, "tot_8m")
            Dim totShort As Double = GetDoubleValue(ccmRow, "tot_billet_short")
            Dim totDurationMin As Double = GetDoubleValue(ccmRow, "tot_Dur")
            Dim totHeats As Integer = Convert.ToInt32(GetDoubleValue(ccmRow, "count_heat_no"), CultureInfo.InvariantCulture)

            Dim totalWeight As Double = CDbl(tot8m * bill_8m_rate + tot12m * bill_12m_rate + totShort * bill_short_rate)
            Dim ladleMetal As Double = GetDoubleValue(ccmRow, "tot_Ldl_Wgt")
            Dim avgProdRate As Double = If(totDurationMin > 0, totalWeight / (totDurationMin / 60.0), 0.0)

            litCcmHeats.Text = totHeats.ToString(CultureInfo.InvariantCulture)
            litCcm12m.Text = FormatNumber(tot12m, "###,##0")
            litCcm8m.Text = FormatNumber(tot8m, "###,##0")
            litCcmShort.Text = FormatNumber(totShort, "###,##0")
            litCcmWeight.Text = FormatTons(totalWeight)
            litCcmProdRate.Text = FormatNumber(avgProdRate, "###,##0.00")
            litCcmDuration.Text = FormatNumber(totDurationMin, "###,##0")
            litCcmLadleMetal.Text = FormatTons(ladleMetal)

            Dim eafChargeForYield As Double = ComputeEafChargeFromDetail(detailTable)
            Dim ccmYield As Double = If(totalWeight > 0 AndAlso eafChargeForYield > 0, totalWeight * 100.0 / eafChargeForYield, 0.0)
            litCcmYield.Text = FormatPercent(ccmYield)

            PrepareChart(chartCcmBillets, "CcmBilletArea", ChartPalette)
            Dim billetSeries As Series = chartCcmBillets.Series.Add("Billets")
            billetSeries.ChartType = SeriesChartType.Column
            billetSeries.IsValueShownAsLabel = True
            billetSeries.Points.AddXY("12 m", tot12m)
            billetSeries.Points.AddXY("8 m", tot8m)
            billetSeries.Points.AddXY("Short", totShort)

            PrepareChart(chartCcmShift, "CcmShiftArea", ChartPalette)
            Dim shiftSeries As Series = chartCcmShift.Series.Add("ShiftShare")
            shiftSeries.ChartType = SeriesChartType.Pie

            Dim shiftCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            For Each row As DataRow In detailTable.Rows
                Dim shiftName As String = row("CCM_SHIFT").ToString()
                If Not String.IsNullOrEmpty(shiftName) Then
                    Dim simplified As String = FormatShiftName(shiftName)
                    If shiftCounts.ContainsKey(simplified) Then
                        shiftCounts(simplified) += 1
                    Else
                        shiftCounts(simplified) = 1
                    End If
                End If
            Next

            If shiftCounts.Count = 0 AndAlso totHeats > 0 Then
                shiftSeries.Points.AddXY(context.ShiftLabel, totHeats)
            Else
                For Each entry In shiftCounts
                    shiftSeries.Points.AddXY(entry.Key, entry.Value)
                Next
            End If

            PrepareChart(chartCcmTrend, "CcmTrendArea", ChartPalette)
            chartCcmTrend.ChartAreas("CcmTrendArea").AxisX.Interval = 1
            chartCcmTrend.ChartAreas("CcmTrendArea").AxisY.Title = "Billet weight (t)"

            Dim trendSeries As Series = chartCcmTrend.Series.Add("Billet Weight")
            trendSeries.ChartType = SeriesChartType.Line
            trendSeries.MarkerStyle = MarkerStyle.Circle
            trendSeries.MarkerSize = 5

            Dim gridTable As DataTable = BuildCcmDetailTable()

            For Each row As DataRow In detailTable.Rows
                Dim heatNum As Integer = Convert.ToInt32(row("HEAT_NUM"), CultureInfo.InvariantCulture)
                Dim durationMin As Double = GetDoubleValue(row, "HEAT_DURATION")
                Dim billetWeight As Double = (GetDoubleValue(row, "BILLET_12M") * bill_12m_rate) +
                                             (GetDoubleValue(row, "BILLET_8M") * bill_8m_rate) +
                                             (GetDoubleValue(row, "BILLET_SHORT") * bill_short_rate)

                trendSeries.Points.AddXY(heatNum, billetWeight)

                Dim newRow As DataRow = gridTable.NewRow()
                newRow("Heat") = heatNum
                newRow("Shift") = FormatShiftName(row("CCM_SHIFT").ToString())
                newRow("Start") = FormatTime(row, "HEAT_TAP_START")
                newRow("End") = FormatTime(row, "HEAT_TAP_FINISH")
                newRow("Duration") = FormatNumber(durationMin, "###,##0")
                newRow("BilletWgt") = billetWeight
                newRow("ProdRate") = If(durationMin > 0, billetWeight / (durationMin / 60.0), 0.0)
                gridTable.Rows.Add(newRow)
            Next

            gvCcmDetail.DataSource = gridTable
            gvCcmDetail.DataBind()
        End Sub

        Private Sub BindLrfSection(ByVal detailTable As DataTable, ByVal lrfSummary As DataTable)
            chartLrfAdditives.Series.Clear()
            chartLrfDistribution.Series.Clear()
            chartLrfTrend.Series.Clear()

            If lrfSummary.Rows.Count = 0 Then
                HideLrfControls()
                Return
            End If

            Dim lrfRow As DataRow = lrfSummary.Rows(0)
            Dim lime As Double = GetDoubleValue(lrfRow, "tot_lime")
            Dim coke As Double = GetDoubleValue(lrfRow, "tot_coke")
            Dim siMn As Double = GetDoubleValue(lrfRow, "tot_Si_mn")
            Dim feMn As Double = GetDoubleValue(lrfRow, "tot_Fe_mn")
            Dim feSi As Double = GetDoubleValue(lrfRow, "tot_Fe_si")
            Dim n2 As Double = GetDoubleValue(lrfRow, "TOT_N2")
            Dim totalPowerMWh As Double = GetDoubleValue(lrfRow, "Tot_kwh") / 1000.0
            Dim avgStir As Double = GetDoubleValue(lrfRow, "AVG_stirring_time")
            Dim avgStartTmp As Double = GetDoubleValue(lrfRow, "AVG_START_TMP")

            Dim totalLiquid As Double = ComputeTotalLiquid(detailTable)
            Dim powerPerTon As Double = If(totalLiquid > 0, GetDoubleValue(lrfRow, "Tot_kwh") / totalLiquid, 0.0)

            litLrfTotalLime.Text = FormatNumber(lime, "###,##0")
            litLrfTotalCoke.Text = FormatNumber(coke, "###,##0")
            litLrfTotalN2.Text = FormatNumber(n2, "###,##0")
            litLrfAvgStir.Text = FormatNumber(avgStir, "###,##0.0")
            litLrfTotalPower.Text = FormatNumber(totalPowerMWh, "###,##0.00")
            litLrfPowerPerTon.Text = FormatNumber(powerPerTon, "###,##0.0")
            litLrfAlloyTotals.Text = FormatNumber(siMn + feMn + feSi, "###,##0")
            litLrfDeltaT.Text = If(avgStartTmp > 0, FormatNumber(avgStartTmp, "###,##0"), "-")

            PrepareChart(chartLrfAdditives, "LrfAdditiveArea", ChartPalette)
            Dim additiveSeries As Series = chartLrfAdditives.Series.Add("Additives")
            additiveSeries.ChartType = SeriesChartType.Column
            additiveSeries.IsValueShownAsLabel = True
            additiveSeries.Points.AddXY("Lime", lime)
            additiveSeries.Points.AddXY("Coke", coke)
            additiveSeries.Points.AddXY("Si-Mn", siMn)
            additiveSeries.Points.AddXY("Fe-Mn", feMn)
            additiveSeries.Points.AddXY("Fe-Si", feSi)

            PrepareChart(chartLrfDistribution, "LrfDistributionArea", ChartPalette)
            Dim distributionSeries As Series = chartLrfDistribution.Series.Add("Alloys")
            distributionSeries.ChartType = SeriesChartType.Pie
            distributionSeries.Points.AddXY("Si-Mn", siMn)
            distributionSeries.Points.AddXY("Fe-Mn", feMn)
            distributionSeries.Points.AddXY("Fe-Si", feSi)

            PrepareChart(chartLrfTrend, "LrfTrendArea", ChartPalette)
            chartLrfTrend.ChartAreas("LrfTrendArea").AxisX.Interval = 1
            chartLrfTrend.ChartAreas("LrfTrendArea").AxisY.Title = "Power (MWh)"
            Dim trendSeries As Series = chartLrfTrend.Series.Add("LRF Power")
            trendSeries.ChartType = SeriesChartType.Line
            trendSeries.MarkerStyle = MarkerStyle.Circle
            trendSeries.MarkerSize = 5

            Dim lrfGrid As DataTable = BuildLrfDetailTable()
            For Each row As DataRow In detailTable.Rows
                Dim heatNum As Integer = Convert.ToInt32(row("HEAT_NUM"), CultureInfo.InvariantCulture)
                Dim lrfPower As Double = GetDoubleValue(row, "LRF_KWH") / 1000.0
                Dim stir As Double = GetDoubleValue(row, "LRF_STIR_TIME")
                Dim limeRow As Double = GetDoubleValue(row, "LRF_LIME")
                Dim alloys As Double = GetDoubleValue(row, "LRF_SI_MN") + GetDoubleValue(row, "LRF_FE_MN") + GetDoubleValue(row, "LRF_FE_SI")

                trendSeries.Points.AddXY(heatNum, lrfPower)

                Dim newRow As DataRow = lrfGrid.NewRow()
                newRow("Heat") = heatNum
                newRow("Ladle") = row("LADLE_NUM")
                newRow("Stir") = FormatNumber(stir, "###,##0.0")
                newRow("Power") = lrfPower
                newRow("Lime") = limeRow
                newRow("Alloys") = alloys
                lrfGrid.Rows.Add(newRow)
            Next

            gvLrfDetail.DataSource = lrfGrid
            gvLrfDetail.DataBind()
        End Sub
        Private Sub ClearVisuals()
            gvEafDetail.DataSource = Nothing
            gvEafDetail.DataBind()
            gvCcmDetail.DataSource = Nothing
            gvCcmDetail.DataBind()
            gvLrfDetail.DataSource = Nothing
            gvLrfDetail.DataBind()

            chartSummaryProduction.Series.Clear()
            chartSummaryEnergy.Series.Clear()
            chartEafMaterials.Series.Clear()
            chartEafEnergy.Series.Clear()
            chartEafTrend.Series.Clear()
            chartCcmBillets.Series.Clear()
            chartCcmShift.Series.Clear()
            chartCcmTrend.Series.Clear()
            chartLrfAdditives.Series.Clear()
            chartLrfDistribution.Series.Clear()
            chartLrfTrend.Series.Clear()
        End Sub

        Private Sub HideEafControls()
            gvEafDetail.DataSource = Nothing
            gvEafDetail.DataBind()
        End Sub

        Private Sub HideCcmControls()
            gvCcmDetail.DataSource = Nothing
            gvCcmDetail.DataBind()
        End Sub

        Private Sub HideLrfControls()
            gvLrfDetail.DataSource = Nothing
            gvLrfDetail.DataBind()
        End Sub

        Private Function BuildEafDetailTable() As DataTable
            Dim table As New DataTable()
            table.Columns.Add("Heat", GetType(Integer))
            table.Columns.Add("Shift", GetType(String))
            table.Columns.Add("StartTime", GetType(String))
            table.Columns.Add("Grade", GetType(String))
            table.Columns.Add("LiquidSteel", GetType(Double))
            table.Columns.Add("Kwh", GetType(Double))
            table.Columns.Add("KwhPerTon", GetType(Double))
            Return table
        End Function

        Private Function BuildCcmDetailTable() As DataTable
            Dim table As New DataTable()
            table.Columns.Add("Heat", GetType(Integer))
            table.Columns.Add("Shift", GetType(String))
            table.Columns.Add("Start", GetType(String))
            table.Columns.Add("End", GetType(String))
            table.Columns.Add("Duration", GetType(String))
            table.Columns.Add("BilletWgt", GetType(Double))
            table.Columns.Add("ProdRate", GetType(Double))
            Return table
        End Function

        Private Function BuildLrfDetailTable() As DataTable
            Dim table As New DataTable()
            table.Columns.Add("Heat", GetType(Integer))
            table.Columns.Add("Ladle", GetType(Object))
            table.Columns.Add("Stir", GetType(String))
            table.Columns.Add("Power", GetType(Double))
            table.Columns.Add("Lime", GetType(Double))
            table.Columns.Add("Alloys", GetType(Double))
            Return table
        End Function

        Private Function ComputeEafChargeFromDetail(ByVal detailTable As DataTable) As Double
            Dim totalChargeKg As Double = 0.0
            For Each row As DataRow In detailTable.Rows
                totalChargeKg += (GetDoubleValue(row, "CHG_WGT") + GetDoubleValue(row, "DRI_WGT") + GetDoubleValue(row, "DRI_FEED")) * 1000.0
                totalChargeKg += GetDoubleValue(row, "HBI_CONS") * 1000.0
                totalChargeKg += GetDoubleValue(row, "PIG_CONS")
            Next
            If totalChargeKg = 0 Then
                Return 0.0
            End If
            Return totalChargeKg / 1000.0
        End Function

        Private Function ComputeTotalLiquid(ByVal detailTable As DataTable) As Double
            Dim total As Double = 0.0
            For Each row As DataRow In detailTable.Rows
                total += GetDoubleValue(row, "ACT_LIQ_WGT")
            Next
            Return total
        End Function

        Private Sub PrepareChart(ByVal chart As Chart, ByVal areaName As String, ByVal palette As ChartColorPalette)
            chart.Series.Clear()
            If Not chart.ChartAreas.Contains(areaName) Then
                chart.ChartAreas.Add(areaName)
            End If
            chart.ChartAreas(areaName).AxisX.MajorGrid.Enabled = False
            chart.ChartAreas(areaName).AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot
            chart.Palette = palette
        End Sub

        Private Function GetDoubleValue(ByVal row As DataRow, ByVal columnName As String) As Double
            If row Is Nothing OrElse Not row.Table.Columns.Contains(columnName) OrElse row.IsNull(columnName) Then
                Return 0.0
            End If
            Return Convert.ToDouble(row(columnName), CultureInfo.InvariantCulture)
        End Function

        Private Function FormatTons(ByVal value As Double) As String
            Return FormatNumber(value, "###,##0.00")
        End Function

        Private Function FormatPercent(ByVal value As Double) As String
            Return value.ToString("0.00' %'", CultureInfo.InvariantCulture)
        End Function

        Private Function FormatNumber(ByVal value As Double, ByVal format As String) As String
            Return value.ToString(format, CultureInfo.InvariantCulture)
        End Function

        Private Function FormatTime(ByVal row As DataRow, ByVal columnName As String) As String
            If row.IsNull(columnName) Then
                Return String.Empty
            End If
            Dim value As DateTime = Convert.ToDateTime(row(columnName), CultureInfo.InvariantCulture)
            Return value.ToString("MM-dd HH:mm", CultureInfo.InvariantCulture)
        End Function

        Private Function FormatShiftName(ByVal rawShift As String) As String
            If String.IsNullOrEmpty(rawShift) Then
                Return String.Empty
            End If
            If rawShift.Length >= 6 AndAlso rawShift.StartsWith("Shift", StringComparison.OrdinalIgnoreCase) Then
                Dim suffix As String = rawShift.Substring(rawShift.Length - 1, 1)
                If Char.IsLetter(suffix, 0) Then
                    Return "Shift " & suffix.ToUpperInvariant()
                End If
            End If
            Return rawShift
        End Function
    End Class
End Namespace
