namespace AyinExcelAddIn.RibbonFunctions

open System
open Microsoft.FSharp.Reflection
open ExcelDna.Integration
open ExcelDna.XlUtils
open FSharpx
open NetOffice.ExcelApi
open NetOffice.ExcelApi.Enums

open AyinExcelAddIn
open AyinExcelAddIn.Utils
open AyinExcelAddIn.Business
open AyinExcelAddIn.Business.Utils
open AyinExcelAddIn.Gui

[<AutoOpen>]
module TearSheet =

    let private perfColHeaders =
        ["ITEM_TYPE"; "ITEM_NAME"; "YYYYMM"; "DEAL_AGE"; "1MO_WALA"; "1MO_COLLAT_BAL";
         "ORIG_BALANCE"; "CURR_BALANCE"; "TRANCHE_1MO_BALANCE"; "TRANCHE_CURRENT_CREDIT_SUPPORT_PCTS";
         "OC_ACTUAL_VALS"; "OC_TARGET_VALS"; "1MO_WAC"; "1MO_NET_WAC"; "1MO_CPR"; "1MO_DELINQ_30_59";
         "1MO_DELINQ_60_PLUS"; "1MO_DELINQ_90_PLUS"; "1MO_BANKRUPT_RATE"; "1MO_FORECLOSURE_RATE";
         "1MO_REO_RATE"; "1MO_ACCUM_NET_LOSS_RATE"; "1MO_NET_LOSS_RATE"; "1MO_CRR"; "1MO_CDR"; "1MO_PRIN_LIQ";
         "SEV"; "WEIGHTED AVG SEV"; "LOSS"; "AVG 6M LOSS"; "PROJ LOSS";
        ]

    let private getCNL symbol g yyyymm =
        let db = Db.nyabsDbCon.GetDataContext()

        match g with
        | null | "0" ->
            query {
                for p in db.Intex.PerfRpt do
                where ((p.Deal = symbol)
                        && (p.Yyyymm = yyyymm)
                        && (p.Itemtype = "DEAL"))
                select (p.Origbalance, p.Accumloss)
            }

        | _          ->
            let gs = g.Split([| ',' |]) |> Array.toSeq
            query {
                for p in db.Intex.PerfRpt do
                where ((p.Deal = symbol)
                        && (p.Yyyymm = yyyymm)
                        && (p.Itemtype = "GROUPNO")
                        && (g.Contains(p.Itemname)))
                select (p.Origbalance, p.Accumloss)
            }

        |> Seq.map (fun (obal, loss) -> (float obal * loss) / 100.)
        |> Seq.fold (+) 0.

    let private loadPerfData (bond: Bond.T) (sht: Worksheet) yyyymm =
        let db = Db.nyabsDbCon.GetDataContext()
        let cusip = bond.cusip.GetOrDefault()
        let deal = bond.deal.GetOrDefault()
        let tranche = bond.tranche.GetOrDefault();

        let group =
           query {
               for ps in db.Intex.PerfStatic do
               where (ps.TrancheCusips = cusip)
               select ps.TrancheGroups
               headOrDefault
           }
        let cnl = getCNL deal.intex_symbol group yyyymm
        if (cnl <> 0.) then
            sht.Cells.[8, 9].Value <- cnl / 1000000.
            sht.Cells.[8, 9].Interior.Color <- System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen)


        if (group <> null) && group.Contains(",") then
            ErrorBox <| "No performance data for " + bond.symbol
        else
            let (itemtype, itemname) =
                match group with
                | null | "0" -> ("DEAL", "*")
                | _          -> ("GROUPNO", group)

            let startYYYYMM = toYYYYMM ((toDate yyyymm).AddMonths(-48))

            let pd =
                query {
                    for pd1 in db.Intex.PerfRpt do
                    join pd2 in db.Intex.PerfRpt on ((pd1.Deal, pd1.Yyyymm) = (pd2.Deal, pd2.Yyyymm))
                    where ((pd1.Deal = deal.intex_symbol)
                           && (pd1.Itemtype = itemtype)
                           && (pd1.Itemname = itemname)
                           && (pd2.Itemtype = "TRNO")
                           && (pd2.Itemname = tranche.name)
                           && (pd1.Yyyymm <= yyyymm)
                           && (pd1.Yyyymm > startYYYYMM))
                    sortBy pd1.Yyyymm
                    select (pd1.Itemtype,
                            pd1.Itemname,
                            pd1.Yyyymm,
                            pd1.Dealage,
                            pd1.``1moWala``,
                            pd1.``1moCollatBal``,
                            pd1.Origbalance,
                            pd1.Curbal,
                            pd2.``1moBalance``,
                            pd2.CurrentCreditSupportPcts,
                            pd1.OcActualVals,
                            pd1.OcTargetVals,
                            pd1.``1moWac``,
                            pd1.``1moNetWac``,
                            pd1.``1moCpr``,
                            pd1.``1moDelinq3059``,
                            pd1.``1moDelinq60Plus``,
                            pd1.``1moDelinq90Plus``,
                            pd1.``1moBankruptRate``,
                            pd1.``1moForeclosureRate``,
                            pd1.``1moReoRate``,
                            pd1.``1moAccumNetLossRate``,
                            pd1.``1moNetLossRate``,
                            pd1.``1moCrr``,
                            pd1.``1moCdr``,
                            pd1.``1moPrinLiq``)
                }
                |> Seq.map (fun t -> FSharpValue.GetTupleFields t |> Array.toList)
                |> Seq.toList

            match pd with
            | [] -> ErrorBox <| "No performance data for " + bond.symbol
            | a :: _  ->
                    let numDataCols = a.Length
                    let cells = sht.Cells
                    let firstRowi = 140
                    let lastRowi = firstRowi + pd.Length
                    // Print the column headers and format it
                    List.iteri (fun i h -> cells.[firstRowi, i + 1].Value <- h) perfColHeaders
                    let perfHeaderRow = sht.Range(cells.[firstRowi, 1], cells.[firstRowi, perfColHeaders.Length])
                    let border = perfHeaderRow.Borders.[XlBordersIndex.xlEdgeBottom]
                    perfHeaderRow.NumberFormat <- "#.00000"
                    perfHeaderRow.Font.Bold <- true
                    perfHeaderRow.Font.Color <- 0x000000
                    border.LineStyle <- XlLineStyle.xlContinuous
                    border.Weight <- 2

                    // Output the performance data
                    pd
                    |> List.iteri (fun i rowData -> List.iteri (fun j d -> cells.[i + firstRowi + 1, j + 1].Value <- d) rowData)

                    // Change the color of the data cells
                    let dataRange = sht.Range(cells.[firstRowi + 1, 1], cells.[lastRowi, numDataCols])
                    dataRange.Font.Color <- 0x7D491F

                    // Calculated columns
                    let sev = @"=IFERROR((((RC[-5]-R[-1]C[-5])%*RC[-20])/RC[-1]*100), 0)"
                    let waSev = @"=IFERROR(SUMPRODUCT(R[-5]C[-1]:RC[-1], R[-5]C[-2]:RC[-2])/SUM(R[-5]C[-2]:RC[-2]), 0)"
                    let loss = @"=IFERROR(((RC[-7]-R[-1]C[-7])%*RC[-22]), 0)"
                    let avg6mLoss = @"=IFERROR(AVERAGE(R[-5]C[-1]:RC[-1]), 0)"

                    [firstRowi + 1 .. lastRowi]
                    |> List.iter (fun i ->
                                      cells.[i, numDataCols + 1].FormulaR1C1 <- sev
                                      if (i <= lastRowi - 5) then cells.[i + 5, numDataCols + 2].FormulaR1C1 <- waSev
                                      cells.[i, numDataCols + 3].FormulaR1C1 <- loss
                                      if (i <= lastRowi - 5) then cells.[i + 5, numDataCols + 4].FormulaR1C1 <- avg6mLoss)


    let private loadTemplate (bond: Bond.T) template =
        let db = Db.nyabsDbCon.GetDataContext()

        let app = new Application(null, ExcelDnaUtil.Application)
        let wkb = app.ActiveWorkbook
        let deal = bond.deal.GetOrDefault()
        let tranche = bond.tranche.GetOrDefault()

        // Get the most recent date for the dealgroup
        try
            let yyyymm =
                query {
                    for dg in db.Intex.DealgrpData do
                    where ((dg.DealId = deal.id) && (dg.Groups = tranche.groups))
                    maxBy dg.Yyyymm
                }

            // Temporarily suspend calculation
            let calcMode = app.Calculation
            app.Calculation <- XlCalculation.xlCalculationManual

            let sht = wkb.Application.Sheets.AddOrReplace(bond.symbol, template)

            sht.Cells.Item(1, 1).Value <- toDate yyyymm
            sht.Cells.Item(4, 1).Value <- bond.cusip.GetOrElse("")

            // Load the performance data into the sheet
            loadPerfData bond sht yyyymm |> ignore

            sht.Calculate |> ignore

            // Restore the calculation mode
            app.Calculation <- calcMode

        with
            | :? InvalidOperationException -> ErrorBox <| "No data for bond: " + bond.symbol

    let createTearSheet () =
        match SimpleInputBox "Bond" "Bond symbol or CUSIP" (Some(bondTickersCol)) with
        | Cancel | Ok "" -> ()
        | Ok bid -> match getDealBond bid with
                    | BondInfo b -> match (b.deal, b.tranche) with
                                    | (Some _, Some _) -> loadTemplate b @"S:\Excel Add In\Templates\Tear Sheet Template.xlsx"
                                    | _ -> ErrorBox <| "Invalid bond identifier: " + bid

                    | _ -> ErrorBox <| "Invalid bond identifier: " + bid


