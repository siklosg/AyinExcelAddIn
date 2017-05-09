namespace AyinExcelAddIn

open System
open System.Windows.Forms
open ExcelDna.Integration
open ExcelDna.XlUtils
open Microsoft.FSharp.Core
open NetOffice.ExcelApi
open NetOffice.OfficeApi
open NetOffice.OfficeApi.Enums

open AyinExcelAddIn.Backoffice
open AyinExcelAddIn.BondFunctions
open AyinExcelAddIn.DealGrp
open AyinExcelAddIn.Gui

type public AyinAddIn() =

    static let mutable _IDPFuncMap =
        [ "BLOOMBERG_SYMBOL", ("bondDbSymbol", bbrSymbol)
          "INTEX_DEAL", ("bondIntexDeal", intexDeal)
          "BLOOMBERG_DEAL", ("bondBbDeal", bbrDeal)
          "CUSIP", ("bondCusip", bondCusip)
          "ORIG_SIZE", ("origSize", origSize)
          "GROUPS", ("bondGroups", bondGroups)
          "COLLAT", ("bondCollat", bondCollat) ]
        |> Map.ofList

    static let mutable _IDHFuncMap = Map.empty
    static member IDPFuncMap = _IDPFuncMap
    static member IDHFuncMap = _IDHFuncMap

    interface IExcelAddIn with

        member this.AutoOpen() =
            let app = new Application(null, ExcelDnaUtil.Application)
            let cb = app.CommandBars.["Cell"]
            cb.Reset()
            let btn = cb.Controls.Add(MsoControlType.msoControlButton) :?> CommandBarButton
            btn.add_ClickEvent (new CommandBarButton_ClickEventHandler(fun _ _ ->
                let sht = app.ActiveSheet :?> Worksheet
                let selection = app.Selection :?> Range

                let name =
                    match getRangeName sht selection with
                    | Some n -> n.NameLocal
                    | None -> ""
                match name.Trim().Split([| "quote_" |], StringSplitOptions.None) with
                | [| _; qn |] -> handleFuncException (fun() -> QuoteFunctions.modifyQuote <| int qn |> ignore)
                | _ -> ()))
            btn.Caption <- "Modify Quote"
            btn.BeginGroup <- true

            try
                // Test the DB connection with a simple query.
                // If we can't connect we'll get an exception
                let db = Db.nyabsDbCon.GetDataContext()

                let dealGrpTags =
                    query {
                        for di in db.Intex.DealgrpDataItems do
                            select di.Tag
                    }
                    |> Seq.toList

                // Populate the IDP and IDH function maps
                do _IDPFuncMap <- dealGrpTags
                                  |> List.fold (fun m t -> m.Add(t, ("dealGrpDP " + t, dealGrpDP t))) _IDPFuncMap
                do _IDHFuncMap <- dealGrpTags
                                  |> List.fold (fun m t -> m.Add(t, ("dealGrpH " + t, dealGrpH t))) _IDHFuncMap
                ()
            with _ ->
                ErrorBox
                <| "No connection to the Ayin database\n" + "Most functions of Ayin Add-in will not be functional"

        // Does nothing
        member this.AutoClose() = ()

    //
    //
    //
    static member help () =
        let version = "2.0.0"
        let releaseDate = "04/24/2017"

        let res =
            try
                let db = Db.nyabsDbCon.GetDataContext()
                Some <| query {
                            for d in db.Intex.Deals do
                                select d.Name
                                headOrDefault
                        }
            with _ -> None

        let conStat =
            match res with
            | None -> "Failed"
            | _ -> "Ok"

        MessageBox.Show("Version: " + version + "\n" + "Released: " + releaseDate + "\n" + "Database connection: " + conStat)
        |> ignore


