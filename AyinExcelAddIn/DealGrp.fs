namespace AyinExcelAddIn

module public DealGrp =

    open System
    open System.Collections.Generic
    open ExcelDna.Integration
    open ExcelDna.XlUtils

    open AyinExcelAddIn
    open AyinExcelAddIn.Utils
    open AyinExcelAddIn.Business
    open AyinExcelAddIn.Business.Utils

   /// The default options for history functions.
    let defaultOpts = "Dir=V; Dts=S; SortBy=1; Ord=D"

    /// Parse the user supplied options. The defaults will be overwritten by these.
    let parseOptArgs (str: string) =
        str.Split([|";"|], StringSplitOptions.RemoveEmptyEntries)
        |> Array.map (fun s -> s.Trim().Split([|"="|], 2, StringSplitOptions.None))
        |> Array.map (function
                      | [| k; v |] -> (k, v)
                      | _          -> failwith "Error parsing options")
        |> Map.ofArray

    /// Comparator for comparing historical data points - (date, value).
    let private histDataCmpr sortBy order (d1, v1) (d2, v2) =
        let ord = if order = "A" then 1 else -1

        match sortBy with
        | "1" -> (compare d1 d2) * ord
        | _   -> (compare v1 v2) * ord

    /// Transforms a functions that returns historical data points to a function
    /// that returns the same data points taking into account the specified options, and
    /// in an Excel ingestible obj[,] array2D.
    let toHistoryFunc (func: 'T -> (DateTime * 'U option) list) (opts: Map<string, string>) =
        fun xs ->
            func xs
            |> List.sortWith (histDataCmpr opts.["SortBy"] opts.["Ord"])
            |> match opts.["Dts"] with
               | "H" -> List.map (fun (_, v) -> [toExcelValueOpt v])
               | _   -> List.map (fun (d, v) -> [toExcelValue d; toExcelValueOpt v])
            |> array2D
            |> match opts.["Dir"] with
               | "H" -> transpose
               | _   -> Operators.id

    // Returns a key-value mapping associated with the greatest key
    // less than or equal to the given key, or None if there
    // is no such key.
    let private floorEntry key lst =
        try
            Some <| List.find (fun (k, _) -> k <= key) lst
        with
            | :? KeyNotFoundException -> None
    ///
    let dealGrpH tag (id, d1, d2) =
        let db = Db.nyabsDbCon.GetDataContext()

        let yyyymm1 = toYYYYMM d1
        let yyyymm2 = toYYYYMM d2

        let res =
            match getDealBond id with
            | BondInfo b -> match (b.deal, b.tranche) with
                            | (Some d, Some t) ->
                                query {
                                    for dg in db.Intex.DealgrpData do
                                    join di in db.Intex.DealgrpDataItems on (dg.DataItemId = di.Id)
                                    where ((dg.DealId = d.id) && (dg.Groups = t.groups) && (di.Tag = tag) && (dg.Yyyymm >= yyyymm1) && (dg.Yyyymm <= yyyymm2))
                                    sortByDescending dg.Yyyymm
                                    select (dg.Yyyymm, dg.Value)
                                } |> Seq.toList
                            | _ -> List.empty

            | DealInfo _ -> List.empty
            | NoInfo     -> List.empty

        let months = Seq.unfold (fun d -> if (d < d1) then None else Some (d, d.AddMonths(-1))) d2 |> Seq.map toYYYYMM

        [
            for m in months do
                match floorEntry m res with
                | Some (_, v) -> yield (toDate m, Some v)
                | None        -> yield (toDate m, None)
        ]



    /// Returns the most recent RBS First Look data point for the
    /// specified tag and id.
    let dealGrpDP tag id =
        let d2 = DateTime.Today
        // Data that is older than 3 months old is considered "too old"
        //let d1 = d2.AddMonths(-3)
        let d1 = DateTime.Parse("2001-01-01")

        match dealGrpH tag (id, d1, d2) with
        | (_, None) :: _   -> ExcelError.ExcelErrorNull |> box
        | (_, Some v) :: _ -> v |> box
        | _                -> ExcelError.ExcelErrorNull |> box

