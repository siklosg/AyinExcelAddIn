namespace AyinExcelAddIn

module public BondFunctions =

    open FSharpx
    open AyinExcelAddIn.Business
    open AyinExcelAddIn.Business.Utils

    ///
    let bbrSymbol id =
        match getDealBond id with
        | DealInfo d -> d.bloomberg_symbol
        | BondInfo b -> b.symbol
        | NoInfo     -> UdfError.InvalidDealBond
        |> box

    ///
    let intexDeal id =
        match getDealBond id with
        | DealInfo d -> d.intex_symbol
        | BondInfo b -> match b.deal with
                        |Some d -> d.intex_symbol
                        | _     -> UdfError.NoData
        | NoInfo     -> UdfError.InvalidDealBond
        |> box

    ///
    let bbrDeal id =
        match getDealBond id with
        | DealInfo d -> d.bloomberg_symbol
        | BondInfo b -> match b.deal with
                        | Some d -> d.bloomberg_symbol
                        | _      -> UdfError.NoData
        | NoInfo     -> UdfError.InvalidDealBond
        |> box

    ///
    let bondCusip id =
        match getDealBond id with
        | BondInfo b -> b.cusip.GetOrElse(UdfError.NoData)
        | DealInfo _ -> UdfError.InvalidBond
        | NoInfo     -> UdfError.InvalidBond
        |> box


    ///
    let origSize id =
        let db = Db.nyabsDbCon.GetDataContext()

        match getDealBond id with
        | DealInfo d -> query {
                            for ds in db.Intex.Deals do
                            where (ds.Name = d.intex_symbol)
                            select ds.OrigSize
                            exactlyOneOrDefault
                        } |> box

        | BondInfo _ -> UdfError.InvalidDeal |> box
        | NoInfo     -> UdfError.InvalidDeal |> box

    ///
    let bondGroups id =
        match getDealBond id with
        | BondInfo b -> match b.tranche with
                        | Some t -> t.groups
                        | _    -> UdfError.NoData
        | DealInfo _ -> UdfError.InvalidBond
        | NoInfo     -> UdfError.InvalidBond
        |> box

    ///
    let bondCollat id =
        match getDealBond id with
        | BondInfo b -> match b.deal with
                        | Some d -> d.collat_type
                        | _    -> "Other"
        | DealInfo _ -> UdfError.InvalidBond
        | NoInfo     -> "Other"
        |> box



