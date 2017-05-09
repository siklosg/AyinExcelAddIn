namespace AyinExcelAddIn.Business

open System.Linq
open System.Runtime.Caching
open AyinExcelAddIn
open AyinExcelAddIn.Backoffice
open AyinExcelAddIn.Utils

///
///
///
type public Deal =
    { id : int
      intex_symbol : string
      bloomberg_symbol : string
      collat_type : string }

///
///
///
type public Tranche =
    { id : int
      name : string
      groups : string
      bloomberg_symbol : string }

///
///
///
module public Bond =
    type T =
        { symbol : string
          cusip : string option
          security : Security.T option
          deal : Deal option
          tranche : Tranche option }


    // Creates a Bond type from a record in the bonds table.
    let internal mkBond ((deal : Db.nyabsDbCon.dataContext.``intex.dealsEntity`` option), 
                         (tranche : Db.nyabsDbCon.dataContext.``intex.tranchesEntity`` option),
                         security : Security.T option) =

        match (deal, tranche, security) with
        | (Some d, Some t, _) ->
            let tickerParts = t.BbgTicker.Split([| ' ' |])
            { symbol = t.BbgTicker.Trim()
              cusip = Some(t.Cusip.Trim())
              security = security
              deal =
                  Some { id = d.Id
                         intex_symbol = d.Name.Trim()
                         bloomberg_symbol = tickerParts.[0] + " " + tickerParts.[1]
                         collat_type = d.CollatType }
              tranche =
                  Some { id = t.Id
                         name = t.Name.Trim()
                         groups =
                             match t.Groups with
                             | null -> ""
                             | _ -> t.Groups.Trim()
                         bloomberg_symbol = t.BbgTicker.Trim() } }
        | (_, _, Some s) ->
            { symbol = s.symbol
              cusip = s.cusip
              security = security
              deal = None
              tranche = None }
        | (_, _, _) -> failwith "This should not have happened"


    // Returns the bond with the specified CUSIP.
    let getBond id =
        let security = Security.getSecurity id
        let db = Db.nyabsDbCon.GetDataContext()

        let (deal, tranche) =
            query {
                for d in db.Intex.Deals do
                    join t in db.Intex.Tranches on (d.Id = t.DealId)
                    where ((t.Cusip = id) || (t.BbgTicker = id))
                    select (d, t)
                    headOrDefault
            }
        match (deal, tranche, security) with
        | (null, null, Left _) -> Left <| "No bond found with id: \"" + id + "\""
        | (null, _, Right s) -> Right <| mkBond (None, None, Some s)
        | (d, t, Left _) -> Right <| mkBond (Some d, Some t, None)
        | (d, t, Right s) -> Right <| mkBond (Some d, Some t, Some s)

type DealBondInfo =
    | DealInfo of Deal
    | BondInfo of Bond.T
    | NoInfo


module public Utils =
    let getDealBond (id : string) =
        let cache = MemoryCache.Default
        let id' = id.Trim()
        if (id' = "") then NoInfo
        elif (cache.Contains(id')) then unbox cache.[id']
        else
            let security =
                match Security.getSecurity id' with
                | Right s -> Some s
                | _ -> None

            let db = Db.nyabsDbCon.GetDataContext()

            // Try to see if it's a bond
            let res =
                query {
                    for d in db.Intex.Deals do
                        join t in db.Intex.Tranches on (d.Id = t.DealId)
                        where ((t.BbgTicker = id') || (t.Cusip = id'))
                        select (d, t)
                        headOrDefault
                }

            let dbi =
                match (security, box res) with
                | (None, null) ->
                    // See if it's a deal
                    let res =
                        query {
                            for d in db.Intex.Deals do
                                join t in db.Intex.Tranches on (d.Id = t.DealId)
                                where ((d.Name = id') || (t.BbgTicker.StartsWith(id')))
                                select (d, t)
                                headOrDefault
                        }
                    if (box res = null) then NoInfo
                    else
                        match res with
                        | (d, t) ->
                            let tickerParts = t.BbgTicker.Split([| ' ' |])
                            DealInfo({ id = d.Id
                                       intex_symbol = d.Name
                                       bloomberg_symbol = tickerParts.[0] + " " + tickerParts.[1]
                                       collat_type = d.CollatType })
                | (Some s, null) ->
                    match s.sectype with
                    | Security.ABS -> BondInfo(Bond.mkBond (None, None, security))
                    | _ -> NoInfo
                | _ -> BondInfo(Bond.mkBond (Some <| fst res, Some <| snd res, security))

            // Add the result to the cache
            cache.Add(id', box dbi, System.DateTimeOffset.Now.AddHours(1.), null) |> ignore
            // Return the result
            dbi

