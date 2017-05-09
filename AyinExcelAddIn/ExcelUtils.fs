namespace ExcelDna

open ExcelDna.Integration
open ExcelDna.Integration.Rtd
open Microsoft.FSharp.Reflection
open NetOffice.ExcelApi
open System
open System.Collections.Generic
open System.Threading

open AyinExcelAddIn.Gui

module public ExcelAsyncUtil =
   /// A helper to pass an F# Async computation to Excel-DNA
   let excelRunAsync functionName parameters async =
       let obsSource =
            ExcelObservableSource(
                fun _ ->
                    {new IExcelObservable with
                        member this.Subscribe observer =
                            // make something like CancellationDisposable
                            let cts = new CancellationTokenSource()
                            let disp = { new IDisposable with member this.Dispose() = cts.Cancel() }

                            // Start the async computation on this thread
                            Async.StartWithContinuations(
                                async,
                                (fun result ->
                                    observer.OnNext(result)
                                    observer.OnCompleted()),
                                (fun ex -> observer.OnError ex),
                                (fun _ -> observer.OnCompleted()),
                                cts.Token)


                            // return the disposable
                            disp
                    })

       ExcelAsyncUtil.Observe(functionName, parameters, obsSource)

   /// A helper to pass an F# IObservable to Excel-DNA
   let excelObserve functionName parameters observable =
       let obsSource =
           ExcelObservableSource(
            fun _ ->
                {new IExcelObservable with
                   member __.Subscribe observer =
                       // Subscribe to the F# observable
                       Observable.subscribe(fun value -> observer.OnNext(value)) observable
                })

       ExcelAsyncUtil.Observe (functionName, parameters, obsSource)

module ArrayResizer =

    let private ResizeJobs = new Queue<ExcelReference>()

    let internal EnqueueResize(caller:ExcelReference, rows:int, columns:int) =
        let target = new ExcelReference(
                            caller.RowFirst,
                            caller.RowFirst + rows - 1,
                            caller.ColumnFirst,
                            caller.ColumnFirst + columns - 1,
                            caller.SheetId)
        ResizeJobs.Enqueue(target)

    let private DoResize(target:ExcelReference) =
        try
            XlCall.Excel(XlCall.xlcEcho, false) |> ignore
            // Get the formula in the first cell of the target
            let (formula: string) =
                XlCall.Excel(XlCall.xlfGetCell, 41, target) |> unbox

            let firstCell = new ExcelReference(
                                    target.RowFirst,
                                    target.RowFirst,
                                    target.ColumnFirst,
                                    target.ColumnFirst,
                                    target.SheetId)

            let isFormulaArray:bool =
                XlCall.Excel(XlCall.xlfGetCell, 49, target) |> unbox

            if isFormulaArray then
                let oldSelectionOnActiveSheet =
                    XlCall.Excel(XlCall.xlfSelection)
                let oldActiveCell = XlCall.Excel(XlCall.xlfActiveCell)
                // Remember old selection and select the first cell of the target
                let firstCellSheet:string =
                    XlCall.Excel(XlCall.xlSheetNm, firstCell) |> unbox

                XlCall.Excel(XlCall.xlcWorkbookSelect,firstCellSheet) |> ignore

                let oldSelectionOnArraySheet = XlCall.Excel(XlCall.xlfSelection)

                XlCall.Excel(XlCall.xlcFormulaGoto,
                            firstCell) |> ignore

                // Extend the selection to the whole array and clear
                XlCall.Excel(XlCall.xlcSelectSpecial, 6) |> ignore

                let (oldArray: ExcelReference) =
                    XlCall.Excel(XlCall.xlfSelection) |> unbox

                oldArray.SetValue(ExcelEmpty.Value) |> ignore

                XlCall.Excel(XlCall.xlcSelect,
                            oldSelectionOnArraySheet) |> ignore

                XlCall.Excel(XlCall.xlcFormulaGoto,
                            oldSelectionOnActiveSheet) |> ignore

            // Get the formula and convert to R1C1 mode
            let isR1C1Mode:bool =
                XlCall.Excel(XlCall.xlfGetWorkspace, 4) |> unbox |> unbox
            let mutable formulaR1C1 = formula

            if not isR1C1Mode then
                // Set the formula into the whole target
                formulaR1C1 <- XlCall.Excel(
                                    XlCall.xlfFormulaConvert,
                                    formula,
                                    true,
                                    false,
                                    ExcelMissing.Value,
                                    firstCell) :?> string
            let ignoredResult = new System.Object()
            let  retval = XlCall.TryExcel(
                            XlCall.xlcFormulaArray,
                            ref ignoredResult,
                            formulaR1C1,
                            target)

            if not (retval = XlCall.XlReturn.XlReturnSuccess)
            then firstCell.SetValue("'" + formula) |> ignore

        finally
                XlCall.Excel(XlCall.xlcEcho, true) |> ignore

    let private DoResizing() =
        while ResizeJobs.Count > 0 do
            DoResize(ResizeJobs.Dequeue())


    /// Resizes array output of Excel UDFs
    [<ExcelFunction>]
    let Resize(array: obj[,]) =
        let (caller: ExcelReference) = XlCall.Excel(XlCall.xlfCaller) |> unbox

        if (caller = null) then
            array
        else
            let rows = array.GetLength(0)
            let columns = array.GetLength(1)

            if ((caller.RowLast - caller.RowFirst + 1 <> rows) ||
                (caller.ColumnLast - caller.ColumnFirst + 1 <> columns)) then
                    EnqueueResize(caller, rows, columns)
                    ExcelAsyncUtil.QueueAsMacro(ExcelAction(DoResizing))
                    array2D [[ box ExcelError.ExcelErrorNA ]]
            else
                    array

module XlCacheUtility =
    [<Literal>]
    let RTDServer = "Utility.StaticRTD"
    let private _objects = new Dictionary<string,obj>()
    let private _tags = new Dictionary<string, int>()

    let lookup handle =
        match _objects.ContainsKey(handle) with
        |true -> _objects.[handle]
        |false -> failwith "handle not found"

    let register tag (o:obj) =
        let counter =
            match _tags.ContainsKey(tag) with
            |true  -> _tags.[tag] + 1
            |false -> 1
        _tags.[tag] <- counter
        let handle = sprintf "[%s.%i]" tag counter
        _objects.[handle] <- o
        handle

    let unregister handle =
        if _objects.ContainsKey(handle)
        then _objects.Remove(handle) |> ignore

type public StaticRTD() =
    inherit ExcelRtdServer()
    let _topics = new Dictionary<ExcelRtdServer.Topic, string>()

    override x.ConnectData(topic, topicInfo, newValues) =
        let name = topicInfo.[0]
        _topics.[topic] <- name
        box name

    override x.DisconnectData(topic: ExcelRtdServer.Topic) =
        _topics.[topic] |> XlCacheUtility.unregister
        _topics.Remove(topic) |> ignore

module XlCache =
    let inline lookup handle = XlCacheUtility.lookup handle |> unbox

    let convertToArrayFunc (func: 'T -> 'U) (args: 'T) =
        if FSharpType.IsTuple(args.GetType()) then
            ((fun xs -> func(unbox <| FSharpValue.MakeTuple(xs, typeof<'T>))),
             FSharpValue.GetTupleFields args)
        else
            ((fun xs -> func(unbox <| xs.[0])), [| box args |])

    let inline register tag (o: obj) =
        o |> XlCacheUtility.register tag
        |> fun name -> XlCall.RTD(XlCacheUtility.RTDServer, null, name)

    ///
    let inline asyncRegister tag func args =
        let (f, p) = convertToArrayFunc func args

        ExcelAsyncUtil.Run(tag, p, fun () -> box <| f p)
        |> function
            |result when result = box ExcelError.ExcelErrorNA -> box ExcelError.ExcelErrorGettingData
            |result -> XlCacheUtility.register tag result
                    |> fun name -> XlCall.RTD(XlCacheUtility.RTDServer, null, name)

    ///
    let inline asyncRun tag func args =
        let (f, p) = convertToArrayFunc func args

        ExcelAsyncUtil.Run(tag, p, fun () -> box <| f p)
        |> function
            |result when result = box ExcelError.ExcelErrorNA -> box ExcelError.ExcelErrorGettingData
            |result -> result

    ///
    let inline asyncRunAndResize tag func args =
        let (f, p) = convertToArrayFunc func args

        ExcelAsyncUtil.Run(tag, p, fun () -> box <| f p)
        |> function
            |result when result = box ExcelError.ExcelErrorNA -> array2D [[box ExcelError.ExcelErrorGettingData]]
            |result -> result :?> obj[,] |> ArrayResizer.Resize

module XlUtils =

    let public toExcelFuncOpt (func: 'T -> 'U option) =
        fun xs ->
            match func xs with
            | Some x -> box x
            | None -> box ExcelError.ExcelErrorNull


    /// Convert an optional value to a value Excel can understand.
    let public toExcelValueOpt =
        function
        | Some x -> box x
        | None -> box ExcelError.ExcelErrorNull

    /// Convert an value to a value Excel can understand. Date values are converted into strings.
    let public toExcelValue (v: obj) =
        match v with
        | :? DateTime as d -> box <| d.ToString("MM/dd/yyyy")
        | _ -> box v


    let public getDateValue (d: obj) =
        match d with
        | :? ExcelMissing -> DateTime.Today
        | :? string       -> DateTime.Parse(string d)
        | :? float        -> DateTime.FromOADate(d :?> float)
        | _               -> failwith "Could not parse date value"

    let public getRangeName (sht: Worksheet) (rng: Range) =
        let app = sht.Application

        try
            sht.Names
            |> Seq.find (fun n -> app.Intersect(n.RefersToRange, rng) <> null)
            |> Some
        with
            | :? System.Collections.Generic.KeyNotFoundException -> None

    // Excecute the given function. Catch and display any exceptions.
    let public handleFuncException func =
        try
            func()
        with
            | e -> ErrorBox <| e.ToString()

    type Sheets with

        member this.AddOrReplace (name:string, ``type``:obj) =
            try
                this.Application.DisplayAlerts <- false
                (this.[name] :?> Worksheet).Delete()
                this.Application.DisplayAlerts <- true
            with
                | _ -> ()

            let sht = this.Add(null, null, 1, ``type``) :?> Worksheet
            sht.Name <- name
            sht

        member this.AddOrReplace (name:string) =
            try
                let sht = this.[name] :?> Worksheet
                sht.Cells.Clear() |> ignore
                sht
            with
                | _ -> let sht = this.Add(null, null, 1) :?> Worksheet
                       sht.Name <- name
                       sht

        member this.AddWithIncrement(name:string) =
            let n =
                this
                |> Seq.map (fun s -> match s with
                                          | :? Worksheet as sht ->
                                                match sht.Name.Split([| (name + " ") |], StringSplitOptions.RemoveEmptyEntries) with
                                                | [||] -> 0
                                                | [| s |] -> try
                                                                int s
                                                             with
                                                                | _ -> 0
                                                | _ -> 0
                                          | _ -> 0)
                |> Seq.max
                |> (+) 1

            let sht = this.Add(null, null, 1) :?> Worksheet
            sht.Name <- name + " " + n.ToString()

            sht

        member this.FindSheet(name:string) =
            try
                Some (this.[name] :?> Worksheet)
            with
                | _ -> None

module public  XlWorksheetRefreshUtil =
    let private _shtRefresherMap = new Dictionary<string, unit -> unit>()

    let private cleanRefreshMap =
        let app = new Application(null, ExcelDnaUtil.Application)
        let shts = app.Sheets

        _shtRefresherMap.Keys
        |> Seq.iter (fun k ->
                            try
                                shts.[k] |> ignore
                            with
                                | _ -> _shtRefresherMap.Remove(k) |> ignore)
        ()

    let public registerShtRefresher shtName func =
        cleanRefreshMap
        _shtRefresherMap.[shtName] <- func

    let public lookup (sht: Worksheet) =
        match _shtRefresherMap.TryGetValue(sht.Name) with
        | (true, func) -> Some func
        | (false, _)   -> None

