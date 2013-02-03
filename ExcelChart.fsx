#r "office.dll"
#r "Microsoft.Office.Interop.Excel"

open Microsoft.Office.Interop.Excel
open System.Runtime.InteropServices

let Attach () = 
    try
        Marshal.GetActiveObject("Excel.Application") 
        :?> Microsoft.Office.Interop.Excel.Application
        |> Some
    with
    | _ -> 
        printfn "Could not attach to Excel"
        None

let Active () =
    let xl = Attach ()
    match xl with
    | None -> None
    | Some(xl) ->
        try
            xl.ActiveWorkbook |> Some   
        with
        | _ ->
            printfn "Could not find active workbook"
            None

let NewChart () =
    let wb = Active ()
    match wb with
    | None ->
        printfn "No workbook"
        None 
    | Some(wb) ->
        try
            let charts = wb.Charts
            charts.Add () :?> Chart |> Some
        with
        | _ -> 
            printfn "Failed to create chart"
            None
         
let WB (name: string) (xl: Microsoft.Office.Interop.Excel.Application Option) =
    match xl with
    | None -> 
        printfn "No Excel instance supplied"
        None
    | Some(xl) ->
        try
            let workbooks = xl.Workbooks
            workbooks.[name] 
            |> Some
        with
        | _ ->
            printfn "Workbook '%s' not found" name
            None

let CHART (name: string) (wb: Workbook Option) =
    match wb with
    | None ->
        printfn "No workbook supplied"
        None 
    | Some(wb) ->
        try
            let charts = wb.Charts
            let chart = charts.Add () :?> Chart
            chart.Location(XlChartLocation.xlLocationAsNewSheet, name)
            |> Some
        with
        | _ -> 
            printfn "Creating chart '%s' failed" name
            None
    
let TITLE (title: string) (chart: Chart) =

    chart.HasTitle <- true
    chart.ChartTitle.Text <- title
    
let LINE data (name: string) (chart: Chart) =

    let seriesCollection = chart.SeriesCollection() :?> SeriesCollection
    let series = seriesCollection.NewSeries()

    let labels, values = data

    series.Values <- values
    series.XValues <- labels
    series.Name <- name
    series.ChartType <- XlChartType.xlLine

    chart

// Plots a function of one parameter over an interval
let PLOT (f: float -> float) over =
    match NewChart () with
    | None -> ignore ()
    | Some(chart) ->
        chart.ChartType <- XlChartType.xlXYScatter
        let min, max = over
        let step = (max - min) / 50.
        let seriesCollection = chart.SeriesCollection() :?> SeriesCollection
        let series = seriesCollection.NewSeries()
        series.XValues <- [| min .. step .. max |]
        series.Values <- [| min .. step .. max |] |> Array.map f 

// Plots surface of a two parameters function over an interval
let SURF f over =
    match NewChart () with
    | None -> ignore ()
    | Some(chart) ->
        let (minX, maxX), (minY, maxY) = over
        let stepX = (maxX - minX) / 20.
        let stepY = (maxY - minY) / 20.
        let seriesCollection = chart.SeriesCollection() :?> SeriesCollection
        for x in minX .. stepX .. maxX do
            let series = seriesCollection.NewSeries()
            series.Name <- (string)x
            series.XValues <- [| minY .. stepY .. maxY |]
            series.Values <- [| minX .. stepX .. maxX |] |> Array.map (fun y -> f x y)
        chart.ChartType <- XlChartType.xlSurfaceWireframe
