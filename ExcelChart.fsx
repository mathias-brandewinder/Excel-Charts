#r "office.dll"
#r "Microsoft.Office.Interop.Excel"

open Microsoft.Office.Interop.Excel
open System.Runtime.InteropServices

// Attach to the running instance of Excel, if any
let Attach () = 
    try
        Marshal.GetActiveObject("Excel.Application") 
        :?> Microsoft.Office.Interop.Excel.Application
        |> Some
    with
    | _ -> 
        printfn "Could not find running instance of Excel"
        None

// Find the Active workbook, if any
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

// Create a new Chart in active workbook
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

// Plots single-argument function(s) over an interval
type Plot (f: float -> float, over: float * float) =
    let mutable functions = [ f ]
    let mutable over = over
    let mutable grain = 50
    let chart = NewChart ()
    let values () = 
        let min, max = over
        let step = (max - min) / (float)grain
        [| min .. step .. max |]
    let draw f =
        match chart with
        | None -> ignore ()
        | Some(chart) -> 
            let seriesCollection = chart.SeriesCollection() :?> SeriesCollection
            let series = seriesCollection.NewSeries()
            let xValues = values ()
            series.XValues <- xValues
            series.Values <- xValues |> Array.map f
    let redraw () =
        match chart with
        | None -> ignore ()
        | Some(chart) ->
            let seriesCollection = chart.SeriesCollection() :?> SeriesCollection            
            let count = seriesCollection.Count
            for s in seriesCollection do s.Delete()
            functions |> List.iter (fun f -> draw f)

    do
        match chart with
        | None -> ignore ()
        | Some(chart) -> 
            chart.ChartType <- XlChartType.xlXYScatter
            let seriesCollection = chart.SeriesCollection() :?> SeriesCollection
            draw f

    member this.Add(f: float -> float) =
        match chart with
        | None -> ignore ()
        | Some(chart) ->
            functions <- f :: functions
            draw f

    member this.Redraw () =
        match chart with
        | None -> ignore ()
        | Some(chart) ->
            let seriesCollection = chart.SeriesCollection() :?> SeriesCollection            
            let count = seriesCollection.Count
            for s in seriesCollection do s.Delete()
            functions |> List.iter (fun f -> draw f)

    member this.Rescale(min, max) =
        over <- (min, max)
        redraw()

    member this.Zoom(zoom: int) =
        grain <- zoom
        redraw()        

// Plots surface of 2-argument function
type Surface (f: float -> float -> float, xOver: (float * float), yOver: (float * float)) =
    let mutable xOver, yOver = xOver, yOver
    let mutable grain = 20
    let chart = NewChart ()
    let values over = 
        let min, max = over
        let step = (max - min) / (float)grain
        [| min .. step .. max |]
    let draw () =
        match chart with
        | None -> ignore ()
        | Some(chart) -> 
            let seriesCollection = chart.SeriesCollection() :?> SeriesCollection
            let xs, ys = values xOver, values yOver
            for x in xs do
                let series = seriesCollection.NewSeries()
                series.Name <- (string)x
                series.XValues <- ys
                series.Values <- ys |> Array.map (f x)
            chart.ChartType <- XlChartType.xlSurfaceWireframe
    let redraw () =
        match chart with
        | None -> ignore ()
        | Some(chart) ->
            let seriesCollection = chart.SeriesCollection() :?> SeriesCollection            
            for s in seriesCollection do s.Delete()
            draw ()
    do
        match chart with
        | None -> ignore ()
        | Some(chart) -> draw ()

    member this.Rescale((xmin, xmax), (ymin, ymax)) =
        xOver <- (xmin, xmax)
        yOver <- (ymin, ymax)
        redraw ()

    member this.Zoom(zoom: int) =
        grain <- zoom
        redraw ()              

// Create XY scatterplot, colored by group
let scatterplot<'a when 'a: equality> (data: (float * float * 'a ) seq) =
    let chart = NewChart ()
    match chart with
    | None -> ignore ()
    | Some(chart) -> 
        let seriesCollection = chart.SeriesCollection() :?> SeriesCollection
        let groups = data |> Seq.map (fun (_, _, g) -> g) |> Seq.distinct
        for group in groups do
            let xs, ys, gs = data |> Seq.filter (fun (_, _, g) -> g = group) |> Seq.toArray |> Array.unzip3
            let series = seriesCollection.NewSeries()
            series.Name <- group.ToString()
            series.XValues <- ys
            series.Values <- xs
        chart.ChartType <- XlChartType.xlXYScatter