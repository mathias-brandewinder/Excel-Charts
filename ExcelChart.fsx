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

// Grab Selected Range, if any
let Selection () =
    let wb = Active ()
    match wb with
    | None -> None
    | Some(wb) ->
        try
            let worksheet = wb.ActiveSheet :?> Worksheet
            let selection = worksheet.UsedRange
            selection.Value2 :?> System.Object [,] 
            |> Array2D.map (fun e -> e.ToString()) |> Some             
        with
        | _ ->
            printfn "Invalid active selection"
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
            for s in seriesCollection do s.Delete() |> ignore
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

    let redraw () =
        match chart with
        | None -> ignore ()
        | Some(chart) ->
            let xl = chart.Application
            xl.ScreenUpdating <- false
            let seriesCollection = chart.SeriesCollection() :?> SeriesCollection            
            for s in seriesCollection do s.Delete()  |> ignore
            let xs, ys = values xOver, values yOver
            for x in xs do
                let series = seriesCollection.NewSeries()
                series.Name <- (string)x
                series.XValues <- ys
                series.Values <- ys |> Array.map (f x)
            chart.ChartType <- XlChartType.xlSurfaceWireframe
            xl.ScreenUpdating <- true

    do
        match chart with
        | None -> ignore ()
        | Some(chart) -> redraw ()

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
        let xl = chart.Application
        xl.ScreenUpdating <- false
        let seriesCollection = chart.SeriesCollection() :?> SeriesCollection
        let groups = data |> Seq.map (fun (_, _, g) -> g) |> Seq.distinct
        for group in groups do
            let xs, ys, _ = data |> Seq.filter (fun (_, _, g) -> g = group) |> Seq.toArray |> Array.unzip3
            let series = seriesCollection.NewSeries()
            series.Name <- group.ToString()
            series.XValues <- xs
            series.Values <- ys
        chart.ChartType <- XlChartType.xlXYScatter
        xl.ScreenUpdating <- true

// Create XY scatterplot, colored by group, with labels
let labeledplot<'a when 'a: equality> (data: (float * float * 'a * string ) seq) =
    let chart = NewChart ()
    match chart with
    | None -> ignore ()
    | Some(chart) -> 
        let xl = chart.Application
        xl.ScreenUpdating <- false
        let seriesCollection = chart.SeriesCollection() :?> SeriesCollection
        let groups = data |> Seq.map (fun (_, _, g, _) -> g) |> Seq.distinct
        for group in groups do
            let filtered = data |> Seq.filter (fun (_, _, g, _) -> g = group) |> Seq.toArray
            let xs = filtered |> Array.map (fun (x, _, _, _) -> x)
            let ys = filtered |> Array.map (fun (_, y, _, _) -> y)
            let ls = filtered |> Array.map (fun (_, _, _, l) -> l)
            let series = seriesCollection.NewSeries()
            series.Name <- group.ToString()
            series.XValues <- xs
            series.Values <- ys
            series.HasDataLabels <- true            
            for i in 1 .. filtered.Length do 
                let point = series.Points(i) :?> Point
                point.DataLabel.Text <- ls.[i-1]
        chart.ChartType <- XlChartType.xlXYScatter
        xl.ScreenUpdating <- true

//// Usage / illustration
//// Fire up Excel, uncomment the next lines,
//// run in FSI, line by line, and go "oh! ah!" ...
//// The funky 3D functions come from Dr. Abdel-Rahman Hedar's awesome page:
//// http://www-optima.amp.i.kyoto-u.ac.jp/member/student/hedar/Hedar_files/TestGO_files/Page364.htm

//let pi = System.Math.PI

//// Simple function plot

//let f x = cos x
//let p = Plot(f, (0., 1.))
//let g x = sin x
//p.Add(g)
//let h x = f(x) * g(x*x)
//p.Add(h)
//p.Rescale(-pi, pi)
//p.Zoom(200)
//let i x = f(x) * g(x) * x
//p.Add(i)
//p.Zoom(500)

//// 3D plots on funky functions
//let booth x y = 
//    pown (x + 2. * y - 7.) 2 
//    + pown (2. * x + y - 5.) 2
//
//let s1 = Surface(booth, (-10., 10.), (-10., 10.))
//
//let branin x y = 
//    pown (y - (5. / (4. * pown pi 2)) * pown x 2 + 5. * x / pi - 6.) 2 
//    + 10.* (1. - 1./ (8. * pi)) * cos(x) + 10.
//
//let s2 = Surface(branin, (-5., 10.), (0., 15.))
//s2.Zoom(50)
//
//let schwefel x y = 
//    - sin(sqrt(abs(x))) * x 
//    - sin(sqrt(abs(y))) * y 
//let s3 = Surface(schwefel, (-1., 1.), (-1., 1.))
//s3.Zoom(50)
//s3.Rescale((-10., 10.), (-10., 10.))
//s3.Rescale((-100., 100.), (-100., 100.))
//s3.Rescale((-1000., 1000.), (-1000., 1000.))