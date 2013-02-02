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
    series.Type <- (int)XlChartType.xlLine

    chart

//let XY data (name: string) (chart: Chart) =
//
//    let grouped = data |> Seq.groupBy snd
//    grouped
//    |> Seq.map (fun (g, el) -> el)
//    |> Seq.m