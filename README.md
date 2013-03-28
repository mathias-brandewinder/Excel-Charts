Excel-Charts
============

Utilities to generate Excel charts from F#

To see the script in action, check out these videos:
* using FSI http://www.youtube.com/watch?v=5loQ7zb5HE8  
* using the Tsunami IDE http://www.youtube.com/watch?v=LGL6uNZOmwo

The script currently supports:
* Plotting 2D and 3D functions,
* Drawing scatterplots,
* Retrieving selected data from the active worksheet.

Example: start Excel and FSI, and go!

```
let f x = cos x    
let plot = Plot(f, (-1., 1.))   
plot.Rescale(-3., 3.)   
let g x = sin x   
plot.Add(g)
```
