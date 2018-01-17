# xlGraph.m
MATLAB API that plots your 2D data in an Excel table

## Motivation

For 2D charts, I always felt excel did a better job than the MATLAB "plot" function. So I wrote an API to create excel charts within a MATLAB script   


## How to use

```bash
x = [0, 1, 2]
y = [0, 1, 2]
xlGraph(x, y)
```

![](examples/xlGraph(x,y).jpg)


The API allows you to plot an unlimited number of series onto an excel chart


```bash
x = [0, 1, 2]
y = [0, 1, 2]

x2 = [9, 7, 6]
y2 = [5, 4, 3]
xlGraph(x, y, x2, y2)
```

![](examples/xlGraph(x,y,x2,y2,x3,y3).jpg)